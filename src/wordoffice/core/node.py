"""
        Node节点，用于生成xml节点
"""
import dataclasses
from typing import Any, Iterable, Iterator

from wordoffice.core.field import AttrField, SubField, SubChoiceField, NonChoiceField
from wordoffice.core.type import attr_type_transforms

_OOXML_ATTR_FIELDS = '_ooxml_attr_fields'
_OOXML_SUB_FIELDS = '_ooxml_sub_fields'

_escape_xml_chars_map = (
	('&', '&amp;'),
	('<', '&lt;'),
	('>', '&gt;'),
	('"', '&quot;'),
	("'", '&apos;'),
)


def escape_xml_chars(text: str) -> str:
	"""
	转义xml特殊字符
	"""
	if text is None:
		return None
	for src, des in _escape_xml_chars_map:
		text = text.replace(src, des)
	return text


def unescape_xml_chars(text: str) -> str:
	"""
	反转义xml特殊字符
	"""
	if text is None:
		return ''
	for src, des in _escape_xml_chars_map:
		text = text.replace(des, src)
	return text


def _build_attr_property(key):
	"""
	构建Attr的property，存储于_OOXML_ATTR_FIELDS，内部键名为@key_name
	"""
	internal_name = '@' + key
	
	def getter(self):
		return self._attrs.get(internal_name)
	
	def setter(self, value):
		self._attrs[internal_name] = value
	
	def deleter(self):
		self._attrs[internal_name] = None
	
	return property(getter, setter, deleter)


def _build_sub_property(storage_name, field_type):
	def getter(self):
		return self._subs.get(storage_name)
	
	def setter(self, value):
		storage = self._subs
		if value is None:
			storage[storage_name] = None
		elif isinstance(value, field_type):
			storage[storage_name] = value
		else:
			if isinstance(value, dict):
				node_value = field_type(**value)
			elif isinstance(value, tuple):
				node_value = field_type(*value)
			else:
				node_value = field_type(value)
			
			storage[storage_name] = node_value
	
	def deleter(self):
		getattr(self, '_subs')[storage_name] = None
	
	return property(getter, setter, deleter)


def _build_sub_choice_item_property(storage_name, field_type):
	def getter(self):
		value = self._subs.get(storage_name)
		return value if isinstance(value, field_type) else None
	
	def setter(self, value):
		storage = getattr(self, '_subs')
		if value is None:
			if isinstance(storage.get(storage_name), field_type):
				storage[storage_name] = None
			return
		if isinstance(value, field_type):
			storage[storage_name] = value
		else:
			if isinstance(value, dict):
				node_value = field_type(**value)
			elif isinstance(value, tuple):
				node_value = field_type(*value)
			else:
				node_value = field_type(value)
			
			storage[storage_name] = node_value
	
	return property(getter, setter)


def _build_sub_choice_property(key_name, choices_type: tuple):
	def getter(self):
		return getattr(self, '_subs').get(key_name)
	
	def setter(self, value):
		storage = getattr(self, '_subs')
		if value is None:
			storage[key_name] = None
		elif isinstance(value, choices_type):
			storage[key_name] = value
		else:
			raise ValueError(f"{self}的sub属性{key_name}类型错误，不应是{type(value)},而是{choices_type}")
	
	def deleter(self):
		getattr(self, '_subs')[key_name] = None
	
	return property(getter, setter, deleter)


class OoxmlNodeBase:
	def to_xml(self) -> str:
		"""生成xml字符串"""
		raise NotImplementedError


class OoxmlNodeMeta(type):
	@staticmethod
	def merge_parent_node(bases):
		attr_fields = {}
		sub_fields = {}
		base_tag_name = None
		for base in bases:
			if _OOXML_ATTR_FIELDS in base.__dict__:
				attr_fields.update(base.__dict__[_OOXML_ATTR_FIELDS])
			if _OOXML_SUB_FIELDS in base.__dict__:
				sub_fields.update(base.__dict__[_OOXML_SUB_FIELDS])
			# 如果父类中有tag属性，则使用父类的tag，如果父类是None则使用类name
			if 'tag' in base.__dict__:
				base_tag_name = base.__dict__['tag'] or base_tag_name
		
		return base_tag_name, attr_fields, sub_fields
	
	def __new__(cls, name, bases, attrs: dict):
		valid_fields_set = []  # 处理过的合法Field
		
		base_tag_name, cls_attr_fields, cls_sub_fields_unordered = cls.merge_parent_node(bases)
		cls_fields_property = {}
		
		for key, value in attrs.items():
			if isinstance(value, NonChoiceField):
				property_obj = None
				
				match value:
					case AttrField():
						if isinstance(value.type, str):
							value.type = (value.type, None)
						cls_attr_fields[key] = value
						property_obj = _build_attr_property(key)
					case SubField():
						cls_sub_fields_unordered[key] = value
						property_obj = _build_sub_property(key, value.type)
				
				cls_fields_property[key] = property_obj
				
				valid_fields_set.append(key)
			
			elif isinstance(value, SubChoiceField):
				
				choice_types = []
				choice_index = value.index
				for choice_key, choice_value in value.choices.items():
					if choice_key in cls_sub_fields_unordered:
						raise ValueError(f"subField不能有同名key")
					
					cur_field: SubField = dataclasses.replace(choice_value)
					cur_field.index = choice_index
					cur_field.conflict = key
					cur_field.required = value.required
					cls_sub_fields_unordered[choice_key] = cur_field
					cls_fields_property[choice_key] = _build_sub_choice_item_property(cur_field.conflict, cur_field.type)
					choice_types.append(cur_field.type)
				
				choice_types = tuple(choice_types)
				cls_sub_fields_unordered[key] = SubField(choice_types, index=choice_index, required=value.required)
				cls_fields_property[key] = _build_sub_choice_property(key, choice_types)
				valid_fields_set.append(key)
		
		# 清除 类变量 中 Fields
		for key in valid_fields_set:
			attrs.pop(key)
		
		attrs[_OOXML_ATTR_FIELDS] = cls_attr_fields
		
		cls_sub_fields_ordered = dict(sorted(cls_sub_fields_unordered.items(), key=lambda x: x[1].index))
		
		attrs[_OOXML_SUB_FIELDS] = cls_sub_fields_ordered
		attrs.update(cls_fields_property)
		
		# 如果没有tag属性，则继承父类的tag
		attrs['tag'] = attrs.get('tag', base_tag_name or name)
		
		children_type = attrs.get('_children_type')
		if children_type is None:
			attrs['_children_type'] = (OoxmlNodeBase,)
		
		return super().__new__(cls, name, bases, attrs)


def children_node_to_xml(children):
	result = []
	for child in children:
		if isinstance(child, OoxmlNodeBase):
			child = child.to_xml()
		if child is None or child == '':
			continue
		result.append(str(child))
	return result


class PureTextNode(OoxmlNodeBase):
	def __init__(self, text):
		self.text = text
	
	def to_xml(self):
		return self.text


class OoxmlNode(OoxmlNodeBase, metaclass=OoxmlNodeMeta):
	"""OOXML节点基类"""
	tag: str = None
	_children_type: tuple = None
	
	def __init__(self, *args, **kwargs):
		attr_fields = getattr(self, _OOXML_ATTR_FIELDS)
		sub_fields = getattr(self, _OOXML_SUB_FIELDS)
		
		self._attrs = {}
		self._subs = {}
		self._children = []
		self._text = None
		
		# 映射所有关键字属性到attr_storage和sub_storage
		for key_name, value in kwargs.copy().items():
			if key_name in attr_fields:
				self.set_attr(key_name, value)
				kwargs.pop(key_name)
			elif key_name in sub_fields:
				self.set_sub(key_name, value)
				kwargs.pop(key_name)
	
	@property
	def tag_name(self):
		"""标签名"""
		return self.tag
	
	@property
	def innerText(self):
		"""设置/获取并自动转义innerText"""
		return unescape_xml_chars(self._text)
	
	@innerText.setter
	def innerText(self, value):
		self._text = escape_xml_chars(value)
	
	def set_attr(self, name: str, value: Any, ooxml_force: bool = True) -> None:
		"""
		设置attr属性
		
		:param name: 属性名
		:param value: 属性值
		:param ooxml_force: 强制优先作为ooxml属性
		"""
		field = super().__getattribute__(_OOXML_ATTR_FIELDS).get(name)
		if field is not None and ooxml_force:
			self._attrs['@' + name] = value
		else:
			self._attrs[name] = value
	
	def get_attr(self, name, ooxml_force=True) -> Any:
		"""
		获取attr属性
		
		:param name: 属性名
		:param ooxml_force: 强制优先作为ooxml属性
		"""
		if ooxml_force:
			storge_key = '@' + name
			return self._attrs.get(storge_key)
		else:
			return self._attrs.get(name)
	
	def set_sub(self, name: str, value: Any) -> None:
		"""
		设置sub属性
		
		:param name: 属性名
		:param value: 属性值
		"""
		field = super().__getattribute__(_OOXML_SUB_FIELDS).get(name)
		if field is not None:
			storage = getattr(self, '_subs')
			storage_key = field.conflict or name
			
			if value is None:
				storage[storage_key] = None
			elif isinstance(value, field.type):
				storage[storage_key] = value
			else:
				if isinstance(value, dict):
					node_value = field.type(**value)
				elif isinstance(value, tuple):
					node_value = field.type(*value)
				else:
					node_value = field.type(value)
				storage[storage_key] = node_value
	
	def get_sub(self, name) -> Any:
		"""
		获取sub属性

		:param name: 属性名
		"""
		return super().__getattribute__('_subs').get(name)
	
	def append(self, node: OoxmlNodeBase) -> None:
		"""追加节点为末尾子节点"""
		self._assert_valid_node(node)
		self._children.append(node)
	
	def extend(self, nodes: Iterable[OoxmlNodeBase]) -> None:
		"""扩展节点集为子节点"""
		for node in nodes:
			self._assert_valid_node(node)
		self._children.extend(nodes)
	
	def insert(self, index: int, node: OoxmlNodeBase) -> None:
		"""指定位置插入节点"""
		self._assert_valid_node(node)
		self._children.insert(index, node)
	
	def insert_before(self, node: int, ref_node: OoxmlNodeBase, tail: bool = True) -> None:
		"""
		插入节点到参考节点的位置
		
		:param node: 要插入的节点
		:param ref_node: 参考节点
		:param tail: 参考节点不存在时是否插入到最后，True为最后，False为最前
		"""
		self._assert_valid_node(node)
		self._assert_valid_node(ref_node)
		try:
			index = self._children.index(ref_node)
			self.insert(index, node)
		except ValueError:
			if tail:
				self.append(node)
			else:
				self.insert(0, node)
	
	def remove(self, node: OoxmlNodeBase) -> None:
		"""移除节点"""
		self._children.remove(node)
	
	def remove_node_type(self, node_type: type[OoxmlNodeBase]) -> None:
		"""移除指定类型的节点"""
		self._children = [child for child in self._children if not isinstance(child, node_type)]
	
	def clear(self) -> None:
		"""清空所有子节点"""
		self._children.clear()
	
	def iter_attrs_original(self) -> Iterator[tuple[str, Any]]:
		"""迭代所有attrs"""
		yield from self._attrs.items()
	
	def iter_attrs_xml(self) -> Iterator[tuple[str, str]]:
		"""迭代所有attrs，并转换为xml格式"""
		fields = getattr(self, _OOXML_ATTR_FIELDS)
		for key, value in self._attrs.items():
			if value is None:
				continue
			if key[0] == '@':
				field: AttrField = fields[key[1:]]
				field_type, field_type_info = field.type
				attr_checker = attr_type_transforms.get(field_type, None)
				if attr_checker is None:
					xml_value = value
				else:
					if field_type_info is None:
						xml_value = attr_checker(value)
					else:
						xml_value = attr_checker(value, **field_type_info)
				
				# 默认值隐藏
				if field.default and xml_value == field.default:
					continue
				
				yield field.xml_name, xml_value
			else:
				yield key, value
	
	def iter_subs(self) -> Iterator[tuple[str, Any]]:
		"""迭代所有sub属性"""
		fields = getattr(self, _OOXML_SUB_FIELDS)
		storage = self._subs
		for key, field in fields.items():
			if field.conflict:
				continue
			value = storage.get(key, None)
			if value is not None:
				yield key, value
	
	def iter_children(self) -> Iterator[OoxmlNodeBase]:
		"""迭代所有子节点"""
		yield from self._children
	
	def _check_required(self, attrs, subs):
		for key, field in getattr(self, _OOXML_ATTR_FIELDS).items():
			if field.required and not field.default and field.xml_name not in attrs:
				assert False, f"{self.__class__}的属性{key}(attr:\"{field.xml_name}\")是必须的"
		for key, field in getattr(self, _OOXML_SUB_FIELDS).items():
			if field.required and key not in subs and not field.conflict:
				assert False, f"{self.__class__}的属性{key}是必须的"
	
	def before_to_xml(self) -> list:
		"""
		:return: [tag, attrs, subs, children, text]列表
		"""
		return [self.tag_name, dict(self.iter_attrs_xml()), dict(self.iter_subs()), list(self.iter_children()), self.innerText]
	
	def to_xml(self) -> str:
		"""生成xml"""
		tag, attrs, subs, children, text = self.before_to_xml()
		self._check_required(attrs, subs)
		
		tag_begin = f'<{tag}'
		tag_end = f'</{tag}>'
		tag_end_open = '>'
		tag_end_closed = '/>'
		
		result_list = [tag_begin]
		attr_list = [f' {key}="{value}"' for key, value in attrs.items()]
		children_list = children_node_to_xml(subs.values()) + children_node_to_xml(children)
		
		if text:
			children_list.insert(0, text)
		
		attr_len = len(attr_list)
		children_len = len(children_list)
		
		if attr_len == 0 and children_len == 0:
			result_list.append(tag_end_closed)
		else:
			result_list.extend(attr_list)
			if children_len > 0:
				result_list.append(tag_end_open)
				result_list.extend(children_list)
				result_list.append(tag_end)
			else:
				result_list.append(tag_end_closed)
		
		return ''.join(result_list)
	
	def _assert_valid_node(self, node):
		if not isinstance(node, self._children_type):
			raise TypeError(f"{node}不是OoxmlNode类型")
	
	def __getitem__(self, indices):
		if isinstance(indices, int):
			return self._children[indices]
		elif isinstance(indices, slice):
			return self._children[indices]
		else:
			raise IndexError("indices必须是切片或row索引")
	
	def __str__(self):
		return self.to_xml()
	
	def __repr__(self):
		return f'<{self.tag_name}/>({self.__class__}, id={id(self)})'
	
	def __len__(self):
		return len(self._children)
