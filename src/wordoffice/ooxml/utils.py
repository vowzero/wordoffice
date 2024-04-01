from wordoffice.core.field import AttrField
from wordoffice.core.node import OoxmlNode, OoxmlNodeMeta

xml_header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


class SimpleValNode(OoxmlNode):
	def __init__(self, val=None, *args, **kwargs):
		"""
		唯一值(val)构造函数
		用于表示只有一个值的节点，从构造函数传入唯一值作为为val属性
		"""
		super().__init__(val=val, *args, **kwargs)


class EmptyHiddenNode(OoxmlNode):
	def to_xml(self):
		res = super().to_xml()
		if res == f'<{self.tag}/>':
			return ''
		return res


class SwitchNode(OoxmlNode):
	"""
	常开布尔节点
	val=True        不输出节点
	val=False       输出节点且不隐藏属性w:val
	"""
	val = AttrField('w:val', 'ooxml_boolean', default='on', required=True)
	_always_on = True
	
	def __init__(self, val):
		super().__init__(val=(val in ('on', '1', 1, 'true', True)))
	
	def to_xml(self):
		if (self._always_on and self.val) or (not self._always_on and not self.val):
			return ""
		
		return super().to_xml()


def create_empty_node(classname, tag_name) -> OoxmlNode:
	return OoxmlNodeMeta(classname, (OoxmlNode,), {"tag": tag_name})


def create_switch_node(class_name, tag_name, always_on=False) -> SwitchNode:
	"""
	构造适用w:val的Ooxml布尔节点
	"""
	return OoxmlNodeMeta(class_name, (SwitchNode,), {'tag': tag_name, '_always_on': always_on})





