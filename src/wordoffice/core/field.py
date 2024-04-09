"""
    用于描述Node的属性和子节点的相关类
"""
from dataclasses import dataclass
from typing import Optional, Any


@dataclass(eq=False)
class AttrField:
	xml_name: str  # 生成xml对应的attr name
	type: str | tuple[str, dict | None]  # 类型检查类
	required: Optional[bool] = False  # 是否必须，必须项应在构造时初始化
	default: Optional[Any] = None  # 默认值隐藏


@dataclass(eq=False)
class SubField:
	type: object  # 子节点类型
	required: Optional[bool] = False  # 是否必须，必须项应在构造时初始化
	index: Optional[int] = 1  # 用于指定子节点的排序索引
	conflict: Optional[str] = False  # 冲突保留存储key_name


NonChoiceField = AttrField | SubField


@dataclass(eq=False)
class SubChoiceField:
	choices: dict[str, Any]  # 子节点类型
	type: Optional[Any] = Any
	index: Optional[int] = 1  # 用于指定子节点的排序索引
	# key_name: Optional[str] = None
	required: Optional[bool] = False  # 是否必须，必须项应在构造时初始化
