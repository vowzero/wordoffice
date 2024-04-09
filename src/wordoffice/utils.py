from decimal import Decimal, ROUND_HALF_UP
from types import MappingProxyType
from common import Measure


def OoxmlDegree(value: int | float) -> int:
	"""
	Ooxml 角度单位转换
	Args:
		value (int|float): 角度值

	Returns:
		int: 转化Ooxml的ST_Angle值
	"""
	return int(value * 60000)


def right_round(num, keep_n=0):
	if isinstance(num, float):
		num = str(num)
	return Decimal(num).quantize((Decimal('0.' + '0' * keep_n)), rounding=ROUND_HALF_UP)


def cn_font(name, ascii='Times New Roman'):
	return {'ascii': ascii, 'eastAsia': name, 'hAnsi': ascii, 'hint': 'eastAsia'}


cn_fontSize = MappingProxyType({
	"初号": Measure.PT(42),
	"小初": Measure.PT(36),
	"一号": Measure.PT(26),
	"小一": Measure.PT(24),
	"二号": Measure.PT(22),
	"小二": Measure.PT(18),
	"三号": Measure.PT(16),
	"小三": Measure.PT(15),
	"四号": Measure.PT(14),
	"小四": Measure.PT(12),
	"五号": Measure.PT(10.5),
	"小五": Measure.PT(9),
	"六号": Measure.PT(7.5),
	"小六": Measure.PT(6.5),
	"七号": Measure.PT(5.5),
	"八号": Measure.PT(5),
})
