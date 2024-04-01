from typing import Callable, cast

from wordoffice.utils import right_round
from wordoffice.common import Measure, Color

attr_type_transforms: dict[str, Callable] = {
	'string': cast(Callable, None)
}


def register_attr_type_transform(func):
	attr_type_transforms[func.__name__] = func


@register_attr_type_transform
def ooxml_boolean(value):
	return 'on' if value in (1, '1', True, 'true', 'on') else 'off'


@register_attr_type_transform
def xsd_boolean(value):
	return 'true' if value in (1, '1', True, 'true', 'on') else 'false'


@register_attr_type_transform
def integer(value):
	return str(int(value))


@register_attr_type_transform
def integer_unsigned(value):
	if isinstance(value, (str | float | int)):
		if int(value) < 0:
			assert False, f"Integer<{value}> must be unsigned"
	return str(int(value))


@register_attr_type_transform
def integer_range(value, range: tuple):
	_value = int(value)
	if range[0] <= _value <= range[1]:
		return str(_value)
	else:
		assert False, f"Integer<{value}> must be in range {range}"


@register_attr_type_transform
def ooxml_color(value):
	return Color(value)


@register_attr_type_transform
def ooxml_long_hex_number(value):
	if isinstance(value, str):
		if value.startswith('#'):
			value = int(value[1:], 16)
		else:
			value = int(value)
	return f'{value:08X}'


@register_attr_type_transform
def enum(value, enums):
	if isinstance(value, enums):
		return str(value)
	if isinstance(value, str):
		return str(enums(value))
	assert False, f"{type(value)}<{value}> is not a valid Enum type{enums}"


@register_attr_type_transform
def ooxml_attr_scale(value):
	if isinstance(value, str) and value[-1] == '%':
		value = int(value[:-1])
	if isinstance(value, int):
		if value < 0 or value > 600:
			assert False, (f"scale must between 0 and 600")
		
		return str(value) + '%'


@register_attr_type_transform
def ooxml_measure(value, type: str, unsigned=False, precision=None):
	match value:
		case int() | float():
			value = getattr(Measure, type.upper())(value)
		case str():
			from_type = type
			if len(value) > 2:
				unit = value[-2:].upper()
				if unit in ('PT', 'CM', 'IN'):
					from_type = unit
					value = value[:-2]
					if from_type == 'IN':
						from_type = 'INCH'
			value = getattr(Measure, from_type.upper())(float(value))
	decimal_value = getattr(value, type)
	if unsigned and decimal_value < 0:
		assert False, f"Measure<{value}> must be unsigned"
	return right_round(decimal_value, precision)


OoxmlTwipsMeasure = ('ooxml_measure', {'type': 'twips', 'unsigned': True, 'precision': 0})
OoxmlSignedTwipsMeasure = ('ooxml_measure', {'type': 'twips', 'unsigned': False, 'precision': 0})

OoxmlHpsMeasure = ('ooxml_measure', {'type': 'hps', 'unsigned': True, 'precision': 0})
OoxmlSignedHpsMeasure = ('ooxml_measure', {'type': 'hps', 'unsigned': False, 'precision': 0})

OoxmlEightPointMeasure = ('ooxml_measure', {'type': 'pt_8', 'unsigned': True, 'precision': 0})

OoxmlPointMeasure = ('ooxml_measure', {'type': 'pt', 'unsigned': True, 'precision': 0})
OoxmlSignedPointMeasure = ('ooxml_measure', {'type': 'pt', 'unsigned': False, 'precision': 0})

OoxmlEmusMeasure = ('ooxml_measure', {'type': 'emus', 'unsigned': False, 'precision': 0})

OoxmlMeasurementOrPercent = 'string'


@register_attr_type_transform
def ooxml_short_hex_number(value):
	if isinstance(value, str):
		if value.startswith('#'):
			value = int(value[1:], 16)
		else:
			value = int(value)
	return f'{value:04X}'
