import re


class Measure:
	__EMUS_TO_IN = 914400
	__EMUS_TO_CM = 360000
	__EMUS_TO_MM = 36000
	__EMUS_TO_PT = 12700
	__EMUS_TO_PX = 12700 * 72 / 96
	__EMUS_TO_PC = 152400
	__EMUS_TO_PI = 152400
	__EMUS_TO_HPS = 6350
	__EMUS_TO_TWIPS = 635
	__EMUS_TO_PT_8 = 1587.5  # EighthPoint
	
	# mm | cm | in | pt | pc | pi
	def __init__(self, value):
		self._emus = value
	
	@classmethod
	def PX(cls, value):
		return cls(value * cls.__EMUS_TO_PX)
	
	@classmethod
	def EMUS(cls, value):
		return cls(value)
	
	@classmethod
	def INCH(cls, value):
		return cls(value * cls.__EMUS_TO_IN)
	
	@classmethod
	def CM(cls, value):
		return cls(value * cls.__EMUS_TO_CM)
	
	@classmethod
	def MM(cls, value):
		return cls(value * cls.__EMUS_TO_MM)
	
	@classmethod
	def PT(cls, value):
		return cls(value * cls.__EMUS_TO_PT)
	
	@classmethod
	def PT_8(cls, value):
		return cls(value * cls.__EMUS_TO_PT_8)
	
	@classmethod
	def PC(cls, value):
		return cls(value * cls.__EMUS_TO_PC)
	
	@classmethod
	def PI(cls, value):
		return cls(value * cls.__EMUS_TO_PI)
	
	@classmethod
	def HPS(cls, value):
		return cls(value * cls.__EMUS_TO_HPS)
	
	@classmethod
	def TWIPS(cls, value):
		return cls(value * cls.__EMUS_TO_TWIPS)
	
	@property
	def px(self):
		return self.emus / self.__EMUS_TO_PX
	
	@px.setter
	def px(self, value):
		self.emus = value * self.__EMUS_TO_PX
	
	@property
	def pt(self):
		return self.emus / self.__EMUS_TO_PT
	
	@pt.setter
	def pt(self, value):
		self.emus = value * self.__EMUS_TO_PT
	
	@property
	def inch(self):
		return self.emus / self.__EMUS_TO_IN
	
	@inch.setter
	def inch(self, value):
		self.emus = value * self.__EMUS_TO_IN
	
	@property
	def cm(self):
		return self.emus / self.__EMUS_TO_CM
	
	@cm.setter
	def cm(self, value):
		self.emus = value * self.__EMUS_TO_CM
	
	@property
	def mm(self):
		return self.emus / self.__EMUS_TO_MM
	
	@mm.setter
	def mm(self, value):
		self.emus = value * self.__EMUS_TO_MM
	
	@property
	def pc(self):
		return self.emus / self.__EMUS_TO_PC
	
	@pc.setter
	def pc(self, value):
		self.emus = value * self.__EMUS_TO_PC
	
	@property
	def pi(self):
		return self.emus / self.__EMUS_TO_PI
	
	@pi.setter
	def pi(self, value):
		self.emus = value * self.__EMUS_TO_PI
	
	@property
	def hps(self):
		return int(self.emus / self.__EMUS_TO_HPS)
	
	@hps.setter
	def hps(self, value):
		self.emus = value * self.__EMUS_TO_HPS
	
	@property
	def twips(self):
		return self.emus / self.__EMUS_TO_TWIPS
	
	@twips.setter
	def twips(self, value):
		self.emus = value * self.__EMUS_TO_TWIPS
	
	@property
	def pt_8(self):
		return self.emus / self.__EMUS_TO_PT_8
	
	@pt_8.setter
	def pt_8(self, value):
		self.emus = value * self.__EMUS_TO_PT_8
	
	@property
	def unsigned(self):
		return self.emus >= 0
	
	@property
	def emus(self):
		return self._emus
	
	@emus.setter
	def emus(self, value):
		self._emus = value
	
	def __str__(self):
		return f"{self.pt}pt"
	
	def __eq__(self, other):
		if isinstance(other, Measure):
			return self.emus == other.emus
		else:
			return False


MeasureZero = Measure.EMUS(0)

_regx_color_hex = re.compile(r'^#?([0-9a-fA-F]{6})$')


def Color(*args):
	if len(args) == 1:
		if isinstance(args[0], tuple):
			r, g, b = args[0]
		elif isinstance(args[0], str):
			if args[0] == 'auto':
				return args[0]
			value = args[0]
			res = _regx_color_hex.match(value)
			xml_value = res.group(1).upper()
			return xml_value
		else:
			assert False, f"{type(args[0])}<{args[0]}> is not a valid Color value"
	else:
		r, g, b = args
	return f'{r:02x}{g:02x}{b:02x}'


ColorAuto = 'auto'
