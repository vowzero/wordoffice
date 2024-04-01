"""
        来自ooxml标准中的枚举类型
"""
from enum import Enum


class EnumStrBase(Enum):
	def __str__(self):
		return str(self.value)
	
	def __repr__(self):
		return str(self.value)


class ST_Border(EnumStrBase):
	NIL = "nil"
	NONE = "none"
	SINGLE = "single"
	THICK = "thick"
	DOUBLE = "double"
	DOTTED = "dotted"
	DASHED = "dashed"
	DOT_DASH = "dotDash"
	DOT_DOT_DASH = "dotDotDash"
	TRIPLE = "triple"
	THIN_THICK_SMALL_GAP = "thinThickSmallGap"
	THICK_THIN_SMALL_GAP = "thickThinSmallGap"
	THIN_THICK_THIN_SMALL_GAP = "thinThickThinSmallGap"
	THIN_THICK_MEDIUM_GAP = "thinThickMediumGap"
	THICK_THIN_MEDIUM_GAP = "thickThinMediumGap"
	THIN_THICK_THIN_MEDIUM_GAP = "thinThickThinMediumGap"
	THIN_THICK_LARGE_GAP = "thinThickLargeGap"
	THICK_THIN_LARGE_GAP = "thickThinLargeGap"
	THIN_THICK_THIN_LARGE_GAP = "thinThickThinLargeGap"
	WAVE = "wave"
	DOUBLE_WAVE = "doubleWave"
	DASH_SMALL_GAP = "dashSmallGap"
	DASH_DOT_STROKED = "dashDotStroked"
	THREE_D_EMBOSS = "threeDEmboss"
	THREE_D_ENGRAVE = "threeDEngrave"
	OUTSET = "outset"
	INSET = "inset"
	APPLES = "apples"
	ARCHED_SCALLOPS = "archedScallops"
	BABY_PACIFIER = "babyPacifier"
	BABY_RATTLE = "babyRattle"
	BALLOONS_3_COLORS = "balloons3Colors"
	BALLOONS_HOT_AIR = "balloonsHotAir"
	BASIC_BLACK_DASHES = "basicBlackDashes"
	BASIC_BLACK_DOTS = "basicBlackDots"
	BASIC_BLACK_SQUARES = "basicBlackSquares"
	BASIC_THIN_LINES = "basicThinLines"
	BASIC_WHITE_DASHES = "basicWhiteDashes"
	BASIC_WHITE_DOTS = "basicWhiteDots"
	BASIC_WHITE_SQUARES = "basicWhiteSquares"
	BASIC_WIDE_INLINE = "basicWideInline"
	BASIC_WIDE_MIDLINE = "basicWideMidline"
	BASIC_WIDE_OUTLINE = "basicWideOutline"
	BATS = "bats"
	BIRDS = "birds"
	BIRDS_FLIGHT = "birdsFlight"
	CABINS = "cabins"
	CAKE_SLICE = "cakeSlice"
	CANDY_CORN = "candyCorn"
	CELTIC_KNOTWORK = "celticKnotwork"
	CERTIFICATE_BANNER = "certificateBanner"
	CHAIN_LINK = "chainLink"
	CHAMPAGNE_BOTTLE = "champagneBottle"
	CHECKED_BAR_BLACK = "checkedBarBlack"
	CHECKED_BAR_COLOR = "checkedBarColor"
	CHECKERED = "checkered"
	CHRISTMAS_TREE = "christmasTree"
	CIRCLES_LINES = "circlesLines"
	CIRCLES_RECTANGLES = "circlesRectangles"
	CLASSICAL_WAVE = "classicalWave"
	CLOCKS = "clocks"
	COMPASS = "compass"
	CONFETTI = "confetti"
	CONFETTI_GRAYS = "confettiGrays"
	CONFETTI_OUTLINE = "confettiOutline"
	CONFETTI_STREAMERS = "confettiStreamers"
	CONFETTI_WHITE = "confettiWhite"
	CORNER_TRIANGLES = "cornerTriangles"
	COUPON_CUTOUT_DASHES = "couponCutoutDashes"
	COUPON_CUTOUT_DOTS = "couponCutoutDots"
	CRAZY_MAZE = "crazyMaze"
	CREATURES_BUTTERFLY = "creaturesButterfly"
	CREATURES_FISH = "creaturesFish"
	CREATURES_INSECTS = "creaturesInsects"
	CREATURES_LADY_BUG = "creaturesLadyBug"
	CROSS_STITCH = "crossStitch"
	CUP = "cup"
	DECO_ARCH = "decoArch"
	DECO_ARCH_COLOR = "decoArchColor"
	DECO_BLOCKS = "decoBlocks"
	DIAMONDS_GRAY = "diamondsGray"
	DOUBLE_D = "doubleD"
	DOUBLE_DIAMONDS = "doubleDiamonds"
	EARTH_1 = "earth1"
	EARTH_2 = "earth2"
	EARTH_3 = "earth3"
	ECLIPSING_SQUARES_1 = "eclipsingSquares1"
	ECLIPSING_SQUARES_2 = "eclipsingSquares2"
	EGGS_BLACK = "eggsBlack"
	FANS = "fans"
	FILM = "film"
	FIRECRACKERS = "firecrackers"
	FLOWERS_BLOCK_PRINT = "flowersBlockPrint"
	FLOWERS_DAISIES = "flowersDaisies"
	FLOWERS_MODERN_1 = "flowersModern1"
	FLOWERS_MODERN_2 = "flowersModern2"
	FLOWERS_PANSY = "flowersPansy"
	FLOWERS_RED_ROSE = "flowersRedRose"
	FLOWERS_ROSES = "flowersRoses"
	FLOWERS_TEACUP = "flowersTeacup"
	FLOWERS_TINY = "flowersTiny"
	GEMS = "gems"
	GINGERBREAD_MAN = "gingerbreadMan"
	GRADIENT = "gradient"
	HANDMADE_1 = "handmade1"
	HANDMADE_2 = "handmade2"
	HEART_BALLOON = "heartBalloon"
	HEART_GRAY = "heartGray"
	HEARTS = "hearts"
	HEEBIE_JEEBIES = "heebieJeebies"
	HOLLY = "holly"
	HOUSE_FUNKY = "houseFunky"
	HYPNOTIC = "hypnotic"
	ICE_CREAM_CONES = "iceCreamCones"
	LIGHT_BULB = "lightBulb"
	LIGHTNING_1 = "lightning1"
	LIGHTNING_2 = "lightning2"
	MAP_PINS = "mapPins"
	MAPLE_LEAF = "mapleLeaf"
	MAPLE_MUFFINS = "mapleMuffins"
	MARQUEE = "marquee"
	MARQUEE_TOOTHED = "marqueeToothed"
	MOONS = "moons"
	MOSAIC = "mosaic"
	MUSIC_NOTES = "musicNotes"
	NORTHWEST = "northwest"
	OVALS = "ovals"
	PACKAGES = "packages"
	PALMS_BLACK = "palmsBlack"
	PALMS_COLOR = "palmsColor"
	PAPER_CLIPS = "paperClips"
	PAPYRUS = "papyrus"
	PARTY_FAVOR = "partyFavor"
	PARTY_GLASS = "partyGlass"
	PENCILS = "pencils"
	PEOPLE = "people"
	PEOPLE_WAVING = "peopleWaving"
	PEOPLE_HATS = "peopleHats"
	POINSETTIAS = "poinsettias"
	POSTAGE_STAMP = "postageStamp"
	PUMPKIN_1 = "pumpkin1"
	PUSH_PIN_NOTE_2 = "pushPinNote2"
	PUSH_PIN_NOTE_1 = "pushPinNote1"
	PYRAMIDS = "pyramids"
	PYRAMIDS_ABOVE = "pyramidsAbove"
	QUADRANTS = "quadrants"
	RINGS = "rings"
	SAFARI = "safari"
	SAWTOOTH = "sawtooth"
	SAWTOOTH_GRAY = "sawtoothGray"
	SCARED_CAT = "scaredCat"
	SEATTLE = "seattle"
	SHADOWED_SQUARES = "shadowedSquares"
	SHARKS_TEETH = "sharksTeeth"
	SHOREBIRD_TRACKS = "shorebirdTracks"
	SKYROCKET = "skyrocket"
	SNOWFLAKE_FANCY = "snowflakeFancy"
	SNOWFLAKES = "snowflakes"
	SOMBRERO = "sombrero"
	SOUTHWEST = "southwest"
	STARS = "stars"
	STARS_TOP = "starsTop"
	STARS_3_D = "stars3d"
	STARS_BLACK = "starsBlack"
	STARS_SHADOWED = "starsShadowed"
	SUN = "sun"
	SWIRLIGIG = "swirligig"
	TORN_PAPER = "tornPaper"
	TORN_PAPER_BLACK = "tornPaperBlack"
	TREES = "trees"
	TRIANGLE_PARTY = "triangleParty"
	TRIANGLES = "triangles"
	TRIANGLE_1 = "triangle1"
	TRIANGLE_2 = "triangle2"
	TRIANGLE_CIRCLE_1 = "triangleCircle1"
	TRIANGLE_CIRCLE_2 = "triangleCircle2"
	SHAPES_1 = "shapes1"
	SHAPES_2 = "shapes2"
	TWISTED_LINES_1 = "twistedLines1"
	TWISTED_LINES_2 = "twistedLines2"
	VINE = "vine"
	WAVELINE = "waveline"
	WEAVING_ANGLES = "weavingAngles"
	WEAVING_BRAID = "weavingBraid"
	WEAVING_RIBBON = "weavingRibbon"
	WEAVING_STRIPS = "weavingStrips"
	WHITE_FLOWERS = "whiteFlowers"
	WOODWORK = "woodwork"
	X_ILLUSIONS = "xIllusions"
	ZANY_TRIANGLES = "zanyTriangles"
	ZIG_ZAG = "zigZag"
	ZIG_ZAG_STITCH = "zigZagStitch"
	CUSTOM = "custom"


class ST_Em(EnumStrBase):
	NONE = 'none'
	DOT = 'dot'
	COMMA = 'comma'
	CIRCLE = 'circle'
	UNDERLINE = 'underline'


class ST_Underline(EnumStrBase):
	SINGLE = 'single'
	WORDS = 'words'
	DOUBLE = 'double'
	THICK = 'thick'
	DOTTED = 'dotted'
	DOTTED_HEAVY = 'dottedHeavy'
	DASH = 'dash'
	DASHED_HEAVY = 'dashedHeavy'
	DASH_LONG = 'dashLong'
	DASH_LONG_HEAVY = 'dashLongHeavy'
	DOT_DASH = 'dotDash'
	DASHDOT_HEAVY = 'dashDotHeavy'
	DOTDOT_DASH = 'dotDotDash'
	DASH_DOTDOT_HEAVY = 'dashDotDotHeavy'
	WAVE = 'wave'
	WAVY_HEAVY = 'wavyHeavy'
	WAVY_DOUBLE = 'wavyDouble'
	NONE = 'none'


class ST_Hint(EnumStrBase):
	EAST_ASIA = 'eastAsia'
	DEFAULT = 'default'


class ST_HighlightColor(EnumStrBase):
	BLACK = 'black'
	BLUE = 'blue'
	CYAN = 'cyan'
	GREEN = 'green'
	MAGENTA = 'magenta'
	RED = 'red'
	YELLOW = 'yellow'
	WHITE = 'white'
	DARK_BLUE = 'darkBlue'
	DARK_CYAN = 'darkCyan'
	DARK_GREEN = 'darkGreen'
	DARK_MAGENTA = 'darkMagenta'
	DARK_RED = 'darkRed'
	DARK_YELLOW = 'darkYellow'
	DARK_GRAY = 'darkGray'
	LIGHT_GRAY = 'lightGray'
	NONE = 'none'


class ST_Shd(EnumStrBase):
	NIL = 'nil'
	CLEAR = 'clear'
	SOLID = 'solid'
	HORZ_STRIPE = 'horzStripe'
	VERT_STRIPE = 'vertStripe'
	REVERSE_DIAGONAL_STRIPE = 'reverseDiagStripe'
	DIAGONAL_STRIPE = 'diagStripe'
	HORZ_CROSS = 'horzCross'
	DIAG_CROSS = 'diagCross'
	THICK_HORZ_STRIPE = 'thinHorzStripe'
	THICK_VERT_STRIPE = 'thinVertStripe'
	THICK_REVERSE_DIAGONAL_STRIPE = 'thinReverseDiagStripe'
	THICK_DIAGONAL_STRIPE = 'thinDiagStripe'
	THICK_HORZ_CROSS = 'thinHorzCross'
	THICK_DIAG_CROSS = 'thinDiagCross'
	PERCENT_5 = 'pct5'
	PERCENT_10 = 'pct10'
	PERCENT_12 = 'pct12'
	PERCENT_15 = 'pct15'
	PERCENT_20 = 'pct20'
	PERCENT_25 = 'pct25'
	PERCENT_30 = 'pct30'
	PERCENT_35 = 'pct35'
	PERCENT_37 = 'pct37'
	PERCENT_40 = 'pct40'
	PERCENT_45 = 'pct45'
	PERCENT_50 = 'pct50'
	PERCENT_55 = 'pct55'
	PERCENT_60 = 'pct60'
	PERCENT_62 = 'pct62'
	PERCENT_65 = 'pct65'
	PERCENT_70 = 'pct70'
	PERCENT_75 = 'pct75'
	PERCENT_80 = 'pct80'
	PERCENT_85 = 'pct85'
	PERCENT_87 = 'pct87'
	PERCENT_90 = 'pct90'
	PERCENT_95 = 'pct95'


class ST_VerticalAlignRun(EnumStrBase):
	BASELINE = 'baseline'
	SUPERSCRIPT = 'superscript'
	SUBSCRIPT = 'subscript'


class ST_Jc(EnumStrBase):
	START = 'start'
	CENTER = 'center'
	END = 'end'
	BOTH = 'both'
	HIGH_KASHIDA = 'highKashida'
	MEDIUM_KASHIDA = 'mediumKashida'
	LOW_KASHIDA = 'lowKashida'
	DISTRIBUTE = 'distribute'
	THAI_DISTRIBUTE = 'thaiDistribute'
	NUM_TAB = 'numTab'


class ST_LineSpacingRule(EnumStrBase):
	AUTO = 'auto'
	EXACT = 'exact'
	AT_LEAST = 'atLeast'


class ST_HdrFtr(EnumStrBase):
	DEFAULT = 'default'
	EVEN = 'even'
	FIRST = 'first'


class ST_DocGrid(EnumStrBase):
	DEFAULT = 'default'
	LINES = 'lines'
	LINES_AND_CHARS = 'linesAndChars'
	SNAP_TO_CHARS = 'snapToChars'


class ST_TextAlignment(EnumStrBase):
	AUTO = 'auto'
	TOP = 'top'
	CENTER = 'center'
	BASELINE = 'baseline'
	BOTTOM = 'bottom'


class ST_LineNumberRestart(EnumStrBase):
	CONTINUOUS = 'continuous'
	NEW_PAGE = 'newPage'
	NEW_SECTION = 'newSection'


class ST_NumberFormat(EnumStrBase):
	DECIMAL = "decimal"
	UPPER_ROMAN = "upperRoman"
	LOWER_ROMAN = "lowerRoman"
	UPPER_LETTER = "upperLetter"
	LOWER_LETTER = "lowerLetter"
	ORDINAL = "ordinal"
	CARDINAL_TEXT = "cardinalText"
	ORDINAL_TEXT = "ordinalText"
	HEX = "hex"
	CHICAGO = "chicago"
	IDEOGRAPH_DIGITAL = "ideographDigital"
	JAPANESE_COUNTING = "japaneseCounting"
	AIUEO = "aiueo"
	IROHA = "iroha"
	DECIMAL_FULL_WIDTH = "decimalFullWidth"
	DECIMAL_HALF_WIDTH = "decimalHalfWidth"
	JAPANESE_LEGAL = "japaneseLegal"
	JAPANESE_DIGITAL_TEN_THOUSAND = "japaneseDigitalTenThousand"
	DECIMAL_ENCLOSED_CIRCLE = "decimalEnclosedCircle"
	DECIMAL_FULL_WIDTH2 = "decimalFullWidth2"
	AIUEO_FULL_WIDTH = "aiueoFullWidth"
	IROHA_FULL_WIDTH = "irohaFullWidth"
	DECIMAL_ZERO = "decimalZero"
	BULLET = "bullet"
	GANADA = "ganada"
	CHOSUNG = "chosung"
	DECIMAL_ENCLOSED_FULLSTOP = "decimalEnclosedFullstop"
	DECIMAL_ENCLOSED_PAREN = "decimalEnclosedParen"
	DECIMAL_ENCLOSED_CIRCLE_CHINESE = "decimalEnclosedCircleChinese"
	IDEOGRAPH_ENCLOSED_CIRCLE = "ideographEnclosedCircle"
	IDEOGRAPH_TRADITIONAL = "ideographTraditional"
	IDEOGRAPH_ZODIAC = "ideographZodiac"
	IDEOGRAPH_ZODIAC_TRADITIONAL = "ideographZodiacTraditional"
	TAIWANESE_COUNTING = "taiwaneseCounting"
	IDEOGRAPH_LEGAL_TRADITIONAL = "ideographLegalTraditional"
	TAIWANESE_COUNTING_THOUSAND = "taiwaneseCountingThousand"
	TAIWANESE_DIGITAL = "taiwaneseDigital"
	CHINESE_COUNTING = "chineseCounting"
	CHINESE_LEGAL_SIMPLIFIED = "chineseLegalSimplified"
	CHINESE_COUNTING_THOUSAND = "chineseCountingThousand"
	KOREAN_DIGITAL = "koreanDigital"
	KOREAN_COUNTING = "koreanCounting"
	KOREAN_LEGAL = "koreanLegal"
	KOREAN_DIGITAL_2 = "koreanDigital2"
	VIETNAMESE_COUNTING = "vietnameseCounting"
	RUSSIAN_LOWER = "russianLower"
	RUSSIAN_UPPER = "russianUpper"
	NONE = "none"
	NUMBER_IN_DASH = "numberInDash"
	HEBREW1 = "hebrew1"
	HEBREW2 = "hebrew2"
	ARABIC_ALPHA = "arabicAlpha"
	ARABIC_ABJAD = "arabicAbjad"
	HINDI_VOWELS = "hindiVowels"
	HINDI_CONSONANTS = "hindiConsonants"
	HINDI_NUMBERS = "hindiNumbers"
	HINDI_COUNTING = "hindiCounting"
	THAI_LETTERS = "thaiLetters"
	THAI_NUMBERS = "thaiNumbers"
	THAI_COUNTING = "thaiCounting"
	BAHT_TEXT = "bahtText"
	DOLLAR_TEXT = "dollarText"
	CUSTOM = "custom"


class ST_ChapterSep(EnumStrBase):
	COLON = "colon"
	EM_DASH = "emDash"
	EN_DASH = "enDash"
	HYPHEN = "hyphen"
	PERIOD = "period"
	PLAIN = "plain"
	SEMICOLON = "semicolon"
	SPACE = "space"
	WORD = "word"
	NOTHING = "nothing"


class ST_PageOrientation(EnumStrBase):
	PORTRAIT = "portrait"
	LANDSCAPE = "landscape"


class ST_TextDirection(EnumStrBase):
	TB = "tb"
	RL = "rl"
	LR = "lr"
	TBV = "tbV"
	RLV = "rlV"
	LRV = "lrV"
	BT_LR = "btLr"
	LR_TB = "lrTb"
	LR_TB_V = "lrTbV"
	TB_LR_V = "tbLrV"
	TB_RL = "tbRl"
	TB_RL_V = "tbRlV"


class ST_SectionMark(EnumStrBase):
	NEXT_PAGE = "nextPage"
	NEXT_COLUMN = "nextColumn"
	CONTINUOUS = "continuous"
	EVEN_PAGE = "evenPage"
	ODD_PAGE = "oddPage"


class ST_VerticalJc(EnumStrBase):
	TOP = "top"
	CENTER = "center"
	BOTH = "both"
	BOTTOM = "bottom"


class ST_VAnchor(EnumStrBase):
	TEXT = "text"
	MARGIN = "margin"
	PAGE = "page"


class ST_HAnchor(EnumStrBase):
	TEXT = "text"
	MARGIN = "margin"
	PAGE = "page"


class ST_XAlign(EnumStrBase):
	LEFT = "left"
	CENTER = "center"
	RIGHT = "right"
	INSIDE = "inside"
	OUTSIDE = "outside"


class ST_YAlign(EnumStrBase):
	TOP = "top"
	CENTER = "center"
	BOTTOM = "bottom"
	INSIDE = "inside"
	OUTSIDE = "outside"


class ST_TblOverlap(EnumStrBase):
	NEVER = "never"
	OVERLAP = "overlap"


class ST_TblWidth(EnumStrBase):
	AUTO = "auto"
	DXA = "dxa"
	NIL = "nil"
	PERCENT = "pct"


class ST_JcTable(EnumStrBase):
	LEFT = "left"
	CENTER = "center"
	RIGHT = "right"
	START = "start"
	END = "end"


class ST_TblLayoutType(EnumStrBase):
	FIXED = "fixed"
	AUTOFIT = "autofit"


class ST_HeightRule(EnumStrBase):
	AUTO = "auto"
	EXACT = "exact"
	AT_LEAST = "atLeast"


class ST_Merge(EnumStrBase):
	CONTINUE = "continue"
	RESTART = "restart"


class ST_StyleType(EnumStrBase):
	PARAGRAPH = "paragraph"
	CHARACTER = "character"
	TABLE = "table"
	NUMBERING = "numbering"


class ST_TblStyleOverrideType(EnumStrBase):
	WHOLE_TABLE = "wholeTable"
	FIRST_ROW = "firstRow"
	LAST_ROW = "lastRow"
	FIRST_COL = "firstCol"
	LAST_COL = "lastCol"
	BAND1_VERT = "band1Vert"
	BAND2_VERT = "band2Vert"
	BAND1_HORZ = "band1Horz"
	BAND2_HORZ = "band2Horz"
	NE_CELL = "neCell"
	NW_CELL = "nwCell"
	SE_CELL = "seCell"
	SW_CELL = "swCell"


class ST_FontFamily(EnumStrBase):
	DECORATIVE = "decorative"
	MODERN = "modern"
	ROMAN = "roman"
	SCRIPT = "script"
	SWISS = "swiss"
	AUTO = "auto"


class ST_FontPitch(EnumStrBase):
	DEFAULT = "default"
	FIXED = "fixed"
	VARIABLE = "variable"


class ST_TileFlipMode(EnumStrBase):
	NONE = "none"
	X = "x"
	Y = "y"
	XY = "xy"


class ST_RectAlignment(EnumStrBase):
	TL = "tl"
	T = "t"
	TR = "tr"
	L = "l"
	CTR = "ctr"
	R = "r"
	BL = "bl"
	B = "b"
	BR = "br"


class ST_BlackWhiteMode(EnumStrBase):
	CLR = "clr"
	AUTO = "auto"
	GRAY = "gray"
	LT_GRAY = "ltGray"
	INV_GRAY = "invGray"
	GRAY_WHITE = "grayWhite"
	BLACK_GRAY = "blackGray"
	BLACK_WHITE = "blackWhite"
	BLACK = "black"
	WHITE = "white"
	HIDDEN = "hidden"


class ST_BlipCompression(EnumStrBase):
	EMAIL = "email"
	SCREEN = "screen"
	PRINT = "print"
	HQPRINT = "hqprint"
	NONE = "none"


class ST_ShapeType(EnumStrBase):
	LINE = "line"
	LINEINV = "lineInv"
	TRIANGLE = "triangle"
	RTTRIANGLE = "rtTriangle"
	RECT = "rect"
	DIAMOND = "diamond"
	PARALLELOGRAM = "parallelogram"
	TRAPEZOID = "trapezoid"
	NONISOSCELESTRAPEZOID = "nonIsoscelesTrapezoid"
	PENTAGON = "pentagon"
	HEXAGON = "hexagon"
	HEPTAGON = "heptagon"
	OCTAGON = "octagon"
	DECAGON = "decagon"
	DODECAGON = "dodecagon"
	STAR4 = "star4"
	STAR5 = "star5"
	STAR6 = "star6"
	STAR7 = "star7"
	STAR8 = "star8"
	STAR10 = "star10"
	STAR12 = "star12"
	STAR16 = "star16"
	STAR24 = "star24"
	STAR32 = "star32"
	ROUNDRECT = "roundRect"
	ROUND1RECT = "round1Rect"
	ROUND2SAMERECT = "round2SameRect"
	ROUND2DIAGRECT = "round2DiagRect"
	SNIPROUNDRECT = "snipRoundRect"
	SNIP1RECT = "snip1Rect"
	SNIP2SAMERECT = "snip2SameRect"
	SNIP2DIAGRECT = "snip2DiagRect"
	PLAQUE = "plaque"
	ELLIPSE = "ellipse"
	TEARDROP = "teardrop"
	HOMEPLATE = "homePlate"
	CHEVRON = "chevron"
	PIEWEDGE = "pieWedge"
	PIE = "pie"
	BLOCKARC = "blockArc"
	DONUT = "donut"
	NOSMOKING = "noSmoking"
	RIGHTARROW = "rightArrow"
	LEFTARROW = "leftArrow"
	UPARROW = "upArrow"
	DOWNARROW = "downArrow"
	STRIPEDRIGHTARROW = "stripedRightArrow"
	NOTCHEDRIGHTARROW = "notchedRightArrow"
	BENTUPARROW = "bentUpArrow"
	LEFTRIGHTARROW = "leftRightArrow"
	UPDOWNARROW = "upDownArrow"
	LEFTUPARROW = "leftUpArrow"
	LEFTRIGHTUPARROW = "leftRightUpArrow"
	QUADARROW = "quadArrow"
	LEFTARROWCALLOUT = "leftArrowCallout"
	RIGHTARROWCALLOUT = "rightArrowCallout"
	UPARROWCALLOUT = "upArrowCallout"
	DOWNARROWCALLOUT = "downArrowCallout"
	LEFTRIGHTARROWCALLOUT = "leftRightArrowCallout"
	UPDOWNARROWCALLOUT = "upDownArrowCallout"
	QUADARROWCALLOUT = "quadArrowCallout"
	BENTARROW = "bentArrow"
	UTURNARROW = "uturnArrow"
	CIRCULARARROW = "circularArrow"
	LEFTCIRCULARARROW = "leftCircularArrow"
	LEFTRIGHTCIRCULARARROW = "leftRightCircularArrow"
	CURVEDRIGHTARROW = "curvedRightArrow"
	CURVEDLEFTARROW = "curvedLeftArrow"
	CURVEDUPARROW = "curvedUpArrow"
	CURVEDDOWNARROW = "curvedDownArrow"
	SWOOSHARROW = "swooshArrow"
	CUBE = "cube"
	CAN = "can"
	LIGHTNINGBOLT = "lightningBolt"
	HEART = "heart"
	SUN = "sun"
	MOON = "moon"
	SMILEYFACE = "smileyFace"
	IRREGULARSEAL1 = "irregularSeal1"
	IRREGULARSEAL2 = "irregularSeal2"
	FOLDEDCORNER = "foldedCorner"
	BEVEL = "bevel"
	FRAME = "frame"
	HALFFRAME = "halfFrame"
	CORNER = "corner"
	DIAGSTRIPE = "diagStripe"
	CHORD = "chord"
	ARC = "arc"
	LEFTBRACKET = "leftBracket"
	RIGHTBRACKET = "rightBracket"
	LEFTBRACE = "leftBrace"
	RIGHTBRACE = "rightBrace"
	BRACKETPAIR = "bracketPair"
	BRACEPAIR = "bracePair"
	STRAIGHTCONNECTOR1 = "straightConnector1"
	BENTCONNECTOR2 = "bentConnector2"
	BENTCONNECTOR3 = "bentConnector3"
	BENTCONNECTOR4 = "bentConnector4"
	BENTCONNECTOR5 = "bentConnector5"
	CURVEDCONNECTOR2 = "curvedConnector2"
	CURVEDCONNECTOR3 = "curvedConnector3"
	CURVEDCONNECTOR4 = "curvedConnector4"
	CURVEDCONNECTOR5 = "curvedConnector5"
	CALLOUT1 = "callout1"
	CALLOUT2 = "callout2"
	CALLOUT3 = "callout3"
	ACCENTCALLOUT1 = "accentCallout1"
	ACCENTCALLOUT2 = "accentCallout2"
	ACCENTCALLOUT3 = "accentCallout3"
	BORDERCALLOUT1 = "borderCallout1"
	BORDERCALLOUT2 = "borderCallout2"
	BORDERCALLOUT3 = "borderCallout3"
	ACCENTBORDERCALLOUT1 = "accentBorderCallout1"
	ACCENTBORDERCALLOUT2 = "accentBorderCallout2"
	ACCENTBORDERCALLOUT3 = "accentBorderCallout3"
	WEDGERECTCALLOUT = "wedgeRectCallout"
	WEDGEROUNDRECTCALLOUT = "wedgeRoundRectCallout"
	WEDGEELLIPSECALLOUT = "wedgeEllipseCallout"
	CLOUDCALLOUT = "cloudCallout"
	CLOUD = "cloud"
	RIBBON = "ribbon"
	RIBBON2 = "ribbon2"
	ELLIPSERIBBON = "ellipseRibbon"
	ELLIPSERIBBON2 = "ellipseRibbon2"
	LEFTRIGHTRIBBON = "leftRightRibbon"
	VERTICALSCROLL = "verticalScroll"
	HORIZONTALSCROLL = "horizontalScroll"
	WAVE = "wave"
	DOUBLEWAVE = "doubleWave"
	PLUS = "plus"
	FLOWCHARTPROCESS = "flowChartProcess"
	FLOWCHARTDECISION = "flowChartDecision"
	FLOWCHARTINPUTOUTPUT = "flowChartInputOutput"
	FLOWCHARTPREDEFINEDPROCESS = "flowChartPredefinedProcess"
	FLOWCHARTINTERNALSTORAGE = "flowChartInternalStorage"
	FLOWCHARTDOCUMENT = "flowChartDocument"
	FLOWCHARTMULTIDOCUMENT = "flowChartMultidocument"
	FLOWCHARTTERMINATOR = "flowChartTerminator"
	FLOWCHARTPREPARATION = "flowChartPreparation"
	FLOWCHARTMANUALINPUT = "flowChartManualInput"
	FLOWCHARTMANUALOPERATION = "flowChartManualOperation"
	FLOWCHARTCONNECTOR = "flowChartConnector"
	FLOWCHARTPUNCHEDCARD = "flowChartPunchedCard"
	FLOWCHARTPUNCHEDTAPE = "flowChartPunchedTape"
	FLOWCHARTSUMMINGJUNCTION = "flowChartSummingJunction"
	FLOWCHARTOR = "flowChartOr"
	FLOWCHARTCOLLATE = "flowChartCollate"
	FLOWCHARTSORT = "flowChartSort"
	FLOWCHARTEXTRACT = "flowChartExtract"
	FLOWCHARTMERGE = "flowChartMerge"
	FLOWCHARTOFFLINESTORAGE = "flowChartOfflineStorage"
	FLOWCHARTONLINESTORAGE = "flowChartOnlineStorage"
	FLOWCHARTMAGNETICTAPE = "flowChartMagneticTape"
	FLOWCHARTMAGNETICDISK = "flowChartMagneticDisk"
	FLOWCHARTMAGNETICDRUM = "flowChartMagneticDrum"
	FLOWCHARTDISPLAY = "flowChartDisplay"
	FLOWCHARTDELAY = "flowChartDelay"
	FLOWCHARTALTERNATEPROCESS = "flowChartAlternateProcess"
	FLOWCHARTOFFPAGECONNECTOR = "flowChartOffpageConnector"
	ACTIONBUTTONBLANK = "actionButtonBlank"
	ACTIONBUTTONHOME = "actionButtonHome"
	ACTIONBUTTONHELP = "actionButtonHelp"
	ACTIONBUTTONINFORMATION = "actionButtonInformation"
	ACTIONBUTTONFORWARDNEXT = "actionButtonForwardNext"
	ACTIONBUTTONBACKPREVIOUS = "actionButtonBackPrevious"
	ACTIONBUTTONEND = "actionButtonEnd"
	ACTIONBUTTONBEGINNING = "actionButtonBeginning"
	ACTIONBUTTONRETURN = "actionButtonReturn"
	ACTIONBUTTONDOCUMENT = "actionButtonDocument"
	ACTIONBUTTONSOUND = "actionButtonSound"
	ACTIONBUTTONMOVIE = "actionButtonMovie"
	GEAR6 = "gear6"
	GEAR9 = "gear9"
	FUNNEL = "funnel"
	MATHPLUS = "mathPlus"
	MATHMINUS = "mathMinus"
	MATHMULTIPLY = "mathMultiply"
	MATHDIVIDE = "mathDivide"
	MATHEQUAL = "mathEqual"
	MATHNOTEQUAL = "mathNotEqual"
	CORNERTABS = "cornerTabs"
	SQUARETABS = "squareTabs"
	PLAQUETABS = "plaqueTabs"
	CHARTX = "chartX"
	CHARTSTAR = "chartStar"
	CHARTPLUS = "chartPlus"


class ST_RelFromH(EnumStrBase):
	MARGIN = "margin"
	PAGE = "page"
	TEXT = "text"
	COLUMN = "column"
	CHARACTER = "character"
	LEFT_MARGIN = "leftMargin"
	RIGHT_MARGIN = "rightMargin"
	INSIDE_MARGIN = "insideMargin"
	OUTSIDE_MARGIN = "outsideMargin"


class ST_AlignH(EnumStrBase):
	LEFT = "left"
	CENTER = "center"
	RIGHT = "right"
	INSIDE = "inside"
	OUTSIDE = "outside"


class ST_RelFromV(EnumStrBase):
	
	MARGIN = "margin"
	PAGE = "page"
	PARAGRAPH = "paragraph"
	LINE = "line"
	TOP_MARGIN = "topMargin"
	BOTTOM_MARGIN = "bottomMargin"
	INSIDE_MARGIN = "insideMargin"
	OUTSIDE_MARGIN = "outsideMargin"


class ST_AlignV(EnumStrBase):
	TOP = "top"
	BOTTOM = "bottom"
	CENTER = "center"
	INSIDE = "inside"
	OUTSIDE = "outside"


class ST_BrType(EnumStrBase):
	PAGE = "page"
	COLUMN = "column"
	TEXT_WRAPPING = "textWrapping"


class ST_BrClear(EnumStrBase):
	NONE = "none"
	LEFT = "left"
	RIGHT = "right"
	ALL = "all"
