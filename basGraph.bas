Attribute VB_Name = "basGraph"
Option Explicit

'Constants for blend modes
Public Const BLM_AVERAGE = 1                  ' Average mode
Public Const BLM_MULTIPLY = 2                 ' Multiply mode
Public Const BLM_SCREEN = 3                   ' Screen mode
Public Const BLM_DARKEN = 4                   ' Darken mode
Public Const BLM_LIGHTEN = 5                  ' Lighten mode
Public Const BLM_DIFFERENCE = 6               ' Difference mode
Public Const BLM_NEGATION = 7                 ' Negation mode
Public Const BLM_EXCLUSION = 8                ' Exclusion mode
Public Const BLM_OVERLAY = 9                  ' Overlay mode
Public Const BLM_HARDLIGHT = 10               ' Hard Light mode
Public Const BLM_SOFTLIGHT = 11               ' Soft Light mode
Public Const BLM_COLORDODGE = 12              ' Color Dodge mode
Public Const BLM_COLORBURN = 13               ' Color Burn mode
Public Const BLM_SOFTDODGE = 14               ' Soft dodge mode
Public Const BLM_SOFTBURN = 15                ' Soft burn mode
Public Const BLM_REFLECT = 16                 ' Reflect mode
Public Const BLM_GLOW = 17                    ' Glow mode
Public Const BLM_FREEZE = 18                  ' Freeze mode
Public Const BLM_HEAT = 19                    ' Heat mode
Public Const BLM_ADDITIVE = 20                ' Additive mode
Public Const BLM_SUBTRACTIVE = 21             ' Subtractive mode
Public Const BLM_INTERPOLATION = 22           ' Interpolation mode
Public Const BLM_STAMP = 23                   ' Stamp mode
Public Const BLM_XOR = 24                     ' XOR mode

'Constants for Histogram functions like GPX_StretchHistogram
Public Const HST_RED = 1                      ' Red
Public Const HST_GREEN = 2                    ' Green
Public Const HST_BLUE = 4                     ' Blue
Public Const HST_COLOR = 7                    ' All the colors
Public Const HST_GRAY = 8                     ' Gray

'Contants for Gradient functions like GPX_Metallic
Public Const GRAD_METALLIC = 1                ' Metallic
Public Const GRAD_GOLD = 2                    ' Gold gradient
Public Const GRAD_ICE = 3                     ' Ice gradient

'***********************************************************************************'
'**********************             Functions               ************************'
'***********************************************************************************'
Declare Function GPX_AdvancedBlur Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Blur As Integer, _
 ByVal Sense As Integer, _
 ByVal Smart As Boolean, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_AllocBufferSize Lib "cjGraph.dll" _
(ByVal PicSrcDC As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_AlphaBlend Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC_1 As Long, _
 ByVal PicSrcDC_2 As Long, _
 ByVal Alpha As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_AmbientLight Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal AmbientColor As Long, _
 ByVal Intensity As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_AntiAlias Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Sensibility As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_AsciiMorph Lib "cjGraph.dll" _
(ByVal PicSrcDC As Long, _
 ByVal Buffer As String, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_BackDropRemoval Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal SelectColor As Long, _
 ByVal SubstituteColor As Long, _
 ByVal Range As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_BackDropRemovalEx Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal SelectColor As Long, _
 ByVal SubstituteColor As Long, _
 ByVal Range As Long, _
 ByVal Top As Boolean, _
 ByVal Left As Boolean, _
 ByVal Right As Boolean, _
 ByVal Botton As Boolean, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_BitBlt Lib "cjGraph.dll" _
(ByVal DestDC As Long, _
 ByVal XDest As Long, _
 ByVal YDest As Long, _
 ByVal Width As Long, _
 ByVal Height As Long, _
 ByVal SrcDC As Long, _
 ByVal XSrc As Long, _
 ByVal YSrc As Long, _
 ByVal RasterOp As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_BlendMode Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC_1 As Long, _
 ByVal PicSrcDC_2 As Long, _
 ByVal Mode As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_BlockWaves Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Amplitude As Integer, _
 ByVal Frequency As Integer, _
 ByVal Mode As Integer, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Blur Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Brightness Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Value As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Canvas Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Canvas As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Caricature Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Cilindrical Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Value As Double, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_CircularWaves Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Amplitude As Integer, _
 ByVal Frequency As Integer, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_CircularWavesEx Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Amplitude As Integer, _
 ByVal Frequency As Integer, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_ColorBalance Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal R As Integer, _
 ByVal G As Integer, _
 ByVal B As Integer, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_ColorRandomize Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal RandValue As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Contrast Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal R As Single, _
 ByVal G As Single, _
 ByVal B As Single, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_DetectBorders Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Border As Long, _
 ByVal BorderColor As Long, _
 ByVal BackColor As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Diffuse Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Emboss Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Depth As Single, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_FarBlur Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Distance As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_FindEdges Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Intensity As Integer, _
 ByVal BW As Integer, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_FishEye Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_FishEyeEx Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Values As Double, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Flip Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Width As Long, _
 ByVal Height As Long, _
 ByVal Horizontal As Long, _
 ByVal Vertical As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Fog Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Fog As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_FourCorners Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Fragment Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Distance As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_FrostGlass Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Frost As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Gamma Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Gamma As Single, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_GlassBlendMode Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC_1 As Long, _
 ByVal PicSrcDC_2 As Long, _
 ByVal Depth As Double, _
 ByVal Direction As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_GrayScale Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Scales As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Hue Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Hue As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Invert Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Intensity As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Make3DEffect Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Normal As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_MediumTones Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Level As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Melt Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Metallic Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Level As Long, _
 ByVal Shift As Long, _
 ByVal Mode As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Mosaic Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Size As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_MotionBlur Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Angle As Double, _
 ByVal Distance As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Neon Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Intensity As Integer, _
 ByVal BW As Integer, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_NotePaper Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Sensibility As Long, _
 ByVal Depth As Long, _
 ByVal Graininess As Long, _
 ByVal Intensity As Long, _
 ByVal Forecolor As Long, _
 ByVal BackColor As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_OilPaint Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal BrushSize As Long, _
 ByVal Smoothness As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_PolarCoordinates Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Flag As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_RadialBlur Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Distance As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_RainDrops Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal DropSize As Long, _
 ByVal Amount As Long, _
 ByVal Coeff As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_RandomicalPoints Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal RandValue As Long, _
 ByVal BackColor As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_ReduceColors Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Levels As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_ReduceTo2Colors Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_ReduceTo8Colors Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Relief Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Rock Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Value As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Roll Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Saturation Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Saturation As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Sepia Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Sharpening Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Value As Single, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Shift Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Shift As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_SmartBlur Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Sensibility As Integer, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_SoftnerBlur Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Solarize Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Invert As Boolean, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Stamp Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Levels As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_StretchHistogram Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Flag As Long, _
 ByVal StretchFactor As Double, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Swirl Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Swirl As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Tile Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal WSize As Long, _
 ByVal HSize As Long, _
 ByVal RandValue As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Tone Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Color As Long, _
 ByVal Tone As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Twirl Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Twirl As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_TwirlEx Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal TwirlMin As Double, _
 ByVal TwirlMax As Double, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_UnsharpMask Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Blur As Integer, _
 ByVal Unsharp As Double, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_Waves Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Amplitude As Long, _
 ByVal Frequency As Long, _
 ByVal FillSides As Byte, _
 ByVal Direction As Byte, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_WebColors Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByRef Response As Long) As Long
'***********************************************************************************'
Declare Function GPX_ZoomBlur Lib "cjGraph.dll" _
(ByVal PicDestDC As Long, _
 ByVal PicSrcDC As Long, _
 ByVal Distance As Long, _
 ByRef Response As Long) As Long

