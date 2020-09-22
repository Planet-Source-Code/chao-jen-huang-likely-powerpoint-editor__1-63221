VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum typShape
    blkRectangle = 0
    blkSquare = 1
    blkEllipse = 2
    blkCircle = 3
    blkRoundRect = 4
    blkRoundSquare = 5
End Enum

Public Enum typControlType
    cjBlock = 0
    cjPicture = 1
    cjTextBox = 2
End Enum

Public Enum typTextAligment
    txt_Left = 0
    txt_Right = 1
    txt_Center = 2
End Enum

Private Type typBlock
    
    '¥H¤U¬°©Ò¦³ª«¥óªº³q¥ÎÄÝ©Ê
    ControlType             As typControlType
    GroupID                 As Integer
    
    Locked                  As Boolean
    Visibile                As Boolean
    
    Name                    As String
    BackColor               As Long
    BorderColor             As Long
    BorderWidth             As Long

    Size                    As RECT
    Shape                   As typShape
    RoundAngel              As Long
    Selected                As Boolean
    
    '¥H¤U¬° Picture ª«¥óªºÄÝ©Ê
    PictureFileName         As String
    PicTransparent          As Boolean
        
    '¥H¤U¬° TextBox ª«¥óªºÄÝ©Ê
    Text                    As String
    TextAligment            As typTextAligment
    FontName                As String
    FontSize                As Long
    FontBold                As Boolean
    FontItlic               As Boolean
    FontUnderline           As Boolean
    FontForeColor           As Long
    
End Type

Private mBlock          As typBlock

Public Property Let Locked(Locked As Boolean)
    mBlock.Locked = Locked
End Property
Public Property Get Locked() As Boolean
    Locked = mBlock.Locked
End Property

Public Property Let TextAligment(Aligment As typTextAligment)
    mBlock.TextAligment = Aligment
End Property
Public Property Get TextAligment() As typTextAligment
    TextAligment = mBlock.TextAligment
End Property

Public Property Let GroupID(GroupID As Integer)
    mBlock.GroupID = GroupID
End Property
Public Property Get GroupID() As Integer
    GroupID = mBlock.GroupID
End Property

Public Property Let Name(Name As String)
    mBlock.Name = Name
End Property
Public Property Get Name() As String
    Name = mBlock.Name
End Property

Public Property Let PicTransparent(Transparent As Boolean)
    mBlock.PicTransparent = Transparent
End Property
Public Property Get PicTransparent() As Boolean
    PicTransparent = mBlock.PicTransparent
End Property

Public Property Let FontName(FontName As String)
    mBlock.FontName = FontName
End Property
Public Property Get FontName() As String
    FontName = mBlock.FontName
End Property

Public Property Let FontSize(FontSize As Long)
    mBlock.FontSize = FontSize
End Property
Public Property Get FontSize() As Long
    FontSize = mBlock.FontSize
End Property

Public Property Let FontBold(FontBold As Boolean)
    mBlock.FontBold = FontBold
End Property
Public Property Get FontBold() As Boolean
    FontBold = mBlock.FontBold
End Property

Public Property Let FontItlic(FontItlic As Boolean)
    mBlock.FontItlic = FontItlic
End Property
Public Property Get FontItlic() As Boolean
    FontItlic = mBlock.FontItlic
End Property

Public Property Let FontUnderline(FontUnderline As Boolean)
    mBlock.FontUnderline = FontUnderline
End Property
Public Property Get FontUnderline() As Boolean
    FontUnderline = mBlock.FontUnderline
End Property

Public Property Let FontForeColor(ForeColor As Long)
    mBlock.FontForeColor = ForeColor
End Property
Public Property Get FontForeColor() As Long
    FontForeColor = mBlock.FontForeColor
End Property

Public Property Let Text(Text As String)
    mBlock.Text = Text
End Property
Public Property Get Text() As String
    Text = mBlock.Text
End Property

Public Property Let RoundAngel(Angel As Long)
    mBlock.RoundAngel = Angel
End Property
Public Property Get RoundAngel() As Long
    RoundAngel = mBlock.RoundAngel
End Property

Public Property Get PictureFileName() As String
    PictureFileName = mBlock.PictureFileName
End Property
Public Property Let PictureFileName(strFileName As String)
    mBlock.PictureFileName = strFileName
End Property

Public Property Let ControlType(ControlType As typControlType)
    mBlock.ControlType = ControlType
End Property
Public Property Get ControlType() As typControlType
    ControlType = mBlock.ControlType
End Property

Public Property Let Selected(blnSelected As Boolean)
    mBlock.Selected = blnSelected
End Property
Public Property Get Selected() As Boolean
    Selected = mBlock.Selected
End Property

Public Property Let Shape(Shape As typShape)
    mBlock.Shape = Shape
End Property
Public Property Get Shape() As typShape
    Shape = mBlock.Shape
End Property

Public Property Let BorderWidth(Width As Integer)
    mBlock.BorderWidth = Width
End Property
Public Property Get BorderWidth() As Integer
    BorderWidth = mBlock.BorderWidth
End Property

Public Property Let BorderColor(Color As Long)
    mBlock.BorderColor = Color
End Property
Public Property Get BorderColor() As Long
    BorderColor = mBlock.BorderColor
End Property

Public Property Let BackColor(Color As Long)
    mBlock.BackColor = Color
End Property
Public Property Get BackColor() As Long
    BackColor = mBlock.BackColor
End Property

Public Property Let Left(Left As Long)
    mBlock.Size.Left = Left
End Property
Public Property Get Left() As Long
    Left = mBlock.Size.Left
End Property

Public Property Let Top(Top As Long)
    mBlock.Size.Top = Top
End Property
Public Property Get Top() As Long
    Top = mBlock.Size.Top
End Property

Public Property Let Width(Width As Long)
    mBlock.Size.Width = Width
End Property
Public Property Get Width() As Long
    Width = mBlock.Size.Width
End Property

Public Property Let Height(Height As Long)
    mBlock.Size.Height = Height
End Property
Public Property Get Height() As Long
    Height = mBlock.Size.Height
End Property

Private Sub Class_Initialize()

    With mBlock
    
        .Locked = False
        .Visibile = True
    
        .ControlType = cjBlock
        .BackColor = vbWhite
        .BorderColor = vbBlack
        .BorderWidth = 1
        .Shape = blkRoundRect
        .RoundAngel = 20
        .Selected = False
        
        .PicTransparent = True
        
        .FontBold = True
        .FontForeColor = vbBlack
        .FontSize = 12
        .FontName = "Times New Roman"
        .TextAligment = txt_Center
        
    End With
    
End Sub

Private Sub Class_Terminate()
    'Set mBlock.Picture = Nothing
End Sub
