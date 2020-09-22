VERSION 5.00
Begin VB.UserControl cjEdt 
   ClientHeight    =   5460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   KeyPreview      =   -1  'True
   ScaleHeight     =   364
   ScaleMode       =   3  '¹³¯À
   ScaleWidth      =   492
   ToolboxBitmap   =   "cjEdt.ctx":0000
   Begin cjEditor.cjRuler vr 
      Height          =   3735
      Left            =   0
      Top             =   240
      Width           =   240
      _extentx        =   423
      _extenty        =   6588
      orientation     =   1
   End
   Begin cjEditor.cjRuler hr 
      Height          =   240
      Left            =   240
      Top             =   0
      Width           =   4455
      _extentx        =   7858
      _extenty        =   423
   End
   Begin VB.HScrollBar hsc 
      Height          =   255
      LargeChange     =   80
      Left            =   3240
      SmallChange     =   10
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.VScrollBar vsc 
      Height          =   855
      LargeChange     =   80
      Left            =   4080
      SmallChange     =   10
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picBG 
      Appearance      =   0  '¥­­±
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   240
      ScaleHeight     =   247
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   294
      TabIndex        =   0
      Top             =   240
      Width           =   4440
      Begin VB.TextBox rtb 
         Appearance      =   0  '¥­­±
         BorderStyle     =   0  '¨S¦³®Ø½u
         BeginProperty Font 
            Name            =   "·s²Ó©úÅé"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   600
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Shape spObj 
         BorderColor     =   &H00404040&
         BorderStyle     =   3  'ÂI½u
         DrawMode        =   6  'Mask Pen Not
         Height          =   495
         Index           =   0
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape spSelect 
         BorderStyle     =   3  'ÂI½u
         DrawMode        =   6  'Mask Pen Not
         Height          =   495
         Left            =   0
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Shape spBG 
      FillStyle       =   0  '¹ê¤ß
      Height          =   3735
      Left            =   360
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "cjEdt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRECT As RECT) As Long
'Private Declare Sub ClientToScreen Lib "user32" (ByVal hwnd As Long, lpp As POINTAPI)
      
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
            
'CreateHatchBrush API ©Ò¨Ï¥Îªºµ§¨êªá¦â
Private Const HS_HORIZONTAL = 0
Private Const HS_VERTICAL = 1
Private Const HS_FDIAGONAL = 2
Private Const HS_BDIAGONAL = 3
Private Const HS_CROSS = 4
Private Const HS_DIAGCROSS = 5
           
Private Const BOXSIZE = 7 '¹Ï§Îªº¼e¤Î±±¨îÂIÂ÷¹Ï§Îªº¶ZÂ÷

Private Enum ControlPoint
    TOPLEFT = 0
    TOPMIDDLE
    TOPRIGHT
    MIDDLERIGHT
    BOTTOMRIGHT
    BOTTOMMIDDLE
    BOTTOMLEFT
    MIDDLELEFT
    TRANSLATE
    OUTSIDE
End Enum

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type typcjEditor
    Size                    As RECT
    PictureBox              As PictureBox
    PictureFileName         As String
    GridShow                As Boolean
    GridScale               As Integer
End Type

Public Enum enuMouseAction
    MoveObject = 0
    MoveSelect = 1
    MoveTopLeftSize = 2
    MoveTopMiddleSize = 3
    MoveTopRightSize = 4
    MoveMiddleLeftSize = 5
    MoveMiddleRightSize = 6
    MoveBottomLeftSize = 7
    MoveBottomMiddleSize = 8
    MoveBottomRightSize = 9
            
    InsertBlock = 10
    InsertPicture = 11
    InsertTextbox = 12
    
    EditTextBox = 13
End Enum

Private mMouseAct           As enuMouseAction

Private cjEditor            As typcjEditor
Private mSelRect            As RECT

Private mBlock()            As clsBlock
Private mBlockPic()         As PictureBox
Private mBlockCnt           As Integer
Private mBlockIndex         As Integer
Private mTextIndex          As Integer

Private mClipboard()        As clsBlock
Private mClipboardPic()     As PictureBox
Private mClipBoardCnt       As Integer

Private mlngGroupX1         As Long
Private mlngGroupY1         As Long
Private mlngGroupX2         As Long
Private mlngGroupY2         As Long

Private mblnDrawObjects     As Boolean

Private Cursor(9)           As Integer ' ·Æ¹«´å¼Ðªº°}¦C

'----------------------------------------- ÄÝ©ó clsBlock ª«¥óªºÄÝ©Ê«Å§i -----------------------------------
Public Property Let objText(Text As String)
    Dim i       As Integer
    For i = 1 To mBlockCnt
        If mBlock(i).Selected Then
            mBlock(i).Text = Text
            Call GetTextBoxProperties(mBlock(i))
        End If
    Next
    Call DrawObjects
End Property
Public Property Get objText() As String
    If mBlockIndex > 0 Then objText = mBlock(mBlockIndex).Text
End Property

Public Property Let objTextAligment(Aligment As typTextAligment)
    Dim i       As Integer
    For i = 1 To mBlockCnt
        If mBlock(i).Selected Then
            mBlock(i).TextAligment = Aligment
            Call GetTextBoxProperties(mBlock(i))
        End If
    Next
    Call DrawObjects
End Property
Public Property Get objTextAligment() As typTextAligment
    If mBlockIndex > 0 Then objTextAligment = mBlock(mBlockIndex).TextAligment
End Property

Public Property Get objHDC() As Long
    If mBlockIndex > 0 Then objHDC = mBlockPic(mBlockIndex).hDC
End Property
Public Property Get objPicture() As Picture
    If mBlockIndex > 0 Then Set objPicture = mBlockPic(mBlockIndex).Picture
End Property

Public Sub objGroup()
    Dim i               As Integer
    Dim ii              As Integer
    Dim iEnd            As Integer
    Dim Block           As clsBlock
    Dim Pic             As Picture
    Dim intGroup        As Integer
        
    '¸s²Õ¨BÆJ¤@¡G¨ú¥X¥Ø«e¥i¥H¥Îªº¸s²Õ½s¸¹
    intGroup = 1
    For i = 1 To mBlockCnt
        If mBlock(i).GroupID >= intGroup Then intGroup = mBlock(i).GroupID + 1
    Next
    
    '¸s²Õ¨BÆJ¤G¡G±N©Ò¦³¿ï¨úªºª«¥ó¶}©l©¹¤W±À¦Ü¿ï¨úª«¥ó³Ì°ª¼hªº¾Fªñ
    '¨D¥X¥i¥H¤W±Àªº³Ì³»¼h
    For i = mBlockCnt To 1 Step -1
        If mBlock(i).Selected Then
            mBlock(i).GroupID = intGroup
            iEnd = i
            Exit For
        End If
    Next
    
    '¶}©l©¹¤W±À
    For i = 1 To iEnd - 1
        
        If mBlock(i).Selected Then
            
            mBlock(i).GroupID = intGroup
            
            For ii = i To iEnd - 1
                
                '§PÂ_¤W¤@¼hªºª«¥ó¬O§_¤w¸g³Q¿ï¨ú¡A§_ªº¸Ü¤~·|©M¤W¤@¼h¥æ´«
                If mBlock(ii + 1).Selected = False Then
                                                                            
                    '¥æ´« clsBlock ª«¥ó
                    Set Block = mBlock(ii)
                    Set mBlock(ii) = mBlock(ii + 1)
                    Set mBlock(ii + 1) = Block
                    Set Block = Nothing
                    
                    '¥æ´« mBlockPic ª«¥ó
                    Set Pic = mBlockPic(ii).Picture
                    mBlockPic(ii).Picture = mBlockPic(ii + 1).Picture
                    mBlockPic(ii + 1).Picture = Pic
                    Set Pic = Nothing
                                        
                End If
                
            Next
                                    
        End If
                
    Next
    
    Call DrawObjects

End Sub

Public Sub objUnGroup()
    Dim i       As Integer
    For i = 1 To mBlockCnt
        If mBlock(i).Selected Then mBlock(i).GroupID = 0
    Next
    Call DrawObjects
End Sub

Public Sub objSetSelectedObject(Index As Integer)
    If mBlockCnt > 0 And Index <= mBlockIndex Then
        mBlockIndex = Index
        mBlock(Index).Selected = True
    End If
    Call DrawObjects
End Sub
Public Sub objCancelSelectedObject(Index As Integer)
    If mBlockCnt > 0 And Index <= mBlockIndex Then
        mBlockIndex = 0
        mBlock(Index).Selected = False
    End If
    Call DrawObjects
End Sub

Public Property Get objSelectedIndex() As Integer
    objSelectedIndex = mBlockIndex
End Property

Public Property Get objSelectedObject() As clsBlock
    If mBlockIndex > 0 Then Set objSelectedObject = mBlock(mBlockIndex)
End Property

Public Sub objPictureTransparent(Transparent As Boolean)
    Dim i       As Integer
    For i = 1 To mBlockCnt
        If mBlock(i).Selected And mBlock(i).ControlType = cjPicture Then
            mBlock(i).PicTransparent = Transparent
        End If
    Next
    Call DrawObjects
End Sub

Public Property Let objFontName(FontName As String)
    Dim i       As Integer
    For i = 1 To mBlockCnt
        If mBlock(i).Selected Then
            mBlock(i).FontName = FontName
            Call GetTextBoxProperties(mBlock(i))
        End If
    Next
    Call DrawObjects
End Property
Public Property Get objFontName() As String
    If mBlockIndex > 0 Then objFontName = mBlock(mBlockIndex).FontName
End Property

Public Property Let objFontBold(Bold As Boolean)
    Dim i       As Integer
    For i = 1 To mBlockCnt
        If mBlock(i).Selected Then
            mBlock(i).FontBold = Bold
            Call GetTextBoxProperties(mBlock(i))
        End If
    Next
    Call DrawObjects
End Property
Public Property Get objFontBold() As Boolean
    If mBlockIndex > 0 Then objFontBold = mBlock(mBlockIndex).FontBold
End Property

Public Property Let objFontUnderline(Underline As Boolean)
    Dim i       As Integer
    For i = 1 To mBlockCnt
        If mBlock(i).Selected Then
            mBlock(i).FontUnderline = Underline
            Call GetTextBoxProperties(mBlock(i))
        End If
    Next
    Call DrawObjects
End Property
Public Property Get objFontUnderline() As Boolean
    If mBlockIndex > 0 Then objFontUnderline = mBlock(mBlockIndex).FontUnderline
End Property

Public Property Let objFontItlic(Itlic As Boolean)
    Dim i       As Integer
    For i = 1 To mBlockCnt
        If mBlock(i).Selected Then
            mBlock(i).FontItlic = Itlic
            Call GetTextBoxProperties(mBlock(i))
        End If
    Next
    Call DrawObjects
End Property
Public Property Get objFontItlic() As Boolean
    If mBlockIndex > 0 Then objFontItlic = mBlock(mBlockIndex).FontItlic
End Property

Public Property Let objFontForeColor(Color As Long)
    Dim i       As Integer
    For i = 1 To mBlockCnt
        If mBlock(i).Selected Then
            mBlock(i).FontForeColor = Color
            Call GetTextBoxProperties(mBlock(i))
        End If
    Next
    Call DrawObjects
End Property
Public Property Get objFontForeColor() As Long
    If mBlockIndex > 0 Then objFontForeColor = mBlock(mBlockIndex).FontForeColor
End Property

Public Property Let objFontSize(Size As Long)
    Dim i       As Integer
    For i = 1 To mBlockCnt
        If mBlock(i).Selected Then
            mBlock(i).FontSize = Size
            Call GetTextBoxProperties(mBlock(i))
        End If
    Next
    Call DrawObjects
End Property
Public Property Get objFontSize() As Long
    If mBlockIndex > 0 Then objFontSize = mBlock(mBlockIndex).FontSize
End Property

Public Property Let objRoundAngel(Angel As Long)
    Dim i       As Integer
    
    For i = 1 To mBlockCnt
        If mBlock(i).Selected And mBlock(i).ControlType = cjBlock Then mBlock(i).RoundAngel = Angel
    Next
    Call DrawObjects
End Property
Public Property Get objRoundAngel() As Long
    If mBlockIndex > 0 Then objRoundAngel = mBlock(mBlockIndex).RoundAngel
End Property

Public Property Let objShpae(Shape As typShape)
    Dim i       As Integer
    
    For i = 1 To mBlockCnt
        If mBlock(i).Selected And mBlock(i).ControlType = cjBlock Then mBlock(i).Shape = Shape
    Next
    Call DrawObjects
End Property
Public Property Get objShpae() As typShape
    If mBlockIndex > 0 Then objShpae = mBlock(mBlockIndex).Shape
End Property

Public Property Let objBorderWidth(Width As Long)
    Dim i       As Integer
    
    For i = 1 To mBlockCnt
        If mBlock(i).Selected And mBlock(i).ControlType = cjBlock Then mBlock(i).BorderWidth = Width
    Next
    Call DrawObjects
End Property

Public Property Get objBorderWidth() As Long
    If mBlockIndex > 0 Then objBorderWidth = mBlock(mBlockIndex).BorderWidth
End Property

Public Property Let objBorderColor(Color As Long)
    Dim i       As Integer
    
    For i = 1 To mBlockCnt
        If mBlock(i).Selected And mBlock(i).ControlType = cjBlock Then mBlock(i).BorderColor = Color
    Next
    Call DrawObjects
End Property
Public Property Get objBorderColor() As Long
    If mBlockIndex > 0 Then objBorderColor = mBlock(mBlockIndex).BorderColor
End Property

Public Property Let objBackColor(Color As Long)
    Dim i       As Integer
    
    For i = 1 To mBlockCnt
        If mBlock(i).Selected And mBlock(i).ControlType = cjBlock Then mBlock(i).BackColor = Color
    Next
    Call DrawObjects
End Property
Public Property Get objBackColor() As Long
    If mBlockIndex > 0 Then objBackColor = mBlock(mBlockIndex).BackColor
End Property

Public Sub objLoadPicture(strFileName As String)

    Dim i       As Integer
    For i = 1 To mBlockCnt
        If mBlock(i).Selected And mBlock(i).ControlType = cjPicture Then
            mBlockPic(i).Picture = LoadPicture(strFileName)
            mBlock(i).PictureFileName = strFileName
        End If
    Next
    
'    If mBlockIndex > 0 Then
'        If mBlock(mBlockIndex).ControlType = cjPicture Then
'            mBlockPic(mBlockIndex).Picture = LoadPicture(strFileName)
'            mBlock(mBlockIndex).PictureFileName = strFileName
'        End If
'    End If
    
    Call DrawObjects
End Sub

Public Property Let objControlType(ControlType As typControlType)

    Dim i       As Integer
    For i = 1 To mBlockCnt
        If mBlock(i).Selected Then mBlock(i).ControlType = ControlType
    Next
    Call DrawObjects
End Property

Public Property Get objControlType() As typControlType
    If mBlockIndex > 0 Then objControlType = mBlock(mBlockIndex).ControlType
End Property

Public Sub objLayerUp()
    Dim Block           As clsBlock
    Dim Pic             As Picture
    Dim i               As Integer
    Dim ii              As Integer
    Dim iGroup          As Integer
    
    If mBlockCnt = 0 Then Exit Sub
        
    '¥Ñ³Ì°ª¼hªº¦¸¤@­Ó¶}©l¡£³Ì°ª¼h¦¸¤£»Ý­n¦A©¹¤W¡¤
    For i = mBlockCnt - 1 To 1 Step -1
        
        If mBlock(i).Selected Then
             
             For ii = i To mBlockCnt - 1
                                          
                If mBlock(ii + 1).Selected = False Then
                                                        
                    If iGroup <> 0 Then
                        If iGroup <> mBlock(ii + 1).GroupID Then Exit For
                    End If
                    If mBlock(ii + 1).GroupID <> 0 Then iGroup = mBlock(ii + 1).GroupID
                    
                    '¥æ´« clsBlock ª«¥ó
                    Set Block = mBlock(ii)
                    Set mBlock(ii) = mBlock(ii + 1)
                    Set mBlock(ii + 1) = Block
                    Set Block = Nothing
    
                    '¥æ´« mBlockPic ª«¥ó
                    Set Pic = mBlockPic(ii).Picture
                    mBlockPic(ii).Picture = mBlockPic(ii + 1).Picture
                    mBlockPic(ii + 1).Picture = Pic
                    Set Pic = Nothing
                    
                    If iGroup = 0 Then Exit For
                                        
                End If
                
             Next
                         
        End If
        
    Next
            
    Call DrawObjects
    
End Sub

Public Sub objLayerTop()
    Dim Block           As clsBlock
    Dim Pic             As Picture
    Dim i               As Integer
    Dim ii              As Integer
    
    If mBlockCnt = 0 Then Exit Sub
    
    '¥Ñ³Ì°ª¼h¦¸©¹¤U§ä¬O§_¦³¤W±Àªºª«¥ó
    For i = mBlockCnt - 1 To 1 Step -1
        If mBlock(i).Selected Then
            
            '§ä¨ì­n¤W±Àªºª«¥ó«á¡A¶}©l±Nª«¥ó±À¦Ü³Ì³»¼h
            For ii = i To mBlockCnt - 1
                
                If mBlock(ii + 1).Selected = False Then
                    '¥æ´« clsBlock ª«¥ó
                    Set Block = mBlock(ii)
                    Set mBlock(ii) = mBlock(ii + 1)
                    Set mBlock(ii + 1) = Block
                    Set Block = Nothing

                    '¥æ´« mBlockPic ª«¥ó
                    Set Pic = mBlockPic(ii).Picture
                    mBlockPic(ii).Picture = mBlockPic(ii + 1).Picture
                    mBlockPic(ii + 1).Picture = Pic
                    Set Pic = Nothing
                
                End If
                
            Next
            
        End If
    Next
            
    Call DrawObjects

End Sub

Public Sub objLayerDown()
    Dim Block           As clsBlock
    Dim Pic             As Picture
    Dim i               As Integer
    Dim ii              As Integer
    Dim iGroup          As Integer
    
    If mBlockCnt = 0 Then Exit Sub
        
    '¥Ñ³Ì§C¼hªº¦¸¤@­Ó¶}©l¡£³Ì§C¼h¦¸¤£»Ý­n¦A©¹¤U¡¤
    For i = 2 To mBlockCnt
        
        If mBlock(i).Selected Then
            
             For ii = i To 2 Step -1
                                          
                If mBlock(ii - 1).Selected = False Then
                                                        
                    If iGroup <> 0 Then
                        If iGroup <> mBlock(ii - 1).GroupID Then Exit For
                    End If
                    If mBlock(ii - 1).GroupID <> 0 Then iGroup = mBlock(ii - 1).GroupID
                    
                    '¥æ´« clsBlock ª«¥ó
                    Set Block = mBlock(ii)
                    Set mBlock(ii) = mBlock(ii - 1)
                    Set mBlock(ii - 1) = Block
                    Set Block = Nothing
    
                    '¥æ´« mBlockPic ª«¥ó
                    Set Pic = mBlockPic(ii).Picture
                    mBlockPic(ii).Picture = mBlockPic(ii - 1).Picture
                    mBlockPic(ii - 1).Picture = Pic
                    Set Pic = Nothing
                    
                    If iGroup = 0 Then Exit For
                                        
                End If
                
            Next
                        
        End If
        
    Next
    
    Call DrawObjects

End Sub

Public Sub objLayerBottom()
    Dim Block           As clsBlock
    Dim Pic             As Picture
    Dim i               As Integer
    Dim ii              As Integer
    
    If mBlockCnt = 0 Then Exit Sub
    
    '¥Ñ³Ì§C¼h¦¸©¹¤W§ä¬O§_¦³¤U±Àªºª«¥ó
    For i = 2 To mBlockCnt
    
        If mBlock(i).Selected Then
                        
            '§ä¨ì­n¤U±Àªºª«¥ó«á¡A¶}©l±Nª«¥ó±À¦Ü³Ì©³¼h
            For ii = i To 2 Step -1
            
                If mBlock(ii - 1).Selected = False Then
                    '¥æ´« clsBlock ª«¥ó
                    Set Block = mBlock(ii)
                    Set mBlock(ii) = mBlock(ii - 1)
                    Set mBlock(ii - 1) = Block
                    Set Block = Nothing

                    '¥æ´« mBlockPic ª«¥ó
                    Set Pic = mBlockPic(ii).Picture
                    mBlockPic(ii).Picture = mBlockPic(ii - 1).Picture
                    mBlockPic(ii - 1).Picture = Pic
                    Set Pic = Nothing
                    
                End If
            
            Next
            
        End If
    Next
                    
    Call DrawObjects
    
End Sub

Public Sub objCut()
    Call objCopy
    Call objDelete
End Sub

Private Sub SetBlockObject(objSrc As clsBlock, objDest As clsBlock)
    
    With objDest
                
        .BackColor = objSrc.BackColor
        .BorderColor = objSrc.BorderColor
        .BorderWidth = objSrc.BorderWidth
        .GroupID = objSrc.GroupID
        .ControlType = objSrc.ControlType
        .Height = objSrc.Height
        .Left = objSrc.Left
        
        .Locked = objSrc.Locked
        
'        .Name = objSrc.Name
        .PicTransparent = objSrc.PicTransparent
        .PictureFileName = objSrc.PictureFileName
        .RoundAngel = objSrc.RoundAngel
        .Selected = objSrc.Selected
        .Shape = objSrc.Shape
        .Top = objSrc.Top
        .Width = objSrc.Width
        
        .Text = objSrc.Text
        .TextAligment = objSrc.TextAligment
        .FontBold = objSrc.FontBold
        .FontForeColor = objSrc.FontForeColor
        .FontItlic = objSrc.FontItlic
        .FontName = objSrc.FontName
        .FontSize = objSrc.FontSize
        
    End With
    
End Sub

Public Sub objCopy()
    Dim i           As Integer
    
    '²MªÅ­ì¥»ªº°Å¶KÃ¯
    For i = 0 To mClipBoardCnt - 1
        Controls.Remove ("picClip" & Trim(i))
    Next
    
    mClipBoardCnt = 0

    For i = 1 To mBlockCnt

        If mBlock(i).Selected Then
                        
            '«Ø¥ß clsBlock °Å¶KÃ¯ª«¥ó
            ReDim Preserve mClipboard(mClipBoardCnt) As clsBlock
            Set mClipboard(mClipBoardCnt) = New clsBlock
            
            Call SetBlockObject(mBlock(i), mClipboard(mClipBoardCnt))
                                            
            '«Ø¥ß picClip °Å¶KÃ¯ª«¥ó
            ReDim Preserve mClipboardPic(mClipBoardCnt)
            Set mClipboardPic(mClipBoardCnt) = Controls.Add("VB.PictureBox", "picClip" & Trim(mClipBoardCnt))
            With mClipboardPic(mClipBoardCnt)
                .Picture = mBlockPic(i).Picture
            End With
            
            mClipBoardCnt = mClipBoardCnt + 1

        End If

    Next
    
End Sub

Public Sub objPaste()

    Dim i           As Integer
    Dim ii          As Integer
    Dim intNewGroup As Integer
    Dim intNowGroup As Integer
            
    intNewGroup = 1

    For i = 1 To mBlockCnt
        mBlock(i).Selected = False
    Next

    For i = 0 To mClipBoardCnt - 1

        mBlockCnt = mBlockCnt + 1

        '«Ø¥ß clsBlock ª«¥ó
        ReDim Preserve mBlock(mBlockCnt)
        
        Set mBlock(mBlockCnt) = New clsBlock
        mClipboard(i).Left = mClipboard(i).Left + 10
        mClipboard(i).Top = mClipboard(i).Top + 10
        
        Call SetBlockObject(mClipboard(i), mBlock(mBlockCnt))
        
        With mBlock(mBlockCnt)
        
            Select Case .ControlType
                Case typControlType.cjBlock
                    .Name = GetObjectDefaultName("Object")
                    
                Case typControlType.cjPicture
                    .Name = GetObjectDefaultName("Picture")
                    
                Case typControlType.cjTextBox
                    .Name = GetObjectDefaultName("TextBox")
                
            End Select
                                    
            '­Y¶K¤Wªºª«¥ó¬°¸s²Õ¡A«h«ü©w·sªº¸s²Õ½s¸¹
            If .GroupID <> 0 Then
            
                If intNowGroup <> .GroupID Then
                
                    intNowGroup = .GroupID
                    
                    '¨ú¥X¥Ø«e¥i¥H¥Îªº¸s²Õ½s¸¹
                    For ii = 1 To mBlockCnt
                        If mBlock(ii).GroupID >= intNewGroup Then intNewGroup = mBlock(ii).GroupID + 1
                    Next
                End If
                
                .GroupID = intNewGroup
            End If
                        
        End With
        
        '«Ø¥ß PictureBox ª«¥ó
        ReDim Preserve mBlockPic(mBlockCnt)
        
        Set mBlockPic(mBlockCnt) = Controls.Add("VB.PictureBox", "picBlock" & Trim(mBlockCnt))
        With mBlockPic(mBlockCnt)
            .BackColor = vbWhite
            .ScaleMode = vbPixels
            .AutoRedraw = True
            .AutoSize = True
            .Picture = mClipboardPic(i).Picture
        End With
        
        Load spObj(mBlockCnt)

    Next
    
    Call DrawObjects
End Sub

Public Sub objDelete()

    Dim i           As Integer
    Dim ii          As Integer
    Dim Block       As clsBlock
    Dim Pic         As Picture
    
    mTextIndex = 0: rtb.Visible = False
    
    '¥Ñ«á­±ªºª«¥ó¶}©l©¹«e§ä
    For i = mBlockCnt To 1 Step -1
    
        If mBlock(i).Selected Then
            
            '­Y§ä¨ì­n§R°£ªºª«¥ó¡A«h±N¸Óª«¥ó¤W±À¦Ü³Ì¤W¼h
            For ii = i To mBlockCnt - 1
                
                '¥æ´« clsBlock ª«¥ó
                Set Block = mBlock(ii)
                Set mBlock(ii) = mBlock(ii + 1)
                Set mBlock(ii + 1) = Block
                Set Block = Nothing
                
                '¥æ´« mBlockPic ª«¥ó
                Set Pic = mBlockPic(ii).Picture
                mBlockPic(ii).Picture = mBlockPic(ii + 1).Picture
                mBlockPic(ii + 1).Picture = Pic
                Set Pic = Nothing
            Next
            
            '§R°£³Ì¤W¼hªºª«¥ó
            Unload spObj(mBlockCnt)
            
            Controls.Remove ("picBlock" & Trim(mBlockCnt))
            
            mBlockCnt = mBlockCnt - 1
            ReDim Preserve mBlock(mBlockCnt)
            
            mBlockIndex = 0
                        
        End If
    
    Next
    
    Call DrawObjects
    
End Sub

Public Sub objSelectAll()
    Dim i           As Integer
    For i = 1 To mBlockCnt
        mBlock(i).Selected = True
    Next
    Call DrawObjects
End Sub

Public Sub objSelectAllCancel()
    Dim i           As Integer
    For i = 1 To mBlockCnt
        mBlock(i).Selected = False
    Next
    Call DrawObjects
End Sub


Public Sub objMoveShift(nXshift As Long, nYshift As Long)
    Dim i       As Integer
    
    For i = 1 To mBlockCnt
        If mBlock(i).Selected Then
            mBlock(i).Left = mBlock(i).Left + nXshift
            mBlock(i).Top = mBlock(i).Top + nYshift
        End If
    Next
    Call DrawObjects
End Sub

'----------------------------------------- µ²§ôÄÝ©ó clsBlock ª«¥óªºÄÝ©Ê«Å§i ----------------------------------

Public Property Let CanDrawObjects(blnDraw As Boolean)
    mblnDrawObjects = blnDraw
End Property
Public Property Get CanDrawObjects() As Boolean
    CanDrawObjects = mblnDrawObjects
End Property

Public Function CreateObject(Left As Long, Top As Long, Width As Long, Height As Long, ControlType As typControlType) As Boolean

    mBlockCnt = mBlockCnt + 1
    
    mBlockIndex = mBlockCnt
    
    '«Ø¥ß clsBlock ª«¥ó
    ReDim Preserve mBlock(mBlockCnt)
    Set mBlock(mBlockCnt) = New clsBlock
    
    With mBlock(mBlockCnt)
        .Left = Left / gsngViewScale
        .Top = Top / gsngViewScale
        .Width = Width / gsngViewScale
        .Height = Height / gsngViewScale

        Select Case ControlType
            Case cjBlock
                .ControlType = cjBlock
                .Name = GetObjectDefaultName("Object")
                
            Case cjPicture
                .ControlType = cjPicture
                .Name = GetObjectDefaultName("Picture")
                
            Case cjTextBox
                .ControlType = cjTextBox
                .Name = GetObjectDefaultName("TextBox")
                
        End Select
        
    End With
    
    '«Ø¥ß PictureBox ª«¥ó
    ReDim Preserve mBlockPic(mBlockCnt)
    Set mBlockPic(mBlockCnt) = Controls.Add("VB.PictureBox", "picBlock" & Trim(mBlockCnt))
                        
    With mBlockPic(mBlockCnt)
        .BackColor = RGB(250, 211, 212)
        .ScaleMode = vbPixels
        .AutoRedraw = True
        .AutoSize = True
    End With
                                    
    Load spObj(mBlockCnt)
    
    CreateObject = True
            
End Function

Public Property Let GridShow(IsShow As Boolean)
    cjEditor.GridShow = IsShow
    Call DrawObjects
End Property
Public Property Get GridShow() As Boolean
    GridShow = cjEditor.GridShow
End Property

Public Property Let GridScale(GridScale As Integer)
    cjEditor.GridScale = GridScale
    Call DrawObjects
End Property
Public Property Get GridScale() As Integer
    GridScale = cjEditor.GridScale
End Property

Public Sub LoadBackgroundPicture(strFileName As String)

    If Len(strFileName) > 0 Then
        cjEditor.PictureFileName = strFileName
        'cjEditor.PictureBox.Picture = LoadPicture(strFileName)
        'Set cjEditor.PictureBox.Picture = Nothing
        'Set picBG.Picture = cjEditor.PictureBox.Picture
        picBG.Picture = LoadPicture(strFileName)
    Else
        cjEditor.PictureFileName = ""
'        Set cjEditor.PictureBox.Picture = Nothing
        Set picBG.Picture = Nothing
        
    End If
    
    picBG.Visible = False
    
    Call UserControl_Resize
    Call DrawObjects
    picBG.Visible = True
    
End Sub
Public Property Get BackgroundPictureFileName() As String
    BackgroundPictureFileName = cjEditor.PictureFileName
End Property

Public Property Let BackColor(Color As Long)
    picBG.BackColor = Color
    Call DrawObjects
End Property
Public Property Get BackColor() As Long
    BackColor = picBG.BackColor
End Property

Public Property Let MouseAction(Action As enuMouseAction)
    mMouseAct = Action
End Property

Public Property Get MouseAction() As enuMouseAction
    MouseAction = mMouseAct
End Property

Public Property Let ClientViewPercent(Percent As Integer)
    
    hr.ViewPercent = Percent
    vr.ViewPercent = Percent
    hr.RulerMode = Millimeters
    vr.RulerMode = Millimeters
        
    '±N¥Ø«e½s¿èªº¤å¦r¤è¶ôÁôÂÃ¨Ã¥B¼g¤J¥Ø«e½s¿èªº¤å¦r
    If rtb.Visible Then
        mBlock(mTextIndex).Text = rtb.Text
        rtb.Visible = False
        mTextIndex = 0
    End If
    
    gsngViewScale = Percent / 100
    
    Call UserControl_Resize
    Call DrawObjects
    
End Property
Public Property Get ClientViewPercent() As Integer
    ClientViewPercent = gsngViewScale * 100
End Property

Public Property Let MillimetersWidth(MM As Long)
    cjEditor.Size.Width = ScaleX(MM, vbMillimeters, vbPixels)
    Call UserControl_Resize
End Property
Public Property Get MillimetersWidth() As Long
    MillimetersWidth = cjEditor.Size.Width
End Property

Public Property Let MillimetersHeight(MM As Long)
    cjEditor.Size.Height = ScaleY(MM, vbMillimeters, vbPixels)
    Call UserControl_Resize
End Property
Public Property Get MillimetersHeight() As Long
    MillimetersHeight = cjEditor.Size.Height
End Property

Private Sub GetTextBoxProperties(objBlock As clsBlock, Optional RefreshText As Boolean = False)

    With objBlock
    
        If RefreshText Then rtb.Text = .Text
                        
        rtb.Font.Size = .FontSize * gsngViewScale
        rtb.Font.Name = .FontName
        rtb.Font.Bold = .FontBold
        rtb.Font.Italic = .FontItlic
        rtb.Font.Underline = .FontUnderline
        rtb.ForeColor = .FontForeColor
                
        picBG.Font.Size = .FontSize * gsngViewScale
        picBG.Font.Name = .FontName
        picBG.Font.Bold = .FontBold
        picBG.Font.Italic = .FontItlic
        picBG.Font.Underline = .FontUnderline
        
        Select Case .TextAligment
            Case typTextAligment.txt_Left
                rtb.Alignment = 0
            Case typTextAligment.txt_Right
                rtb.Alignment = 1
            Case typTextAligment.txt_Center
                rtb.Alignment = 2
        End Select
        
        rtb.Move .Left * gsngViewScale, .Top * gsngViewScale, .Width * gsngViewScale, .Height * gsngViewScale
        
    End With

End Sub

Private Sub picBG_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i                       As Integer
    Dim ii                      As Integer
    Dim iTmp                    As Integer
    Dim Left                    As Single
    Dim Top                     As Single
    Dim Height                  As Single
    Dim Width                   As Single
    
    Dim iArray()                As Integer
    Dim blnSkip                 As Boolean
    
    Dim blnIsFirstSelected      As Boolean
    Dim blnIsObjSelected        As Boolean

    With mSelRect
        .Left = x
        .Top = y
        spSelect.Move x, y, 0, 0
    End With
        
'    mMouseAct = MoveSelect '±N·Æ¹«°Ê§@­«¸m¬°¿ï¨ú
    picBG.Cls
    mBlockIndex = 0
    
    If Button = vbLeftButton Or Button = vbRightButton Then
                
        ReDim Preserve iArray(0)
        spSelect.Visible = True
        
        '§PÂ_¬O§_³B©ó½Õ¾ãª«¥ó¤j¤pªº¤u¨ã¤W
        For i = 1 To mBlockCnt
     
            With mBlock(i)
                                
                If .Selected Then
                    
                    If .GroupID = 0 Then
                    
                        'Âà´«¹ïÀ³ªº®y¼Ð¤ñ¨Ò
                        Left = .Left * gsngViewScale
                        Top = .Top * gsngViewScale
                        Width = .Width * gsngViewScale
                        Height = .Height * gsngViewScale
                    Else
                    
                        'Á×§K­«½Æ³B²z¸s²Õªºµ{¦¡°Ï¬q blnSkip
                        blnSkip = False
                        Left = 999: Top = 999
                        Width = 0: Height = 0
                        
                        For ii = 1 To UBound(iArray)
                            If iArray(ii) = .GroupID Then
                                blnSkip = True
                                Exit For
                            End If
                        Next
                        
                        If blnSkip = False Then
                                                                
                            ReDim Preserve iArray(UBound(iArray) + 1)
                            iArray(UBound(iArray)) = .GroupID
                                                                
                            '¨ú±o¸s²Õª«¥óªº½Õ¾ã¾nÂI
                            For ii = 1 To mBlockCnt
                                If .GroupID = mBlock(ii).GroupID Then
                                    If mBlock(ii).Left < Left Then Left = mBlock(ii).Left
                                    If mBlock(ii).Top < Top Then Top = mBlock(ii).Top
                                    If mBlock(ii).Width + mBlock(ii).Left > Width Then Width = mBlock(ii).Width + mBlock(ii).Left
                                    If mBlock(ii).Height + mBlock(ii).Top > Height Then Height = mBlock(ii).Height + mBlock(ii).Top
                                End If
                            Next
                            
                            Left = Left * gsngViewScale
                            Top = Top * gsngViewScale
                            Width = Width * gsngViewScale - Left
                            Height = Height * gsngViewScale - Top
                            
                            mlngGroupX1 = Left
                            mlngGroupY1 = Top
                            mlngGroupX2 = Left + Width
                            mlngGroupY2 = Top + Height
                            
                        End If
                    
                    End If
                    
                    '½Õ¾ã¤W¥ª
                    If x >= Left - BOXSIZE - 1 And y >= Top - BOXSIZE - 1 And _
                        x <= Left - 1 And y <= Top - 1 Then
                                            
                        mMouseAct = MoveTopLeftSize
                        mBlockIndex = i
                                                                              
                        rtb.Visible = False
                        mTextIndex = 0
                        .Text = rtb.Text
                        
                        RaiseEvent MouseDown(Button, Shift, x / gsngViewScale, y / gsngViewScale)
                        
                        Exit Sub
                    End If
                    
                    '½Õ¾ã¤W¤¤
                    If x >= (Left * 2 + Width) / 2 - BOXSIZE / 2 And y >= Top - BOXSIZE - 1 And _
                        x <= (Left * 2 + Width) / 2 + BOXSIZE / 2 And y <= Top - 1 Then
                    
                        mMouseAct = MoveTopMiddleSize
                        mBlockIndex = i
                                                               
                        rtb.Visible = False
                        mTextIndex = 0
                        .Text = rtb.Text
                        
                        RaiseEvent MouseDown(Button, Shift, x / gsngViewScale, y / gsngViewScale)
                        
                        Exit Sub
                    End If
                    
                    '½Õ¾ã¤W¥k
                    If x >= Left + Width And y >= Top - BOXSIZE - 1 And _
                        x <= Left + Width + BOXSIZE And y <= Top - 1 Then
                    
                        mMouseAct = MoveTopRightSize
                        mBlockIndex = i
                                                                               
                        rtb.Visible = False
                        mTextIndex = 0
                        .Text = rtb.Text
                                                
                        RaiseEvent MouseDown(Button, Shift, x / gsngViewScale, y / gsngViewScale)
                                                                               
                        Exit Sub
                    End If
                    
                    '½Õ¾ã¤¤¥ª
                    If x >= Left - BOXSIZE - 1 And y >= (Top * 2 + Height) / 2 - BOXSIZE / 2 And _
                        x <= Left - 1 And y <= (Top * 2 + Height) / 2 + BOXSIZE / 2 Then
                        
                        mMouseAct = MoveMiddleLeftSize
                        mBlockIndex = i
                                                                                                       
                        rtb.Visible = False
                        mTextIndex = 0
                        .Text = rtb.Text
                                                
                        RaiseEvent MouseDown(Button, Shift, x / gsngViewScale, y / gsngViewScale)
                                                                                                       
                        Exit Sub
                    End If
    
                    '½Õ¾ã¤¤¥k
                    If x >= Left + Width And y >= (Top * 2 + Height) / 2 - BOXSIZE / 2 And _
                        x <= Left + Width + BOXSIZE And y <= (Top * 2 + Height) / 2 + BOXSIZE / 2 Then
                        
                        mMouseAct = MoveMiddleRightSize
                        mBlockIndex = i
                        
                        rtb.Visible = False
                        mTextIndex = 0
                        .Text = rtb.Text
                        
                        RaiseEvent MouseDown(Button, Shift, x / gsngViewScale, y / gsngViewScale)
                                                
                        Exit Sub
                    End If
    
                    '½Õ¾ã¤U¥ª
                    If x >= Left - BOXSIZE - 1 And y >= Top + Height And _
                        x <= Left - 1 And y <= Top + Height + BOXSIZE Then
                        
                        mMouseAct = MoveBottomLeftSize
                        mBlockIndex = i
                                                                               
                        rtb.Visible = False
                        mTextIndex = 0
                        .Text = rtb.Text
                        
                        RaiseEvent MouseDown(Button, Shift, x / gsngViewScale, y / gsngViewScale)
                                                            
                        Exit Sub
                    End If
    
                    '½Õ¾ã¤U¤¤
                    If x >= (Left * 2 + Width) / 2 - BOXSIZE / 2 And y >= Top + Height And _
                        x <= (Left * 2 + Width) / 2 + BOXSIZE / 2 And y <= Top + Height + BOXSIZE Then
                        
                        mMouseAct = MoveBottomMiddleSize
                        mBlockIndex = i
                        
                        rtb.Visible = False
                        mTextIndex = 0
                        .Text = rtb.Text
                        
                        RaiseEvent MouseDown(Button, Shift, x / gsngViewScale, y / gsngViewScale)
                        
                        Exit Sub
                    End If
    
                    '½Õ¾ã¤U¥k
                    If x >= Left + Width And y >= Top + Height And _
                        x <= Left + Width + BOXSIZE And y <= Top + Height + BOXSIZE Then
                        
                        mMouseAct = MoveBottomRightSize
                        mBlockIndex = i
                                                                       
                        rtb.Visible = False
                        mTextIndex = 0
                        .Text = rtb.Text
                        
                        RaiseEvent MouseDown(Button, Shift, x / gsngViewScale, y / gsngViewScale)
                        
                        Exit Sub
                    End If
                    
                End If
                
            End With
        
        Next
        
        '¥Î¨Ó§PÂ_¥Ø«e«ö¤Uªºª«¥ó¬O§_¬° Mouse_Up ¥Ñ½d³ò¿ï¨ú¿ï¾Üªºª«¥ó(s)¡A
        '­Y¬Oªº¸Ü blnIsFirstSelected = False ­Y§_ªº¸Ü blnIsFirstSelected = True
        '¥»¬qµ{¦¡½X¥Î©ó·í¨Ï¥ÎªÌ·Q­n±N¥ý«e Mouse_Up ¿ï¨úªºª«¥ó¡A¸s²Õ²¾°Ê¡C
        
        blnIsFirstSelected = True
        blnIsObjSelected = False
        
        '¤£¬O²Ä¤@¦¸«ö¤Uª«¥óªº³B²zµ{§Ç
        For i = mBlockCnt To 1 Step -1
        
            With mBlock(i)
                'Âà´«¹ïÀ³ªº®y¼Ð¤ñ¨Ò
                Left = .Left * gsngViewScale
                Top = .Top * gsngViewScale
                Width = .Width * gsngViewScale
                Height = .Height * gsngViewScale
                
                '§PÂ_ª«¥ó¬O§_¦b¤§«eªº´N¤w¸g¿ï¨ú¹L
                If x >= Left And y >= Top _
                        And x <= Left + Width And y <= Top + Height Then
                    
                    mBlockIndex = i
                    
                    If mBlock(i).Selected Then
                        mMouseAct = MoveObject
                                                                        
                        spObj(i).Move Left, Top, Width, Height
                        blnIsFirstSelected = False
                                                
                        '­Y«ö¦íshift·|¨ú®ø¥Ø«eÂI¿ïªºª«¥ó
                        If Shift = 1 Then
                            .Selected = False
                            If .GroupID <> 0 Then
                                For ii = 1 To mBlockCnt
                                    If .GroupID = mBlock(ii).GroupID Then mBlock(ii).Selected = False
                                Next
                            End If
                        End If
                                     
                        '­Y¿ï¨ú¨ì¤å¦r¤è¶ô¤º®e¡A«h¨ú®ø¨ä¥Lª«¥óªº¿ï¨ú¥H¤Îµe¥X½Õ¾ã¾nÂI
                        If .ControlType = cjTextBox Then
                        
                            mMouseAct = EditTextBox
                            mTextIndex = i
                            
                            For ii = 1 To mBlockCnt
                                If ii <> i Then
                                    mBlock(ii).Selected = False
                                    spObj(ii).Visible = False
                                End If
                            Next
                                                                                    
                            If rtb.Visible = False Then Call GetTextBoxProperties(mBlock(i), True)
                                                                                                                  
                            rtb.Visible = True
                            rtb.SetFocus
                                                            
                            Exit For
                                                            
                        End If
                                                                        
                        Exit For
                        
                    Else
                        '­Y«ö¦íshift·|¿ï¨ú¥Ø«eÂI¿ïªºª«¥ó
                        If Shift = 1 Then
                            blnIsFirstSelected = False
                            .Selected = True
                            If .GroupID <> 0 Then
                                For ii = 1 To mBlockCnt
                                    If .GroupID = mBlock(ii).GroupID Then mBlock(ii).Selected = True
                                Next
                            End If
                            
                        Else
                            blnIsFirstSelected = True
                        End If
                        
                        Exit For
                        
                    End If
                    
                Else
                    '§PÂ_¿ï¨úªº¬O§_¬°¤å¦r¤è¶ôªº¥~³òÃä¬É
                    If .ControlType = cjTextBox Then
                        If (x >= Left And y >= Top - BOXSIZE And x <= Left + Width And y <= Top) Or _
                            (x >= Left - BOXSIZE And y >= Top And x <= Left And y <= Top + Height) Or _
                            (x >= Left And y >= Top + Height And x <= Left + Width And y <= Top + Height + BOXSIZE) Or _
                            (x >= Left + Width And y >= Top And x <= Left + Width + BOXSIZE And y <= Top + Height) And .Selected Then
                                                        
                            mBlockIndex = i
                                                                                    
                            mMouseAct = MoveObject '±N·Æ¹«°Ê§@§óÅÜ¬°²¾°Êª«¥ó
                                                                        
                            spObj(i).Move Left, Top, Width, Height
                            rtb.Visible = False
                            
                            blnIsFirstSelected = False
                            
                        End If
                    End If
                                                                                                
                End If
                
            End With
        Next
        
        '²Ä¤@¦¸¿ï¨ú¨ì clsBlock ª«¥ó
        If blnIsFirstSelected Then
                
            For i = mBlockCnt To 1 Step -1
                        
                With mBlock(i)
                    
                     If blnIsObjSelected Then '¤w¸g¿ï¨ú³Ì¤W¼hª«¥ó¡A¤£¦A¿ï¨ú¤U­±ªºª«¥ó
                        
                        '­Yª«¥ó¤£¬O¸s²Õ«h²M°£¿ï¨ú¡F­Yª«¥ó¬O¸s²Õ¥B¸s²Õ½s¸¹»P¥Ø«e¿ï¨ú¸s²Õ(iTmp)¤£¬Û¦P¡A¤]²MªÅ
                        If .GroupID <> iTmp Or .GroupID = 0 Then
                            .Selected = False
                            spObj(i).Visible = False
                        End If
                     Else
                     
                        'Âà´«¹ïÀ³ªº®y¼Ð¤ñ¨Ò
                        Left = .Left * gsngViewScale
                        Top = .Top * gsngViewScale
                        Width = .Width * gsngViewScale
                        Height = .Height * gsngViewScale
                    
                        If x >= Left And y >= Top And x <= Left + Width And y <= Top + Height Then
                                                                                                   
                            mMouseAct = MoveObject '±N·Æ¹«°Ê§@§óÅÜ¬°²¾°Êª«¥ó
                            
                            mBlockIndex = i
                            blnIsObjSelected = True
                                                                                    
                            .Selected = True '¼Ð¥Üª«¥ó¤w¸g³Q¿ï¨ú
                            
                            spObj(i).Move Left, Top, Width, Height
                            spObj(i).Visible = True
                            
                            If rtb.Visible And mTextIndex > 0 Then
                                mBlock(mTextIndex).Text = rtb.Text
                                rtb.Visible = False
                            End If
                            
                            '­Y¿ï¨ú¨ì¤å¦r¤è¶ô¤º®e¡A«h¨ú®ø¨ä¥Lª«¥óªº¿ï¨ú¥H¤Îµe¥X½Õ¾ã¾nÂI
                            If .ControlType = cjTextBox Then
                            
                                mMouseAct = EditTextBox
                                mTextIndex = i
                                
                                For ii = 1 To mBlockCnt
                                    If ii <> i Then
                                        mBlock(ii).Selected = False
                                        spObj(ii).Visible = False
                                    End If
                                Next
                                                                                                                                                                                            
                                If rtb.Visible = False Then Call GetTextBoxProperties(mBlock(i), True)
                                
                                rtb.Visible = True
                                rtb.SetFocus
                                Exit For
                                
                            End If
                            
                            If .GroupID <> 0 Then
                                For ii = 1 To mBlockCnt '¨ú±o¸s²Õªºª«¥ó
                                    iTmp = .GroupID
                                    If .GroupID = mBlock(ii).GroupID Then
                                        mBlock(ii).Selected = True
                                    End If
                                Next
                            End If
                            
                        Else
                            
                            If mTextIndex > 0 Then
                                mBlock(mTextIndex).Text = rtb.Text
                                rtb.Visible = False
                            End If
                            
                            '§PÂ_¿ï¨úªº¬O§_¬°¤å¦r¤è¶ôªº¥~³òÃä¬É
                            If .ControlType = cjTextBox Then
                                If (x >= Left And y >= Top - BOXSIZE And x <= Left + Width And y <= Top) Or _
                                    (x >= Left - BOXSIZE And y >= Top And x <= Left And y <= Top + Height) Or _
                                    (x >= Left And y >= Top + Height And x <= Left + Width And y <= Top + Height + BOXSIZE) Or _
                                    (x >= Left + Width And y >= Top And x <= Left + Width + BOXSIZE And y <= Top + Height) Then
                                    
                                    mMouseAct = MoveObject '±N·Æ¹«°Ê§@§óÅÜ¬°²¾°Êª«¥ó
                                    
                                    mBlockIndex = i
                                    mTextIndex = i
                                    
                                    blnIsObjSelected = True
                                                                                            
                                    .Selected = True '¼Ð¥Üª«¥ó¤w¸g³Q¿ï¨ú
                                    
                                    spObj(i).Move Left, Top, Width, Height
                                    spObj(i).Visible = True
                                Else
                                    .Selected = False
                                    spObj(i).Visible = False
                                    
                                End If
                                        
                            Else
                                .Selected = False
                                spObj(i).Visible = False
                            End If
                        End If
                    
                    End If
                    
                End With
            Next
            
        End If
                
    End If
    
    Call DrawObjects
    
    RaiseEvent MouseDown(Button, Shift, x / gsngViewScale, y / gsngViewScale)
      
End Sub

Private Sub picBG_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim i                       As Integer
    Dim iGroup                  As Integer
    Dim Left                    As Single
    Dim Top                     As Single
    Dim Height                  As Single
    Dim Width                   As Single
    
    hr.MouseMoved x * Screen.TwipsPerPixelX
    vr.MouseMoved y * Screen.TwipsPerPixelY
    
    If Button = vbLeftButton Then
        
        Select Case mMouseAct
            
            '²¾°Êªº¬Oª«¥ó¦ì¸m
            Case MoveObject
                For i = 1 To mBlockCnt
                    With mBlock(i)
                                                                    
                        If .Selected Then
                        
                            Left = .Left * gsngViewScale
                            Top = .Top * gsngViewScale
                            Width = .Width * gsngViewScale
                            Height = .Height * gsngViewScale
                            
                            spObj(i).Visible = True
                            spObj(i).Move Left + (x - mSelRect.Left), Top + (y - mSelRect.Top)
                            
                        End If
                    End With
                Next
                                    
            '²¾°Êªº¬O ¿ï¨ú½d³ò ¥H¤Î ¥[¤J clsBlock ª«¥óªº½d³ò
            Case MoveSelect, InsertBlock, InsertPicture, InsertTextbox
                
                With mSelRect
                
                    '§PÂ_·Æ¹««ö¤U¿ï¨úªº½d³ò
                    
                    '¥ª¤W¨ì¥k¤U
                    If .Left < x And .Top < y Then
                        .Width = x - .Left
                        .Height = y - .Top
                        spSelect.Move .Left, .Top, .Width, .Height
                    End If
                    
                    '¥ª¤U¨ì¥k¤W
                    If .Left < x And .Top > y Then
                        .Width = x - .Left
                        .Height = .Top - y
                        spSelect.Move .Left, y, .Width, .Height
                    End If
                                
                    '¥k¤U¨ì¥ª¤U
                    If .Left > x And .Top < y Then
                        .Width = .Left - x
                        .Height = y - .Top
                        spSelect.Move x, .Top, .Width, .Height
                    End If
                    
                    '¥k¤U¨ì¥ª¤W
                    If .Left > x And .Top > y Then
                        .Width = .Left - x
                        .Height = .Top - y
                        spSelect.Move x, y, .Width, .Height
                    End If
            
                End With
                        
            '²¾°Êª«¥ó¤W¥ªªº¤j¤p
            Case MoveTopLeftSize
                For i = 1 To mBlockCnt
                    If mBlock(i).Selected Then
                    
                        Left = mBlock(i).Left * gsngViewScale
                        Top = mBlock(i).Top * gsngViewScale
                        Width = mBlock(i).Width * gsngViewScale
                        Height = mBlock(i).Height * gsngViewScale
                    
                        With spObj(i)
                            
                            If mBlock(i).GroupID = 0 Then
                                                                            
                                If Left - (mBlock(mBlockIndex).Left * gsngViewScale - x) > 0 Then
                                    If Width + mBlock(mBlockIndex).Left * gsngViewScale - x > 5 Then
                                        .Width = Width + mBlock(mBlockIndex).Left * gsngViewScale - x
                                        .Left = Left - (mBlock(mBlockIndex).Left * gsngViewScale - x)
                                    End If
                                End If
                                
                                If Top - (mBlock(mBlockIndex).Top * gsngViewScale - y) > 0 Then
                                    If Height + (mBlock(mBlockIndex).Top * gsngViewScale - y) > 5 Then
                                        .Height = Height + (mBlock(mBlockIndex).Top * gsngViewScale - y)
                                        .Top = Top - (mBlock(mBlockIndex).Top * gsngViewScale - y)
                                    End If
                                End If
                                
                            Else
                                If Width - (x - mlngGroupX1) > 5 Then
                                    .Width = Width - (x - mlngGroupX1)
                                    .Left = Left + (x - mlngGroupX1)
                                End If
                                If Height - (y - mlngGroupY1) > 5 Then
                                    .Height = Height - (y - mlngGroupY1)
                                    .Top = Top + (y - mlngGroupY1)
                                End If
                            End If
                            
                            .Move .Left, .Top, .Width, .Height
                            .Visible = True
                            
                        End With
                    End If
                Next
                
            '²¾°Êª«¥ó¤W¤¤ªº¤j¤p
            Case MoveTopMiddleSize
                For i = 1 To mBlockCnt
                    If mBlock(i).Selected Then
                    
                        Top = mBlock(i).Top * gsngViewScale
                        Height = mBlock(i).Height * gsngViewScale
                    
                        With spObj(i)
                            
                            If mBlock(i).GroupID = 0 Then
                            
                                If Top - (mBlock(mBlockIndex).Top * gsngViewScale - y) > 0 Then
                                    If Height + (mBlock(mBlockIndex).Top * gsngViewScale - y) > 5 Then
                                        .Height = Height + (mBlock(mBlockIndex).Top * gsngViewScale - y)
                                        .Top = Top - (mBlock(mBlockIndex).Top * gsngViewScale - y)
                                    End If
                                End If
                            Else
                                If Height - (y - mlngGroupY1) > 5 Then
                                    .Height = Height - (y - mlngGroupY1)
                                    .Top = Top + (y - mlngGroupY1)
                                End If
                            End If
                            
                            .Move .Left, .Top, .Width, .Height
                            .Visible = True
                            
                        End With
                    End If
                Next
            
            Case MoveTopRightSize
                                        
                For i = 1 To mBlockCnt
                
                    If mBlock(i).Selected Then
                                              
                        With mBlock(i)
                            Top = .Top * gsngViewScale
                            Width = .Width * gsngViewScale
                            Height = .Height * gsngViewScale
                        End With
                                                                    
                        With spObj(i)
                                                                    
                            If mBlock(i).GroupID = 0 Then
                                                
                                If Width + (x - mBlock(mBlockIndex).Left * gsngViewScale - mBlock(mBlockIndex).Width * gsngViewScale) > 5 Then
                                    .Width = Width + (x - mBlock(mBlockIndex).Left * gsngViewScale - mBlock(mBlockIndex).Width * gsngViewScale)
                                End If
                                
                                If Height + (mBlock(mBlockIndex).Top * gsngViewScale - y) > 5 Then
                                    .Height = Height + (mBlock(mBlockIndex).Top * gsngViewScale - y)
                                    .Top = Top - (mBlock(mBlockIndex).Top * gsngViewScale - y)
                                End If
                            Else
                                If Width + x - mlngGroupX2 > 5 Then .Width = Width + x - mlngGroupX2
                                If Height - (y - mlngGroupY1) > 5 Then
                                    .Height = Height - (y - mlngGroupY1)
                                    .Top = Top + (y - mlngGroupY1)
                                End If
                            End If
                        
                            .Move .Left, .Top, .Width, .Height
                            .Visible = True
                        End With
                    End If
                Next
                  
            Case MoveMiddleLeftSize
                For i = 1 To mBlockCnt
                    If mBlock(i).Selected Then
                    
                        Left = mBlock(i).Left * gsngViewScale
                        Width = mBlock(i).Width * gsngViewScale
                        
                        With spObj(i)
                                 
                            If mBlock(i).GroupID = 0 Then
                                 
                                If Width + mBlock(mBlockIndex).Left * gsngViewScale - x > 5 Then
                                    .Width = Width + mBlock(mBlockIndex).Left * gsngViewScale - x
                                    .Left = Left - (mBlock(mBlockIndex).Left * gsngViewScale - x)
                                End If
                            Else
                                If Width - (x - mlngGroupX1) > 5 Then
                                    .Width = Width - (x - mlngGroupX1)
                                    .Left = Left + (x - mlngGroupX1)
                                End If
                            End If
                            .Move .Left, .Top, .Width, .Height
                            .Visible = True
                            
                        End With
                    End If
                Next
                        
            Case MoveMiddleRightSize
                For i = 1 To mBlockCnt
                    If mBlock(i).Selected Then
                            
                        Width = mBlock(i).Width * gsngViewScale
                        
                        With spObj(i)
                            If mBlock(i).GroupID = 0 Then
                                If Width + (x - mBlock(mBlockIndex).Left * gsngViewScale - mBlock(mBlockIndex).Width) * gsngViewScale > 5 Then
                                    .Width = Width + (x - mBlock(mBlockIndex).Left * gsngViewScale - mBlock(mBlockIndex).Width * gsngViewScale)
                                End If
                            Else
                                If Width + x - mlngGroupX2 > 5 Then .Width = Width + x - mlngGroupX2
                            End If
                            
                            .Move .Left, .Top, .Width, .Height
                            .Visible = True
                            
                        End With
                    End If
                Next
                                        
            Case MoveBottomLeftSize
                For i = 1 To mBlockCnt
                    If mBlock(i).Selected Then
                    
                        Left = mBlock(i).Left * gsngViewScale
                        Width = mBlock(i).Width * gsngViewScale
                        Height = mBlock(i).Height * gsngViewScale
                        
                        With spObj(i)
                                                                 
                            If mBlock(i).GroupID = 0 Then
                                If Width + mBlock(mBlockIndex).Left * gsngViewScale - x > 5 Then
                                    .Width = Width + mBlock(mBlockIndex).Left * gsngViewScale - x
                                    .Left = Left - (mBlock(mBlockIndex).Left * gsngViewScale - x)
                                End If
                                If Height + (y - mBlock(mBlockIndex).Height * gsngViewScale - mBlock(mBlockIndex).Top * gsngViewScale) > 5 Then
                                    .Height = Height + (y - mBlock(mBlockIndex).Height * gsngViewScale - mBlock(mBlockIndex).Top * gsngViewScale)
                                End If
                            Else
                                If Width - (x - mlngGroupX1) > 5 Then
                                    .Width = Width - (x - mlngGroupX1)
                                    .Left = Left + (x - mlngGroupX1)
                                End If
                                If Height + y - mlngGroupY2 > 5 Then
                                    .Height = Height + y - mlngGroupY2
                                End If
                            End If
                            
                            .Move .Left, .Top, .Width, .Height
                            .Visible = True
                            
                        End With
                    End If
                Next
                                    
            Case MoveBottomMiddleSize
            
                For i = 1 To mBlockCnt
                
                    If mBlock(i).Selected Then
                    
                        Height = mBlock(i).Height * gsngViewScale
                        
                        With spObj(i)
                            If mBlock(i).GroupID = 0 Then
                                If Height + (y - mBlock(mBlockIndex).Height * gsngViewScale - mBlock(mBlockIndex).Top * gsngViewScale) > 5 Then
                                    .Height = Height + (y - mBlock(mBlockIndex).Height * gsngViewScale - mBlock(mBlockIndex).Top * gsngViewScale)
                                End If
                            Else
                                If Height + y - mlngGroupY2 > 5 Then .Height = Height + y - mlngGroupY2
                            End If
                            .Move .Left, .Top, .Width, .Height
                            .Visible = True
                            
                        End With
                    End If
                Next
                
            Case MoveBottomRightSize
                For i = 1 To mBlockCnt
                    If mBlock(i).Selected Then
                    
                        Width = mBlock(i).Width * gsngViewScale
                        Height = mBlock(i).Height * gsngViewScale
                        
                        With spObj(i)
                            
                            If mBlock(i).GroupID = 0 Then
                                                            
                                If Width + (x - mBlock(mBlockIndex).Left * gsngViewScale - mBlock(mBlockIndex).Width * gsngViewScale) > 5 Then
                                    .Width = Width + (x - mBlock(mBlockIndex).Left * gsngViewScale - mBlock(mBlockIndex).Width * gsngViewScale)
                                End If
                                If Height + (y - mBlock(mBlockIndex).Height * gsngViewScale - mBlock(mBlockIndex).Top * gsngViewScale) > 5 Then
                                    .Height = Height + (y - mBlock(mBlockIndex).Height * gsngViewScale - mBlock(mBlockIndex).Top * gsngViewScale)
                                End If
                            Else
                                If Width + x - mlngGroupX2 > 5 Then .Width = Width + x - mlngGroupX2
                                If Height + y - mlngGroupY2 > 5 Then .Height = Height + y - mlngGroupY2
                            End If
                            
                            .Move .Left, .Top, .Width, .Height
                            .Visible = True
                            
                        End With
                    End If
                Next
                                                  
        End Select
        
    End If
    
    RaiseEvent MouseMove(Button, Shift, x / gsngViewScale, y / gsngViewScale)
    
End Sub

Private Sub picBG_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim i                       As Integer
    Dim ii                      As Integer
    Dim iGroup                  As Integer
    
    Dim Left                    As Single
    Dim Top                     As Single
    Dim Height                  As Single
    Dim Width                   As Single
    
    '­pºâ¥X¿ï¨ú½d³òªº¥¿½T¤j¤p
    With mSelRect
        '¥ª¤U¨ì¥k¤W
        If .Left < x And .Top > y Then
            .Top = y
        End If
        
        '¥k¤W¨ì¥ª¤U
        If .Left > x And .Top < y Then
            .Left = x
        End If
        
        '¥ª¤U¨ì¥k¤W
        If .Left > x And .Top > y Then
            .Left = x
            .Top = y
        End If
        
    End With
    
    Select Case mMouseAct
    
        '©ñ¶}ªº¬Oª«¥ó½Õ¾ã¹L«áªº¤j¤p
        Case MoveTopLeftSize, MoveTopMiddleSize, MoveTopRightSize, _
                MoveMiddleLeftSize, MoveMiddleRightSize, _
                MoveBottomLeftSize, MoveBottomMiddleSize, MoveBottomRightSize
            
            mMouseAct = MoveSelect '±N·Æ¹«°Ê§@­«¸m¬°¿ï¨ú
            
            For i = 1 To mBlockCnt
                
                spObj(i).Visible = False
                
                With mBlock(i)
                    
                    If .Selected Then
                        .Left = spObj(i).Left / gsngViewScale
                        .Top = spObj(i).Top / gsngViewScale
                        .Width = spObj(i).Width / gsngViewScale
                        .Height = spObj(i).Height / gsngViewScale
                    End If
                    
                End With
            
            Next

            Call DrawObjects
            
        '©ñ¶}ªº¬Oª«¥ó
        Case MoveObject
                        
            mMouseAct = MoveSelect '±N·Æ¹«°Ê§@­«¸m¬°¿ï¨ú
                        
            For i = 1 To mBlockCnt
            
                With mBlock(i)
                                        
                    If .Selected Then
                                        
                        .Left = spObj(i).Left / gsngViewScale
                        .Top = spObj(i).Top / gsngViewScale
                        .Width = spObj(i).Width / gsngViewScale
                        .Height = spObj(i).Height / gsngViewScale
                        
                        spObj(i).Visible = False
                        
                    End If
                    
                End With
                
            Next
            
            Call DrawObjects
                        
        '©ñ¶}ªº¬O½d³ò¿ï¨ú
        Case MoveSelect
        
            spSelect.Visible = False
            picBG.Cls
            
            '§PÂ_¬O§_³z¹L °é¿ï½d³ò ¿ï¨ú¨ì clsBlock ª«¥ó
            For i = 1 To mBlockCnt
            
                With mBlock(i)
                
                    spObj(i).Visible = False
                    
                    Left = .Left * gsngViewScale
                    Top = .Top * gsngViewScale
                    Width = .Width * gsngViewScale
                    Height = .Height * gsngViewScale
                    
                    If Left >= mSelRect.Left And Top >= mSelRect.Top _
                        And Left + Width <= mSelRect.Left + mSelRect.Width And Top + Height <= mSelRect.Top + mSelRect.Height Then
                    
                        mMouseAct = MoveObject '±N·Æ¹«°Ê§@§óÅÜ¬°²¾°Êª«¥ó
                        
                        mBlockIndex = i
                        
                        .Selected = True '¼Ð¥Üª«¥ó¤w¸g³Q¿ï¨ú
                        spObj(i).Move Left, Top, Width, Height
                        
                        If .ControlType = cjTextBox Then Call DrawTextBox(mBlock(i), Left, Top, Width, Height)
                                                                                    
                        If iGroup <> .GroupID Then
                            iGroup = .GroupID
                            
                            '©¹©³¼h§ä¬O§_¦³¬Û¦Pªº¸s²Õ
                            For ii = i To 1 Step -1
                                If mBlock(ii).GroupID = iGroup Then
                                    mBlock(ii).Selected = True
                                Else
                                    Exit For
                                End If
                            Next
                            
                        End If
                                                                       
                    Else
                    
                        If Shift <> 1 Then .Selected = False '­Y¨S«ö¦íshift«h·|¨ú®ø¤£¦A¿ï¨ú½d³ò¤ºªºª«¥ó
                        
                        '­Yª«¥ó¬°¸s²Õ¡A«hÀË¬d¬O§_ÂI¿ï¤F¸s²Õ¤¤ªº¥ô¤@ª«¥ó
                        If .GroupID <> 0 Then
                            '©¹©³¼h§ä¬O§_¦³¬Û¦Pªº¸s²Õ
                            For ii = i To 1 Step -1
                                If mBlock(ii).GroupID = .GroupID Then
                                
                                    If mBlock(ii).Selected Then .Selected = True
                                
                                Else
                                    Exit For
                                    
                                End If
                            Next
                        End If
                        
                    End If
                    
                End With
                
            Next
            
            Call DrawAdjustFrame
            
        '©ñ¶}ªº¬O´¡¤J clsBlock
        Case InsertBlock, InsertPicture, InsertTextbox
        
            spSelect.Visible = False
            
            '¤j©ó¤@©wªº°Ï°ì¤~·|·s¼W clsBlock ª«¥ó
            If mSelRect.Width > 5 And mSelRect.Height > 5 Or mMouseAct = InsertTextbox Then
                                        
                Select Case mMouseAct
                    Case InsertBlock
                        Call CreateObject(mSelRect.Left, mSelRect.Top, mSelRect.Width, mSelRect.Height, cjBlock)
                        
                    Case InsertPicture
                        Call CreateObject(mSelRect.Left, mSelRect.Top, mSelRect.Width, mSelRect.Height, cjPicture)
                        
                    Case InsertTextbox
                        Call CreateObject(mSelRect.Left, mSelRect.Top, mSelRect.Width, mSelRect.Height, cjTextBox)
                                            
                End Select
                
                mBlock(mBlockCnt).Selected = True
                
                mMouseAct = MoveSelect '±N·Æ¹«°Ê§@­«¸m¬°¿ï¨ú
                
                Call DrawObjects
                
            End If
        
        Case EditTextBox
            mMouseAct = MoveSelect '±N·Æ¹«°Ê§@­«¸m¬°¿ï¨ú
                                            
    End Select
    
    mSelRect.Width = 0
    mSelRect.Height = 0
    
    RaiseEvent MouseUp(Button, Shift, x / gsngViewScale, y / gsngViewScale)
                                
End Sub

Private Function GetObjectDefaultName(strPrefix As String) As String
    
    Dim i       As Integer
    Dim iLen    As Integer
    Dim iCnt    As Integer
    Dim iNum    As Integer
    
    iNum = 1
    iCnt = 1
    iLen = Len(strPrefix)
    
    For i = 1 To mBlockCnt
        If Left(mBlock(i).Name, iLen) = strPrefix Then
            iCnt = iCnt + 1
            If IsNumeric(Mid(mBlock(i).Name, iLen + 1)) Then
                If iCnt <> Val(Mid(mBlock(i).Name, iLen + 1)) Then iNum = iCnt
            End If
        End If
    Next
    'iNum = iNum + 1
    GetObjectDefaultName = strPrefix & Trim(iNum)
    
End Function

Private Sub DrawAdjustFrame()
    
    Dim i           As Integer
    Dim ii          As Integer
    Dim Left        As Single
    Dim Top         As Single
    Dim Height      As Single
    Dim Width       As Single
    
    Dim iArray()    As Integer
    Dim blnSkip     As Boolean
    
    Dim hBrush      As Long, hPen As Long
    Dim hOldBrush As Long, hOldPen As Long
    
    ' «Ø¥ß Brush(¹Ï¨ê) ª«¥ó
    hBrush = CreateSolidBrush(vbBlue)
    ' «Ø¥ß Pen(µ§) ª«¥ó
    hPen = CreatePen(vbInsideSolid, 1, vbWhite)
    ' ¿ï¨úª«¥ó:Brush ¤Î Pen
    hOldBrush = SelectObject(picBG.hDC, hBrush)
    hOldPen = SelectObject(picBG.hDC, hPen)
    
    ReDim Preserve iArray(0)
    
    For i = 1 To mBlockCnt
        
        With mBlock(i)
        
            If .Selected Then
            
                If mBlock(i).GroupID = 0 Then
                
                    '¿ï¨ú«D¸s²Õª«¥ó
                    Left = .Left * gsngViewScale
                    Top = .Top * gsngViewScale
                    Height = .Height * gsngViewScale
                    Width = .Width * gsngViewScale
                                
                Else
                
                    'Á×§K­«½Æ³B²z¸s²Õªºµ{¦¡°Ï¬q blnSkip
                    blnSkip = False
                    Left = 999: Top = 999
                    Width = 0: Height = 0
                    
                    For ii = 1 To UBound(iArray)
                        If iArray(ii) = .GroupID Then
                            blnSkip = True
                            Exit For
                        End If
                    Next
                    
                    If blnSkip = False Then
                                                            
                        ReDim Preserve iArray(UBound(iArray) + 1)
                        iArray(UBound(iArray)) = .GroupID
                                                            
                        '¿ï¨ú¸s²Õª«¥ó
                        For ii = 1 To mBlockCnt
                            If .GroupID = mBlock(ii).GroupID Then
                                If mBlock(ii).Left < Left Then Left = mBlock(ii).Left
                                If mBlock(ii).Top < Top Then Top = mBlock(ii).Top
                                If mBlock(ii).Width + mBlock(ii).Left > Width Then Width = mBlock(ii).Width + mBlock(ii).Left
                                If mBlock(ii).Height + mBlock(ii).Top > Height Then Height = mBlock(ii).Height + mBlock(ii).Top
                            End If
                        Next
                        
                        Left = Left * gsngViewScale
                        Top = Top * gsngViewScale
                        Width = Width * gsngViewScale - Left
                        Height = Height * gsngViewScale - Top
                        
                    End If
                    
                End If
                                
                '¤W¥ª
                Rectangle picBG.hDC, Left - BOXSIZE - 1, Top - BOXSIZE - 1, Left - 1, Top - 1
                '¤W¤¤
                Rectangle picBG.hDC, (Left * 2 + Width) / 2 - BOXSIZE / 2, Top - BOXSIZE - 1, (Left * 2 + Width) / 2 + BOXSIZE / 2, Top - 1
                '¤W¥k
                Rectangle picBG.hDC, Left + Width, Top - BOXSIZE - 1, Left + Width + BOXSIZE, Top - 1
                '¤¤¥ª
                Rectangle picBG.hDC, Left - BOXSIZE - 1, (Top * 2 + Height) / 2 - BOXSIZE / 2, Left - 1, (Top * 2 + Height) / 2 + BOXSIZE / 2
                '¤¤¥k
                Rectangle picBG.hDC, Left + Width, (Top * 2 + Height) / 2 - BOXSIZE / 2, Left + Width + BOXSIZE, (Top * 2 + Height) / 2 + BOXSIZE / 2
                '¤U¥ª
                Rectangle picBG.hDC, Left - BOXSIZE - 1, Top + Height, Left - 1, Top + Height + BOXSIZE
                '¤U¤¤
                Rectangle picBG.hDC, (Left * 2 + Width) / 2 - BOXSIZE / 2, Top + Height, (Left * 2 + Width) / 2 + BOXSIZE / 2, Top + Height + BOXSIZE
                '¤U¥k
                Rectangle picBG.hDC, Left + Width, Top + Height, Left + Width + BOXSIZE, Top + Height + BOXSIZE

            End If
            
        End With
    Next

    ' ÁÙ­ì­ì¥ýªº Brush ¤Î Pen
    SelectObject picBG.hDC, hOldBrush
    SelectObject picBG.hDC, hOldPen
    
    ' §R°£ Pen ¤Î Brush ª«¥ó
    DeleteObject hBrush
    DeleteObject hPen
        
End Sub

Private Sub DrawObjects()
    
    Dim i                   As Integer
    Dim ii                  As Integer
    Dim Left                As Single
    Dim Top                 As Single
    Dim Height              As Single
    Dim Width               As Single
    
    If mblnDrawObjects = False Then Exit Sub
    
    picBG.AutoRedraw = True
    picBG.Cls
    
    If cjEditor.GridShow Then
        For i = 0 To picBG.ScaleWidth Step cjEditor.GridScale
            For ii = 0 To picBG.ScaleHeight Step cjEditor.GridScale
                SetPixel picBG.hDC, i, ii, vbBlack
            Next
        Next
    End If
    
    For i = 1 To mBlockCnt
    
        With mBlock(i)
                                    
            Left = .Left * gsngViewScale
            Top = .Top * gsngViewScale
            Height = .Height * gsngViewScale
            Width = .Width * gsngViewScale
            
            spObj(i).Move Left, Top, Width, Height
                        
            Select Case .ControlType
                Case typControlType.cjBlock
                    Call DrawShape(mBlock(i), Left, Top, Width, Height)
                                        
                Case typControlType.cjPicture
                    If Len(.PictureFileName) > 0 Then
                        Call DrawPicture(mBlockPic(i), Left, Top, Width, Height, .PicTransparent, False)
                    Else
                        Call DrawPicture(mBlockPic(i), Left, Top, Width, Height, .PicTransparent, True)
                    End If
                                        
                Case typControlType.cjTextBox
                    Call DrawTextBox(mBlock(i), Left, Top, Width, Height)
            End Select
            
        End With
    
    Next
        
    picBG.Refresh
    picBG.AutoRedraw = False
    
    Call DrawAdjustFrame
    
End Sub

Private Sub DrawTextBox(clsObject As clsBlock, Left As Single, Top As Single, Width As Single, Height As Single)

    Dim Font        As LOGFONT
    
    Dim hOldFont    As Long
    Dim hFont       As Long
    Dim sRECT       As RECT

    Dim hBrush As Long, hPen As Long
    Dim hOldBrush As Long, hOldPen As Long
    Dim hOldTextColor   As Long
    
    With clsObject
            
        If clsObject.Selected And clsObject.GroupID = 0 Then
                              
            If mMouseAct = EditTextBox Then
                hBrush = CreateHatchBrush(HS_BDIAGONAL, vbBlack) ' «Ø¥ß Brush(¹Ï¨ê) ª«¥ó
            Else
                hBrush = CreateHatchBrush(HS_FDIAGONAL, vbBlack) ' «Ø¥ß Brush(¹Ï¨ê) ª«¥ó
            End If
            
            hPen = CreatePen(vbInvisible, 0, vbBlack) ' «Ø¥ß Pen(µ§) ª«¥ó
            '¿ï¨úª«¥ó:Brush ¤Î Pen
            hOldBrush = SelectObject(picBG.hDC, hBrush)
            hOldPen = SelectObject(picBG.hDC, hPen)

            Rectangle picBG.hDC, Left - BOXSIZE, Top - BOXSIZE, Left + Width + BOXSIZE, Top + BOXSIZE / 2 - 1
            Rectangle picBG.hDC, Left - BOXSIZE, Top + Height, Left + Width + BOXSIZE, Height + Top + BOXSIZE
            Rectangle picBG.hDC, Left - BOXSIZE, Top, Left + BOXSIZE / 2 - 1, Height + Top
            Rectangle picBG.hDC, Left + Width, Top, Left + Width + BOXSIZE, Height + Top

            ' ÁÙ­ì­ì¥ýªº Brush ¤Î Pen
            SelectObject picBG.hDC, hOldBrush
            SelectObject picBG.hDC, hOldPen
            ' §R°£ Pen ¤Î Brush ª«¥ó
            DeleteObject hBrush
            DeleteObject hPen
            
        End If
        
        If Len(.Text) = 0 Then
            'µe¥X¤å¦r¤è¶ôªºÃä½t
            hPen = CreatePen(vbDot, 0, vbBlack) '«Ø¥ß Pen(µ§) ª«¥ó
            hOldPen = SelectObject(picBG.hDC, hPen) '¿ï¨úª«¥óPen
            Rectangle picBG.hDC, Left - 1, Top - 1, Width + Left + 1, Height + Top + 1
            SelectObject picBG.hDC, hOldPen ' ÁÙ­ì­ì¥ýªºPen
            DeleteObject hPen ' §R°£ Pen ª«¥ó
            
        Else
            'Font Size of VB, Convert to API rule is (vb.Size * -20) / Screen.TwipsPerPixelY
            Font.lfHeight = (.FontSize * -20) / Screen.TwipsPerPixelY * gsngViewScale
            Font.lfWidth = (0 * -20) / Screen.TwipsPerPixelY
            Font.lfCharSet = DEFAULT_CHARSET
            Font.lfEscapement = 0 * 10
            If .FontBold Then Font.lfWeight = 700 Else Font.lfWeight = 400
            If .FontItlic Then Font.lfItalic = 1 Else Font.lfItalic = 0
            If .FontUnderline Then Font.lfUnderline = 1 Else Font.lfUnderline = 0
                                                                                                    
            ' LOGFONT ¸ê®Æµ²ºcªº³]©w
            RtlMoveMemory Font.lfFaceName(0), _
                           ByVal .FontName, _
                           LenB(StrConv(.FontName, vbFromUnicode)) + 1

'            Font.lfStrikeOut = .Font.lfStrikeOut
'
            ' «Ø¥ß¦r«¬ª«¥ó
            hFont = CreateFontIndirect(Font)
            ' ¿ï¨ú¦r«¬ª«¥ó
            hOldFont = SelectObject(picBG.hDC, hFont)

            sRECT.Left = Left
            sRECT.Top = Top
            sRECT.Width = Width + Left
            sRECT.Height = Top + Height

            hOldTextColor = GetTextColor(picBG.hDC)
            Call SetTextColor(picBG.hDC, .FontForeColor)
            
            Select Case .TextAligment
                Case typTextAligment.txt_Left
'                    DrawText picBG.hDC, .Text, LenB(StrConv(.Text, vbFromUnicode)), sRECT, DT_CALCRECT Or DT_WORDBREAK
                    DrawText picBG.hDC, .Text, LenB(StrConv(.Text, vbFromUnicode)), sRECT, DT_LEFT Or DT_WORDBREAK
                    
                Case typTextAligment.txt_Right
'                    DrawText picBG.hDC, .Text, LenB(StrConv(.Text, vbFromUnicode)), sRECT, DT_CALCRECT Or DT_WORDBREAK
                    DrawText picBG.hDC, .Text, LenB(StrConv(.Text, vbFromUnicode)), sRECT, DT_RIGHT Or DT_WORDBREAK
                    
                Case typTextAligment.txt_Center
'                    DrawText picBG.hDC, .Text, LenB(StrConv(.Text, vbFromUnicode)), sRECT, DT_CALCRECT Or DT_WORDBREAK
                    DrawText picBG.hDC, .Text, LenB(StrConv(.Text, vbFromUnicode)), sRECT, DT_CENTER Or DT_WORDBREAK
                    
            End Select
            
            rtb.Height = sRECT.Height - sRECT.Top
            
            Call SetTextColor(picBG.hDC, hOldTextColor)
                                
            ' ÁÙ­ì¦r«¬
            SelectObject picBG.hDC, hOldFont
            ' §R°£¦r«¬
            DeleteObject hFont
            
        End If
    
    End With
End Sub

Private Sub DrawShape(clsObject As clsBlock, Left As Single, Top As Single, Width As Single, Height As Single)

    Dim hBrush As Long, hPen As Long
    Dim hOldBrush As Long, hOldPen As Long
    
    With clsObject
    
        If .BorderWidth > 0 Then
        
            ' «Ø¥ß Brush(¹Ï¨ê) ª«¥ó
            hBrush = CreateSolidBrush(.BackColor)
            ' «Ø¥ß Pen(µ§) ª«¥ó
            hPen = CreatePen(vbInsideSolid, .BorderWidth, .BorderColor)
            ' ¿ï¨úª«¥ó:Brush ¤Î Pen
            hOldBrush = SelectObject(picBG.hDC, hBrush)
            hOldPen = SelectObject(picBG.hDC, hPen)
        
            Select Case .Shape
                Case blkRectangle   '¯x§Î
                    Rectangle picBG.hDC, Left, Top, Left + Width, Top + Height
                    
                Case blkSquare      '¥¿¤è§Î
                
                Case blkEllipse     '¾ò¶ê
                    Ellipse picBG.hDC, Left, Top, Left + Width, Top + Height
                    
                Case blkCircle      '¶ê§Î
                
                Case blkRoundRect   '¶ê¨¤¯x§Î
                    RoundRect picBG.hDC, Left, Top, Left + Width, Top + Height, .RoundAngel, .RoundAngel
                    
                Case blkRoundSquare '¶ê¨¤¥¿¤è§Î
                    
            End Select
            
            ' ÁÙ­ì­ì¥ýªº Brush ¤Î Pen
            SelectObject picBG.hDC, hOldBrush
            SelectObject picBG.hDC, hOldPen
            
            ' §R°£ Pen ¤Î Brush ª«¥ó
            DeleteObject hBrush
            DeleteObject hPen
            
        End If
        
    End With
    
End Sub

Private Sub DrawPicture(picObject As PictureBox, Left As Single, Top As Single, Width As Single, Height As Single, _
                        blnTransparent As Boolean, blnDrawRECT As Boolean)
                        
    Dim hPen                As Long
    Dim hOldPen             As Long
    Dim lngPointColor       As Long
    
    With picObject
                
        If blnDrawRECT Then
            
            '­Y¨S¦³¹Ï§Î«hµe¥Xª«¥óªºÃä½t
            hPen = CreatePen(vbDot, 0, vbBlack) '«Ø¥ß Pen(µ§) ª«¥ó
            hOldPen = SelectObject(picBG.hDC, hPen) '¿ï¨úª«¥óPen
            Rectangle picBG.hDC, Left, Top, Width + Left, Height + Top
            SelectObject picBG.hDC, hOldPen ' ÁÙ­ì­ì¥ýªºPen
            DeleteObject hPen ' §R°£ Pen ª«¥ó
                
        End If
                
        If blnTransparent Then
            lngPointColor = .Point(.ScaleLeft, .ScaleTop)
            TransparentBlt picBG.hDC, Left, Top, Width, Height, _
                .hDC, .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight, lngPointColor
        Else
            TransparentBlt picBG.hDC, Left, Top, Width, Height, _
                .hDC, .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight, RGB(5, 10, 6)
        End If
                
    End With
End Sub

Private Sub picBG_Paint()
    Call DrawAdjustFrame
End Sub

Private Sub rtb_KeyUp(KeyCode As Integer, Shift As Integer)
        
    If rtb.Width < picBG.TextWidth(rtb.Text) Then
        rtb.Width = picBG.TextWidth(rtb.Text)
        mBlock(mBlockIndex).Width = rtb.Width
    End If
                
    If rtb.Height < picBG.TextHeight(rtb.Text) Then
        rtb.Height = picBG.TextHeight(rtb.Text)
        mBlock(mBlockIndex).Height = picBG.TextHeight(rtb.Text)
    End If
                   
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If rtb.Visible = False Then RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If rtb.Visible = False Then RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If rtb.Visible = False Then RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    
    With UserControl
    
        hr.Move vr.Width, .ScaleTop, .ScaleWidth, 15
        vr.Move .ScaleLeft, hr.Height, 15, .ScaleHeight
        
        picBG.Move vr.Left + vr.Width, hr.Top + hr.Height, cjEditor.Size.Width * gsngViewScale, cjEditor.Size.Height * gsngViewScale
        spBG.Move picBG.Left + 5, picBG.Top + 5, picBG.Width, picBG.Height
                
        If picBG.Width > .ScaleWidth - vr.Width Then
            If vsc.Visible Then
                hsc.Move .ScaleLeft, .ScaleHeight - hsc.Height, .ScaleWidth - vsc.Width
            Else
                hsc.Move .ScaleLeft, .ScaleHeight - hsc.Height, .ScaleWidth
            End If
            hsc.Min = -vr.Width
            hsc.Max = picBG.Width - .ScaleWidth + vsc.Width + 5
            hsc.Value = hsc.Min
            hsc.Visible = True
        Else
            hsc.Visible = False
        End If
        
        If picBG.Height > .ScaleHeight - hr.Height Then
            If hsc.Visible Then
                vsc.Move .ScaleWidth - vsc.Width, .ScaleTop, vsc.Width, .ScaleHeight - hsc.Height
            Else
                vsc.Move .ScaleWidth - vsc.Width, .ScaleTop, vsc.Width, .ScaleHeight
            End If
            vsc.Min = -hr.Height
            vsc.Max = picBG.Height - .ScaleHeight + hsc.Height + 5
            vsc.Value = vsc.Min
            vsc.Visible = True
        Else
            vsc.Visible = False
        End If
                
    End With
    
End Sub

Private Sub UserControl_Initialize()

    '·Æ¹«ªº¾Þ§@¬°¿ï¨ú
    mMouseAct = MoveSelect

    'Åã¥Ü¤ñ¨Ò¬° 100%
    gsngViewScale = 1
    
    '°ÊºA¸ü¤J PictureBox ¨Ó·í cjEditor ªº¹Ï§Î½w½Ä°Ï
    Set cjEditor.PictureBox = Controls.Add("VB.PictureBox", "cjPictureBox")
    cjEditor.PictureBox.AutoRedraw = True
    cjEditor.PictureBox.AutoSize = True
    
    cjEditor.GridShow = True
    cjEditor.GridScale = 8
        
    mblnDrawObjects = True
    
    '³]©w¦U­Ó±±¨îÂIªº·Æ¹«´å¼Ð
    Cursor(TOPLEFT) = vbSizeNWSE
    Cursor(TOPMIDDLE) = vbSizeNS
    Cursor(TOPRIGHT) = vbSizeNESW
    Cursor(MIDDLERIGHT) = vbSizeWE
    Cursor(BOTTOMRIGHT) = vbSizeNWSE
    Cursor(BOTTOMMIDDLE) = vbSizeNS
    Cursor(BOTTOMLEFT) = vbSizeNESW
    Cursor(MIDDLELEFT) = vbSizeWE
    Cursor(TRANSLATE) = vbSizePointer
    Cursor(OUTSIDE) = vbArrow
    
End Sub

Private Sub UserControl_Terminate()

    Dim i       As Integer
    
    If mClipBoardCnt > 0 Then
        For i = 0 To mClipBoardCnt - 1
            Set mClipboard(i) = Nothing
            Controls.Remove ("picClip" & Trim(i))
        Next
    End If
    
    If mBlockCnt > 0 Then
        For i = 0 To mBlockCnt
            Set mBlock(i) = Nothing
            If i > 0 Then Controls.Remove ("picBlock" & Trim(i)) 'picBlock based on 1
        Next
    End If
            
    Controls.Remove ("cjPictureBox")
    
End Sub

Private Sub hsc_Change()
    picBG.Left = -hsc.Value
End Sub

Private Sub hsc_Scroll()
    picBG.Left = -hsc.Value
End Sub

Private Sub vsc_Change()
    picBG.Top = -vsc.Value
End Sub

Private Sub vsc_Scroll()
    picBG.Top = -vsc.Value
End Sub
