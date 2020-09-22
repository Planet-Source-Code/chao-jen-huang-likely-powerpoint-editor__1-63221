VERSION 5.00
Begin VB.UserControl cjEdt 
   ClientHeight    =   5460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   KeyPreview      =   -1  'True
   ScaleHeight     =   364
   ScaleMode       =   3  '像素
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
      Appearance      =   0  '平面
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   240
      ScaleHeight     =   247
      ScaleMode       =   3  '像素
      ScaleWidth      =   294
      TabIndex        =   0
      Top             =   240
      Width           =   4440
      Begin VB.TextBox rtb 
         Appearance      =   0  '平面
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
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
         BorderStyle     =   3  '點線
         DrawMode        =   6  'Mask Pen Not
         Height          =   495
         Index           =   0
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape spSelect 
         BorderStyle     =   3  '點線
         DrawMode        =   6  'Mask Pen Not
         Height          =   495
         Left            =   0
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Shape spBG 
      FillStyle       =   0  '實心
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
            
'CreateHatchBrush API 所使用的筆刷花色
Private Const HS_HORIZONTAL = 0
Private Const HS_VERTICAL = 1
Private Const HS_FDIAGONAL = 2
Private Const HS_BDIAGONAL = 3
Private Const HS_CROSS = 4
Private Const HS_DIAGCROSS = 5
           
Private Const BOXSIZE = 7 '圖形的寬及控制點離圖形的距離

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

Private Cursor(9)           As Integer ' 滑鼠游標的陣列

'----------------------------------------- 屬於 clsBlock 物件的屬性宣告 -----------------------------------
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
        
    '群組步驟一：取出目前可以用的群組編號
    intGroup = 1
    For i = 1 To mBlockCnt
        If mBlock(i).GroupID >= intGroup Then intGroup = mBlock(i).GroupID + 1
    Next
    
    '群組步驟二：將所有選取的物件開始往上推至選取物件最高層的鄰近
    '求出可以上推的最頂層
    For i = mBlockCnt To 1 Step -1
        If mBlock(i).Selected Then
            mBlock(i).GroupID = intGroup
            iEnd = i
            Exit For
        End If
    Next
    
    '開始往上推
    For i = 1 To iEnd - 1
        
        If mBlock(i).Selected Then
            
            mBlock(i).GroupID = intGroup
            
            For ii = i To iEnd - 1
                
                '判斷上一層的物件是否已經被選取，否的話才會和上一層交換
                If mBlock(ii + 1).Selected = False Then
                                                                            
                    '交換 clsBlock 物件
                    Set Block = mBlock(ii)
                    Set mBlock(ii) = mBlock(ii + 1)
                    Set mBlock(ii + 1) = Block
                    Set Block = Nothing
                    
                    '交換 mBlockPic 物件
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
        
    '由最高層的次一個開始﹝最高層次不需要再往上﹞
    For i = mBlockCnt - 1 To 1 Step -1
        
        If mBlock(i).Selected Then
             
             For ii = i To mBlockCnt - 1
                                          
                If mBlock(ii + 1).Selected = False Then
                                                        
                    If iGroup <> 0 Then
                        If iGroup <> mBlock(ii + 1).GroupID Then Exit For
                    End If
                    If mBlock(ii + 1).GroupID <> 0 Then iGroup = mBlock(ii + 1).GroupID
                    
                    '交換 clsBlock 物件
                    Set Block = mBlock(ii)
                    Set mBlock(ii) = mBlock(ii + 1)
                    Set mBlock(ii + 1) = Block
                    Set Block = Nothing
    
                    '交換 mBlockPic 物件
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
    
    '由最高層次往下找是否有上推的物件
    For i = mBlockCnt - 1 To 1 Step -1
        If mBlock(i).Selected Then
            
            '找到要上推的物件後，開始將物件推至最頂層
            For ii = i To mBlockCnt - 1
                
                If mBlock(ii + 1).Selected = False Then
                    '交換 clsBlock 物件
                    Set Block = mBlock(ii)
                    Set mBlock(ii) = mBlock(ii + 1)
                    Set mBlock(ii + 1) = Block
                    Set Block = Nothing

                    '交換 mBlockPic 物件
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
        
    '由最低層的次一個開始﹝最低層次不需要再往下﹞
    For i = 2 To mBlockCnt
        
        If mBlock(i).Selected Then
            
             For ii = i To 2 Step -1
                                          
                If mBlock(ii - 1).Selected = False Then
                                                        
                    If iGroup <> 0 Then
                        If iGroup <> mBlock(ii - 1).GroupID Then Exit For
                    End If
                    If mBlock(ii - 1).GroupID <> 0 Then iGroup = mBlock(ii - 1).GroupID
                    
                    '交換 clsBlock 物件
                    Set Block = mBlock(ii)
                    Set mBlock(ii) = mBlock(ii - 1)
                    Set mBlock(ii - 1) = Block
                    Set Block = Nothing
    
                    '交換 mBlockPic 物件
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
    
    '由最低層次往上找是否有下推的物件
    For i = 2 To mBlockCnt
    
        If mBlock(i).Selected Then
                        
            '找到要下推的物件後，開始將物件推至最底層
            For ii = i To 2 Step -1
            
                If mBlock(ii - 1).Selected = False Then
                    '交換 clsBlock 物件
                    Set Block = mBlock(ii)
                    Set mBlock(ii) = mBlock(ii - 1)
                    Set mBlock(ii - 1) = Block
                    Set Block = Nothing

                    '交換 mBlockPic 物件
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
    
    '清空原本的剪貼簿
    For i = 0 To mClipBoardCnt - 1
        Controls.Remove ("picClip" & Trim(i))
    Next
    
    mClipBoardCnt = 0

    For i = 1 To mBlockCnt

        If mBlock(i).Selected Then
                        
            '建立 clsBlock 剪貼簿物件
            ReDim Preserve mClipboard(mClipBoardCnt) As clsBlock
            Set mClipboard(mClipBoardCnt) = New clsBlock
            
            Call SetBlockObject(mBlock(i), mClipboard(mClipBoardCnt))
                                            
            '建立 picClip 剪貼簿物件
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

        '建立 clsBlock 物件
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
                                    
            '若貼上的物件為群組，則指定新的群組編號
            If .GroupID <> 0 Then
            
                If intNowGroup <> .GroupID Then
                
                    intNowGroup = .GroupID
                    
                    '取出目前可以用的群組編號
                    For ii = 1 To mBlockCnt
                        If mBlock(ii).GroupID >= intNewGroup Then intNewGroup = mBlock(ii).GroupID + 1
                    Next
                End If
                
                .GroupID = intNewGroup
            End If
                        
        End With
        
        '建立 PictureBox 物件
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
    
    '由後面的物件開始往前找
    For i = mBlockCnt To 1 Step -1
    
        If mBlock(i).Selected Then
            
            '若找到要刪除的物件，則將該物件上推至最上層
            For ii = i To mBlockCnt - 1
                
                '交換 clsBlock 物件
                Set Block = mBlock(ii)
                Set mBlock(ii) = mBlock(ii + 1)
                Set mBlock(ii + 1) = Block
                Set Block = Nothing
                
                '交換 mBlockPic 物件
                Set Pic = mBlockPic(ii).Picture
                mBlockPic(ii).Picture = mBlockPic(ii + 1).Picture
                mBlockPic(ii + 1).Picture = Pic
                Set Pic = Nothing
            Next
            
            '刪除最上層的物件
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

'----------------------------------------- 結束屬於 clsBlock 物件的屬性宣告 ----------------------------------

Public Property Let CanDrawObjects(blnDraw As Boolean)
    mblnDrawObjects = blnDraw
End Property
Public Property Get CanDrawObjects() As Boolean
    CanDrawObjects = mblnDrawObjects
End Property

Public Function CreateObject(Left As Long, Top As Long, Width As Long, Height As Long, ControlType As typControlType) As Boolean

    mBlockCnt = mBlockCnt + 1
    
    mBlockIndex = mBlockCnt
    
    '建立 clsBlock 物件
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
    
    '建立 PictureBox 物件
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
        
    '將目前編輯的文字方塊隱藏並且寫入目前編輯的文字
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
        
'    mMouseAct = MoveSelect '將滑鼠動作重置為選取
    picBG.Cls
    mBlockIndex = 0
    
    If Button = vbLeftButton Or Button = vbRightButton Then
                
        ReDim Preserve iArray(0)
        spSelect.Visible = True
        
        '判斷是否處於調整物件大小的工具上
        For i = 1 To mBlockCnt
     
            With mBlock(i)
                                
                If .Selected Then
                    
                    If .GroupID = 0 Then
                    
                        '轉換對應的座標比例
                        Left = .Left * gsngViewScale
                        Top = .Top * gsngViewScale
                        Width = .Width * gsngViewScale
                        Height = .Height * gsngViewScale
                    Else
                    
                        '避免重複處理群組的程式區段 blnSkip
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
                                                                
                            '取得群組物件的調整駐點
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
                    
                    '調整上左
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
                    
                    '調整上中
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
                    
                    '調整上右
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
                    
                    '調整中左
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
    
                    '調整中右
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
    
                    '調整下左
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
    
                    '調整下中
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
    
                    '調整下右
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
        
        '用來判斷目前按下的物件是否為 Mouse_Up 由範圍選取選擇的物件(s)，
        '若是的話 blnIsFirstSelected = False 若否的話 blnIsFirstSelected = True
        '本段程式碼用於當使用者想要將先前 Mouse_Up 選取的物件，群組移動。
        
        blnIsFirstSelected = True
        blnIsObjSelected = False
        
        '不是第一次按下物件的處理程序
        For i = mBlockCnt To 1 Step -1
        
            With mBlock(i)
                '轉換對應的座標比例
                Left = .Left * gsngViewScale
                Top = .Top * gsngViewScale
                Width = .Width * gsngViewScale
                Height = .Height * gsngViewScale
                
                '判斷物件是否在之前的就已經選取過
                If x >= Left And y >= Top _
                        And x <= Left + Width And y <= Top + Height Then
                    
                    mBlockIndex = i
                    
                    If mBlock(i).Selected Then
                        mMouseAct = MoveObject
                                                                        
                        spObj(i).Move Left, Top, Width, Height
                        blnIsFirstSelected = False
                                                
                        '若按住shift會取消目前點選的物件
                        If Shift = 1 Then
                            .Selected = False
                            If .GroupID <> 0 Then
                                For ii = 1 To mBlockCnt
                                    If .GroupID = mBlock(ii).GroupID Then mBlock(ii).Selected = False
                                Next
                            End If
                        End If
                                     
                        '若選取到文字方塊內容，則取消其他物件的選取以及畫出調整駐點
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
                        '若按住shift會選取目前點選的物件
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
                    '判斷選取的是否為文字方塊的外圍邊界
                    If .ControlType = cjTextBox Then
                        If (x >= Left And y >= Top - BOXSIZE And x <= Left + Width And y <= Top) Or _
                            (x >= Left - BOXSIZE And y >= Top And x <= Left And y <= Top + Height) Or _
                            (x >= Left And y >= Top + Height And x <= Left + Width And y <= Top + Height + BOXSIZE) Or _
                            (x >= Left + Width And y >= Top And x <= Left + Width + BOXSIZE And y <= Top + Height) And .Selected Then
                                                        
                            mBlockIndex = i
                                                                                    
                            mMouseAct = MoveObject '將滑鼠動作更變為移動物件
                                                                        
                            spObj(i).Move Left, Top, Width, Height
                            rtb.Visible = False
                            
                            blnIsFirstSelected = False
                            
                        End If
                    End If
                                                                                                
                End If
                
            End With
        Next
        
        '第一次選取到 clsBlock 物件
        If blnIsFirstSelected Then
                
            For i = mBlockCnt To 1 Step -1
                        
                With mBlock(i)
                    
                     If blnIsObjSelected Then '已經選取最上層物件，不再選取下面的物件
                        
                        '若物件不是群組則清除選取；若物件是群組且群組編號與目前選取群組(iTmp)不相同，也清空
                        If .GroupID <> iTmp Or .GroupID = 0 Then
                            .Selected = False
                            spObj(i).Visible = False
                        End If
                     Else
                     
                        '轉換對應的座標比例
                        Left = .Left * gsngViewScale
                        Top = .Top * gsngViewScale
                        Width = .Width * gsngViewScale
                        Height = .Height * gsngViewScale
                    
                        If x >= Left And y >= Top And x <= Left + Width And y <= Top + Height Then
                                                                                                   
                            mMouseAct = MoveObject '將滑鼠動作更變為移動物件
                            
                            mBlockIndex = i
                            blnIsObjSelected = True
                                                                                    
                            .Selected = True '標示物件已經被選取
                            
                            spObj(i).Move Left, Top, Width, Height
                            spObj(i).Visible = True
                            
                            If rtb.Visible And mTextIndex > 0 Then
                                mBlock(mTextIndex).Text = rtb.Text
                                rtb.Visible = False
                            End If
                            
                            '若選取到文字方塊內容，則取消其他物件的選取以及畫出調整駐點
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
                                For ii = 1 To mBlockCnt '取得群組的物件
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
                            
                            '判斷選取的是否為文字方塊的外圍邊界
                            If .ControlType = cjTextBox Then
                                If (x >= Left And y >= Top - BOXSIZE And x <= Left + Width And y <= Top) Or _
                                    (x >= Left - BOXSIZE And y >= Top And x <= Left And y <= Top + Height) Or _
                                    (x >= Left And y >= Top + Height And x <= Left + Width And y <= Top + Height + BOXSIZE) Or _
                                    (x >= Left + Width And y >= Top And x <= Left + Width + BOXSIZE And y <= Top + Height) Then
                                    
                                    mMouseAct = MoveObject '將滑鼠動作更變為移動物件
                                    
                                    mBlockIndex = i
                                    mTextIndex = i
                                    
                                    blnIsObjSelected = True
                                                                                            
                                    .Selected = True '標示物件已經被選取
                                    
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
            
            '移動的是物件位置
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
                                    
            '移動的是 選取範圍 以及 加入 clsBlock 物件的範圍
            Case MoveSelect, InsertBlock, InsertPicture, InsertTextbox
                
                With mSelRect
                
                    '判斷滑鼠按下選取的範圍
                    
                    '左上到右下
                    If .Left < x And .Top < y Then
                        .Width = x - .Left
                        .Height = y - .Top
                        spSelect.Move .Left, .Top, .Width, .Height
                    End If
                    
                    '左下到右上
                    If .Left < x And .Top > y Then
                        .Width = x - .Left
                        .Height = .Top - y
                        spSelect.Move .Left, y, .Width, .Height
                    End If
                                
                    '右下到左下
                    If .Left > x And .Top < y Then
                        .Width = .Left - x
                        .Height = y - .Top
                        spSelect.Move x, .Top, .Width, .Height
                    End If
                    
                    '右下到左上
                    If .Left > x And .Top > y Then
                        .Width = .Left - x
                        .Height = .Top - y
                        spSelect.Move x, y, .Width, .Height
                    End If
            
                End With
                        
            '移動物件上左的大小
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
                
            '移動物件上中的大小
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
    
    '計算出選取範圍的正確大小
    With mSelRect
        '左下到右上
        If .Left < x And .Top > y Then
            .Top = y
        End If
        
        '右上到左下
        If .Left > x And .Top < y Then
            .Left = x
        End If
        
        '左下到右上
        If .Left > x And .Top > y Then
            .Left = x
            .Top = y
        End If
        
    End With
    
    Select Case mMouseAct
    
        '放開的是物件調整過後的大小
        Case MoveTopLeftSize, MoveTopMiddleSize, MoveTopRightSize, _
                MoveMiddleLeftSize, MoveMiddleRightSize, _
                MoveBottomLeftSize, MoveBottomMiddleSize, MoveBottomRightSize
            
            mMouseAct = MoveSelect '將滑鼠動作重置為選取
            
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
            
        '放開的是物件
        Case MoveObject
                        
            mMouseAct = MoveSelect '將滑鼠動作重置為選取
                        
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
                        
        '放開的是範圍選取
        Case MoveSelect
        
            spSelect.Visible = False
            picBG.Cls
            
            '判斷是否透過 圈選範圍 選取到 clsBlock 物件
            For i = 1 To mBlockCnt
            
                With mBlock(i)
                
                    spObj(i).Visible = False
                    
                    Left = .Left * gsngViewScale
                    Top = .Top * gsngViewScale
                    Width = .Width * gsngViewScale
                    Height = .Height * gsngViewScale
                    
                    If Left >= mSelRect.Left And Top >= mSelRect.Top _
                        And Left + Width <= mSelRect.Left + mSelRect.Width And Top + Height <= mSelRect.Top + mSelRect.Height Then
                    
                        mMouseAct = MoveObject '將滑鼠動作更變為移動物件
                        
                        mBlockIndex = i
                        
                        .Selected = True '標示物件已經被選取
                        spObj(i).Move Left, Top, Width, Height
                        
                        If .ControlType = cjTextBox Then Call DrawTextBox(mBlock(i), Left, Top, Width, Height)
                                                                                    
                        If iGroup <> .GroupID Then
                            iGroup = .GroupID
                            
                            '往底層找是否有相同的群組
                            For ii = i To 1 Step -1
                                If mBlock(ii).GroupID = iGroup Then
                                    mBlock(ii).Selected = True
                                Else
                                    Exit For
                                End If
                            Next
                            
                        End If
                                                                       
                    Else
                    
                        If Shift <> 1 Then .Selected = False '若沒按住shift則會取消不再選取範圍內的物件
                        
                        '若物件為群組，則檢查是否點選了群組中的任一物件
                        If .GroupID <> 0 Then
                            '往底層找是否有相同的群組
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
            
        '放開的是插入 clsBlock
        Case InsertBlock, InsertPicture, InsertTextbox
        
            spSelect.Visible = False
            
            '大於一定的區域才會新增 clsBlock 物件
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
                
                mMouseAct = MoveSelect '將滑鼠動作重置為選取
                
                Call DrawObjects
                
            End If
        
        Case EditTextBox
            mMouseAct = MoveSelect '將滑鼠動作重置為選取
                                            
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
    
    ' 建立 Brush(圖刷) 物件
    hBrush = CreateSolidBrush(vbBlue)
    ' 建立 Pen(筆) 物件
    hPen = CreatePen(vbInsideSolid, 1, vbWhite)
    ' 選取物件:Brush 及 Pen
    hOldBrush = SelectObject(picBG.hDC, hBrush)
    hOldPen = SelectObject(picBG.hDC, hPen)
    
    ReDim Preserve iArray(0)
    
    For i = 1 To mBlockCnt
        
        With mBlock(i)
        
            If .Selected Then
            
                If mBlock(i).GroupID = 0 Then
                
                    '選取非群組物件
                    Left = .Left * gsngViewScale
                    Top = .Top * gsngViewScale
                    Height = .Height * gsngViewScale
                    Width = .Width * gsngViewScale
                                
                Else
                
                    '避免重複處理群組的程式區段 blnSkip
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
                                                            
                        '選取群組物件
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
                                
                '上左
                Rectangle picBG.hDC, Left - BOXSIZE - 1, Top - BOXSIZE - 1, Left - 1, Top - 1
                '上中
                Rectangle picBG.hDC, (Left * 2 + Width) / 2 - BOXSIZE / 2, Top - BOXSIZE - 1, (Left * 2 + Width) / 2 + BOXSIZE / 2, Top - 1
                '上右
                Rectangle picBG.hDC, Left + Width, Top - BOXSIZE - 1, Left + Width + BOXSIZE, Top - 1
                '中左
                Rectangle picBG.hDC, Left - BOXSIZE - 1, (Top * 2 + Height) / 2 - BOXSIZE / 2, Left - 1, (Top * 2 + Height) / 2 + BOXSIZE / 2
                '中右
                Rectangle picBG.hDC, Left + Width, (Top * 2 + Height) / 2 - BOXSIZE / 2, Left + Width + BOXSIZE, (Top * 2 + Height) / 2 + BOXSIZE / 2
                '下左
                Rectangle picBG.hDC, Left - BOXSIZE - 1, Top + Height, Left - 1, Top + Height + BOXSIZE
                '下中
                Rectangle picBG.hDC, (Left * 2 + Width) / 2 - BOXSIZE / 2, Top + Height, (Left * 2 + Width) / 2 + BOXSIZE / 2, Top + Height + BOXSIZE
                '下右
                Rectangle picBG.hDC, Left + Width, Top + Height, Left + Width + BOXSIZE, Top + Height + BOXSIZE

            End If
            
        End With
    Next

    ' 還原原先的 Brush 及 Pen
    SelectObject picBG.hDC, hOldBrush
    SelectObject picBG.hDC, hOldPen
    
    ' 刪除 Pen 及 Brush 物件
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
                hBrush = CreateHatchBrush(HS_BDIAGONAL, vbBlack) ' 建立 Brush(圖刷) 物件
            Else
                hBrush = CreateHatchBrush(HS_FDIAGONAL, vbBlack) ' 建立 Brush(圖刷) 物件
            End If
            
            hPen = CreatePen(vbInvisible, 0, vbBlack) ' 建立 Pen(筆) 物件
            '選取物件:Brush 及 Pen
            hOldBrush = SelectObject(picBG.hDC, hBrush)
            hOldPen = SelectObject(picBG.hDC, hPen)

            Rectangle picBG.hDC, Left - BOXSIZE, Top - BOXSIZE, Left + Width + BOXSIZE, Top + BOXSIZE / 2 - 1
            Rectangle picBG.hDC, Left - BOXSIZE, Top + Height, Left + Width + BOXSIZE, Height + Top + BOXSIZE
            Rectangle picBG.hDC, Left - BOXSIZE, Top, Left + BOXSIZE / 2 - 1, Height + Top
            Rectangle picBG.hDC, Left + Width, Top, Left + Width + BOXSIZE, Height + Top

            ' 還原原先的 Brush 及 Pen
            SelectObject picBG.hDC, hOldBrush
            SelectObject picBG.hDC, hOldPen
            ' 刪除 Pen 及 Brush 物件
            DeleteObject hBrush
            DeleteObject hPen
            
        End If
        
        If Len(.Text) = 0 Then
            '畫出文字方塊的邊緣
            hPen = CreatePen(vbDot, 0, vbBlack) '建立 Pen(筆) 物件
            hOldPen = SelectObject(picBG.hDC, hPen) '選取物件Pen
            Rectangle picBG.hDC, Left - 1, Top - 1, Width + Left + 1, Height + Top + 1
            SelectObject picBG.hDC, hOldPen ' 還原原先的Pen
            DeleteObject hPen ' 刪除 Pen 物件
            
        Else
            'Font Size of VB, Convert to API rule is (vb.Size * -20) / Screen.TwipsPerPixelY
            Font.lfHeight = (.FontSize * -20) / Screen.TwipsPerPixelY * gsngViewScale
            Font.lfWidth = (0 * -20) / Screen.TwipsPerPixelY
            Font.lfCharSet = DEFAULT_CHARSET
            Font.lfEscapement = 0 * 10
            If .FontBold Then Font.lfWeight = 700 Else Font.lfWeight = 400
            If .FontItlic Then Font.lfItalic = 1 Else Font.lfItalic = 0
            If .FontUnderline Then Font.lfUnderline = 1 Else Font.lfUnderline = 0
                                                                                                    
            ' LOGFONT 資料結構的設定
            RtlMoveMemory Font.lfFaceName(0), _
                           ByVal .FontName, _
                           LenB(StrConv(.FontName, vbFromUnicode)) + 1

'            Font.lfStrikeOut = .Font.lfStrikeOut
'
            ' 建立字型物件
            hFont = CreateFontIndirect(Font)
            ' 選取字型物件
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
                                
            ' 還原字型
            SelectObject picBG.hDC, hOldFont
            ' 刪除字型
            DeleteObject hFont
            
        End If
    
    End With
End Sub

Private Sub DrawShape(clsObject As clsBlock, Left As Single, Top As Single, Width As Single, Height As Single)

    Dim hBrush As Long, hPen As Long
    Dim hOldBrush As Long, hOldPen As Long
    
    With clsObject
    
        If .BorderWidth > 0 Then
        
            ' 建立 Brush(圖刷) 物件
            hBrush = CreateSolidBrush(.BackColor)
            ' 建立 Pen(筆) 物件
            hPen = CreatePen(vbInsideSolid, .BorderWidth, .BorderColor)
            ' 選取物件:Brush 及 Pen
            hOldBrush = SelectObject(picBG.hDC, hBrush)
            hOldPen = SelectObject(picBG.hDC, hPen)
        
            Select Case .Shape
                Case blkRectangle   '矩形
                    Rectangle picBG.hDC, Left, Top, Left + Width, Top + Height
                    
                Case blkSquare      '正方形
                
                Case blkEllipse     '橢圓
                    Ellipse picBG.hDC, Left, Top, Left + Width, Top + Height
                    
                Case blkCircle      '圓形
                
                Case blkRoundRect   '圓角矩形
                    RoundRect picBG.hDC, Left, Top, Left + Width, Top + Height, .RoundAngel, .RoundAngel
                    
                Case blkRoundSquare '圓角正方形
                    
            End Select
            
            ' 還原原先的 Brush 及 Pen
            SelectObject picBG.hDC, hOldBrush
            SelectObject picBG.hDC, hOldPen
            
            ' 刪除 Pen 及 Brush 物件
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
            
            '若沒有圖形則畫出物件的邊緣
            hPen = CreatePen(vbDot, 0, vbBlack) '建立 Pen(筆) 物件
            hOldPen = SelectObject(picBG.hDC, hPen) '選取物件Pen
            Rectangle picBG.hDC, Left, Top, Width + Left, Height + Top
            SelectObject picBG.hDC, hOldPen ' 還原原先的Pen
            DeleteObject hPen ' 刪除 Pen 物件
                
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

    '滑鼠的操作為選取
    mMouseAct = MoveSelect

    '顯示比例為 100%
    gsngViewScale = 1
    
    '動態載入 PictureBox 來當 cjEditor 的圖形緩衝區
    Set cjEditor.PictureBox = Controls.Add("VB.PictureBox", "cjPictureBox")
    cjEditor.PictureBox.AutoRedraw = True
    cjEditor.PictureBox.AutoSize = True
    
    cjEditor.GridShow = True
    cjEditor.GridScale = 8
        
    mblnDrawObjects = True
    
    '設定各個控制點的滑鼠游標
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
