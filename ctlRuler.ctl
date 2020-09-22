VERSION 5.00
Begin VB.UserControl cjRuler 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   CanGetFocus     =   0   'False
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2160
   ScaleHeight     =   405
   ScaleWidth      =   2160
   ToolboxBitmap   =   "ctlRuler.ctx":0000
   Begin VB.Menu mnuMenu 
      Caption         =   "Kontext"
      Visible         =   0   'False
      Begin VB.Menu mnuMode 
         Caption         =   "Centimeter"
         Index           =   0
      End
      Begin VB.Menu mnuMode 
         Caption         =   "Inch"
         Index           =   1
      End
      Begin VB.Menu mnuMode 
         Caption         =   "Pixel * 100"
         Index           =   2
      End
      Begin VB.Menu mnuMode 
         Caption         =   "Twip * 1000"
         Index           =   3
      End
   End
End
Attribute VB_Name = "cjRuler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private OLE
Private X1 As Single
Private RScale As Long
'Private MSize As Single

Public Enum RulerModeConst
    Millimeters = 0
    Inch = 1
    Pixel = 2
    Twips = 3
End Enum

Public Enum asrOrientationConstants
    asrHorizontal = 0
    asrVertival = 1
End Enum

Private m_Mode              As RulerModeConst
Private m_Orientation       As asrOrientationConstants
Private msngViewPercent     As Single

Public Property Let ViewPercent(Percent As Single)
    msngViewPercent = Percent / CSng(100)
End Property

Public Property Get ViewPercent() As Single
    ViewPercent = msngViewPercent * CSng(100)
End Property

Public Property Get RulerMode() As RulerModeConst
    RulerMode = m_Mode
End Property

Public Property Let RulerMode(New_Mode As RulerModeConst)
    m_Mode = New_Mode
    Select Case m_Mode
        Case 0
            RScale = 570 * msngViewPercent
        Case 1
            RScale = 1440 * msngViewPercent
        Case 2
            RScale = Screen.TwipsPerPixelX * msngViewPercent * 100
        Case 3
            RScale = 1000 * msngViewPercent
    End Select
    UserControl.Cls
    DrawRuler
    PropertyChanged "RulerMode"
End Property

Public Property Get Orientation() As asrOrientationConstants
    Orientation = m_Orientation
End Property
Public Property Let Orientation(New_Val As asrOrientationConstants)
    m_Orientation = New_Val
    UserControl.Cls
    DrawRuler
    PropertyChanged "Orientation"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(New_Val As OLE_COLOR)
    UserControl.BackColor = New_Val
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(New_Val As OLE_COLOR)
    UserControl.ForeColor = New_Val
    PropertyChanged "ForeColor"
End Property

Private Sub DrawRuler()
    Dim Sincr As Single
    'Scalemode is in TWIPS 1440 per inch
    Dim i As Integer
    'Number of segment across form
    Sincr = RScale / 10
    With UserControl
        If m_Orientation = asrHorizontal Then
            Do While Sincr < .ScaleWidth
                'Number of sections
                For i = 1 To 10
                    'Size of Tics
                    If i = 10 Then
                        UserControl.Line (Sincr, 0)-(Sincr, .ScaleHeight)
                        .CurrentY = 0
                        'UserControl.Print CStr(Int(Sincr / RScale))
                        
                    ElseIf i = Int(10 * 0.5) Then
                        UserControl.Line (Sincr, .ScaleHeight - _
                        (.ScaleHeight * 0.5))-(Sincr, .ScaleHeight)
                    Else
                        UserControl.Line (Sincr, .ScaleHeight - _
                        (.ScaleHeight * 0.125))-(Sincr, .ScaleHeight)
                    End If
                    Sincr = Sincr + (RScale / 10)
                Next
            Loop
        Else
            Do While Sincr < .ScaleHeight
                'Number of sections
                For i = 1 To 10
                    'Size of Tics
                    If i = 10 Then
                        'Einheiten schreiben
                        UserControl.Line (0, Sincr)-(.ScaleHeight, Sincr)
                        .CurrentX = 0
'                        UserControl.Print CStr(Int(Sincr / RScale))
                    ElseIf i = Int(10 * 0.5) Then
                        '50%
                        UserControl.Line (.ScaleWidth - _
                        (.ScaleWidth * 0.5), Sincr)-(.ScaleWidth, Sincr)
                    Else
                        UserControl.Line (.ScaleWidth - _
                        (.ScaleWidth * 0.125), Sincr)-(.ScaleWidth, Sincr)
                    End If
                    Sincr = Sincr + (RScale / 10)
                Next
            Loop
        End If
    End With
End Sub

Public Sub MouseMoved(x As Single)

    With UserControl
        .DrawMode = 6
        If m_Orientation = asrHorizontal Then
            UserControl.Line (x, 0)-(x, .ScaleHeight)
            If X1 > 0 Then
                UserControl.Line (X1, 0)-(X1, .ScaleHeight)
            End If
            X1 = x
        Else
            UserControl.Line (0, x)-(.ScaleWidth, x)
            If X1 > 0 Then
                UserControl.Line (0, X1)-(.ScaleWidth, X1)
            End If
            X1 = x
        End If
        .DrawMode = 13
    End With
End Sub

Private Sub mnuMenu_Click()
    Dim i As Integer
    For i = 0 To mnuMode.Count - 1
        mnuMode(i).Checked = False
    Next i
    mnuMode(m_Mode).Checked = True
End Sub

Private Sub mnuMode_Click(Index As Integer)
    RulerMode = Index
End Sub

Private Sub UserControl_Initialize()
    msngViewPercent = 1
    RScale = 570
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Button = 2 Then UserControl.PopupMenu mnuMenu
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Orientation = .ReadProperty("Orientation", 0)
        RulerMode = .ReadProperty("RulerMode", 0)
        UserControl.BackColor = .ReadProperty("BackColor", &H80000018)
        UserControl.ForeColor = .ReadProperty("ForeColor", &H80000012)
    End With
End Sub

Private Sub UserControl_Resize()
    UserControl.Cls
    X1 = 0
    'Draw Ruler 16ths of an inch
    DrawRuler

End Sub

'Private Sub UserControl_Show()
'    DrawRuler
'End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Orientation", m_Orientation, 0
        .WriteProperty "RulerMode", m_Mode, 0
        .WriteProperty "BackColor", UserControl.BackColor, &H80000018
        .WriteProperty "ForeColor", UserControl.ForeColor, &H80000012
    End With
End Sub
