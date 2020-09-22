VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPaint 
   BorderStyle     =   4  '³æ½u©T©w¤u¨ãµøµ¡
   Caption         =   " Painting"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      Height          =   3015
      Left            =   240
      ScaleHeight     =   197
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   213
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
   Begin MSComctlLib.Slider sld 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Max             =   20
      TickStyle       =   3
   End
   Begin VB.CommandButton cmdSelGraph 
      Caption         =   "¿ï¥Î¯S®Ä"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
End
Attribute VB_Name = "frmPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSelectedGraph      As Integer



Private Sub cmdSelGraph_Click()
    mSelectedGraph = 31
End Sub

Private Sub pic_Click()
'    Dim hBitmap      As Long
'
'    hBitmap = frmTst.cjEdt.objhBitmap
'
'    pic.AutoRedraw = True
'    pic.Cls
'
'    DrawBitmap pic.hDC, pic.ScaleLeft, pic.ScaleTop, pic.ScaleWidth, pic.ScaleHeight, _
'                hBitmap, 100, 100
'    pic.Refresh
'    pic.AutoRedraw = False
End Sub

Private Sub sld_Change()
    
    Dim lngRet      As Long

    Select Case mSelectedGraph
        Case 31
            GPX_Waves pic.hDC, pic.hDC, sld.Value, sld.Value, sld.Value, True, lngRet
    End Select
End Sub
