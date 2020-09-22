VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B1285566-69F2-4974-9F90-690280681391}#9.2#0"; "cjEditor.ocx"
Begin VB.Form frmTst 
   Caption         =   "cjEditor"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10935
   Icon            =   "frmTst.frx":0000
   ScaleHeight     =   6840
   ScaleWidth      =   10935
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   WindowState     =   2  '³Ì¤j¤Æ
   Begin MSComctlLib.ImageList ils 
      Left            =   6240
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   33
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":0CCA
            Key             =   "Pointer"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":19A4
            Key             =   "Block"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":1D3E
            Key             =   "ZorderTop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":1E98
            Key             =   "ZorderBottom"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":1FF2
            Key             =   "CBackground"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":2CCC
            Key             =   "CBackColor"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":3266
            Key             =   "ZorderUp"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":3600
            Key             =   "ZorderDown"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":399A
            Key             =   "Picture"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":3CED
            Key             =   "Text"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":3E47
            Key             =   "InsertBlock"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":3FA1
            Key             =   "Ellipse"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":40FB
            Key             =   "RectAngle"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":4255
            Key             =   "RoundRect"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":43AF
            Key             =   "BorderWidth"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":4509
            Key             =   "FillBackColor"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":4663
            Key             =   "LoadPic"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":47BD
            Key             =   "BorderColor"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":4B0F
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":4C69
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":4DC3
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":4F1D
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":5077
            Key             =   "RoundAngel"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":5411
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":556B
            Key             =   "Itlic"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":56C5
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":581F
            Key             =   "InsertPic"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":5BB9
            Key             =   "TextBox"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":6A0B
            Key             =   "Transparent"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":6B65
            Key             =   "FontColor"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":70FF
            Key             =   "TLeft"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":7499
            Key             =   "TRight"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTst.frx":7833
            Key             =   "TCenter"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   6240
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin cjEditor.cjEdt cjEdt 
      Height          =   3855
      Left            =   480
      TabIndex        =   6
      Top             =   480
      Width           =   5655
      _extentx        =   9975
      _extenty        =   6800
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  '¹ï»ôªí³æ¤U¤è
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   6465
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   2
            Enabled         =   0   'False
            TextSave        =   "Áä½L¤W Caps Lock ªºª¬ºA"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "Áä½L¤Wªº Insert Áäª¬ºA"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6800
            Key             =   "Obj"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   6800
            Key             =   "Msg"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tblObj 
      Align           =   3  '¹ï»ôªí³æ¥ª¤è
      Height          =   6105
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   10769
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ils"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Block"
            Object.ToolTipText     =   "Change to Object"
            ImageKey        =   "Block"
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Picture"
            Object.ToolTipText     =   "Change to Picture"
            ImageKey        =   "Picture"
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TextBox"
            Object.ToolTipText     =   "Change to Textbox"
            ImageKey        =   "Text"
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RoundRect"
            Object.ToolTipText     =   "RoundRect Angle"
            ImageKey        =   "RoundRect"
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RectAngle"
            Object.ToolTipText     =   "RectAngle"
            ImageKey        =   "RectAngle"
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ellipse"
            Object.ToolTipText     =   "Ellipse"
            ImageKey        =   "Ellipse"
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RoundAngel"
            Object.ToolTipText     =   "RoundAngel"
            ImageKey        =   "RoundAngel"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BorderColor"
            Object.ToolTipText     =   "Border Color"
            ImageKey        =   "BorderColor"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BorderWidth"
            Object.ToolTipText     =   "Border Width"
            ImageKey        =   "BorderWidth"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BackColor"
            Object.ToolTipText     =   "Border Fill Color"
            ImageKey        =   "FillBackColor"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LoadPic"
            Object.ToolTipText     =   "Load Picture"
            ImageKey        =   "LoadPic"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Transparent"
            Object.ToolTipText     =   "Transparent Picture"
            ImageKey        =   "Transparent"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Slider sld 
      Height          =   255
      Left            =   10800
      TabIndex        =   1
      ToolTipText     =   "View Scale Percent"
      Top             =   105
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      Max             =   200
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
   End
   Begin MSComctlLib.Toolbar tbl 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ils"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   35
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tblMouseSelect"
            Object.ToolTipText     =   "Pointer"
            ImageKey        =   "Pointer"
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tblInsertBlock"
            Object.ToolTipText     =   "Insert Object"
            ImageKey        =   "InsertBlock"
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "InsertPic"
            Object.ToolTipText     =   "Insert Picture"
            ImageKey        =   "InsertPic"
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TextBox"
            Object.ToolTipText     =   "Insert Textbox"
            ImageKey        =   "TextBox"
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LayerTop"
            Object.ToolTipText     =   "Pull to Top"
            ImageKey        =   "ZorderTop"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LayerBottom"
            Object.ToolTipText     =   "Push to Bottom"
            ImageKey        =   "ZorderBottom"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LayerUp"
            Object.ToolTipText     =   "Pull Up"
            ImageKey        =   "ZorderUp"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LayerDown"
            Object.ToolTipText     =   "Push Down"
            ImageKey        =   "ZorderDown"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CBackColor"
            Object.ToolTipText     =   "Background Color"
            ImageKey        =   "CBackColor"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CBackGround"
            Object.ToolTipText     =   "Background Picture"
            ImageKey        =   "CBackground"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Itlic"
            Object.ToolTipText     =   "Itlic"
            ImageKey        =   "Itlic"
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FontColor"
            Object.ToolTipText     =   "Font Color"
            ImageKey        =   "FontColor"
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TLeft"
            Object.ToolTipText     =   "Aligment Left"
            ImageKey        =   "TLeft"
            Style           =   2
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TCenter"
            Object.ToolTipText     =   "Aligment Center"
            ImageKey        =   "TCenter"
            Style           =   2
         EndProperty
         BeginProperty Button35 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TRight"
            Object.ToolTipText     =   "Aligment Right"
            ImageKey        =   "TRight"
            Style           =   2
         EndProperty
      EndProperty
      Begin VB.ComboBox cboSize 
         Height          =   300
         Left            =   7320
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   5
         Top             =   20
         Width           =   735
      End
      Begin VB.ComboBox cboFont 
         Appearance      =   0  '¥­­±
         Height          =   300
         Left            =   5400
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   4
         Top             =   20
         Width           =   1900
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGroupObjects 
         Caption         =   "&Group"
         Begin VB.Menu mnuGroup 
            Caption         =   "&Group"
         End
         Begin VB.Menu mnuUnGroup 
            Caption         =   "&Cancel Group"
         End
      End
      Begin VB.Menu mnuLayer 
         Caption         =   "&Layer"
         Begin VB.Menu mnuLayerTop 
            Caption         =   "Pull to &Top"
         End
         Begin VB.Menu mnuLayerBottom 
            Caption         =   "Push to &Bottom"
         End
         Begin VB.Menu mnuLayerUp 
            Caption         =   "Pull &Up"
         End
         Begin VB.Menu mnuLayerDown 
            Caption         =   "Push &Down"
         End
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuObjClass 
         Caption         =   "Object &Class"
         Begin VB.Menu mnuBlock 
            Caption         =   "&Object"
         End
         Begin VB.Menu mnuPicture 
            Caption         =   "&Picture"
         End
         Begin VB.Menu mnuTextBox 
            Caption         =   "&Textbox"
         End
      End
      Begin VB.Menu mnuPicTransparent 
         Caption         =   "&Transparent"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGridShow 
         Caption         =   "Show &Grid"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuGridScale 
         Caption         =   "Grid &Scale"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackground 
         Caption         =   "&Background"
         Begin VB.Menu mnuBackgroundPic 
            Caption         =   "Load &Picture"
         End
         Begin VB.Menu mnuBackgroundColor 
            Caption         =   "Background &Color"
         End
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLocked 
         Caption         =   "&Locked"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuInsertObj 
         Caption         =   "Insert &Object"
      End
      Begin VB.Menu mnuInsertPic 
         Caption         =   "Insert &Picture"
      End
      Begin VB.Menu mnuInsertTextBox 
         Caption         =   "Insert &Textbox"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPaint 
         Caption         =   "Picture &Filter..."
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About cjEditor"
      End
   End
End
Attribute VB_Name = "frmTst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFont_Click()
    cjEdt.objFontName = cboFont.Text
End Sub

Private Sub cboSize_Click()
    cjEdt.objFontSize = Val(cboSize.Text)
End Sub

Private Sub cjEdt_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngShift        As Long

    If Shift = 1 Then lngShift = 1 Else lngShift = 10
    Select Case KeyCode
        Case vbKeyUp
            cjEdt.objMoveShift 0, -lngShift
        Case vbKeyDown
            cjEdt.objMoveShift 0, lngShift
        Case vbKeyLeft
            cjEdt.objMoveShift -lngShift, 0
        Case vbKeyRight
            cjEdt.objMoveShift lngShift, 0
    End Select
            
End Sub

Private Sub cjEdt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    '§PÂ_ª«¥óºØÃþ
    If cjEdt.objSelectedIndex > 0 Then
        mnuBlock.Checked = False
        mnuPicture.Checked = False
        mnuTextBox.Checked = False
    
        Select Case cjEdt.objSelectedObject.ControlType
            Case typControlType.cjBlock
                tblObj.Buttons("Block").Value = tbrPressed
                mnuBlock.Checked = True
                
            Case typControlType.cjPicture
                tblObj.Buttons("Picture").Value = tbrPressed
                mnuPicture.Checked = True
                
            Case typControlType.cjTextBox
                tblObj.Buttons("TextBox").Value = tbrPressed
                mnuTextBox.Checked = True
                
        End Select
    End If
    
    '§PÂ_ª«¥ó¥~«¬
    tblObj.Buttons("RoundAngel").Enabled = False
    
    If cjEdt.objSelectedIndex > 0 Then
        Select Case cjEdt.objSelectedObject.Shape
            Case typShape.blkRoundRect: tblObj.Buttons("RoundRect").Value = tbrPressed: tblObj.Buttons("RoundAngel").Enabled = True
            Case typShape.blkRectangle: tblObj.Buttons("RectAngle").Value = tbrPressed
            Case typShape.blkEllipse: tblObj.Buttons("Ellipse").Value = tbrPressed
        End Select
    End If
    
End Sub

Private Sub cjEdt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stBar.Panels("Msg").Text = "®y¼Ð¡G" & Round(ScaleX(x, vbPixels, vbCentimeters), 1) & " cm  " & Round(ScaleY(y, vbPixels, vbCentimeters), 1) & " cm"
End Sub

Private Sub cjEdt_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case cjEdt.MouseAction
        Case cjEditor.MoveSelect
            tbl.Buttons("tblMouseSelect").Value = tbrPressed
        Case cjEditor.InsertBlock
            tbl.Buttons("tblInsertBlock").Value = tbrPressed
    End Select
    
    If cjEdt.objSelectedIndex > 0 Then
        stBar.Panels("Obj").Text = "ª«¥ó¦WºÙ: " & cjEdt.objSelectedObject.Name
        
        If cjEdt.objSelectedObject.PicTransparent Then
            mnuPicTransparent.Checked = True
            tblObj.Buttons("Transparent").Value = tbrPressed
        Else
            mnuPicTransparent.Checked = False
            tblObj.Buttons("Transparent").Value = tbrUnpressed
        End If
        
        mnuGroupObjects.Enabled = True
        If cjEdt.objSelectedObject.GroupID > 0 Then
            mnuGroup.Enabled = False
            mnuUnGroup.Enabled = True
        Else
            mnuGroup.Enabled = True
            mnuUnGroup.Enabled = False
        End If
                
        
    Else
        stBar.Panels("Obj").Text = ""
        mnuGroupObjects.Enabled = False
        mnuPicTransparent.Checked = False
        tblObj.Buttons("Transparent").Value = tbrUnpressed
                        
    End If
    
    If Button = vbRightButton Then Me.PopupMenu mnuEdit
    
End Sub

Private Sub Form_Load()
    
    Dim i       As Integer
    
    For i = 8 To 96
        cboSize.AddItem i
    Next
    cboSize.Text = 12

    For i = 0 To Screen.FontCount - 1
        cboFont.AddItem Screen.Fonts(i)
    Next
    cboFont.Text = "Times New Roman"

    cjEdt.MillimetersWidth = 194
    cjEdt.MillimetersHeight = 150
    sld.Value = 100
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    cjEdt.Move Me.ScaleLeft + tblObj.Width, Me.ScaleTop + tbl.Height, Me.ScaleWidth - tblObj.Width, Me.ScaleHeight - tbl.Height - stBar.Height
End Sub

Private Sub mnuAbout_Click()
    MsgBox "¼¶¼g§@ªÌ¡G¶À¬L¤¯" & vbCrLf & _
            "³nÅéª©¥»¡G" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
            "ª©Åv©Ò¦³¡Ghttp://www.cj.idv.tw" & vbCrLf & _
            "ÄÀ¥X¤é´Á¡G" & "24-Aug-2005"
End Sub

Private Sub mnuBackgroundColor_Click()
    On Error Resume Next
    Dlg.ShowColor
    If Err.Number <> cdlCancel Then cjEdt.BackColor = Dlg.Color
    On Error GoTo 0
End Sub

Private Sub mnuBackgroundPic_Click()
    On Error Resume Next
    Dlg.Filter = "©Ò¦³¹ÏÀÉ (*.bmp;*.ico;*.wmf;*.jpg;*.gif)|*.bmp;*.ico;*.wmf;*.jpg;*.gif|ÂI°}¹Ï (*.bmp)|*.bmp|¹Ï¥Ü¤å¥ó(*.ico)|*.ico|¤¤Ä~ÀÉ (*.wmf)|.wmf|JpegÀÉ (*.jpg)|*.jpg|GifÀÉ (*.gif)|*.gif"
    Dlg.ShowOpen
    If Err.Number <> cdlCancel Then
        Call cjEdt.LoadBackgroundPicture(Dlg.FileName)
    Else
        Call cjEdt.LoadBackgroundPicture("")
    End If
    On Error GoTo 0
End Sub

Private Sub mnuBlock_Click()
    cjEdt.objControlType = cjBlock
End Sub

Private Sub mnuCopy_Click()
    cjEdt.objCopy
End Sub

Private Sub mnuCut_Click()
    cjEdt.objCut
End Sub

Private Sub mnuDelete_Click()
    cjEdt.objDelete
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuGridScale_Click()
    cjEdt.GridScale = InputBox("½Ð¿é¤J©w¦ìÂI¤j¤p", , cjEdt.GridScale)
End Sub

Private Sub mnuGridShow_Click()
    mnuGridShow.Checked = Not mnuGridShow.Checked
    cjEdt.GridShow = mnuGridShow.Checked
End Sub

Private Sub mnuGroup_Click()
    cjEdt.objGroup
End Sub

Private Sub mnuInsertObj_Click()
    cjEdt.MouseAction = InsertBlock
End Sub

Private Sub mnuInsertPic_Click()
    cjEdt.MouseAction = InsertPicture
End Sub

Private Sub mnuInsertTextBox_Click()
    cjEdt.MouseAction = InsertTextbox
End Sub

Private Sub mnuLayerBottom_Click()
    cjEdt.objLayerBottom
End Sub

Private Sub mnuLayerDown_Click()
    cjEdt.objLayerDown
End Sub

Private Sub mnuLayerTop_Click()
    cjEdt.objLayerTop
End Sub

Private Sub mnuLayerUp_Click()
    cjEdt.objLayerUp
End Sub

Private Sub mnuOpen_Click()

'    Dim FNo             As Integer
'    Dim sngViewScale    As Single
'
'    On Error Resume Next
'    Dlg.Filter = "PMLS ³]©wÀÉ (*.txt|*.txt"
'    Dlg.ShowOpen
'    If Err.Number <> cdlCancel Then
'
'        FNo = FreeFile
'
'        On Error GoTo 0
'        cjEdt.Visible = False
'        sngViewScale = cjEdt.ClientViewPercent
'        cjEdt.ClientViewPercent = 100
'        Call LoadUnicodePMLS(FNo, StrReverse(Mid(StrReverse(Dlg.FileName), InStr(1, StrReverse(Dlg.FileName), "\"))))
'        cjEdt.ClientViewPercent = sngViewScale
'        cjEdt.Visible = True
'        Close #FNo
'
'    End If
    
End Sub

Private Sub mnuPaste_Click()
    cjEdt.objPaste
End Sub

Private Sub mnuPicTransparent_Click()
    mnuPicTransparent.Checked = Not mnuPicTransparent.Checked
    If mnuPicTransparent.Checked Then
        cjEdt.objPictureTransparent True
        tblObj.Buttons("Transparent").Value = tbrPressed
    Else
        cjEdt.objPictureTransparent False
        tblObj.Buttons("Transparent").Value = tbrUnpressed
    End If
End Sub

Private Sub mnuPicture_Click()
     cjEdt.objControlType = cjPicture
End Sub

Private Sub mnuSelectAll_Click()
    cjEdt.objSelectAll
End Sub

Private Sub mnuTextBox_Click()
    cjEdt.objControlType = cjTextBox
End Sub

Private Sub mnuUnGroup_Click()
    cjEdt.objUnGroup
End Sub

Private Sub sld_Change()
    cjEdt.ClientViewPercent = sld.Value
End Sub

Private Sub tbl_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "tblMouseSelect"
            cjEdt.MouseAction = MoveSelect
        Case "tblInsertBlock"
            mnuInsertObj_Click
        Case "InsertPic"
            mnuInsertPic_Click
        Case "TextBox"
            mnuInsertTextBox_Click
        Case "LayerTop"
            mnuLayerTop_Click
        Case "LayerBottom"
            mnuLayerBottom_Click
        Case "LayerUp"
            mnuLayerUp_Click
        Case "LayerDown"
            mnuLayerDown_Click
        Case "CBackColor"
            mnuBackgroundColor_Click
        Case "CBackGround"
            mnuBackgroundPic_Click
        Case "Cut"
            mnuCut_Click
        Case "Copy"
            mnuCopy_Click
        Case "Paste"
            mnuPaste_Click
        Case "Delete"
            mnuDelete_Click
        Case "Bold"
            cjEdt.objFontBold = Not cjEdt.objFontBold
        Case "Itlic"
            cjEdt.objFontItlic = Not cjEdt.objFontItlic
        Case "Underline"
            cjEdt.objFontUnderline = Not cjEdt.objFontUnderline
        Case "FontColor"
            Dlg.ShowColor
            cjEdt.objFontForeColor = Dlg.Color
        Case "TLeft"
            cjEdt.objTextAligment = txt_Left
        Case "TCenter"
            cjEdt.objTextAligment = txt_Center
        Case "TRight"
            cjEdt.objTextAligment = txt_Right
    End Select
    
End Sub

Private Sub tblObj_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Err_tblObj_ButtonClick

    Select Case Button.Key
        Case "Block"
            mnuBlock_Click
        Case "Picture"
            mnuPicture_Click
        Case "TextBox"
            mnuTextBox_Click
        Case "RectAngle"
            cjEdt.objShpae = blkRectangle
        Case "RoundRect"
            cjEdt.objShpae = blkRoundRect
        Case "RoundAngel"
            If cjEdt.objRoundAngel > 0 Then cjEdt.objRoundAngel = InputBox("½Ð¿é¤J¶ê¨¤ªº«×¼Æ¡G", , cjEdt.objRoundAngel)
        Case "Ellipse"
            cjEdt.objShpae = blkEllipse
        Case "BorderColor"
            On Error Resume Next
            Dlg.ShowColor
            If Err.Number <> cdlCancel Then cjEdt.objBorderColor = Dlg.Color
            On Error GoTo Err_tblObj_ButtonClick
        Case "BorderWidth"
            cjEdt.objBorderWidth = InputBox("½Ð¿é¤J®Ø½uªº¼e«×¡G", , cjEdt.objBorderWidth)
        Case "BackColor"
            On Error Resume Next
            Dlg.ShowColor
            If Err.Number <> cdlCancel Then cjEdt.objBackColor = Dlg.Color
            On Error GoTo Err_tblObj_ButtonClick
        Case "LoadPic"
            On Error Resume Next
            Dlg.Filter = "©Ò¦³¹ÏÀÉ (*.bmp;*.ico;*.wmf;*.jpg;*.gif)|*.bmp;*.ico;*.wmf;*.jpg;*.gif|ÂI°}¹Ï (*.bmp)|*.bmp|¹Ï¥Ü¤å¥ó(*.ico)|*.ico|¤¤Ä~ÀÉ (*.wmf)|.wmf|JpegÀÉ (*.jpg)|*.jpg|GifÀÉ (*.gif)|*.gif"
            Dlg.ShowOpen
            If Err.Number <> cdlCancel Then
                Call cjEdt.objLoadPicture(Dlg.FileName)
            Else
                Call cjEdt.objLoadPicture("")
            End If
            On Error GoTo Err_tblObj_ButtonClick
        
        Case "Transparent"
            mnuPicTransparent_Click
        
    End Select
    
    Exit Sub
    
Err_tblObj_ButtonClick:
    MsgBox Err.Description
    
End Sub
