VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   11040
   ClientLeft      =   3990
   ClientTop       =   -3240
   ClientWidth     =   19080
   FillColor       =   &H00FD8866&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C00000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Sao Viet Accounting Software"
   Picture         =   "frmMain.frx":424A
   ScaleHeight     =   11040
   ScaleWidth      =   19080
   Tag             =   "11"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbToolBar 
      Height          =   630
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1111
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   11
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "TaiKhoan"
            Object.ToolTipText     =   "Tµi kho¶n"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "NgoaiTe"
            Object.ToolTipText     =   "Nguyªn tÖ"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Kho"
            Object.ToolTipText     =   "Kho"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "VatTu"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "LuuChuyen"
            Object.ToolTipText     =   "L­u chuyÓn"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "DuPhong"
            Object.ToolTipText     =   "Dù phßng"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "TaiSan"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "CN"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "TongHop"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ThanhPham"
            Object.ToolTipText     =   "Thµnh phÈm"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "KetThuc"
            Object.ToolTipText     =   "Tho¸t"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.Timer timerChaytb 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   11760
      Top             =   1320
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   240
      TabIndex        =   80
      Top             =   10320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   13080
      TabIndex        =   78
      Top             =   1440
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox picPopup 
      Height          =   1335
      Left            =   7440
      ScaleHeight     =   1275
      ScaleWidth      =   1755
      TabIndex        =   74
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
      Begin VB.Label lblItem 
         Caption         =   "Label4"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   76
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblItem 
         Caption         =   "Label4"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   75
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.PictureBox picMenu 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      ScaleHeight     =   435
      ScaleWidth      =   10155
      TabIndex        =   72
      Top             =   0
      Visible         =   0   'False
      Width           =   10215
      Begin VB.Label lblMenu 
         Caption         =   "Label4"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   73
         Top             =   120
         Width           =   615
      End
   End
   Begin ComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   57
      Top             =   10650
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   688
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   12347
            MinWidth        =   12347
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "22/09/26"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   16080
      TabIndex        =   69
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer timerBackup 
      Interval        =   60000
      Left            =   9480
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "T¹o c«ng ty míi"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13560
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H0000C000&
      Caption         =   "NhËp chøng tõ"
      DragIcon        =   "frmMain.frx":177B0
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   1440
      MaskColor       =   &H0000C000&
      Picture         =   "frmMain.frx":284AA
      Style           =   1  'Graphical
      TabIndex        =   60
      Tag             =   "Voucher"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H0000C000&
      Caption         =   "Sæ kÕ to¸n"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   1440
      MaskColor       =   &H0000C000&
      Picture         =   "frmMain.frx":2DC8C
      Style           =   1  'Graphical
      TabIndex        =   59
      Tag             =   "Detail Report"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H0000C000&
      Caption         =   "B¸o c¸o thuÕ& tµi chÝnh"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   1440
      MaskColor       =   &H0000C000&
      Picture         =   "frmMain.frx":3346E
      Style           =   1  'Graphical
      TabIndex        =   58
      Tag             =   "Financial Report"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   9480
      OLEDropMode     =   1  'Manual
      TabIndex        =   31
      Top             =   4440
      Visible         =   0   'False
      Width           =   1515
      Begin VB.CheckBox chk 
         BackColor       =   &H00FFC0C0&
         Caption         =   "48"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FFC0C0&
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   840
         MaskColor       =   &H00000000&
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CDT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FFC0C0&
         Caption         =   "SX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2040
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command 
      Caption         =   "KÕ to¸n      Chñ ®Çu t­"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton Command 
      Caption         =   "B¸o c¸o   Qu¶n trÞ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   4
      Left            =   1800
      Picture         =   "frmMain.frx":38C50
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton Command 
      Caption         =   "Ch­¬ng tr×nh theo &yªu cÇu doanh nghiÖp"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   840
      TabIndex        =   2
      Tag             =   "Customized Report"
      Top             =   360
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   960
      TabIndex        =   10
      Top             =   5520
      Visible         =   0   'False
      Width           =   2520
      Begin VB.OptionButton OptNN 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ViÖt"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   240
         MaskColor       =   &H00000000&
         TabIndex        =   11
         Tag             =   "Vietnamese"
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton OptNN 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Anh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   3
         Left            =   1320
         MaskColor       =   &H00000000&
         TabIndex        =   12
         Tag             =   "English"
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8760
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Timer CTTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11280
      Top             =   2400
   End
   Begin VB.PictureBox imlIcons 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   13680
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   15
      Top             =   6240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin Crystal.CrystalReport Rpt 
      Left            =   3000
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "1.6"
      BeginProperty Font 
         Name            =   "VNI-Lithos"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0017A0D4&
      Height          =   570
      Index           =   21
      Left            =   18600
      TabIndex        =   87
      Tag             =   "Accounting Software Company"
      Top             =   1600
      Width           =   1095
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "LbCty 5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   17
      Left            =   5160
      TabIndex        =   86
      Tag             =   "1"
      Top             =   6600
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   20
      Left            =   3840
      TabIndex        =   85
      Tag             =   "Address"
      Top             =   6600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "LbCty 5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   16
      Left            =   5160
      TabIndex        =   84
      Tag             =   "1"
      Top             =   6240
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   19
      Left            =   3840
      TabIndex        =   83
      Tag             =   "Address"
      Top             =   6240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "LbCty 5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   15
      Left            =   6360
      TabIndex        =   82
      Tag             =   "1"
      Top             =   5880
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nhµ cung cÊp HDDT :"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   18
      Left            =   3840
      TabIndex        =   81
      Tag             =   "Address"
      Top             =   5880
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6600
      TabIndex        =   79
      Top             =   1800
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      Height          =   255
      Left            =   13320
      TabIndex        =   77
      Top             =   1560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Version 9.9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   71
      Top             =   9000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Version 9.9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   70
      Top             =   9000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   1155
      Left            =   13080
      Picture         =   "frmMain.frx":396D2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2400
   End
   Begin VB.Image Image1 
      Height          =   1725
      Left            =   17040
      Picture         =   "frmMain.frx":41CB1
      Top             =   360
      Width           =   1500
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Lo¹i h×nh ho¹t ®éng:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   17
      Left            =   3840
      TabIndex        =   68
      Tag             =   "Activies"
      Top             =   4830
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "B¹n cã thÓ dïng víi giíi h¹n 200 chøng tõ, møc doanh thu hai tû"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0017A0D4&
      Height          =   435
      Index           =   16
      Left            =   8280
      TabIndex        =   67
      Top             =   8760
      Width           =   8055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "NhÊn vµo ®©y  ®Ó t¹o c«ng ty míi trªn mµn h×nh"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0017A0D4&
      Height          =   285
      Index           =   0
      Left            =   9960
      TabIndex        =   65
      Top             =   8520
      Width           =   5415
   End
   Begin VB.Label txtdungthu 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   450
      Left            =   6600
      TabIndex        =   0
      Top             =   1560
      Width           =   9255
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "fax"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   14
      Left            =   7200
      TabIndex        =   64
      Top             =   3500
      Width           =   1215
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "N¨m tµi chÝnh"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   13
      Left            =   9000
      TabIndex        =   63
      Tag             =   "Financial Year"
      Top             =   3500
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   8400
      TabIndex        =   62
      Tag             =   "Bank VND Account"
      Top             =   3500
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "H¹ch to¸n theo:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   15
      Left            =   3840
      TabIndex        =   61
      Tag             =   "Tel"
      Top             =   4395
      Width           =   1695
   End
   Begin VB.Image img 
      Height          =   495
      Left            =   11760
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Phaàn meàm keá toaùn vietstar"
      BeginProperty Font 
         Name            =   "VNI-Lithos"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0017A0D4&
      Height          =   690
      Index           =   1
      Left            =   4920
      TabIndex        =   56
      Tag             =   "Accounting Software Company"
      Top             =   480
      Width           =   8415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "N¨m tµi chÝnh:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   10
      Left            =   6600
      TabIndex        =   55
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   9
      Left            =   6600
      TabIndex        =   54
      Top             =   3500
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TØnh thµnh:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   8
      Left            =   14280
      TabIndex        =   53
      Tag             =   "Province"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sè ®Þa ®iÓm kinh doanh:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   7
      Left            =   3840
      TabIndex        =   52
      Tag             =   "Activies"
      Top             =   5300
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "M· sè thuÕ:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   6
      Left            =   3840
      TabIndex        =   51
      Tag             =   "Tax Code"
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "§iÖn tho¹i:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   5
      Left            =   3840
      TabIndex        =   50
      Tag             =   "Tel"
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nh©n viªn triÓn khai"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   4
      Left            =   9840
      TabIndex        =   49
      Tag             =   "District"
      Top             =   8280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "§Þa chØ:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   3
      Left            =   3840
      TabIndex        =   48
      Tag             =   "Address"
      Top             =   2500
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "C«ng ty:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   2
      Left            =   10680
      TabIndex        =   47
      Tag             =   "Employee"
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tªn C«ng ty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   0
      Left            =   3960
      TabIndex        =   46
      Top             =   1920
      Width           =   10695
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "§Þa chØ"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   2
      Left            =   4800
      TabIndex        =   45
      Top             =   2520
      Width           =   9375
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tªn C«ng ty"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   10
      Left            =   5520
      TabIndex        =   44
      Top             =   4400
      Width           =   4215
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "®iÖn tho¹i"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   3
      Left            =   5280
      TabIndex        =   43
      Top             =   3500
      Width           =   1455
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "M· sè thuÕ"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   8
      Left            =   5280
      TabIndex        =   42
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TØnh thµnh"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   11
      Left            =   6120
      TabIndex        =   41
      Top             =   4830
      Width           =   7335
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Index           =   4
      Left            =   9840
      TabIndex        =   40
      Top             =   7800
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "N¨m tµi chÝnh"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   7
      Left            =   8280
      TabIndex        =   39
      Tag             =   "Financial Year"
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "§¬n vÞ triÓn khai: C«ng ty TNHH DV ThuÕ Sao ViÖt"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   12
      Left            =   9840
      TabIndex        =   38
      Top             =   7200
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "§c: 640 Tr­¬ng C«ng §Þnh, Tp Vòng Tµu, §t 090 3839 678"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   13
      Left            =   9840
      TabIndex        =   37
      Top             =   7560
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   14
      Left            =   9840
      TabIndex        =   36
      Top             =   7920
      Width           =   2295
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":42D99
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":440AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":453BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":466CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":479E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":48CF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":49945
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4B497
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4C149
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4D54B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4F8BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4FCAF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "§¬n vÞ"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   11
      Left            =   9240
      TabIndex        =   30
      Top             =   1800
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lbCty 
      Caption         =   "LbCty 9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   13440
      TabIndex        =   29
      Top             =   4920
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label 
      Caption         =   "§¬n vÞ ph¸t hµnh:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   13560
      TabIndex        =   28
      Top             =   4200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   16
      Left            =   12240
      TabIndex        =   27
      Top             =   2640
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lbCty 
      AutoSize        =   -1  'True
      Caption         =   "LbCty 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   13920
      TabIndex        =   14
      Tag             =   "0"
      Top             =   5520
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Th«ng tin doanh nghiÖp:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   1
      Left            =   3840
      TabIndex        =   8
      Tag             =   "This product Ý is licensed to"
      Top             =   1560
      Width           =   7920
   End
   Begin VB.Label email 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   10
      Left            =   6840
      TabIndex        =   9
      Top             =   6720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lbCty 
      Caption         =   "LbCty 12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   12
      Left            =   12000
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Lb 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   2
      Left            =   6600
      TabIndex        =   26
      Top             =   6000
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Lb 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   4920
      TabIndex        =   21
      Top             =   6960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Lb 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Index           =   0
      Left            =   3120
      TabIndex        =   20
      Tag             =   "Model"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ngµnh nghÒ"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   5880
      TabIndex        =   25
      Tag             =   "Profession"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Lo¹i h×nh DN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   5880
      TabIndex        =   19
      Tag             =   "Class"
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Phiªn b¶n"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   3240
      TabIndex        =   18
      Tag             =   "Version"
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   22
      Left            =   12960
      TabIndex        =   6
      Top             =   4320
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label 
      Caption         =   "TØnh, thµnh phè"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2.45745e5
      TabIndex        =   13
      Tag             =   "Province"
      Top             =   5700
      Width           =   1455
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "LbCty 6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   6480
      TabIndex        =   23
      Tag             =   "1"
      Top             =   5300
      Width           =   1935
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "LbCty 5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   5400
      TabIndex        =   22
      Tag             =   "1"
      Top             =   3975
      Width           =   7575
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tµi kho¶n Ngo¹i tÖ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   3480
      TabIndex        =   24
      Tag             =   "Bank F.C. Account"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sè tµi kho¶n:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   7
      Left            =   3840
      TabIndex        =   17
      Tag             =   "Bank VND Account"
      Top             =   3970
      Width           =   1335
   End
   Begin VB.Menu mnuHethong 
      Caption         =   "Th«ng sè"
      Tag             =   "&System"
      WindowList      =   -1  'True
      Begin VB.Menu mnHT 
         Caption         =   "&TÖp d÷ liÖu mÆc ®Þnh..."
         Index           =   2
         Tag             =   "Default data file"
      End
      Begin VB.Menu mnHT 
         Caption         =   "&NÐn tÖp d÷ liÖu..."
         Index           =   3
         Tag             =   "Compress data file..."
      End
      Begin VB.Menu mnHT 
         Caption         =   "Më tÖ&p d÷ liÖu nÐn..."
         Index           =   4
         Tag             =   "Open compressed data file"
      End
      Begin VB.Menu mnHT 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnHT 
         Caption         =   "F"
         Index           =   10
         Tag             =   "Font convert"
         Visible         =   0   'False
      End
      Begin VB.Menu mnHT 
         Caption         =   "Th«ng sè hÖ thèng"
         Index           =   11
         Tag             =   "Options"
      End
      Begin VB.Menu mnHT 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnHT 
         Caption         =   "&Danh s¸ch ng­êi sö dông"
         Index           =   13
         Tag             =   "User List"
      End
      Begin VB.Menu mnHT 
         Caption         =   "§Æt mËt &khÈu"
         Index           =   14
         Tag             =   "Change Password"
      End
      Begin VB.Menu mnHT 
         Caption         =   "-"
         Index           =   23
      End
      Begin VB.Menu mnHT 
         Caption         =   "§æi ng­êi sö dôn&g"
         Index           =   24
         Tag             =   "Log off"
      End
      Begin VB.Menu mnHT 
         Caption         =   "KÕt thóc c&h­¬ng tr×nh"
         Index           =   25
         Tag             =   "Quit"
      End
   End
   Begin VB.Menu mnDuLieu 
      Caption         =   "NhËp sè d­ ®Çu kú"
      Index           =   1
      Tag             =   "&Tools"
      Begin VB.Menu mnDL 
         Caption         =   "KiÓm tra &nhËp xuÊt tån"
         Index           =   0
         Tag             =   "Inventory Check-Up"
      End
      Begin VB.Menu mnDL 
         Caption         =   "KiÓm tra hÖ thèng &tµi kho¶n"
         Index           =   1
         Tag             =   "Account Check-Up"
      End
      Begin VB.Menu mnDL 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnDL 
         Caption         =   "Xö &lý sè liÖu..."
         Index           =   3
         Tag             =   "Run SQL Query..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnDL 
         Caption         =   "&Xo¸ ph¸t sinh th¸ng"
         Index           =   4
         Tag             =   "Delete data in month"
         Begin VB.Menu mnXoa 
            Caption         =   "Sè d­ ®Çu n¨m"
            Index           =   0
         End
         Begin VB.Menu mnXoa 
            Caption         =   "1"
            Index           =   1
         End
         Begin VB.Menu mnXoa 
            Caption         =   "2"
            Index           =   2
         End
         Begin VB.Menu mnXoa 
            Caption         =   "3"
            Index           =   3
         End
         Begin VB.Menu mnXoa 
            Caption         =   "4"
            Index           =   4
         End
         Begin VB.Menu mnXoa 
            Caption         =   "5"
            Index           =   5
         End
         Begin VB.Menu mnXoa 
            Caption         =   "6"
            Index           =   6
         End
         Begin VB.Menu mnXoa 
            Caption         =   "7"
            Index           =   7
         End
         Begin VB.Menu mnXoa 
            Caption         =   "8"
            Index           =   8
         End
         Begin VB.Menu mnXoa 
            Caption         =   "9"
            Index           =   9
         End
         Begin VB.Menu mnXoa 
            Caption         =   "10"
            Index           =   10
         End
         Begin VB.Menu mnXoa 
            Caption         =   "11"
            Index           =   11
         End
         Begin VB.Menu mnXoa 
            Caption         =   "12"
            Index           =   12
         End
      End
      Begin VB.Menu mnDL 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnDL 
         Caption         =   "&ChuyÓn sang n¨m míi"
         Index           =   6
         Tag             =   "Convert to new Financial Year"
      End
      Begin VB.Menu mnDL 
         Caption         =   "N¨&m tµi chÝnh"
         Index           =   7
         Tag             =   "Select Financial Year"
         Visible         =   0   'False
         Begin VB.Menu mnNam 
            Caption         =   "0"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "1"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "2"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "3"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "4"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "5"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "6"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "7"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "8"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "9"
            Index           =   9
         End
      End
      Begin VB.Menu mnDL 
         Caption         =   "Nguyªn tÖ..."
         Index           =   8
         Tag             =   "Posting Vouchers..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnDL 
         Caption         =   "Chøng tõ &kÕt chuyÓn..."
         Index           =   9
         Tag             =   "Posting Vouchers..."
      End
      Begin VB.Menu mnDL 
         Caption         =   "&Ph©n bæ chi phÝ..."
         Index           =   10
         Tag             =   "Expenses Allocation..."
      End
      Begin VB.Menu mnDL 
         Caption         =   "KÕt c&huyÓn sè liÖu..."
         Index           =   11
         Tag             =   "Monthly Conversion"
      End
      Begin VB.Menu mnDL 
         Caption         =   "Kh&o¸ sè liÖu th¸ng"
         Index           =   12
         Tag             =   "Clost data in month"
         Begin VB.Menu mnk 
            Caption         =   "Sè d­ ®Çu n¨m"
            Index           =   0
         End
         Begin VB.Menu mnk 
            Caption         =   "1"
            Index           =   1
         End
         Begin VB.Menu mnk 
            Caption         =   "2"
            Index           =   2
         End
         Begin VB.Menu mnk 
            Caption         =   "3"
            Index           =   3
         End
         Begin VB.Menu mnk 
            Caption         =   "4"
            Index           =   4
         End
         Begin VB.Menu mnk 
            Caption         =   "5"
            Index           =   5
         End
         Begin VB.Menu mnk 
            Caption         =   "6"
            Index           =   6
         End
         Begin VB.Menu mnk 
            Caption         =   "7"
            Index           =   7
         End
         Begin VB.Menu mnk 
            Caption         =   "8"
            Index           =   8
         End
         Begin VB.Menu mnk 
            Caption         =   "9"
            Index           =   9
         End
         Begin VB.Menu mnk 
            Caption         =   "10"
            Index           =   10
         End
         Begin VB.Menu mnk 
            Caption         =   "11"
            Index           =   11
         End
         Begin VB.Menu mnk 
            Caption         =   "12"
            Index           =   12
         End
      End
      Begin VB.Menu mnDL 
         Caption         =   "ChuyÓn d÷ liÖu ®Çu kú"
         Index           =   13
         Tag             =   "ChuyÓn d÷ liÖu ®Çu kú"
      End
      Begin VB.Menu mnDL 
         Caption         =   "Khai b¸o mÉu biÓu song ng÷"
         Index           =   14
         Tag             =   "Financial Report Description"
      End
   End
   Begin VB.Menu mnVatTu 
      Caption         =   "&VËt t­, hµng ho¸"
      Tag             =   "&Product and Contruction Cost"
      Begin VB.Menu mnVT 
         Caption         =   "&Ph©n lo¹i vËt t­..."
         Index           =   0
         Tag             =   "Classification..."
      End
      Begin VB.Menu mnVT 
         Caption         =   "Danh s¸ch vËt t­ hµng ho¸..."
         Index           =   1
         Tag             =   "Import-Export Source List..."
      End
      Begin VB.Menu mnVT 
         Caption         =   "&Kªnh ph©n phèi..."
         Index           =   2
         Tag             =   "Import-Export Source List..."
      End
      Begin VB.Menu mnVT 
         Caption         =   "L­ chuyÓn né bé..."
         Index           =   3
         Tag             =   "Import-Export Source List..."
      End
      Begin VB.Menu mnVT 
         Caption         =   "Thµnh phÈm hoµn thµnh trong kú..."
         Index           =   4
         Tag             =   "Import-Export Source List..."
      End
      Begin VB.Menu mnVT 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnVT 
         Caption         =   "Thªm kho hµng"
         Index           =   6
         Tag             =   "Opeining Balance"
      End
      Begin VB.Menu mnVT 
         Caption         =   "&Tån kho ®Çu kú..."
         Index           =   7
         Tag             =   "Opeining Balance"
      End
      Begin VB.Menu mnVT 
         Caption         =   "TÝnh l¹i gi¸ xuÊt kho trong th¸ng..."
         Index           =   8
         Tag             =   "Recalculate cost of material in month..."
      End
      Begin VB.Menu mnVT 
         Caption         =   "TÝnh gi¸ vèn hµng &b¸n..."
         Index           =   9
         Tag             =   "Recalculate cost of sold gooods"
      End
      Begin VB.Menu mnVT 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnVT 
         Caption         =   "KiÓm kª tån kho cuèi ngµy"
         Index           =   11
      End
      Begin VB.Menu mnVT 
         Caption         =   "Ph©n &lo¹i c«ng tr×nh, s¶n phÈm"
         Index           =   12
         Tag             =   "Classification of Product and Contructions"
      End
      Begin VB.Menu mnVT 
         Caption         =   "&Chi tiÕt c«ng tr×nh, s¶n phÈm"
         Index           =   13
         Tag             =   "Product and Contruction List"
      End
      Begin VB.Menu mnVT 
         Caption         =   "Tµi kho¶n &doanh thu"
         Index           =   14
         Tag             =   "Turnover Account of Finished Contructions"
      End
      Begin VB.Menu mnVT 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnVT 
         Caption         =   "§Æt/Bá TK theo dâi chi tiÕt"
         Index           =   16
         Tag             =   "Set Account"
      End
      Begin VB.Menu mnVT 
         Caption         =   "-"
         Index           =   17
         Visible         =   0   'False
      End
      Begin VB.Menu mnVT 
         Caption         =   "Tõ kho¸ theo tªn hµng"
         Index           =   18
      End
      Begin VB.Menu mnVT 
         Caption         =   "Gi¸ vèn hµng nhËp khÈu"
         Index           =   19
         Visible         =   0   'False
      End
      Begin VB.Menu mnVT 
         Caption         =   "Tù ®éng nhËp kho"
         Index           =   21
         Begin VB.Menu mnVTdt 
            Caption         =   "Theo ngµy"
            Index           =   1
         End
         Begin VB.Menu mnVTdt 
            Caption         =   "Theo th¸ng"
            Index           =   2
         End
      End
      Begin VB.Menu mnVT 
         Caption         =   "Xo¸ tªn vËt t­ thõa"
         Index           =   22
      End
   End
   Begin VB.Menu mnCongno 
      Caption         =   "C«n&g nî"
      Tag             =   "&Liability"
      Begin VB.Menu mnCN 
         Caption         =   "&Ph©n lo¹i"
         Index           =   0
         Tag             =   "Classification..."
      End
      Begin VB.Menu mnCN 
         Caption         =   "&Danh s¸ch"
         Index           =   1
         Tag             =   "Items"
      End
      Begin VB.Menu mnCN 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnCN 
         Caption         =   "&Sè d­ ®Çu kú"
         Index           =   3
         Tag             =   "Opening Balance"
      End
      Begin VB.Menu mnCN 
         Caption         =   "Danh s¸ch &Hîp ®ång"
         Index           =   4
         Tag             =   "Contract List"
      End
      Begin VB.Menu mnCN 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnCN 
         Caption         =   "Ph©n lo¹i &nh©n viªn b¸n hµng"
         Index           =   6
         Tag             =   "Salesman Classification"
      End
      Begin VB.Menu mnCN 
         Caption         =   "Danh s¸ch nh©n &viªn b¸n hµng"
         Index           =   7
         Tag             =   "Salesman List"
      End
      Begin VB.Menu mnCN 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnCN 
         Caption         =   "§Æt/Bá TK theo dâi chi tiÕt"
         Index           =   11
         Tag             =   "Set Account"
      End
      Begin VB.Menu mnCN 
         Caption         =   "Xo¸ tªn c«ng nî thõa"
         Index           =   12
      End
   End
   Begin VB.Menu mnTSCD 
      Caption         =   "Tµi &s¶n cè ®Þnh"
      Index           =   12
      Tag             =   "Fixed &Assets"
      Begin VB.Menu mnTS 
         Caption         =   "Ph©n lo¹i &tµi s¶n..."
         Index           =   0
         Tag             =   "Classification of Assets..."
      End
      Begin VB.Menu mnTS 
         Caption         =   "Ph©n lo¹i &chøng tõ..."
         Index           =   1
         Tag             =   "Classification of Voucher..."
      End
      Begin VB.Menu mnTS 
         Caption         =   "Danh s¸ch TSCD..."
         Index           =   2
         Tag             =   "Classification of Voucher..."
      End
      Begin VB.Menu mnTS 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnTS 
         Caption         =   "&N­íc s¶n xuÊt..."
         Index           =   4
         Tag             =   "Country List..."
      End
      Begin VB.Menu mnTS 
         Caption         =   "T×nh tr¹ng &sö dông..."
         Index           =   5
         Tag             =   "Conjuncture List..."
      End
      Begin VB.Menu mnTS 
         Caption         =   "§èi t­îng &qu¶n lý..."
         Index           =   6
         Tag             =   "Administrative Object..."
      End
      Begin VB.Menu mnTS 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnTS 
         Caption         =   "Tµi s¶n ®Çu &kú..."
         Index           =   8
         Tag             =   "Opening Balance..."
      End
      Begin VB.Menu mnTS 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnTS 
         Caption         =   "§Æt/Bá TK chi phÝ khÊu hao"
         Index           =   10
         Tag             =   "Set Account"
      End
      Begin VB.Menu mnTS 
         Caption         =   "Danh s¸ch tµi s¶n cè ®Þnh"
         Index           =   11
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Trî &gióp"
      Tag             =   "&Help"
      Begin VB.Menu mnuHLP 
         Caption         =   "&Giíi thiÖu"
         Index           =   1
         Tag             =   "&About"
      End
      Begin VB.Menu mnuHLP 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuHLP 
         Caption         =   "&Tµi liÖu..."
         Index           =   3
         Tag             =   "&Directory..."
      End
      Begin VB.Menu mnuHLP 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuHLP 
         Caption         =   "&T¹o c«ng ty míi"
         Index           =   5
         Tag             =   "&New"
      End
   End
   Begin VB.Menu mnversion 
      Caption         =   "Version"
      Begin VB.Menu mnversionChild 
         Caption         =   "Ver1.1"
         Index           =   1
      End
   End
   Begin VB.Menu mnTT 
      Caption         =   "TT"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateToolhelp32Snapshot Lib "Kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "Kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "Kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private Declare Function TerminateProcess Lib "Kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Const PROCESS_TERMINATE = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const MAX_PATH = 260
Private hasUPdate As Boolean
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private m_lngLeft As Long
Private m_lngTop As Long
Private m_lngWidth As Long
Private m_lngHeight As Long
Private Const SW_SHOW = 5
Private Const SW_HIDE = 0
Private Const SW_RESTORE = 9
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205
Private Declare Function GetTickCount Lib "Kernel32" () As Long

Private Type versionInfo
    Version As String
    fileId As String
    fileName As String
    ReleaseDate As String
    Description As String
End Type
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
                                           (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
                                            ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long


Private Type MenuInfo
    id As Long
    Name As String
End Type
Private arrItems() As MenuInfo
Private mnItemCount As Integer

Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" _
                                       (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, _
                                        ByVal nMaxCount As Long, ByVal wFlag As Long) As Long

Private Const MF_BYCOMMAND As Long = &H0&

Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" ( _
                                          ByVal dwMessage As Long, _
                                          lpData As NOTIFYICONDATA) As Long

Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private LastVersion As String
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Dim nid As NOTIFYICONDATA

Public typeprint As Integer
Private g_NcmBackup As NONCLIENTMETRICS
Private g_HasBackup As Boolean
Public keyhasregistry As String
Private Declare Function MulDiv Lib "Kernel32" ( _
                                ByVal nNumber As Long, _
                                ByVal nNumerator As Long, _
                                ByVal nDenominator As Long) As Long
Private Const SPIF_UPDATEINIFILE As Long = &H1
Private Const SPIF_SENDCHANGE As Long = &H2
Private Const SPIF_SENDWININICHANGE As Long = &H2   ' Ð?ng nghia
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSY = 90   ' Thêm cái này
Private Const SPI_GETNONCLIENTMETRICS = 41
Private Const SPI_SETNONCLIENTMETRICS = 42
Private Const WM_SETTEXT = &HC
Private Const WM_SETTINGCHANGE As Long = &H1A
Private Const HWND_BROADCAST As Long = &HFFFF&
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 32
End Type

Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSmCaptionWidth As Long
    iSmCaptionHeight As Long
    lfSmCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function DefWindowProcW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long


'Module tieng viet
Dim MenuCount As Integer
Dim CurrentMenu As Integer
Dim itemCount As Integer
Dim fontmenu As String, fontcaption As String, ncm As NONCLIENTMETRICS
Attribute fontcaption.VB_VarUserMemId = 1073938448
Attribute ncm.VB_VarUserMemId = 1073938448
'end Module tieng viet
Public gSubMenu As frmSubMenu
Attribute gSubMenu.VB_VarUserMemId = 1073938438
Public gCurrentMenu As Integer
Attribute gCurrentMenu.VB_VarUserMemId = 1073938439

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function GetAdaptersInfo Lib "iphlpapi" (lpAdapterInfo As Any, lpSize As Long) As Long

Public Tudongtinhgiavon As Boolean
Attribute Tudongtinhgiavon.VB_VarUserMemId = 1073938440

Private Declare Function GetModuleFileName Lib "Kernel32" Alias "GetModuleFileNameA" _
                                           (ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Private Declare Function FindFirstFile Lib "Kernel32" Alias "FindFirstFileA" _
                                       (ByVal lpFilename As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindNextFile Lib "Kernel32" Alias "FindNextFileA" _
                                      (ByVal hFindFile As Long, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function FindClose Lib "Kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                                      (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
                                       ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As Currency
    ftLastAccessTime As Currency
    ftLastWriteTime As Currency
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternateFileName As String * 14
End Type


Dim ret
Attribute ret.VB_VarUserMemId = 1073938441
Dim m_nonClientMetrics As NONCLIENTMETRICS
Attribute m_nonClientMetrics.VB_VarUserMemId = 1073938442

Dim m_fontCaption As String * 32
Attribute m_fontCaption.VB_VarUserMemId = 1073938443

Dim m_fontSmCaption As String * 32
Attribute m_fontSmCaption.VB_VarUserMemId = 1073938444
Dim m_fontMenu As String * 32
Attribute m_fontMenu.VB_VarUserMemId = 1073938445
Dim m_fontMessage As String * 32
Attribute m_fontMessage.VB_VarUserMemId = 1073938446
Dim m_fontStatus As String * 32
Attribute m_fontStatus.VB_VarUserMemId = 1073938447
Dim m_fontIcon As String * 32
Attribute m_fontIcon.VB_VarUserMemId = 1073938448
Dim pProcessEnable As Boolean
Attribute pProcessEnable.VB_VarUserMemId = 1073938449

Private Const MaxNamTC = 9
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)


Private Declare Function LDBUser_GetUsers Lib "MSLDBUSR.DLL" (lpszUserBuffer() As String, ByVal lpszFilename As String, ByVal nOptions As Long) As Integer

Private Const OptLDBLoggedUsers = &H2
'Sub chutieng_viet()
'Dim MyUnicodeText
' Set MyUnicodeText = New Class1
'        ' Read Unicode Text from file txtFileName and display in TextBox1(0)
'       'TextBox1(0).Text = MyUnicodeText.ReadUnicode(txtFileName)
'         UVowels = mDOMVowels.ReadUnicode(GetLocalDirectory & "UnicodeVowels.xml")
'       LbCongty.Caption = MyUnicodeText.ReadUnicode("D:\soft\sv\Accounting\config.xml")
'End Sub
Private Sub AddItem(id As Integer, Name As String)
    mnItemCount = mnItemCount + 1
    ReDim Preserve arrItems(1 To mnItemCount)
    arrItems(mnItemCount).id = id
    arrItems(mnItemCount).Name = Name
End Sub
Private Function FindItemByID(id As Integer) As String
    Dim i As Integer
    For i = 1 To mnItemCount
        If arrItems(i).id = id Then
            FindItemByID = arrItems(i).Name
            Exit Function
        End If
    Next i
    FindItemByID = "Not Found"
End Function

Private Sub AddTray()
    m_lngLeft = Me.Left
    m_lngTop = Me.Top
    m_lngWidth = Me.Width
    m_lngHeight = Me.Height
    nid.cbSize = Len(nid)
    nid.hwnd = Me.hwnd
    nid.uID = 1
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallbackMessage = WM_MOUSEMOVE
    nid.hIcon = Me.Icon.Handle
    nid.szTip = "?ng d?ng VB6" & Chr$(0)

    Shell_NotifyIcon NIM_ADD, nid

End Sub
Private Sub Command_Click(Index As Integer)
    frmMain.typeprint = 0
    Select Case Index
    Case 0:

        'FrmChungtu.Autonhapkho
        If User_Right = 2 Then
            NoRight 0
            Exit Sub
        End If
        If pSTOP = 1 Then
            MsgBox VString(pTenCty), vbCritical, App.ProductName
            Exit Sub
        End If
        pPhieu = 0
        ' frmTaiLieu.Show 1
        FrmChungtu.Show vbModal
        Set FrmChungtu = Nothing
    Case 1:
        If User_Right = 1 Then
            NoRight 0
            Exit Sub
        End If
        FBcKt.Show 1
    Case 2:
        If User_Right = 1 Then
            NoRight 0
            Exit Sub
        End If
        FBcTC.Show 1
    Case 3:
        RunCT
    Case 4:
        If User_Right = 2 Then
            NoRight 0
            Exit Sub
        End If
        pPhieu = 1
        FrmChungtu.Show 1
        Set FrmChungtu = Nothing
    Case 5:
        If User_Right = 1 Then
            NoRight 0
            Exit Sub
        End If
        FrmBCQT.Show 1
    Case 6:
        FrmCDT.Show 1
    End Select
    HienThongBao "", 1
End Sub
Public Function GetMenuCaption(ByVal hMenu As Long, ByVal MenuID As Long) As String
    Dim buffer As String
    Dim length As Long

    ' T?o buffer d? ch?a caption
    buffer = Space$(256)

    ' G?i API d? l?y caption
    length = GetMenuString(hMenu, MenuID, buffer, Len(buffer), MF_BYCOMMAND)

    ' C?t chu?i theo d? dài th?c t?
    If length > 0 Then
        GetMenuCaption = Left$(buffer, length)
    Else
        GetMenuCaption = ""
    End If
End Function
Private Sub CommandButton1_Click()

End Sub

Private Sub Command_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Index = 0 Then
    Command(0).BackColor = 8438015
    Command(1).BackColor = &HC000&
    Command(2).BackColor = &HC000&
    ElseIf Index = 1 Then
    Command(1).BackColor = 8438015
    Command(0).BackColor = &HC000&
    Command(2).BackColor = &HC000&
    ElseIf Index = 2 Then
    Command(2).BackColor = 8438015
    Command(1).BackColor = &HC000&
    Command(3).BackColor = &HC000&
    
  End If
End Sub

Private Sub Command1_Click()
Dim fso As New FileSystemObject
      '  MsgBox pCurDir + "DATA"
      Dim duong_dan As String
     ' duong_dan = Mid(pCurDir, 1, Len(pCurDir) - 1) + CStr(Minute(Now)) + CStr(Second(Now))
      
     duong_dan = Mid(Mid(pCurDir, 1, Len(pCurDir) - 1), 1, InStrRev(Mid(pCurDir, 1, Len(pCurDir) - 1), "\")) + "VietStar_" + CStr(Minute(Now)) + CStr(Second(Now))
     
     MkDir duong_dan
      MkDir duong_dan + "\data"
     If Len(Dir(pCurDir + "REPORTS\QD48.MDB")) = 0 Then
      fso.CopyFile pCurDir + "REPORTS\QD15.MDB", duong_dan + "\Data\QD15.mdb", True
     Else
      fso.CopyFile pCurDir + "REPORTS\QD48.MDB", duong_dan + "\Data\QD48.mdb", True
     End If
     
      fso.CopyFolder pCurDir + "REPORTS", duong_dan + "\REPORTS"
      fso.CopyFolder pCurDir + "Tailieu", duong_dan + "\Tailieu"
      fso.CopyFile pCurDir + "Dummy.xml", duong_dan + "\Dummy.xml"
      fso.CopyFile pCurDir + "UnicodeVowels.xml", duong_dan + "\UnicodeVowels.xml"
      fso.CopyFile pCurDir + "VNIVowelMap.txt", duong_dan + "\VNIVowelMap.txt"
    '  fso.CopyFile pCurDir + "VietStar.exe", duong_dan + "\VietStar.exe"
      
      fso.CopyFile duong_dan + "\REPORTS\vietstar.exe", duong_dan + "\VietStar.exe"
      CreateShortCut duong_dan + "\VietStar.exe", "VietStar_" + CStr(Minute(Now)) + CStr(Second(Now))
      MsgBox "B¹n ®· t¹o míi thµnh c«ng, icon ®· cã ngoµi mµn h×nh:" & vbNewLine & duong_dan
      Shell "EXPLORER.EXE " & duong_dan + "\VietStar.exe"
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = 8438015
End Sub

Private Sub Command2_Click()
    Taifilecapnhat
End Sub

Private Sub Command3_Click()
    ftmtest.Show vbModal
End Sub

Private Sub CTTimer_Timer()
    If pProcessEnable Then
        pProcessEnable = False
        XuLyChungtu
        pProcessEnable = True
    End If
End Sub

Private Sub email_Click(Index As Integer)
    Select Case Index
        Case 0:
            ShellExecute hwnd, "open", "mailto:" + email(Index).Caption, vbNullString, vbNullString, 0
        Case 1:
            ShellExecute ByVal 0&, "open", email(Index).Caption, vbNullString, vbNullString, 3
        Case 2:
            ShellExecute hwnd, "open", "ypager.exe", vbNullString, Left(pWinDir, 2) + "\Program Files\Yahoo!\Messenger", 1
    End Select
End Sub

Private Sub File1_Click()
'hung
End Sub
 

Private Function ParseJson(json As String, key As String) As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim parts() As String
    parts = Split(json, ",")

    Dim part As Variant
    For Each part In parts
        Dim kv() As String
        kv = Split(part, ":")
        If UBound(kv) = 1 Then
            dict(Trim(Replace(kv(0), """", ""))) = Trim(Replace(kv(1), """", ""))
        End If
    Next part

    ParseJson = dict(key)
End Function

Public Function GetMacAddress() As String
    Const OFFSET_LENGTH As Long = 400
    Dim lSize As Long
    Dim baBuffer() As Byte
    Dim lIdx As Long
    Dim sRetVal As String

    Call GetAdaptersInfo(ByVal 0, lSize)
    If lSize <> 0 Then
        ReDim baBuffer(0 To lSize - 1) As Byte
        Call GetAdaptersInfo(baBuffer(0), lSize)
        Call CopyMemory(lSize, baBuffer(OFFSET_LENGTH), 4)
        For lIdx = OFFSET_LENGTH + 4 To OFFSET_LENGTH + 4 + lSize - 1
            sRetVal = IIf(LenB(sRetVal) <> 0, sRetVal & ":", vbNullString) & Right$("0" & Hex$(baBuffer(lIdx)), 2)
        Next
    End If
    GetMacAddress = sRetVal
End Function


Private Sub AuToNhapTP()

    Dim rs_tonkho As Recordset
    Dim X As Integer
    Dim Query As String
    Dim monthCurrent As Integer
    monthCurrent = month(Date)    ' L?y tháng hi?n t?i
    X = monthCurrent


    Query = "SELECT * FROM TonKho WHERE MaTaiKhoan = 38 AND Luong_" & X & " < 0"

    Set rs_tonkho = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)

    If Not rs_tonkho.EOF Then
        'tim hoa don cu va phan loai theo ngay
        MsgBox rs_tonkho!Luong_8 & "" & rs_tonkho!Luong_9

        ' Di chuy?n d?n b?n ghi ti?p theo
        rs_tonkho.MoveNext
    End If
End Sub
Private Sub UpdateIIS()

    On Error GoTo ErrorHandler

    ' ===== TRY =====
    Screen.MousePointer = vbHourglass
    DoEvents

    Dim path As String
    path = Environ("PUBLIC") & "\SaoVietPublic"

    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If

    Dim fso As Object
    Dim sourcePath As String
    Dim destPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    sourcePath = "\\192.168.1.90\Ke toan 2025 New\2 Copi vao dung 4\SaovietITax\Public"
    destPath = Environ("PUBLIC") & "\SaoVietPublic"

    If Dir(destPath, vbDirectory) = "" Then MkDir destPath
    ProgressBar1.Value = 20
    DoEvents
    ' Copy file
    fso.CopyFile sourcePath & "\*.*", destPath & "\", True
    ProgressBar1.Value = 30
    DoEvents
    ' Copy folder
    fso.CopyFolder sourcePath & "\*", destPath & "\", True
    ProgressBar1.Value = 50
    DoEvents

CleanExit:
    ' ===== FINALLY =====
    Screen.MousePointer = vbDefault
    If Not fso Is Nothing Then Set fso = Nothing
    Exit Sub

ErrorHandler:
    ' ===== CATCH =====
    MsgBox "L?i Update IIS: " & Err.Description, vbCritical

    ' Ghi log (r?t nên có)
    Open App.path & "\log_update.txt" For Append As #1
    Print #1, Now & " - " & Err.number & " - " & Err.Description
    Close #1

    Resume CleanExit
End Sub
Public Sub Taifilecapnhat()
    ProgressBar1.Value = 10
    DoEvents
    UpdateIIS

    Dim originPaths As String
    originPaths = App.path
    Dim serverpath As String
    serverpath = originPaths & "\Hoadon\serverpath.txt"
    Dim uncPath As String
    uncPath = ReadTxt(serverpath)


    Dim sourceFile As String
    Dim DestFolder As String
    Dim destFile As String
    sourceFile = uncPath & "\update.exe"
    DestFolder = App.path & "\Hoadon"
    ' Ðu?ng d?n file dích d?y d?
    destFile = DestFolder & "\update.exe"
    On Error GoTo ErrorHandler

    ' T?o thu m?c Hoadon n?u chua t?n t?i
    If Dir(DestFolder, vbDirectory) = "" Then
        MkDir DestFolder
    End If

    ' Copy file t? server v? (n?u file ngu?n t?n t?i và khác v?i file dích)
    If Dir(sourceFile) <> "" Then
        ' Ch? copy n?u file dích chua t?n t?i ho?c khác kích thu?c/ngày gi?
        If Dir(destFile) = "" Then
            FileCopy sourceFile, destFile
            MsgBox "Ðã t?i update.exe v? thu m?c Hoadon thành công!", vbInformation
        Else
            ' So sánh kích thu?c d? tránh copy th?a (tùy ch?n nâng cao)
            If 1 < 2 Then
                Kill destFile  ' Xóa file cu tru?c khi copy dè
                FileCopy sourceFile, destFile
                ' MsgBox "Ðã c?p nh?t file update.exe m?i!", vbInformation
            Else
                MsgBox "File update.exe dã là phiên b?n m?i nh?t.", vbInformation  ' Có th? b? n?u không mu?n thông báo
            End If
        End If
    Else
        MsgBox "Không tìm th?y file update.exe trên server!" & vbCrLf & sourceFile, vbExclamation
    End If

    ' Ti?p t?c ch?y chuong trình bình thu?ng
    ' === CH?Y FILE UPDATE.EXE SAU KHI T?I XONG ===
    Dim i As Integer

    For i = 50 To 100
        ProgressBar1.Value = i
        DoEvents
    Next i
    Dim Result As Long
    Result = ShellExecute(0, "open", destFile, "", DestFolder, 0)
ErrorHandler:
    'MsgBox "L?i khi t?i file update.exe:" & vbCrLf & Err.Description, vbCritical
    ProgressBar1.Value = 100
    DoEvents
End Sub



Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Ð?i con tr? khi hover
    Image2.MousePointer = vbSizeNS
    ' Load icon t? file .ico ho?c .cur
End Sub


Public Sub SetMenuUnicode(hMenu As Long, pos As Long, textU As String)
    Dim mii As MENUITEMINFOW
    Dim s As String

    s = textU & vbNullChar

    With mii
        .cbSize = Len(mii)
        .fMask = MIIM_STRING Or MIIM_FTYPE
        .fType = MFT_STRING
        .dwTypeData = StrPtr(s)
        .cch = Len(textU)
    End With

    SetMenuItemInfoW hMenu, pos, 1, mii
End Sub
Public Sub Chayngam()
    Static FirstRun As Boolean

    If Not FirstRun Then
         ''FirstRun = True
          AddTray
          Me.Hide
         Command_Click 0
    End If
End Sub
Public Sub ChayTb()
    'timerChaytb.Enabled = True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
        Case vbKeyN:
            RFocus Command(0)
            Command_Click 0
        Case vbKeyC:
            RFocus Command(1)
            Command_Click 1
        Case vbKeyT:
            RFocus Command(2)
            Command_Click 2
        End Select
    End If

    If (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyQ Then XoaQuery

    If (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyD Then
        ChDir pCurDir + "DATA"
        Recycle "K*" + "_" + CStr(LbCty(0).tag) + ".SAS"
    End If

    If (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyF Then
        FontSetUp
        pKhongDau = 1 - pKhongDau
        SetFont Me
        If pKhongDau = 1 Then
            Label(14).Caption = ABCtoKDau(Label(14).Caption)
            Label(26).Caption = ABCtoKDau(Label(26).Caption)
        End If
    End If

    If (Shift And vbAltMask) > 0 And (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyR Then
        If MsgBox("Xo¸ tÊt c¶ Relations?", vbYesNo + vbCritical, App.ProductName) = vbYes Then DeleteRel
    End If

    If (Shift And vbCtrlMask) > 0 And img.Picture <> 0 And pVersion = 1 Then
        Select Case KeyCode
        Case vbKeyLeft: img.Left = img.Left - 10
        Case vbKeyRight: img.Left = img.Left + 10
        Case vbKeyUp: img.Top = img.Top - 10
        Case vbKeyDown: img.Top = img.Top + 10
        End Select
    End If

    If (Shift And vbShiftMask) > 0 And img.Picture <> 0 And pVersion = 1 Then
        Select Case KeyCode
        Case vbKeyLeft: img.Width = img.Width - 10
        Case vbKeyRight: img.Width = img.Width + 10
        Case vbKeyUp: img.Height = img.Height - 10
        Case vbKeyDown: img.Height = img.Height + 10
        End Select
    End If
    Dim rs As Recordset
    If (Shift And vbAltMask) > 0 And (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyK Then
        ExecuteSQL5 ("Update tbLicensekey set Type=1, Year =2 ,Totals =1000000")
    End If
  
    If (Shift And vbAltMask) > 0 And (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyU Then
        Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW License.* FROM License", dbOpenSnapshot)
        'ExecuteSQL5 "update License set TenCty = '" + ModSAS.Federo16(rs!TenCty, CStr(rs!NamTC)) + "',DiaChi = '" + ModSAS.Federo16(rs!DiaChi, CStr(rs!NamTC)) + "',MaSoThue = '" + ModSAS.Federo16(rs!masothue, CStr(rs!NamTC)) + "',CMP = '" + ModSAS.Federo16(IIf(IsNull(rs!CMP), "", rs!CMP), CStr(rs!NamTC)) + "'"
        Dim ma_so_so As String
        'ma_so_so = ModSAS.Federo16Decrypt("dad`dccefucgcqcici", opotion_1)
        'ma_so_so = "1@35^7*9)1"
        'SetPsw pDataPath, pPSW, ma_so_so
        'SaveSetting "MyApp", "Settings", "FirstRun", "False"

        ExecuteSQL5 ("Update tbLicensekey set Type=-1")
        ExecuteSQL5 ("Update License set CMG=249991")
        WSpace.Close
        End
    End If
    '    If (Shift And vbAltMask) > 0 And (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyO Then
    '            SetPsw pDataPath, pPSW, "unlock"
    '            WSpace.Close
    '            End
    '    End If
End Sub
Private Function GetFontName(fontFace As String) As String
    On Error Resume Next
    Dim nullPos As Long
    nullPos = InStr(fontFace, vbNullChar)
    If nullPos > 0 Then
        GetFontName = Left$(fontFace, nullPos - 1)
    Else
        GetFontName = fontFace
    End If
End Function
Private Function ReadTxt(FilePath As String) As String
    Dim fileNumber As Integer
    Dim content As String

    On Error GoTo ErrorHandler

    ' Ki?m tra file t?n t?i
    If Dir(FilePath) = "" Then
        ReadTxt = "File không t?n t?i: " & FilePath
        Exit Function
    End If

    ' L?y file number
    fileNumber = FreeFile

    ' M? file d? d?c
    Open FilePath For Input As #fileNumber

    ' Ð?c toàn b? n?i dung
    If LOF(fileNumber) > 0 Then
        content = Input$(LOF(fileNumber), #fileNumber)
    End If

    ' Ðóng file
    Close #fileNumber

    ' Tr? v? n?i dung
    ReadTxt = content
    Exit Function

ErrorHandler:
    ReadTxt = "L?i d?c file: " & Err.Description
    On Error Resume Next
    Close #fileNumber
End Function
Private Sub Kiemtraphienbanht()

    Dim originPaths As String
    originPaths = App.path
    Dim serverpath As String

    Dim txtPath2 As String
    txtPath2 = originPaths & "\" & "Hoadon\version.txt"
    Dim content3 As String
    content3 = ReadTxt(txtPath2)
    mnversion.Caption = "Version " & content3
    Dim v1 As Double
    Dim v2 As Double

    v1 = CDbl(content3)
    If LastVersion = "" Then
        Exit Sub
    End If
    v2 = CDbl(LastVersion)

    If v1 = v2 Then
        Image2.Visible = False
    Else
        Image2.Visible = True
    End If
    Exit Sub
    serverpath = originPaths & "\Hoadon\serverpath.txt"

    Dim uncPath As String
    uncPath = ReadTxt(serverpath)
    If InStr(1, uncPath, "Versions", vbTextCompare) > 0 Then
        txtPath2 = originPaths & "\" & "Hoadon\version.txt"
        content3 = ReadTxt(txtPath2)
        mnversion.Caption = "Version " & content3
        'Image2.Visible = True
        Exit Sub
    End If

    Dim txtPath As String
    txtPath = uncPath & "\" & "Tools\version.txt"
    Dim content As String
    content = ReadTxt(txtPath)
    mnversion.Caption = "Version " & content

    Dim originPath As String
    originPath = App.path & "\Hoadon\version.txt"
    Dim content2 As String
    content2 = ReadTxt(originPath)
    'Label2.Caption = content2
    'Image2.Visible = False
    If content <> content2 Then
        'Image2.Visible = True
        ' Else
        'Image2.Visible = False
        'Image2_Click

    End If
End Sub



Private Sub StringToByteArray(ByVal str As String, byteArray() As Byte)
    Dim i As Long
    Dim maxLen As Long

    maxLen = UBound(byteArray)
    For i = 0 To maxLen
        byteArray(i) = 0
    Next i

    For i = 1 To Len(str)
        If i - 1 > maxLen Then Exit For
        byteArray(i - 1) = Asc(Mid$(str, i, 1))
    Next i
    byteArray(i) = 0    ' Null terminator
End Sub
Private Sub SetFontNameSimple(faceName() As Byte, ByVal FontName As String)
    Dim i As Long
    For i = 0 To 31
        faceName(i) = 0
    Next i

    For i = 1 To Len(FontName)
        If i > 32 Then Exit For
        faceName(i - 1) = Asc(Mid$(FontName, i, 1))
    Next i
End Sub

Private Sub ChangeFont()
    If findwindowpartial("Microsoft Word") = 0 And findwindowpartial("Microsoft Excel") = 0 Then

        DoEvents

        m_nonClientMetrics.cbSize = Len(m_nonClientMetrics)
        ret = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, Len(m_nonClientMetrics), m_nonClientMetrics, 0)

        ' FIX L?I L?N 2 + FIX CH? NH?: dùng Trim$ thay vì InStr
        m_fontCaption = Trim$(m_nonClientMetrics.lfCaptionFont.lfFaceName)
        m_fontSmCaption = Trim$(m_nonClientMetrics.lfSmCaptionFont.lfFaceName)
        m_fontMenu = Trim$(m_nonClientMetrics.lfMenuFont.lfFaceName)
        m_fontStatus = Trim$(m_nonClientMetrics.lfStatusFont.lfFaceName)

        Dim sNewFont As String
        sNewFont = sFONTNAME & vbNullChar

        ' Ch? thay 4 ph?n an toàn nh?t, gi? nguyên height ? ch? không b? nh?
        m_nonClientMetrics.lfCaptionFont.lfFaceName = sNewFont
        m_nonClientMetrics.lfSmCaptionFont.lfFaceName = sNewFont
        m_nonClientMetrics.lfMenuFont.lfFaceName = sNewFont
        m_nonClientMetrics.lfStatusFont.lfFaceName = sNewFont

        ' Áp d?ng thay d?i
    ret = SystemParametersInfo(SPI_SETNONCLIENTMETRICS, Len(m_nonClientMetrics), m_nonClientMetrics, 0)

    ' Thêm do?n code sau d? thay d?i font ch? c?a menu
    Me.FontName = sFONTNAME
    Me.FontSize = 12

    DoEvents
End If
End Sub
Public Function U(ParamArray Codes() As Variant) As String
    Dim i As Long
    For i = LBound(Codes) To UBound(Codes)
        U = U & ChrW(Codes(i))
    Next
End Function
Private Sub LoadMenuform()

    
    Dim hMenu As Long
    Dim hSub As Long
    Dim hSub2 As Long
    Dim hSub3 As Long
    Dim hDuLieu As Long


    hMenu = GetMenu(Me.hwnd)
    ' =====================================================
    ' MENU C?P 1
    ' =====================================================
    SetMenuUnicode hMenu, 0, U(84, 104, 244, 110, 103, 32, 115, 7889)                ' Thông s?
    'SetMenuUnicode hMenu, 1, U(78, 104, 7853, 112, 32, 115, 7889, 32, 100, 432, 32, 273, 7847, 117, 32, 107, 7923)    ' Nh?p s? du d?u k?
    SetMenuUnicode hMenu, 1, U(88, 7917, 32, 108, 253)


    SetMenuUnicode hMenu, 2, U(86, 7853, 116, 32, 116, 432, 44, 32, 104, 224, 110, 103, 32, 104, 243, 97)    ' V?t tu, hàng hóa
    SetMenuUnicode hMenu, 3, U(67, 244, 110, 103, 32, 110, 7907)                     ' Công n?
    SetMenuUnicode hMenu, 4, U(84, 224, 105, 32, 115, 7843, 110, 32, 99, 7889, 32, 273, 7883, 110, 104)    ' Tài s?n c? d?nh
    SetMenuUnicode hMenu, 5, U(84, 114, 7907, 32, 103, 105, 250, 112)                ' Tr? giúp

    ' =====================================================
    ' THÔNG So sub
    ' =====================================================
    hSub = GetSubMenu(hMenu, 0)

    'SetMenuUnicode mnHT(2), 0, U(84, 7879, 112, 32, 100, 7919, 32, 108, 105, 7879, 117, 32, 109, 7863, 99, 32, 273, 7883, 110, 104)
    SetMenuUnicode hSub, 0, U(84, 7879, 112, 32, 100, 7919, 32, 108, 105, 7879, 117, 32, 109, 7863, 99, 32, 273, 7883, 110, 104)
    SetMenuUnicode hSub, 1, U(78, 233, 110, 32, 116, 7879, 112, 32, 100, 7919, 32, 108, 105, 7879, 117)
    SetMenuUnicode hSub, 2, U(77, 7903, 32, 116, 7879, 112, 32, 100, 7919, 32, 108, 105, 7879, 117, 32, 110, 233, 110)
    SetMenuUnicode hSub, 4, U(84, 104, 244, 110, 103, 32, 115, 7889, 32, 104, 7879, 32, 116, 104, 7889, 110, 103)
    SetMenuUnicode hSub, 6, U(68, 97, 110, 104, 32, 115, 225, 99, 104, 32, 110, 103, 432, 7901, 105, 32, 115, 7917, 32, 100, 7909, 110, 103)
    SetMenuUnicode hSub, 7, U(272, 7863, 116, 32, 109, 7853, 116, 32, 107, 104, 7849, 117)
    SetMenuUnicode hSub, 9, U(272, 7893, 105, 32, 110, 103, 432, 7901, 105, 32, 115, 7917, 32, 100, 7909, 110, 103)
    SetMenuUnicode hSub, 10, U(75, 7871, 116, 32, 116, 104, 250, 99, 32, 99, 104, 432, 417, 110, 103, 32, 116, 114, 236, 110, 104)

    ' =====================================================
    ' Xu ly sub
    ' =====================================================
    hSub = GetSubMenu(hMenu, 1)
    SetMenuUnicode hSub, 0, U(75, 105, 7875, 109, 32, 116, 114, 97, 32, 110, 104, 7853, 112, 32, 120, 117, 7845, 116, 32, 116, 7891, 110)
    SetMenuUnicode hSub, 1, U(75, 105, 7875, 109, 32, 116, 114, 97, 32, 104, 7879, 32, 116, 7889, 110, 103, 32, 116, 224, 105, 32, 107, 104, 111, 7843, 110)
    'SetMenuUnicode hSub, 2, U(88, 7917, 32, 108, 253, 32, 115, 7889, 32, 108, 105, 7879, 117)
    SetMenuUnicode hSub, 3, U(88, 243, 97, 32, 112, 104, 225, 116, 32, 115, 105, 110, 104, 32, 116, 104, 225, 110, 103)
    hSub2 = GetSubMenu(hSub, 3)
    SetMenuUnicode hSub2, 0, U(83, 7889, 32, 100, 432, 32, 273, 7847, 117, 32, 110, 259, 109)

    SetMenuUnicode hSub, 5, U(67, 104, 117, 121, 7875, 110, 32, 115, 97, 110, 103, 32, 110, 259, 109, 32, 109, 7899, 105)
    SetMenuUnicode hSub, 6, U(67, 104, 7917, 110, 103, 32, 116, 7915, 32, 107, 7871, 116, 32, 99, 104, 117, 121, 7875, 110)
    SetMenuUnicode hSub, 7, U(80, 104, 226, 110, 32, 98, 7893, 32, 99, 104, 105, 32, 112, 104, 237)
    SetMenuUnicode hSub, 8, U(75, 7871, 116, 32, 99, 104, 117, 121, 7875, 110, 32, 115, 7889, 32, 108, 105, 7879, 117)
    SetMenuUnicode hSub, 9, U(75, 104, 243, 97, 32, 115, 7889, 32, 108, 105, 7879, 117, 32, 116, 104, 225, 110, 103)
    SetMenuUnicode hSub, 10, U(67, 104, 117, 121, 7875, 110, 32, 100, 7919, 32, 108, 105, 7879, 117, 32, 273, 7847, 117, 32, 107, 7923)
    SetMenuUnicode hSub, 11, U(75, 104, 97, 105, 32, 98, 225, 111, 32, 109, 7851, 117, 32, 98, 105, 7875, 117, 32, 115, 111, 110, 103, 32, 110, 103, 7919)
    hSub3 = GetSubMenu(hSub, 8)
    SetMenuUnicode hSub3, 0, U(83, 7889, 32, 100, 432, 32, 99, 117, 7889, 105, 32, 110, 259, 109)
    ' =====================================================
    ' V?T TU – HÀNG HÓA
    ' =====================================================
    hSub = GetSubMenu(hMenu, 2)
    SetMenuUnicode hSub, 0, U(80, 104, 226, 110, 32, 108, 111, 7841, 105, 32, 118, 7853, 116, 32, 116, 432)
    SetMenuUnicode hSub, 1, U(68, 97, 110, 104, 32, 115, 225, 99, 104, 32, 118, 7853, 116, 32, 116, 432, 32, 104, 224, 110, 103, 32, 104, 243, 97)
    SetMenuUnicode hSub, 2, U(75, 234, 110, 104, 32, 112, 104, 226, 110, 32, 112, 104, 7889, 105)
    SetMenuUnicode hSub, 3, U(76, 432, 117, 32, 99, 104, 117, 121, 7875, 110, 32, 110, 7897, 105, 32, 98, 7897)
    SetMenuUnicode hSub, 4, U(84, 104, 224, 110, 104, 32, 112, 104, 7849, 109, 32, 104, 111, 224, 110, 32, 116, 104, 224, 110, 104, 32, 116, 114, 111, 110, 103, 32, 107, 7923)
    SetMenuUnicode hSub, 6, U(84, 104, 234, 109, 32, 107, 104, 111, 32, 104, 224, 110, 103)
    SetMenuUnicode hSub, 7, U(84, 7891, 110, 32, 107, 104, 111, 32, 273, 7847, 117, 32, 107, 7923)
    SetMenuUnicode hSub, 8, U(84, 237, 110, 104, 32, 108, 7841, 105, 32, 103, 105, 225, 32, 120, 117, 7845, 116, 32, 107, 104, 111, 32, 116, 114, 111, 110, 103, 32, 116, 104, 225, 110, 103)
    SetMenuUnicode hSub, 9, U(84, 237, 110, 104, 32, 103, 105, 225, 32, 118, 7889, 110, 32, 104, 224, 110, 103, 32, 98, 225, 110)
    ' Ki?m kê t?n kho cu?i ngày
    SetMenuUnicode hSub, 11, U(75, 105, 7875, 109, 32, 107, 234, 32, 116, 7891, 110, 32, 107, 104, 111, 32, 99, 117, 7889, 105, 32, 110, 103, 224, 121)
    'SetMenuUnicode hSub, 11, U(75, 105, 7875, 109, 32, 107, 234, 32, 116, 7891, 110, 32, 107, 104, 111, 32, 99, 117, 7889, 105, 32, 110, 103, 224, 121)

    SetMenuUnicode hSub, 12, U(80, 104, 226, 110, 32, 108, 111, 7841, 105, 32, 99, 244, 110, 103, 32, 116, 114, 236, 110, 104, 44, 32, 115, 7843, 110, 32, 112, 104, 7849, 109)
    SetMenuUnicode hSub, 13, U(67, 104, 105, 32, 116, 105, 7871, 116, 32, 99, 244, 110, 103, 32, 116, 114, 236, 110, 104, 44, 32, 115, 7843, 110, 32, 112, 104, 7849, 109)
    SetMenuUnicode hSub, 14, U(84, 224, 105, 32, 107, 104, 111, 7843, 110, 32, 100, 111, 97, 110, 104, 32, 116, 104, 117)
    SetMenuUnicode hSub, 16, U(272, 7863, 116, 32, 98, 7887, 32, 84, 75, 32, 116, 104, 101, 111, 32, 100, 245, 105, 32, 116, 104, 101, 111, 32, 99, 104, 105, 32, 116, 105, 7871, 116)
    'SetMenuUnicode hSub, 17, U(68, 97, 110, 104, 32, 115, 225, 99, 104, 32, 118, 7853, 116, 32, 116, 432, 32, 104, 224, 110, 103, 32, 104, 111, 225)
    SetMenuUnicode hSub, 17, U(84, 7915, 32, 107, 104, 111, 225, 32, 116, 104, 101, 111, 32, 116, 234, 110, 32, 104, 224, 110, 103)
    SetMenuUnicode hSub, 18, U(84, 7921, 32, 273, 7897, 110, 103, 32, 110, 104, 7853, 112, 32, 107, 104, 111)
    SetMenuUnicode hSub, 19, U(88, 111, 225, 32, 118, 7853, 116, 32, 116, 432, 32, 116, 104, 7915, 97)
    ' =====================================================
    ' CÔNG N?
    ' =====================================================
    hSub = GetSubMenu(hMenu, 3)
    SetMenuUnicode hSub, 0, U(80, 104, 226, 110, 32, 108, 111, 7841, 105)
    SetMenuUnicode hSub, 1, U(68, 97, 110, 104, 32, 115, 225, 99, 104)
    SetMenuUnicode hSub, 3, U(83, 7889, 32, 100, 432, 32, 273, 7847, 117, 32, 107, 7923)
    SetMenuUnicode hSub, 4, U(68, 97, 110, 104, 32, 115, 225, 99, 104, 32, 72, 7907, 112, 32, 273, 7891, 110, 103)
    SetMenuUnicode hSub, 9, U(272, 7863, 116, 32, 98, 7887, 32, 116, 104, 101, 111, 32, 100, 245, 105, 32, 116, 104, 101, 111, 32, 99, 104, 105, 32, 116, 105, 7871, 116)
    SetMenuUnicode hSub, 10, U(88, 111, 225, 32, 99, 244, 110, 103, 32, 110, 7907, 32, 116, 104, 7915, 97)

    ' =====================================================
    ' TÀI S?N C? Ð?NH
    ' =====================================================
    hSub = GetSubMenu(hMenu, 4)
    SetMenuUnicode hSub, 0, U(80, 104, 226, 110, 32, 108, 111, 7841, 105, 32, 116, 224, 105, 32, 115, 7843, 110)
    SetMenuUnicode hSub, 1, U(80, 104, 226, 110, 32, 108, 111, 7841, 105, 32, 99, 104, 7917, 110, 103, 32, 116, 7915)
    SetMenuUnicode hSub, 2, U(68, 97, 110, 104, 32, 115, 225, 99, 104, 32, 84, 83, 67, 272)
    SetMenuUnicode hSub, 4, U(78, 432, 7899, 99, 32, 115, 7843, 110, 32, 120, 117, 7845, 116)
    SetMenuUnicode hSub, 5, U(84, 236, 110, 104, 32, 116, 114, 7841, 110, 103, 32, 115, 7917, 32, 100, 7909, 110, 103)
    SetMenuUnicode hSub, 6, U(272, 7889, 105, 32, 116, 432, 7901, 110, 103, 32, 113, 117, 7843, 110, 32, 108, 253)
    SetMenuUnicode hSub, 8, U(84, 224, 105, 32, 115, 7843, 110, 32, 273, 7847, 117, 32, 107, 7923)
    SetMenuUnicode hSub, 10, U(272, 7863, 116, 32, 98, 7887, 32, 84, 75, 32, 99, 104, 105, 32, 112, 104, 237, 32, 107, 104, 7845, 117, 32, 104, 97, 111)
    SetMenuUnicode hSub, 11, U(68, 97, 110, 104, 32, 115, 225, 99, 104, 32, 116, 224, 105, 32, 115, 7843, 110, 32, 99, 7889, 32, 273, 7883, 110, 104)

    ' =====================================================
    ' TR? GIÚP
    ' =====================================================
    hSub = GetSubMenu(hMenu, 5)
    SetMenuUnicode hSub, 0, U(71, 105, 7899, 105, 32, 116, 104, 105, 7879, 117)
    SetMenuUnicode hSub, 2, U(84, 224, 105, 32, 108, 105, 7879, 117)
    SetMenuUnicode hSub, 4, U(84, 7841, 111, 32, 99, 244, 110, 103, 32, 116, 121, 32, 109, 7899, 105)
End Sub
 
Public Sub HideSubMenu()

    gCurrentMenu = -1

    If Not gSubMenu Is Nothing Then
        Unload gSubMenu
        Set gSubMenu = Nothing
    End If

End Sub
Public Sub SetMenuFont(FontName As String, PointSize As Integer, Optional Bold As Boolean = False)
    Dim ncm As NONCLIENTMETRICS
    Dim hDC As Long
    Dim PixelsPerInchY As Long
    Dim lfHeight As Long

    ncm.cbSize = Len(ncm)

    ' L?y thông tin hi?n t?i
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, Len(ncm), ncm, 0

    ' Tính lfHeight t? PointSize (âm d? là point size)
    hDC = GetDC(0)  ' Desktop DC
    PixelsPerInchY = GetDeviceCaps(hDC, LOGPIXELSY)
    ReleaseDC 0, hDC

    lfHeight = -MulDiv(PointSize, PixelsPerInchY, 72)   ' Công th?c chu?n (MulDiv c?n Declare n?u chua có)
    ' Ho?c don gi?n n?u DPI ~96: lfHeight = -CLng(PointSize * 1.333)

    With ncm.lfMenuFont
        .lfFaceName = Left$(FontName & String$(32, vbNullChar), 32)   ' Copy tên font + nulls
        .lfHeight = lfHeight
        .lfWeight = IIf(Bold, 700, 400)   ' FW_BOLD = 700, FW_NORMAL = 400
        ' .lfItalic = 1 n?u mu?n italic, v.v.
    End With

    ' Ð?ng b? caption n?u mu?n
    With ncm.lfCaptionFont
        .lfFaceName = Left$(FontName & String$(32, vbNullChar), 32)
        .lfHeight = lfHeight
        .lfWeight = IIf(Bold, 700, 400)
    End With
    ' Áp d?ng cho MessageBox
    With ncm.lfMessageFont
        .lfFaceName = Left$(FontName & String$(32, vbNullChar), 32)
        .lfHeight = lfHeight
        .lfWeight = IIf(Bold, 700, 400)
    End With

    ' Áp d?ng thay d?i
    SystemParametersInfo SPI_SETNONCLIENTMETRICS, Len(ncm), ncm, SPIF_SENDCHANGE

    ' Broadcast d? t?t c? window c?p nh?t
    SendMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
End Sub
Public Sub RestoreSysFont()
    If g_HasBackup = False Then Exit Sub

    SystemParametersInfo SPI_SETNONCLIENTMETRICS, _
                         Len(g_NcmBackup), g_NcmBackup, _
                         SPIF_UPDATEINIFILE Or SPIF_SENDCHANGE

    SendMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
End Sub
Public Sub BackupSysFont()
    g_NcmBackup.cbSize = Len(g_NcmBackup)
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, Len(g_NcmBackup), g_NcmBackup, 0
    g_HasBackup = True
End Sub
Public Sub CheckAndCreateTBCpu()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbCpu"
    tableExists = False


    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.Name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        ' Username
        Set fld = tdf.CreateField("Name", dbText, 255)
        fld.Required = False
        fld.AllowZeroLength = True
        tdf.Fields.Append fld
 
        ' Thêm b?ng vào CSDL
        DBKetoan.TableDefs.Append tdf
    End If
End Sub
Public Sub CheckAndCreateTBInvoiceTemplate()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbInvoiceTemplate"
    tableExists = False


    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.Name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        ' ID
        Set fld = tdf.CreateField("ID", dbText, 255)
        fld.Required = False
        fld.AllowZeroLength = True
        tdf.Fields.Append fld
        
         ' Code
        Set fld = tdf.CreateField("Code", dbText, 255)
        fld.Required = False
        fld.AllowZeroLength = True
        tdf.Fields.Append fld
        
         ' Name
        Set fld = tdf.CreateField("Name", dbText, 255)
        fld.Required = False
        fld.AllowZeroLength = True
        tdf.Fields.Append fld
 
        ' Thêm b?ng vào CSDL
        DBKetoan.TableDefs.Append tdf
    End If
End Sub
Public Sub CheckAndCreateTBInvoice()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbInvoiceInfo"
    tableExists = False


    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.Name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        ' Url
        Set fld = tdf.CreateField("Url", dbText, 255)
        fld.Required = False
        fld.AllowZeroLength = True
        tdf.Fields.Append fld
        ' Url
        Set fld = tdf.CreateField("Username", dbText, 255)
        fld.Required = False
        fld.AllowZeroLength = True
        tdf.Fields.Append fld
        'Password
        ' Url
        Set fld = tdf.CreateField("Password", dbText, 255)
        fld.Required = False
        fld.AllowZeroLength = True
        tdf.Fields.Append fld
        
 
        ' Thêm b?ng vào CSDL
        DBKetoan.TableDefs.Append tdf
    End If
End Sub
Private Sub menuversion()
    mnItemCount = 0
    ReDim arrItems(1 To 1)    ' Kh?i t?o m?ng


    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    Set folder = fso.GetFolder("\\192.168.1.90\Ke toan 2025 New\1 Copi vao dung 1\Versions")
    Dim i As Integer
    i = 2
    For Each subFolder In folder.SubFolders
        ' Debug.Print subFolder.Name
        Load mnversionChild(i)
        mnversionChild(i).Caption = subFolder.Name
        mnversionChild(i).Visible = True
        Call AddItem(i, subFolder.Name)
        i = i + 1
    Next

    Set subFolder = Nothing
    Set folder = Nothing
    Set fso = Nothing

    'Menu m?u
    'mnversionChild(1).Visible = False
    mnversionChild(1).Visible = False
End Sub
Private Sub KiemtraversionDrive()
    Dim fileId As String
    fileId = "1fXXVQCw2XHusJ9y5zwIt0jPJTCL4dxc2"
    '1fXXVQCw2XHusJ9y5zwIt0jPJTCL4dxc2
    Dim fileName As String
    fileName = "version.config"

    Dim exeDir As String
    exeDir = App.path  ' D:\DA3\DA3
    Dim extractFolder As String
    extractFolder = exeDir & "\Tools\Debug"

    Dim zipPath As String
    zipPath = extractFolder & "\" & fileName  ' D:\DA3\DA3\Tools\Debug\VietstarDriver.zip
    Dim retryCount As Integer
    Dim downloadUrl As String
    Dim Result As Long
    ' Xóa file cu
    If Dir(zipPath) <> "" Then Kill zipPath

    Dim urls(1) As String
    urls(0) = "https://drive.google.com/uc?export=download&id=" & fileId & "&confirm=t"
    urls(1) = "https://drive.google.com/uc?export=download&id=" & fileId

    Dim success As Boolean
    success = False

    ' Result = URLDownloadToFile(0, urls(0), zipPath, 0, 0)
    If DownloadFileWithTimeout("https://drive.google.com/uc?export=download&id=" & fileId, App.path & "\file.zip", 5) Then
        MsgBox "T?i thành công!"
        Result = 0
    Else
        MsgBox "Loi khong the cap nhat version moi"
        Exit Sub
    End If

    If Result = 0 Then
        success = True
    Else
        success = False
    End If

    If Not success Then
        'MsgBox "Không th? t?i file!", vbExclamation
        Exit Sub
    End If


    'Load menu
    Dim json As String
    Dim versions As Collection
    Dim i As Long

    mnItemCount = 0
    ReDim arrItems(1 To 1)    ' Kh?i t?o m?ng
    json = ReadJsonFile(App.path & "\Tools\Debug\version.config")

    Set versions = GetVersions(json)
    Dim j As Integer
    j = 2
    For i = 1 To versions.count

        If CStr(versions(i)) <> "1.1" Then
            Debug.Print versions(i)
            Load mnversionChild(j)
            mnversionChild(j).Caption = versions(i)
            mnversionChild(j).Visible = True
            Call AddItem(j, versions(i))
        End If
        j = j + 1
    Next

    mnversionChild(1).Visible = False

    'Lay version hien tai

    Dim txtPath As String
    txtPath = App.path & "\" & "HoaDon\version.txt"
    Dim content As String
    content = ReadTxt(txtPath)

    'Kiem tra version hien tai có trong danh sách không co thi tai ban cap nhat moi nhat
    Dim versionData As versionInfo
    Docfilejson content, versionData
    If versionData.Version = "" Then
        'Lay version moi nhat
        DownloadAndRun
    End If
    Docfilejson "1.1", versionData
    LastVersion = Replace(versionData.fileName, """", "")
End Sub
Private Function DownloadFileWithWinHTTP(ByVal url As String, ByVal savePath As String, Optional ByVal TimeoutSeconds As Long = 60) As Boolean
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim stream As Object
    Dim StartTime As Long

    ' T?o d?i tu?ng WinHTTP
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' M? k?t n?i
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"

    ' Set timeout
    http.setTimeouts TimeoutSeconds * 1000, TimeoutSeconds * 1000, TimeoutSeconds * 1000, TimeoutSeconds * 1000

    ' G?i request
    http.send

    ' Ki?m tra status
    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1
        stream.Open
        stream.Write http.responseBody
        stream.SaveToFile savePath, 2
        stream.Close
        DownloadFileWithWinHTTP = True
    Else
        MsgBox "Status: " & http.Status & vbCrLf & "URL: " & url
        DownloadFileWithWinHTTP = False
    End If

    Exit Function
ErrorHandler:
    ' MsgBox "L?i: " & Err.Description & vbCrLf & "URL: " & URL
    Dim s As String
    s = ChrW(67) & ChrW(243) & ChrW(32) & ChrW(108) & ChrW(7895) & ChrW(105) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(111) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(113) & ChrW(117) & ChrW(225) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(236) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(99) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7853) & ChrW(116) & ChrW(44) & ChrW(32) & ChrW(118) & ChrW(117) & ChrW(105) & ChrW(32) & ChrW(108) & ChrW(242) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(7917) & ChrW(32) & ChrW(108) & ChrW(7841) & ChrW(105) & ChrW(32) & ChrW(115) & ChrW(97) & ChrW(117)
    MessageBoxW Me.hwnd, StrPtr(s), StrPtr("Thông báo"), vbOKOnly

    DownloadFileWithWinHTTP = False
End Function
Private Function DownloadFileWithTimeout(ByVal url As String, ByVal LocalFile As String, Optional ByVal TimeoutSeconds As Long = 5) As Boolean
    On Error GoTo ErrorHandler
    Dim http As Object, stream As Object
    Dim StartTime As Long

    StartTime = GetTickCount
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' Set timeout (cái này không an thua v?i Google Drive)
    http.setTimeouts TimeoutSeconds * 1000, TimeoutSeconds * 1000, TimeoutSeconds * 1000, TimeoutSeconds * 1000

    ' M? k?t n?i - Ð?ng b? (False)
    http.Open "GET", url, False
    http.send

    ' **QUAN TR?NG: Ki?m tra timeout liên t?c**
    Do While http.readyState <> 4
        ' Ki?m tra timeout m?i 100ms
        If (GetTickCount - StartTime) / 1000 > TimeoutSeconds Then
            Debug.Print "TIMEOUT sau " & TimeoutSeconds & " giây"
            DownloadFileWithTimeout = False
            Exit Function
        End If
        DoEvents
        Sleep 100  ' Ch? 100ms r?i ki?m tra ti?p
    Loop

    ' N?u hoàn thành trong th?i gian cho phép
    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1
        stream.Open
        stream.Write http.responseBody
        stream.SaveToFile LocalFile, 2
        stream.Close
        DownloadFileWithTimeout = True
    Else
        DownloadFileWithTimeout = False
    End If

    Exit Function
ErrorHandler:
    DownloadFileWithTimeout = False
End Function
Private Function GetVersions(ByVal json As String) As Collection

    Dim col As New Collection
    Dim p As Long
    Dim pEnd As Long
    Dim pVersions As Long
    Dim level As Long
    Dim ch As String
    Dim key As String

    pVersions = InStr(1, json, """versions""", vbTextCompare)

    If pVersions = 0 Then
        Set GetVersions = col
        Exit Function
    End If

    p = InStr(pVersions, json, "{")

    If p = 0 Then
        Set GetVersions = col
        Exit Function
    End If

    p = p + 1
    level = 1

    Do While p <= Len(json)

        ch = Mid$(json, p, 1)

        If ch = "}" Then
            level = level - 1

            If level = 0 Then Exit Do

        ElseIf ch = "{" Then
            level = level + 1

        ElseIf ch = """" And level = 1 Then

            pEnd = InStr(p + 1, json, """")

            If pEnd = 0 Then Exit Do

            key = Mid$(json, p + 1, pEnd - p - 1)

            'Sau key ph?i là :
            If Trim$(Mid$(json, pEnd + 1, 1)) = ":" Then

                col.Add key

            End If

            p = pEnd

        End If

        p = p + 1

    Loop

    Set GetVersions = col

End Function

Private Sub Docfilejson(ByVal versionid As String, ByRef versionData As versionInfo)

    Dim json As String
    Dim pVersion As Long

    json = ReadJsonFile(App.path & "\Tools\Debug\configVersion.json")

    If json = "" Then
        MsgBox "Không d?c du?c file JSON"
        Exit Sub
    End If

    'Tìm dúng version du?c truy?n vào
    pVersion = InStr(1, json, """" & versionid & """", vbTextCompare)

    If pVersion = 0 Then
        'MsgBox "Không tìm th?y version " & versionid
        Exit Sub
    End If

    versionData.Version = versionid

    versionData.fileId = GetJsonValue(json, "fileId", pVersion)
    versionData.fileName = GetJsonValue(json, "fileName", pVersion)
    versionData.ReleaseDate = GetJsonValue(json, "releaseDate", pVersion)
    versionData.Description = GetJsonValue(json, "description", pVersion)

End Sub
Private Function ReadJsonFile(ByVal FilePath As String) As String
    Dim f As Integer
    Dim s As String

    If Dir(FilePath) = "" Then Exit Function

    f = FreeFile

    Open FilePath For Input As #f
        s = Input$(LOF(f), f)
    Close #f

    ReadJsonFile = s
End Function

Private Function GetJsonValue(ByVal json As String, _
                              ByVal PropertyName As String, _
                              ByVal startPos As Long) As String

    Dim p As Long
    Dim pColon As Long
    Dim pStart As Long
    Dim pEnd As Long
    Dim SearchText As String

    SearchText = """" & PropertyName & """"

    p = InStr(startPos, json, SearchText, vbTextCompare)

    If p = 0 Then Exit Function

    pColon = InStr(p + Len(SearchText), json, ":")

    If pColon = 0 Then Exit Function

    ' B? kho?ng tr?ng và d?u "
    pStart = pColon + 1

    Do While pStart <= Len(json)
        If Mid$(json, pStart, 1) <> " " _
           And Mid$(json, pStart, 1) <> vbTab _
           And Mid$(json, pStart, 1) <> vbCr _
           And Mid$(json, pStart, 1) <> vbLf _
           And Mid$(json, pStart, 1) <> """" Then
            Exit Do
        End If

        pStart = pStart + 1
    Loop

    ' N?u giá tr? b?t d?u b?ng "
    If Mid$(json, pColon + 1, 1) = """" Then

        pStart = pColon + 2

        pEnd = InStr(pStart, json, """")

        If pEnd = 0 Then Exit Function

        GetJsonValue = Mid$(json, pStart, pEnd - pStart)

    Else

        pEnd = InStr(pStart, json, ",")

        If pEnd = 0 Then
            pEnd = InStr(pStart, json, "}")
        End If

        If pEnd = 0 Then Exit Function

        GetJsonValue = Trim$(Mid$(json, pStart, pEnd - pStart))

    End If

End Function
Private Sub DownloadAndRun()
    On Error GoTo ErrorHandler

    Dim exeDir As String
    Dim fileId As String
    Dim fileName As String
    Dim zipPath As String
    Dim extractFolder As String
    Dim exePath As String
    Dim downloadUrl As String
    Dim Result As Long
    Dim retryCount As Integer

    ' ===== THI?T L?P ÐU?NG D?N =====
    exeDir = App.path  ' D:\DA3\DA3
    fileId = "1qj3mjQkA1o6QEsellIQ09_vrzq1ETCOL"
    fileName = "VietstarDriver.zip"

    ' Thu m?c gi?i nén (Tools\Debug)
    extractFolder = exeDir & "\Tools\Debug"
    zipPath = extractFolder & "\" & fileName  ' D:\DA3\DA3\Tools\Debug\VietstarDriver.zip
    exePath = extractFolder & "\VietstarDriver.exe"

    ' Xóa file cu
    If Dir(zipPath) <> "" Then Kill zipPath
    ' T?o thu m?c n?u chua có
    If Dir(extractFolder, vbDirectory) = "" Then
        MkDir extractFolder
    End If
    ' ===== T?I FILE =====
    Dim urls(2) As String
    urls(0) = "https://drive.google.com/uc?export=download&id=" & fileId
    urls(1) = "https://drive.google.com/uc?export=download&id=" & fileId & "&confirm=1"
    urls(2) = "https://drive.google.com/uc?export=download&id=" & fileId & "&confirm=t"

    Dim success As Boolean
    success = False

    For retryCount = 0 To 2
        downloadUrl = urls(retryCount)
        Result = URLDownloadToFile(0, downloadUrl, zipPath, 0, 0)

        If Result = 0 Then
            If FileLen(zipPath) > 1024 Then
                success = True
                Exit For
            Else
                Kill zipPath
            End If
        End If

        Sleep 1000
    Next retryCount

    If Not success Then
        MsgBox "Không th? t?i file!", vbExclamation
        Exit Sub
    End If

    ' ===== GI?I NÉN VÀO TOOLS\DEBUG =====
    If Dir(zipPath) <> "" Then
        UnzipFile zipPath, extractFolder  ' Gi?i nén vào Tools\Debug
        If Dir(zipPath) <> "" Then Kill zipPath
    End If

    ' ===== CH?Y FILE EXE =====
    If Dir(exePath) <> "" Then
        Shell exePath, vbNormalFocus
        Unload Me
        End
    Else
        MsgBox "Không tìm th?y file EXE!", vbExclamation
    End If

    Exit Sub

ErrorHandler:
    MsgBox "L?i: " & Err.Description, vbCritical
End Sub
Private Sub UnzipFile(zipPath As String, extractFolder As String)
    On Error GoTo ErrorHandler

    Dim cmd As String
    Dim wsh As Object
    Dim Result As Long


    ' Ki?m tra file ZIP
    If Dir(zipPath) = "" Then
        MsgBox "Không tìm th?y file ZIP: " & zipPath, vbExclamation
        Exit Sub
    End If

    ' T?o thu m?c dích
    If Dir(extractFolder, vbDirectory) = "" Then
        MkDir extractFolder
    End If

    ' Ki?m tra file ZIP
    If FileLen(zipPath) < 1024 Then
        MsgBox "File ZIP b? l?i ho?c quá nh?!", vbExclamation
        Kill zipPath
        Exit Sub
    End If

    ' PowerShell
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " & _
          """Expand-Archive -LiteralPath '" & Replace(zipPath, "'", "''") & _
          "' -DestinationPath '" & Replace(extractFolder, "'", "''") & _
          "' -Force"""

    Set wsh = CreateObject("WScript.Shell")

    ' 0 = ?n c?a s?
    ' True = CH? PowerShell ch?y xong
    Result = wsh.Run(cmd, 0, True)

    If Result <> 0 Then
        MsgBox "Gi?i nén th?t b?i. Mã l?i: " & Result, vbExclamation
        Exit Sub
    End If

    ' ========== THÊM DÒNG NÀY Ð? XÓA FILE ZIP ==========
    If Dir(zipPath) <> "" Then
        Kill zipPath
    End If
    ' ===================================================
    Set wsh = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "L?i gi?i nén: " & Err.number & " - " & Err.Description, vbExclamation
End Sub
' ===== HÀM KI?M TRA FILE ÐANG B? KHÓA =====
Private Function IsFileLocked(FilePath As String) As Boolean
    On Error GoTo Locked
    
    Open FilePath For Binary Access Read Lock Read As #1
    Close #1
    IsFileLocked = False
    Exit Function
    
Locked:
    IsFileLocked = True
End Function

' ========== HÀM CHÍNH: KI?M TRA VERSION T? GITHUB ==========
Private Sub KiemtraversionGitHub()
    Dim json As String
    Dim latestVer As String
    Dim currentVer As String
    Dim fileName As String
    Dim fileId As String
    Dim downloadUrl As String
    Dim savePath As String
    Dim exeDir As String
    Dim extractFolder As String
    Dim versionInfo As versionInfo

    ' 1. Ð?c file JSON t? GitHub
    json = GetConfigFromGitHub()
    If json = "" Then
        MsgBox "Không th? d?c file c?u hình t? GitHub!", vbCritical
        Exit Sub
    End If

    ' 2. L?y danh sách versions
    Dim versions As Collection
    Set versions = GetVersionsFromJSON(json)

    If versions.count = 0 Then
        MsgBox "Không tìm th?y version nào trong file c?u hình!", vbCritical
        Exit Sub
    End If

    ' 3. L?y version m?i nh?t (version cu?i cùng trong danh sách)
    latestVer = versions(versions.count)

    ' 4. Ð?c version hi?n t?i
    Dim txtPath As String
    txtPath = App.path & "\HoaDon\version.txt"
    currentVer = ReadTxt(txtPath)

    ' 5. L?y thông tin c?a latestVer
    versionInfo = GetVersionInfoByVersion(json, latestVer)

    ' 6. So sánh version
    If latestVer <> currentVer Then
        fileName = versionInfo.fileName
        fileId = versionInfo.fileId
    End If
    LoadMenuFromJSON json
End Sub
Private Sub LoadMenuFromJSON(ByVal json As String)
    Dim versions As Collection
    Dim i As Long
    Dim j As Integer

    ' Xóa menu cu (gi? nguyên item 0 và 1)
    On Error Resume Next
    For i = 2 To mnversionChild.count - 1
        Unload mnversionChild(i)
    Next
    On Error GoTo 0

    ' L?y danh sách version
    Set versions = GetVersionsFromJSON(json)

    j = 2
    For i = 1 To versions.count
        If CStr(versions(i)) <> "1.1" Then
            Debug.Print versions(i)
            Load mnversionChild(j)
            mnversionChild(j).Caption = versions(i)
            mnversionChild(j).Visible = True
            Call AddItem(j, versions(i))
        End If
        j = j + 1
    Next

    mnversionChild(1).Visible = False
End Sub
Private Sub ExtractZipFile(ByVal zipPath As String, ByVal extractPath As String)
    Dim ShellApp As Object
    Dim ZipFile As Object
    Dim DestFolder As Object

    ' T?o d?i tu?ng Shell
    Set ShellApp = CreateObject("Shell.Application")

    ' M? file zip
    Set ZipFile = ShellApp.Namespace(zipPath)

    ' T?o thu m?c dích n?u chua có
    If Dir(extractPath, vbDirectory) = "" Then
        MkDir extractPath
    End If

    ' M? thu m?c dích
    Set DestFolder = ShellApp.Namespace(extractPath)

    ' Gi?i nén t?t c? các file (tham s? 16 = không hi?n th? h?p tho?i)
    DestFolder.CopyHere ZipFile.Items, 16
End Sub

' ========== HÀM Ð?C FILE JSON T? GITHUB ==========
Private Function GetConfigFromGitHub() As String
    On Error GoTo ErrorHandler
    Dim http As Object
    Dim url As String
    
    ' Dùng link raw t? nhánh master
    url = "https://raw.githubusercontent.com/n93studioenter/SaovietGit/master/configVersion.json"
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send
    
    If http.Status = 200 Then
        GetConfigFromGitHub = http.responseText
    Else
        GetConfigFromGitHub = ""
    End If
    Exit Function
ErrorHandler:
    GetConfigFromGitHub = ""
End Function

' ========== HÀM L?Y DANH SÁCH VERSION T? JSON ==========
Private Function GetVersionsFromJSON(ByVal json As String) As Collection
    Dim versions As New Collection
    Dim startPos As Long
    Dim endPos As Long
    Dim tempPos As Long
    Dim Version As String
    
    ' Tìm ph?n "versions": { ... }
    startPos = InStr(json, """versions""")
    If startPos > 0 Then
        startPos = InStr(startPos, json, "{")
        startPos = InStr(startPos, json, "{") + 1
        endPos = InStrRev(json, "}")
        
        ' Duy?t qua t?ng version
        Do While startPos < endPos And startPos > 0
            ' Tìm d?u " d? b?t d?u tên version
            startPos = InStr(startPos, json, """")
            If startPos = 0 Or startPos > endPos Then Exit Do
            
            startPos = startPos + 1
            endPos = InStr(startPos, json, """")
            Version = Mid(json, startPos, endPos - startPos)
            
            ' B? qua các key không ph?i version (fileName, fileId, v.v.)
            If Version <> "fileName" And Version <> "fileId" And _
               Version <> "releaseDate" And Version <> "description" Then
                If Version <> "" Then
                    versions.Add Version
                End If
            End If
            
            startPos = endPos + 1
        Loop
    End If
    
    Set GetVersionsFromJSON = versions
End Function

' ========== HÀM L?Y THÔNG TIN VERSION ==========
Private Function GetVersionInfoByVersion(ByVal json As String, ByVal Version As String) As versionInfo
    Dim info As versionInfo
    Dim startPos As Long
    Dim endPos As Long
    Dim versionPattern As String
    Dim key As String
    Dim Value As String
    
    versionPattern = """" & Version & """" & ":"
    startPos = InStr(json, versionPattern)
    
    If startPos > 0 Then
        ' Tìm v? trí b?t d?u c?a object
        startPos = InStr(startPos, json, "{")
        
        ' L?y toàn b? object c?a version này
        endPos = FindMatchingBrace(json, startPos)
        
        If endPos > startPos Then
            Dim block As String
            block = Mid(json, startPos, endPos - startPos + 1)
            
            ' Parse fileName
            info.fileName = ParseJSONValue(block, "fileName")
            ' Parse fileId
            info.fileId = ParseJSONValue(block, "fileId")
            ' Parse releaseDate
            info.ReleaseDate = ParseJSONValue(block, "releaseDate")
            ' Parse description
            info.Description = ParseJSONValue(block, "description")
        End If
    End If
    
    info.Version = Version
    GetVersionInfoByVersion = info
End Function

' ========== HÀM TÌM D?U NGO?C NH?N TUONG ?NG ==========
Private Function FindMatchingBrace(ByVal Text As String, ByVal startPos As Long) As Long
    Dim count As Integer
    Dim i As Long
    
    count = 1
    i = startPos + 1
    
    Do While i <= Len(Text) And count > 0
        If Mid(Text, i, 1) = "{" Then
            count = count + 1
        ElseIf Mid(Text, i, 1) = "}" Then
            count = count - 1
            If count = 0 Then
                FindMatchingBrace = i
                Exit Function
            End If
        End If
        i = i + 1
    Loop
    
    FindMatchingBrace = startPos
End Function

' ========== HÀM PARSE GIÁ TR? T? JSON ==========
Private Function ParseJSONValue(ByVal jsonBlock As String, ByVal key As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim pattern As String
    
    pattern = """" & key & """" & ":"
    startPos = InStr(jsonBlock, pattern)
    
    If startPos > 0 Then
        startPos = InStr(startPos, jsonBlock, ":")
        startPos = InStr(startPos, jsonBlock, """") + 1
        endPos = InStr(startPos, jsonBlock, """")
        If endPos > startPos Then
            ParseJSONValue = Mid(jsonBlock, startPos, endPos - startPos)
        End If
    End If
End Function

 

' ========== HÀM GHI FILE TEXT ==========
Private Sub WriteTxt(ByVal FilePath As String, ByVal content As String)
    On Error Resume Next
    Dim fileNo As Integer
    fileNo = FreeFile
    Open FilePath For Output As #fileNo
    Print #fileNo, content
    Close #fileNo
End Sub
Private Function DownloadJSONFromGitHub() As String
    On Error GoTo ErrorHandler
    Dim http As Object
    Dim url As String

    ' Link file JSON trên GitHub (nhánh master)
    url = "https://raw.githubusercontent.com/n93studioenter/SaovietGit/master/configVersion.json"

    ' T?o d?i tu?ng WinHTTP
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' G?i request
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send

    ' Ki?m tra k?t qu?
    If http.Status = 200 Then
        DownloadJSONFromGitHub = http.responseText
    Else
        DownloadJSONFromGitHub = ""
    End If

    Exit Function
ErrorHandler:
    DownloadJSONFromGitHub = ""
End Function
Private Sub DownloadReportFromGitHub()
    Dim downloaDriver As String
    Dim savePathDriver As String
    downloaDriver = "https://github.com/n93studioenter/SaovietGit/releases/download/latest-config/Reports.zip"
    savePathDriver = App.path & "\Tools\Debug" & "\" & "Reports" & ".zip"
    Dim extractFolder As String
    extractFolder = App.path & "\REPORTS"
    If Dir(savePathDriver) <> "" Then Kill savePathDriver
    frmSplash.Label2.Caption = "TiÕn hµnh t¶i file REPORTS..."
    DoEvents
    If DownloadFileWithWinHTTP(downloaDriver, savePathDriver, 10) Then
        frmSplash.Label2.Caption = "TiÕn hµnh gi¶i nÐn file REPORTS..."
        DoEvents
        UnzipFile savePathDriver, extractFolder
    Else
        Exit Sub
    End If
End Sub
Private Sub KillWebDriverUsingShell()
    Dim killedCount As Integer
    Shell "taskkill /F /IM chromedriver.exe /T", vbHide
End Sub
Private Sub KiemtraversionGitHub2()
    KillWebDriverUsingShell
    Dim exeDir As String
    Dim extractFolder As String
    Dim jsonPath As String
    Dim json As String
    Dim versions As Collection
    Dim i As Long
    Dim j As Integer

    ' 1. Ðu?ng d?n luu file JSON
    exeDir = App.path
    extractFolder = exeDir & "\Tools\Debug"

    ' T?o thu m?c n?u chua có
    If Dir(extractFolder, vbDirectory) = "" Then
        MkDir extractFolder
    End If

    jsonPath = extractFolder & "\configVersion.json"

    ' 2. Xóa file cu n?u có
    If Dir(jsonPath) <> "" Then Kill jsonPath

    frmSplash.Label2.Caption = "§ang t¶i file configVersion..."
    DoEvents
    ' 3. T?i file JSON t? GitHub
    Dim jsonContent As String
    jsonContent = DownloadJSONFromGitHub()

    If jsonContent = "" Then
        MsgBox "Không th? t?i file c?u hình t? GitHub!", vbCritical
        Exit Sub
    End If

    ' 4. Luu file JSON xu?ng máy
    Open jsonPath For Output As #1
    Print #1, jsonContent
    Close #1

    ' 5. Load menu t? file JSON v?a t?i
    mnItemCount = 0
    ReDim arrItems(1 To 1)
    json = ReadJsonFile(jsonPath)

    Set versions = GetVersions(json)
    j = 2
    For i = 1 To versions.count
        If CStr(versions(i)) <> "1.1" Then
            Debug.Print versions(i)
            Load mnversionChild(j)
            mnversionChild(j).Caption = versions(i)
            mnversionChild(j).Visible = True
            Call AddItem(j, versions(i))
        End If
        j = j + 1
    Next

    mnversionChild(1).Visible = False

    ' 6. L?y version hi?n t?i
    Dim txtPath As String
    txtPath = App.path & "\HoaDon\version.txt"
    Dim content As String
    content = ReadTxt(txtPath)

    ' 7. L?y version m?i nh?t t? JSON
    Dim latestVer As String
    latestVer = GetLatestVersionFromJSON(jsonContent)

    Dim versionData As versionInfo
    Docfilejson "1.1", versionData
    LastVersion = Replace(versionData.fileName, """", "")
    frmSplash.Label2.Caption = "Version hiÖn t¹i " & LastVersion
    DoEvents
    'Kiem tra hien tai co khong
    Dim versioncurrentData As versionInfo
    Docfilejson content, versioncurrentData
    
    DownloadReportFromGitHub
    ' 8. So sánh và x? lý c?p nh?t
    'Or versioncurrentData.Version = ""
    If val(content) <> val(LastVersion) And versioncurrentData.Version = "" Then
        hasUPdate = True
        frmSplash.Label2.Caption = "§ang cËp nhËt øng dông, vui lßng chê..."
        ' H?i ngu?i dùng
        Dim s As String
        s = ChrW(272) & ChrW(227) & ChrW(32) & ChrW(99) & ChrW(243) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(105) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(98) & ChrW(7843) & ChrW(110) & ChrW(32) & ChrW(109) & ChrW(7899) & ChrW(105) & ChrW(44) & ChrW(32) & ChrW(98) & ChrW(7841) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(243) & ChrW(32) & ChrW(109) & ChrW(117) & ChrW(7889) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7853) & ChrW(116)
        Dim xn As String
        xn = ChrW(88) & ChrW(225) & ChrW(99) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7853) & ChrW(110)
        'Tai VietstarDriver ve truoc

        Dim downloaDriver As String
        Dim savePathDriver As String
        downloaDriver = "https://github.com/n93studioenter/SaovietGit/releases/download/latest-config/VietstarDriver.zip"
        savePathDriver = App.path & "\Tools\Debug" & "\" & "VietstarDriver" & ".zip"
        Dim savepathexe As String
        savepathexe = App.path & "\Tools\Debug" & "\" & "VietstarDriver" & ".exe"
        If Dir(savePathDriver) <> "" Then Kill savePathDriver
        If Dir(savepathexe) <> "" Then Kill savepathexe
        frmSplash.Label2.Caption = "TiÕn hµnh t¶i file VietstarDriver..."
        DoEvents
        If DownloadFileWithWinHTTP(downloaDriver, savePathDriver, 10) Then
            frmSplash.Label2.Caption = "TiÕn hµnh gi¶i nÐn file VietstarDriver..."
            DoEvents
            UnzipFile savePathDriver, extractFolder
        Else
            Exit Sub
        End If

        ' L?y thông tin file c?n t?i
        Dim fileName As String
        fileName = GetFileNameByVersion(jsonContent, latestVer)

        If fileName <> "" Then
            ' T?o link t?i t? GitHub Release
            Dim downloadUrl As String
            Dim savePath As String
            downloadUrl = "https://github.com/n93studioenter/SaovietGit/releases/download/latest-config/1.1.zip"
            'downloadUrl = "https://github.com/n93studioenter/SaovietGit/releases/download/" & latestVer & "/" & fileName
            savePath = extractFolder & "\" & fileName & ".zip"

            ' Xóa file cu
            If Dir(savePath) <> "" Then Kill savePath

            ' T?i file
            frmSplash.Label2.Caption = "TiÕn hµnh t¶i file version míi..."
            DoEvents
            If DownloadFileWithWinHTTP(downloadUrl, savePath, 60) Then
                s = ChrW(67) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7853) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(224) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(99) & ChrW(244) & ChrW(110) & ChrW(103)
                'MessageBoxW Me.hwnd, StrPtr(s), StrPtr("Thông báo"), vbOKOnly
                If IsProcessRunning("ToolTaiHD.exe") Then
                    'MsgBox "ToolTaiHD.exe dang ch?y! Ti?n hành dóng..."
                    CloseProcessByName "ToolTaiHD.exe"
                End If
                frmSplash.Label2.Caption = "TiÕn hµnh t¶i tool auto..."
                DoEvents
                TaiToolAuto
                'UpdateApplication savePath

                ' C?p nh?t version.txt
                WriteTxt txtPath, LastVersion
                frmSplash.Label2.Caption = "§ang gi¶i nÐn c¸c tËp tin...."
                DoEvents
                ' Gi?i nén n?u là zip
                UnzipFile savePath, extractFolder

                Dim exePath As String
                exePath = App.path & "\Tools\Debug" & "\" & "VietstarDriver" & ".exe"
                'Dong chuong trinh va mo file update
                If Dir(exePath) <> "" Then
                    'Shell exePath, vbNormalFocus
                    ShellExecute 0, "open", exePath, "", App.path & "\Tools\Debug", SW_SHOW

                    Unload Me
                    End
                Else
                    MsgBox "Không tìm th?y file EXE!", vbExclamation
                End If
            Else
                MsgBox "T?i file th?t b?i!", vbCritical
            End If
        End If
        ' If MessageBoxW(Me.hwnd, StrPtr(s), StrPtr(xn), vbYesNo + vbExclamation) = vbYes Or 1 < 2 Then

        'End If

    Else
        Image2.Visible = False
    End If
End Sub
Private Function FindNewExe(ByVal folder As String) As String
    Dim exeFile As String
    Dim currentExe As String

    ' File dang ch?y
    currentExe = "VietstarDriver.exe"

    exeFile = Dir(folder & "\*.exe")

    Do While exeFile <> ""
        ' Tìm file có ch?a "vietstar" VÀ KHÁC "VietstarDriver.exe"
        If InStr(1, LCase(exeFile), "vietstar", vbTextCompare) > 0 Then
            If LCase(exeFile) <> LCase(currentExe) Then
                FindNewExe = folder & "\" & exeFile
                Exit Function
            End If
        End If
        exeFile = Dir
    Loop

    FindNewExe = ""
End Function
Private Sub UpdateApplication(ByVal zipPath As String)
    Dim batchPath As String
    Dim currentExe As String
    Dim batchContent As String

    currentExe = App.ExeName & ".exe"
    batchPath = App.path & "\update.bat"

    ' Batch file - ch? l?y file .exe (không l?y .exe.config)
   batchContent = "@echo off" & vbCrLf & _
                "echo Dang cap nhat..." & vbCrLf & _
                "cd /d """ & App.path & """" & vbCrLf & _
                "taskkill /f /im """ & currentExe & """ 2>nul" & vbCrLf & _
                "timeout /t 2 /nobreak > nul" & vbCrLf & _
                "del """ & App.path & "\" & currentExe & """ 2>nul" & vbCrLf & _
                "powershell -Command ""Expand-Archive -Path '" & zipPath & "' -DestinationPath '" & App.path & "\Tools\Debug' -Force""" & vbCrLf & _
                "dir /b """ & App.path & "\Tools\Debug\*"" > debug.txt" & vbCrLf & _
                "for /f ""delims="" %%i in ('dir /b /o-d """ & App.path & "\Tools\Debug\*.exe"" 2^>nul') do (" & vbCrLf & _
                "    echo %%i >> debug.txt" & vbCrLf & _
                "    move """ & App.path & "\Tools\Debug\%%i"" """ & App.path & "\" & currentExe & """" & vbCrLf & _
                "    goto done" & vbCrLf & _
                ")" & vbCrLf & _
                ":done" & vbCrLf & _
                "start """" """ & App.path & "\" & currentExe & """" & vbCrLf & _
                "del """ & zipPath & """ 2>nul" & vbCrLf & _
                "del debug.txt 2>nul" & vbCrLf & _
                "del %0"
    Open batchPath For Output As #1
    Print #1, batchContent
    Close #1

    Shell batchPath, vbHide
    End
End Sub
' ========== HÀM L?Y VERSION M?I NH?T T? JSON ==========
Private Function GetLatestVersionFromJSON(ByVal json As String) As String
    Dim versions As Collection
    Dim latest As String
    Dim i As Integer
    
    Set versions = GetVersions(json)
    
    If versions.count > 0 Then
        ' L?y version cu?i cùng (m?i nh?t)
        latest = versions(versions.count)
    End If
    
    GetLatestVersionFromJSON = latest
End Function

' ========== HÀM L?Y FILENAME THEO VERSION ==========
Private Function GetFileNameByVersion(ByVal json As String, ByVal Version As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim versionPattern As String

    versionPattern = """" & Version & """" & ":"
    startPos = InStr(json, versionPattern)

    If startPos > 0 Then
        ' Tìm "fileName"
        startPos = InStr(startPos, json, """fileName""")
        If startPos > 0 Then
            startPos = InStr(startPos, json, ":")
            startPos = InStr(startPos, json, """") + 1
            endPos = InStr(startPos, json, """")
            If endPos > startPos Then
                GetFileNameByVersion = Mid(json, startPos, endPos - startPos)
            End If
        End If
    End If
End Function
Private Sub TaiToolAuto()
    On Error Resume Next  ' B? qua m?i l?i và ti?p t?c ch?y

    Dim downloaDriver As String
    Dim savePathDriver As String
    Dim unzipResult As Boolean

    ' Ki?m tra thu m?c t?n t?i
    If Dir("D:\ToolAuto\", vbDirectory) <> "" Then
        downloaDriver = "https://github.com/n93studioenter/SaovietGit/releases/download/latest-config/ToolAuto.zip"
        savePathDriver = "D:\ToolAuto.zip"
        ' T?i file
        If DownloadFileWithWinHTTP(downloaDriver, savePathDriver, 60) Then
            ' Gi?i nén
            UnzipFile savePathDriver, "D:\ToolAuto"

            'gan taskbar
            'PinToTaskbar "D:\ToolAuto\ToolTaiHD.exe"
        End If
    Else
        If Dir("D:\Ke toan 2025 New\ToolAuto\", vbDirectory) <> "" Then
            downloaDriver = "https://github.com/n93studioenter/SaovietGit/releases/download/latest-config/ToolAuto.zip"
            savePathDriver = "D:\Ke toan 2025 New\ToolAuto.zip"
            ' T?i file
            If DownloadFileWithWinHTTP(downloaDriver, savePathDriver, 60) Then
                ' Gi?i nén
                UnzipFile savePathDriver, "D:\Ke toan 2025 New\ToolAuto"
            End If
            Exit Sub
        End If
        downloaDriver = "https://github.com/n93studioenter/SaovietGit/releases/download/latest-config/ToolAutoFull.zip"
        savePathDriver = "D:\ToolAuto.zip"
        ' T?i file
        If DownloadFileWithWinHTTP(downloaDriver, savePathDriver, 60) Then
            ' Gi?i nén
            UnzipFile savePathDriver, "D:\ToolAuto"
            'PinToTaskbar "D:\ToolAuto\ToolTaiHD.exe"
        End If
    End If

    On Error GoTo 0  ' T?t ch? d? b? qua l?i
    ' Không có b?t k? thông báo hay d?u hi?u nào du?c hi?n th?
End Sub
Private Sub PinToTaskbar(ByVal ExeFullPath As String)
    Dim fso As Object
    Dim WshShell As Object
    Dim ts As Object
    Dim ps1Path As String
    Dim AppName As String
    Dim Script As String
    Dim TaskBarFolder As String
    Dim file As Object
    Dim AlreadyPinned As Boolean
    Dim sc As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set WshShell = CreateObject("WScript.Shell")

    AppName = fso.GetBaseName(ExeFullPath)
    AlreadyPinned = False

    ' === Ki?m tra dã ghim chua ===
    TaskBarFolder = WshShell.ExpandEnvironmentStrings("%AppData%") & _
                    "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"

    If fso.FolderExists(TaskBarFolder) Then
        For Each file In fso.GetFolder(TaskBarFolder).Files
            If LCase(fso.GetExtensionName(file.Name)) = "lnk" Then
                On Error Resume Next
                Set sc = WshShell.CreateShortCut(file.path)
                If LCase(sc.TargetPath) = LCase(ExeFullPath) Then
                    AlreadyPinned = True
                End If
                Set sc = Nothing
                On Error GoTo 0
            End If
            If AlreadyPinned Then Exit For
        Next
    End If

    ' N?u dã ghim r?i thì thoát luôn
    If AlreadyPinned Then
        MsgBox "Ðã du?c ghim vào Taskbar r?i.", vbInformation
        Set fso = Nothing
        Set WshShell = Nothing
        Exit Sub
    End If

    ' === Chua ghim thì m?i ch?y ===
    ps1Path = Environ("TEMP") & "\PinTaskbar_" & AppName & ".ps1"

    Script = "$exe = '" & ExeFullPath & "'" & vbCrLf
    Script = Script & "$name = '" & AppName & "'" & vbCrLf
    Script = Script & "$lnk = ""$env:APPDATA\Microsoft\Windows\Start Menu\Programs\$name.lnk""" & vbCrLf
    Script = Script & "$Wsh = New-Object -ComObject WScript.Shell" & vbCrLf
    Script = Script & "$s = $Wsh.CreateShortcut($lnk)" & vbCrLf
    Script = Script & "$s.TargetPath = $exe" & vbCrLf
    Script = Script & "$s.WorkingDirectory = [IO.Path]::GetDirectoryName($exe)" & vbCrLf
    Script = Script & "$s.IconLocation = $exe" & vbCrLf
    Script = Script & "$s.Save()" & vbCrLf
    Script = Script & "$xml = @""" & vbCrLf
    Script = Script & "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
    Script = Script & "<LayoutModificationTemplate xmlns=""http://schemas.microsoft.com/Start/2014/LayoutModification""" & vbCrLf
    Script = Script & " xmlns:defaultlayout=""http://schemas.microsoft.com/Start/2014/FullDefaultLayout""" & vbCrLf
    Script = Script & " xmlns:start=""http://schemas.microsoft.com/Start/2014/StartLayout""" & vbCrLf
    Script = Script & " xmlns:taskbar=""http://schemas.microsoft.com/Start/2014/TaskbarLayout"" Version=""1"">" & vbCrLf
    Script = Script & "<CustomTaskbarLayoutCollection PinListPlacement=""Replace"">" & vbCrLf
    Script = Script & "<defaultlayout:TaskbarLayout>" & vbCrLf
    Script = Script & "<taskbar:TaskbarPinList>" & vbCrLf
    Script = Script & "<taskbar:DesktopApp DesktopApplicationLinkPath=""$lnk""/>" & vbCrLf
    Script = Script & "</taskbar:TaskbarPinList>" & vbCrLf
    Script = Script & "</defaultlayout:TaskbarLayout>" & vbCrLf
    Script = Script & "</CustomTaskbarLayoutCollection>" & vbCrLf
    Script = Script & "</LayoutModificationTemplate>" & vbCrLf
    Script = Script & """@" & vbCrLf
    Script = Script & "$xmlPath = ""$env:LOCALAPPDATA\Microsoft\Windows\Shell\LayoutModification.xml""" & vbCrLf
    Script = Script & "Set-Content -Path $xmlPath -Value $xml -Encoding UTF8" & vbCrLf
    Script = Script & "Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue" & vbCrLf
    Script = Script & "Start-Sleep -Seconds 2" & vbCrLf
    Script = Script & "Remove-Item ""$env:APPDATA\Microsoft\Windows\Shell\*.dat"" -Force -ErrorAction SilentlyContinue" & vbCrLf
    Script = Script & "Remove-Item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband' -Recurse -Force -ErrorAction SilentlyContinue" & vbCrLf
    Script = Script & "Start-Process explorer.exe" & vbCrLf

    Set ts = fso.CreateTextFile(ps1Path, True)
    ts.Write Script
    ts.Close

    WshShell.Run "powershell -NoProfile -ExecutionPolicy Bypass -File """ & ps1Path & """", 0, True

    On Error Resume Next
    fso.DeleteFile ps1Path, True
    On Error GoTo 0

    Set ts = Nothing
    Set WshShell = Nothing
    Set fso = Nothing

    MsgBox "Ðã ghim vào Taskbar!", vbInformation
End Sub

Private Function GetExeName(proc As PROCESSENTRY32) As String
    Dim pos As Integer
    pos = InStr(1, proc.szExeFile, vbNullChar)
    If pos > 0 Then
        GetExeName = RTrim(Left(proc.szExeFile, pos - 1))
    Else
        GetExeName = RTrim(proc.szExeFile)
    End If
End Function
Public Sub CloseProcessByName(ByVal ExeName As String)
    Dim hSnapshot As Long
    Dim proc As PROCESSENTRY32
    Dim hProcess As Long
    Dim found As Long
    Dim sName As String
    Dim isFound As Boolean
    
    ' T?o snapshot
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If hSnapshot = 0 Then
        MsgBox "Không th? t?o snapshot!", vbCritical
        Exit Sub
    End If
    
    ' ?? QUAN TR?NG: Gán kích thu?c C? Ð?NH
    proc.dwSize = 296
    
    found = Process32First(hSnapshot, proc)
    isFound = False
    
    Do While found <> 0
        sName = GetExeName(proc)
        
        ' Debug: In ra danh sách
        ' Debug.Print sName
        
        If UCase(sName) = UCase(ExeName) Then
            isFound = True
            hProcess = OpenProcess(PROCESS_TERMINATE, False, proc.th32ProcessID)
            
            If hProcess <> 0 Then
                If TerminateProcess(hProcess, 0) <> 0 Then
                    Debug.Print "? Ðã dóng: " & sName & " (PID: " & proc.th32ProcessID & ")"
                Else
                    Debug.Print "? Không th? dóng: " & sName
                End If
                CloseHandle hProcess
            Else
                Debug.Print "? Không th? m? ti?n trình: " & sName
            End If
        End If
        
        found = Process32Next(hSnapshot, proc)
    Loop
    
    CloseHandle hSnapshot
    
    If Not isFound Then
        Debug.Print "? Không tìm th?y: " & ExeName
    End If
End Sub

' ============================================
' HÀM KI?M TRA TI?N TRÌNH
' ============================================
Public Function IsProcessRunning(ByVal ExeName As String) As Boolean
    Dim hSnapshot As Long
    Dim proc As PROCESSENTRY32
    Dim found As Long
    Dim sName As String

    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If hSnapshot = 0 Then
        IsProcessRunning = False
        Exit Function
    End If

    proc.dwSize = 296
    found = Process32First(hSnapshot, proc)

    Do While found <> 0
        sName = GetExeName(proc)
        If UCase(sName) = UCase(ExeName) Then
            IsProcessRunning = True
            CloseHandle hSnapshot
            Exit Function
        End If
        found = Process32Next(hSnapshot, proc)
    Loop

    CloseHandle hSnapshot
    IsProcessRunning = False
End Function
Private Sub Form_Load()
'Kiem tra xem co file versiontool.txt
    Dim vbpath As String
    Dim fileName As String
    Dim FileNum As Integer
    Dim fullPath As String

    ' Ðu?ng d?n thu m?c c?n ki?m tra
    vbpath = App.path   ' Thay d?i du?ng d?n c?a b?n

    ' T?o du?ng d?n d?y d? d?n file
    fileName = vbpath & "\Hoadon\versiontool.txt"
    If Dir(fileName) <> "" Then
    Else
        ' T?o file m?i
        FileNum = FreeFile
        Open fileName For Output As #FileNum
        Print #FileNum, "Version: 1.0"  ' Ghi n?i dung m?c d?nh
        Print #FileNum, "Created: " & Now
        Print #FileNum, "This file was automatically created."
        Close #FileNum
        TaiToolAuto
    End If


    KiemtraversionGitHub2
    If hasUPdate Then
        Exit Sub
    End If
    'menuversion
    'KiemtraversionDrive
    'DownloadAndRun
    ExecuteSQL5 "DELETE FROM tbGetphieu WHERE MaCT = '0'"

    ExecuteSQL5_Themmoi ("ALTER TABLE ChungTu ADD tencongty text")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbRegister ADD tk155 text")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbCpu ADD PcName text")
    frmMain.sbStatusBar.Panels(4).ToolTipText = "Log On Time: " + Format(Time, "hh:mm:ss")
    Label5.Caption = "  (" & ABCtoVNI("Phiªn b¶n dïng thö") & ")"
    'PopMenu1.BindMenu Me.hwnd, Me.Menu
    'MsgBox TypeName(PopMenu1)

    'Call BackupSysFont
    'Call SetMenuFont("Tahoma", 11, True)
    'RestoreSysFont
    'RestoreSysFont
    ' Fake menu bar
    ' picFakeMenu.Align = vbAlignTop
    'picFakeMenu.Height = 360
    'picFakeMenu.BackColor = CLR_MENU_NORMAL

    'InitFakeMenu
    'gCurrentMenu = -1
    ExecuteSQL5_Themmoi ("ALTER TABLE license  ADD AuToNK NUMBER")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbimport  ADD Khautruthue NUMBER")
    ExecuteSQL5_Themmoi ("ALTER TABLE ChungTu  ADD SoPSGoc NUMBER")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbimportdetail  ADD SoPSGoc NUMBER")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbimportdetail  ADD VAT NUMBER")
    Dim check162 As String
    check162 = SelectSQL("SELECT SoHieu AS F1 FROM HeThongTK where SoHieu = '621' ")
    If check162 = 0 Then

        'mnTT.Caption = "TT-46/2025-BTC"
    Else
        'mnTT.Caption = "TT-99/2025-BTC"
        'mnTT.Visible = False
    End If
    mnTT.Caption = "TT133/2016/TT-BTC"
    LoadMenuform
    Kiemtraphienbanht
    'Taifilecapnhat


    Dim X1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer

    'ResetSystemFontToDefault
    'ChangeFont
    ' L?y handle d?n menu c?a form
    If pVersion = 2 Then
        Label(19).Visible = False
        '  img.Top = 3120
        '  img.Left = 360
    End If

    If pVersion > 0 Then
        If Len(Dir(pCurDir + "LOGO.JPG")) > 0 Then
            On Error Resume Next
            Set img.Picture = LoadPicture(pCurDir + "LOGO.JPG")
            On Error GoTo 0
            If img.Picture <> 0 Then
                If pVersion <> 2 Then
                    X1 = GetSetting(IniPath, "Logo", "X1", 0)
                    y1 = GetSetting(IniPath, "Logo", "Y1", 0)
                    If X1 <> 0 Then
                        img.Left = X1
                        img.Top = y1
                    End If
                    x2 = GetSetting(IniPath, "Logo", "X2", 0)
                    y2 = GetSetting(IniPath, "Logo", "Y2", 0)
                    If x2 <> 0 Then
                        img.Width = x2
                        img.Height = y2
                    End If
                End If
                img.Visible = True
            End If
        End If
    End If

    HienThongBao pDataPath, 2
    '  dlgCommonDialog.InitDir = pCurDir + "DATA"


    GetLicense

    LietKeTep

    On Error Resume Next
    ' Rpt.WindowShowPrintSetupBtn = True
    '   Rpt.WindowShowGroupTree = True
    '  Rpt.WindowShowSearchBtn = True
    ' Rpt.WindowShowZoomCtl = True
    On Error GoTo 0

    setMDSettings

    Select Case pProcessMode
    Case 2: pProcessEnable = True
        Me.Caption = Me.Caption + " - SERVER Application"
        CTTimer.Enabled = True
    Case 1: pProcessEnable = False
        Me.Caption = Me.Caption + " - CLIENT Application"
    End Select

    Mask_D = GetShortDateFormat

    'chutieng_viet
    ExecuteSQL5 "UPDATE HeThongTK set MaTC = MaSo where MaTC <> MaSo"
    ExecuteSQL_them_bang "SoLoThuoc1"
    ExecuteSQL_them_bang "SoLoThuoc"

    ExecuteSQL_them_query "VatTuNhap", "SELECT mavattu, solo, handung, sum(SoPS2No) AS soluong" _
                                     & " From chungtu " _
                                     & " Where maloai = 1 and len(solo) > 0 " _
                                     & " GROUP BY mavattu, solo, handung " _
                                     & " UNION select SOLOTHUOC.mavattu,SOLOTHUOC.solo,SOLOTHUOC.handung, SOLOTHUOC.soluong " _
                                     & " from SOLOTHUOC "
    ExecuteSQL_them_query "VatTuXuat", "SELECT mavattu, solo, handung, sum(SoPS2Co) AS soluong FROM chungtu WHERE MaLoai=2 and len(solo) > 0 GROUP BY mavattu, solo, handung"

    Dim sqqq As String
    sqqq = "SELECT a.* " _
         & " FROM (SELECT vattu.maso AS mavattu, VATTU.SOHIEU, VATTU.TenVattu, IIf(IsNull(VatTuNhap.SOLO),' ',VatTuNhap.SOLO) AS solo, IIf(IsNull(VatTuNhap.Handung),' ',VatTuNhap.Handung) AS Handung, VatTuNhap.soluong AS soluongnhap, IIf(IsNull(VatTuXuat.soluong),0,VatTuXuat.soluong) AS soluongxuat, iif(isnull(VatTuNhap.soluong),0,VatTuNhap.soluong)-IIf(IsNull(VatTuXuat.soluong),0,VatTuXuat.soluong) AS conlai FROM (VATTU LEFT JOIN VatTuNhap ON VatTuNhap.MAVATTU=VATTU.MASO) LEFT JOIN VatTuXuat ON (VatTuNhap.mavattu=VatTuXuat.mavattu) AND (VatTuNhap.solo=VatTuXuat.solo) AND (VatTuNhap.handung=VatTuXuat.handung))  AS a " _
         & " Where a.conlai > 0 " _
         & " ORDER BY a.handung "
    ExecuteSQL_them_query "DanhSachVatTu", sqqq


    LbCty(4).Visible = False
    'If Year(DateTime.Date) < 2018 Then
    'Label3(12).Caption = "§¬n vÞ triÓn khai: Lª V¨n L¸y"
    'Label3(13).Caption = "Sè ®iÖn tho¹i: 093 3415 959"
    'End If
    ban_quyen = 0
    If (boolean_kiemtra() = False) Then
        'frmMain.txtdungthu.Caption = ABCtoVNI("PhÇn mÒm hÕt h¹n dïng, vui lßng liªn hÖ víi nhµ cung cÊp!")
        'frmMain.txtdungthu.Caption = "PhÇn mÒm hÕt h¹n dïng, vui lßng liªn hÖ víi nhµ cung cÊp!"
        Label5.Caption = "  (" & ABCtoVNI("PhÇn mÒm hÕt h¹n dïng") & ")"
        Dim demslchungtu As Double
        demslchungtu = SelectSQL( _
                       "SELECT COUNT(*) AS F1 " & _
                       "FROM (" & _
                     " SELECT MaCT FROM ChungTu " & _
                     " WHERE SoHieu NOT LIKE '*GV*' " & _
                     " GROUP BY MaCT" & _
                       ") AS T" _
                     )
        Dim demdoanhthu As Double
        demdoanhthu = SelectSQL("SELECT sum(duco_12) as F1 from hethongtk where sohieu ='511' ")
        If (demslchungtu > 200 Or demdoanhthu > 2000000000#) Then
            ban_quyen = 1
        Else
            frmMain.txtdungthu.Caption = ""
        End If
    End If
    Loadactive
End Sub
Public Sub Loadactive()

    Dim counttkinvoice As Long
    counttkinvoice = SelectSQL("select count(*) AS f1 from tbInvoiceInfo")

    If counttkinvoice = 1 Then
        Label3(18).Visible = True
        LbCty(15).Visible = True
        Label3(19).Visible = True
        Label3(20).Visible = True
        LbCty(16).Visible = True
        LbCty(17).Visible = True
        LbCty(15).Caption = SelectSQL("select Url AS f1 from  tbInvoiceInfo")
        LbCty(16).Caption = SelectSQL("select Username AS f1 from  tbInvoiceInfo")
        LbCty(17).Caption = SelectSQL("select Password AS f1 from  tbInvoiceInfo")
    End If
    CheckAndCreateTBCpu
    ExecuteSQL5_Themmoi ("ALTER TABLE HoaDon ADD IdNhap text")
    ExecuteSQL5_Themmoi ("ALTER TABLE HoaDon ADD IdTemplate text")
    ExecuteSQL5_Themmoi ("ALTER TABLE HoaDon ADD StatusPH text")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbCpu ADD PcName text")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbRegister ADD Printername text")
    ExecuteSQL5_Themmoi ("ALTER TABLE tbInvoiceTemplate ADD KHHD text")
    Dim cmg As Long
    cmg = SelectSQL("select CMG AS f1 from  License")
    If cmg = 249991 Then
        'Label5.Visible = True

    Else
        Label5.Visible = False
    End If
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtu  ADD nhanban number")
    'Tudongtinhgiavon = True
    'Kiemtraphienban
    ' FindLatestExe
    Label1.Left = 0
    Label1.Top = (Me.ScaleHeight * 95 / 100)


    Image1.Left = (Me.ScaleWidth * 87 / 100)
    Label3(21).Left = Image1.Left + Image1.Width + 40
    Image1.Top = (Me.ScaleHeight * 5 / 100)
     Label3(21).Top = (Me.ScaleHeight * 5 / 100) + 1250
     
    'Command1.Left = (Me.ScaleWidth * 90 / 100)
    Command1.Left = Me.ScaleWidth * 0.957 - Command1.Width

    Command1.Top = (Me.ScaleHeight * 80 / 100)
    Label3(0).Alignment = vbRightJustify
    Label3(0).Left = Me.ScaleWidth * 0.957 - Label3(0).Width

    Label3(0).Top = (Me.ScaleHeight * 85 / 100)
    Label3(16).Alignment = vbRightJustify
    Label3(16).Left = Me.ScaleWidth * 0.96 - Label3(16).Width
    Label3(16).Top = (Me.ScaleHeight * 88 / 100)


    ExecuteSQL5_Themmoi ("ALTER TABLE HeThongTK DROP COLUMN KyHieu")
    ExecuteSQL5_Themmoi ("ALTER TABLE ChungTu  ADD NgayImport Datetime")
    ExecuteSQL5_Themmoi ("ALTER TABLE license  ADD tenhoadon text")
    ExecuteSQL5 ("ALTER TABLE license ALTER COLUMN TaiKhoanVN TEXT(200)")
    ExecuteSQL5 ("ALTER TABLE license ALTER COLUMN DiaChi TEXT(255)")
    ExecuteSQL5 ("ALTER TABLE license ALTER COLUMN FAX TEXT(200)")
    ExecuteSQL5_Themmoi ("ALTER TABLE license  ADD col711 text")
    ExecuteSQL5_Themmoi ("ALTER TABLE license  ADD col711ra text")
    ExecuteSQL5_Themmoi ("ALTER TABLE license  ADD skiperror text")
    ExecuteSQL5 "Update License set skiperror=0 where skiperror='...'"
    ' ExecuteSQL5 ("ALTER TABLE Vattu ALTER COLUMN TenVattu MEMO")
    'ExecuteSQL5 ("UPDATE HOADON SET KyHieu = '01GTKT3/001' WHERE KYHIEU = '...'")
    'mnDuLieu.Caption = "Xö lý"

    StationList
End Sub
Public Function ExecuteSQL_them_query(Ten As String, sql As String, Optional msg As Boolean = True) As Integer
    On Error GoTo ErrLock
    DBKetoan.CreateQueryDef Ten, sql
    Debug.Print sql
    On Error GoTo 0
    ExecuteSQL_them_query = 0
    Exit Function
ErrLock:
    'MsgBox Err.Description
End Function
Public Function ExecuteSQL_them_bang(Ten As String, Optional msg As Boolean = True) As Integer
      On Error GoTo ErrLock
      ExecuteSQL5_Themmoi ("create table " + Ten + " (MaVatTu number,SoLo Text,HanDung datetime,SoLuong Number)")
      On Error GoTo 0
      ExecuteSQL_them_bang = 0
      Exit Function
ErrLock:
'MsgBox Err.Description
End Function
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command(1).BackColor = &HC000&    ' &H808000
    Command(0).BackColor = &HC000&    '&H808000
    Command(2).BackColor = &HC000&    '&H808000
    Command1.BackColor = &HC0FFFF
    
    Dim msg As Long
    msg = X / Screen.TwipsPerPixelX
End Sub
Private Sub ShowApp()
        ' ? HI?N FORM B?NG API
    ShowWindow Me.hwnd, SW_RESTORE  ' Restore n?u dang minimize
    ShowWindow Me.hwnd, SW_SHOW     ' Hi?n form
    
    Me.WindowState = vbNormal
    Me.SetFocus
    SetForegroundWindow Me.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call BackupSysFont
    'kiem tra he thong tai khoan truoc
    'KiemTraTaiKhoan
    Screen.MousePointer = 11
    HienThongBao "KÕt thóc ch­¬ng tr×nh kÕ to¸n!", 1

    CloseUp
    Recycle pCurDir + "*.BMP"

    If pVersion > 0 And img.Visible Then
        SaveSetting IniPath, "Logo", "X1", img.Left
        SaveSetting IniPath, "Logo", "Y1", img.Top
        SaveSetting IniPath, "Logo", "X2", img.Width
        SaveSetting IniPath, "Logo", "Y2", img.Height
    End If


    m_nonClientMetrics.lfCaptionFont.lfFaceName = m_fontCaption
    m_nonClientMetrics.lfCaptionFont.lfFaceName = m_fontSmCaption
    m_nonClientMetrics.lfMenuFont.lfFaceName = m_fontMenu
    m_nonClientMetrics.lfMessageFont.lfFaceName = m_fontMessage
    m_nonClientMetrics.lfStatusFont.lfFaceName = m_fontStatus

    ret = SystemParametersInfo(SPI_SETNONCLIENTMETRICS, Len(m_nonClientMetrics), m_nonClientMetrics, 0)
    ' m_logFont.lfFaceName = m_fontIcon
    'ret = SystemParametersInfo(SPI_SETICONTITLELOGFONT, Len(m_logFont), m_logFont, 0)

    Recycle pCurDir + "DATA\backup1.MDB"

    SetPsw pDataPath, pPSW, ""
    On Error Resume Next
    DBEngine.CompactDatabase pDataPath, pCurDir + "DATA\backup1.MDB"

    On Error GoTo 0
    If Len(Dir(pCurDir + "DATA\backup1.MDB")) > 0 Then
        Recycle pDataPath
        FileCopy pCurDir + "DATA\backup1.MDB", pDataPath
        SetPsw pCurDir + "DATA\backup1.MDB", "", pPSW
    End If
    'pPSW = "1@35^7*9)"
    SetPsw pDataPath, "", pPSW
    '========================

    ' Recycle pCurDir + "DATA\AJZIP.MDB"

    Recycle pCurDir + "DATA\backup2.MDB"

    SetPsw pDataPath, pPSW, ""
    On Error Resume Next
    '  DBEngine.CompactDatabase pDataPath, pCurDir + "DATA\AJZIP.MDB"
    DBEngine.CompactDatabase pDataPath, pCurDir + "DATA\backup2.MDB"

    On Error GoTo 0
    If Len(Dir(pCurDir + "DATA\backup2.MDB")) > 0 Then
        Recycle pDataPath
        FileCopy pCurDir + "DATA\backup2.MDB", pDataPath
        SetPsw pCurDir + "DATA\backup2.MDB", "", pPSW
    End If
    SetPsw pDataPath, "", pPSW
    '========================
    restoreSettings

    Screen.MousePointer = 0

    End
    Set App = Nothing

End Sub


Private Sub Image2_Click()
    Screen.MousePointer = vbHourglass
    Dim i As Integer
    ProgressBar1.Visible = True
    Label4.Visible = True
    Label4.Caption = "Update..."
    DoEvents
    'Taifilecapnhat
    'Capnhatdata
    'GetlastVersion
    Label4.Visible = False
    Dim fileNumber As Integer
    Dim FilePath As String
    Dim content As String

    'Cap nhat file version
    FilePath = App.path & "\\HoaDon\\\version.txt"
    fileNumber = FreeFile
    content = "1.1"
    On Error Resume Next
    Open FilePath For Output As #fileNumber
    If Err.number = 0 Then
        Print #fileNumber, content;
        Close #fileNumber
        'MsgBox "Ðã ghi dè file version.txt thành công!", vbInformation
    Else
        MsgBox "L?i khi ghi dè file!", vbExclamation
    End If
    KiemtraversionGitHub3
End Sub
Public Sub GetlastVersion()
    Dim fileNumber As Integer
    Dim FilePath As String
    Dim content As String
    FilePath = App.path & "\\HoaDon\\\version.txt"
    fileNumber = FreeFile
    On Error Resume Next
    Open FilePath For Output As #fileNumber
    If Err.number = 0 Then
        Print #fileNumber, "1.0";
        Close #fileNumber
        'MsgBox "Ðã ghi dè file version.txt thành công!", vbInformation
    Else
        MsgBox "L?i khi ghi dè file!", vbExclamation
    End If
    DownloadAndRun
End Sub
Public Sub Capnhatdata()
    Dim fileNumber As Integer
    Dim FilePath As String
    Dim content As String
    FilePath = App.path & "\\HoaDon\\\serverpath.txt"
    content = "\\192.168.1.90\Ke toan 2025 New\1 Copi vao dung 1"
    fileNumber = FreeFile
    On Error Resume Next
    Open FilePath For Output As #fileNumber
    If Err.number = 0 Then
        Print #fileNumber, content;
        Close #fileNumber
        'MsgBox "Ðã ghi dè file version.txt thành công!", vbInformation
    Else
        MsgBox "L?i khi ghi dè file!", vbExclamation
    End If

    'Cap nhat file version
    FilePath = App.path & "\\HoaDon\\\version.txt"
    fileNumber = FreeFile
    On Error Resume Next
    Open FilePath For Output As #fileNumber
    If Err.number = 0 Then
        Print #fileNumber, FindItemByID(Index);
        Close #fileNumber
        'MsgBox "Ðã ghi dè file version.txt thành công!", vbInformation
    Else
        MsgBox "L?i khi ghi dè file!", vbExclamation
    End If
    CopyAndRun

End Sub

Private Sub lbCty_Click(Index As Integer)
    Clipboard.Clear
    Clipboard.SetText LbCty(Index).Caption
End Sub

Public Sub mnCn_Click(Index As Integer)
    If Index = 12 Then
        Dim sMsg As String
        Dim sCap As String
        Dim response As Long

        sMsg = _
        ChrW(66) & ChrW(7841) & ChrW(110) & " " & _
               ChrW(99) & ChrW(243) & " " & _
               ChrW(109) & ChrW(117) & ChrW(7889) & " " & _
               ChrW(120) & ChrW(111) & ChrW(225) & " " & _
               ChrW(116) & ChrW(7845) & ChrW(116) & " " & _
               ChrW(99) & ChrW(7843) & " " & _
               ChrW(99) & ChrW(244) & ChrW(110) & ChrW(103) & " " & _
               ChrW(110) & ChrW(7907) & " " & _
               ChrW(107) & ChrW(104) & ChrW(244) & ChrW(110) & ChrW(103) & " " & _
               ChrW(99) & ChrW(243) & " " & _
               ChrW(112) & ChrW(104) & ChrW(225) & ChrW(116) & " " & _
               ChrW(115) & ChrW(105) & ChrW(110) & ChrW(104) & "?"
        sCap = _
        ChrW(88) & ChrW(225) & ChrW(99) & " " & _
               ChrW(110) & ChrW(104) & ChrW(7853) & ChrW(110) & " " & _
               ChrW(120) & ChrW(111) & ChrW(225)
        ' xoá
        response = MessageBoxW(Me.hwnd, StrPtr(sMsg), StrPtr(sCap), vbYesNo Or vbQuestion)
        'Dim response As Integer
        'response = MsgBox("B¹n cã muèn xo¸ tÊt c¶ c«ng nî kh«ng cã  ph¸t sinh ?", vbYesNo + vbQuestion, "X¸c nhËn xo¸")
        If response = vbYes Then
            'Kiem tra cong no khach hang

            Dim sql As String
            sql = "Delete FROM KhachHang  WHERE MaSo NOT IN (SELECT MaKhachHang FROM HoaDon) AND MaSo NOT IN (SELECT MaKhachHang FROM SoDuKhachHang)"
            ExecuteSQL5 sql

            sql = "DELETE FROM SoDuKhachHang " & _
                  "WHERE MaKhachHang IN (" & _
                "    SELECT MaKhachHang " & _
                "    FROM SoDuKhachHang sdk " & _
                "    WHERE (" & _
                "        No_1 = 0 AND Co_1 = 0 " & _
                "        AND No_2 = 0 AND Co_2 = 0 " & _
                "        AND No_3 = 0 AND Co_3 = 0 " & _
                "        AND No_4 = 0 AND Co_4 = 0 " & _
                "        AND No_5 = 0 AND Co_5 = 0 " & _
                "        AND No_6 = 0 AND Co_6 = 0 " & _
                "        AND No_7 = 0 AND Co_7 = 0 " & _
                "        AND No_8 = 0 AND Co_8 = 0 " & _
                "        AND No_9 = 0 AND Co_9 = 0 " & _
                "        AND No_10 = 0 AND Co_10 = 0 " & _
                "        AND No_11 = 0 AND Co_11 = 0 " & _
                "        AND No_12 = 0 AND Co_12 = 0 " & _
                "    ) " & _
                "    AND DuNo_12 + DuCo_12 = 0 " & _
                "    AND NOT EXISTS (SELECT 1 FROM HoaDon hd WHERE hd.MaKhachHang = sdk.MaKhachHang) " & _
                "    AND NOT EXISTS (SELECT 1 FROM ChungTu ct WHERE ct.MaKH = sdk.MaKhachHang) " & _
                "    AND NOT EXISTS (SELECT 1 FROM ChungTu ct WHERE ct.MaKHC = sdk.MaKhachHang) " & _
                "    AND NOT EXISTS (SELECT 1 FROM ChungTuLQ ctlq WHERE ctlq.MaKH = sdk.MaKhachHang) " & _
                  ")"

            ExecuteSQL5 sql
            Dim ss As String
            ss = ChrW(88) & ChrW(111) & ChrW(225) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(224) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(99) & ChrW(244) & ChrW(110) & ChrW(103)
            MessageBoxW Me.hwnd, StrPtr(ss), StrPtr("Thông báo"), vbOKOnly

            sql = "DELETE FROM KhachHang " & _
                  "WHERE NOT EXISTS (SELECT 1 FROM HoaDon WHERE MaKhachHang = KhachHang.MaSo) " & _
                  "AND NOT EXISTS (SELECT 1 FROM ChungTu WHERE MaKH = KhachHang.MaSo) " & _
                  "AND NOT EXISTS (SELECT 1 FROM ChungTu WHERE MaKHC = KhachHang.MaSo) " & _
                  "AND NOT EXISTS (SELECT 1 FROM ChungTuLQ WHERE MaKH = KhachHang.MaSo) " & _
                  "AND MaSo IN ( " & _
                "    SELECT MaKhachHang " & _
                "    FROM SoDuKhachHang " & _
                "    WHERE No_1 = 0 AND Co_1 = 0 " & _
                "       AND No_2 = 0 AND Co_2 = 0 " & _
                "       AND No_3 = 0 AND Co_3 = 0 " & _
                "       AND No_4 = 0 AND Co_4 = 0 " & _
                "       AND No_5 = 0 AND Co_5 = 0 " & _
                "       AND No_6 = 0 AND Co_6 = 0 " & _
                "       AND No_7 = 0 AND Co_7 = 0 " & _
                "       AND No_8 = 0 AND Co_8 = 0 " & _
                "       AND No_9 = 0 AND Co_9 = 0 " & _
                "       AND No_10 = 0 AND Co_10 = 0 " & _
                "       AND No_11 = 0 AND Co_11 = 0 " & _
                "       AND No_12 = 0 AND Co_12 = 0 " & _
                "    GROUP BY MaKhachHang " & _
                "    HAVING SUM(DuNo_12 + DuCo_12) = 0 " & _
                  ")"
            ExecuteSQL5 sql
            'MsgBox "Xo¸ thµnh c«ng !"
            Dim s As String
            s = ChrW(88) & ChrW(111) & ChrW(225) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(224) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(99) & ChrW(244) & ChrW(110) & ChrW(103)
            MessageBoxW Me.hwnd, StrPtr(s), StrPtr("Thông báo"), vbOKOnly

        End If
        Exit Sub
    End If
    If Index = 3 Or Index = 9 Then
        If Not KtraMKAdmin Then Exit Sub
    End If
    Select Case Index
    Case 0:
        frmPhanLoaiVT.tag = 2
        frmPhanLoaiVT.Show 1
    Case 1:
        FrmKhachHang.Show vbModal
    Case 3:
        If ChoDieuChinhDauKy Then
            If pCongNoHD = 0 Then
                FKHDauKy.Show vbModal
            Else
                FKHDauKy2.Show vbModal
            End If
        End If
    Case 4:
        FrmHD.Show vbModal
    Case 6:
        frmPhanLoaiVT.tag = 4
        frmPhanLoaiVT.Show 1
    Case 7:
        FrmNhanVien.Show 1
    Case 9:
        If KtraMKAdmin Then FrmLS.Show 1
    Case 11:
        If KtraMKAdmin Then DatTKCN
    End Select
    HienThongBao "", 1
End Sub

Private Sub mnDL_Click(Index As Integer)

    If Index = 1 Then
        Dim strSQL As String
        ' Cú pháp cho MS Access
        strSQL = "UPDATE HoaDon SET MatHang = UCase(Mid(MatHang, 1, 1)) & Mid(MatHang, 2)"
        ExecuteSQL5 strSQL
    End If


    Dim sql As String

    If User_Right <> 0 Or (Me.tag Mod 10 = 0) Or (User_Right = 2) Then
        NoRight 0
        Exit Sub
    End If
    Me.MousePointer = 11


    Select Case Index
    Case 0:
        If Not STDetail Then
            NoRight 1
            GoTo KT
        End If
        KiemTraVatTu
        '            Dim i  As Integer
        '            Dim rs As Recordset
        '            Set rs = DBKetoan.OpenRecordset("SELECT mavattu,sum(luong_0) as luong from VTdaunam ", dbOpenSnapshot)
        '            For i = 0 To rs.RecordCount
        '            ExecuteSQL5 ("update solothuoc set ")
        '            Next
        '
        '            SoLoThuoc

    Case 20:

        FrmNguyente.Show 1

    Case 1:
        KiemTraTaiKhoan

    Case 3:
        If FPsw.GetPswX() = "UCDIT" Then
            sql = FrmGetStr.GetString("LÖnh xö lý:", App.ProductName)
            If Len(sql) > 0 Then ExecuteSQL5 sql
        End If
    Case 6:
        Dim rs_ktra As Recordset
        Dim Query As String
        Dim rst As String
        Query = "SELECT *  FROM tbLicensekey "
        Set rs_ktra = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
        If Not rs_ktra.EOF Then
            ' Duy?t qua t?t c? các b?n ghi
            Do While Not rs_ktra.EOF
                If rs_ktra!Type = 2 Then
                    Dim resultArray() As String
                    resultArray = Split(rs_ktra!Year, "|")
                    Dim Chk As Integer
                    Chk = (CInt(resultArray(0)) - 1) + CInt(resultArray(1)) - pNamTC
                    If Chk <= 0 Then
                        'MsgBox "Gãi d÷ liÖu theo n¨m ®· hÕt, vui lßng liªn hÖ ®Ó ®­îc chuyÓn sang n¨m míi"
                        Dim s As String
                        s = ChrW(71) & ChrW(243) & ChrW(105) & ChrW(32) & ChrW(100) & ChrW(7919) & ChrW(32) & ChrW(108) & ChrW(105) & ChrW(7879) & ChrW(117) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(101) & ChrW(111) & ChrW(32) & ChrW(110) & ChrW(259) & ChrW(109) & ChrW(32) & ChrW(273) & ChrW(227) & ChrW(32) & ChrW(104) & ChrW(7871) & ChrW(116) & ChrW(44) & ChrW(32) & ChrW(118) & ChrW(117) & ChrW(105) & ChrW(32) & ChrW(108) & ChrW(242) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(108) & ChrW(105) & ChrW(234) & ChrW(110) & ChrW(32) & ChrW(104) & ChrW(7879) & ChrW(32) & ChrW(273) & ChrW(7875) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(117) & ChrW(121) & ChrW(7875) & ChrW(110) & ChrW(32) & ChrW(115) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(110) & ChrW(259) & ChrW(109) & ChrW(32) & ChrW(109) & ChrW(7899) & ChrW(105)
                        MessageBoxW Me.hwnd, StrPtr(s), StrPtr("Thông báo"), vbOKOnly

                        Me.MousePointer = 0
                        Exit Sub
                    End If
                End If
                If rs_ktra!Type <> 2 And rs_ktra!Type <> 1 Then
                    'MsgBox "§¨ng ký license ®Ó ®­îc thùc hiÖn chøc n¨ng nµy"
                    
                    s = ChrW(208) & ChrW(259) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(107) & ChrW(253) & ChrW(32) & ChrW(108) & ChrW(105) & ChrW(99) & ChrW(101) & ChrW(110) & ChrW(115) & ChrW(101) & ChrW(32) & ChrW(273) & ChrW(7875) & ChrW(32) & ChrW(273) & ChrW(432) & ChrW(7907) & ChrW(99) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(7921) & ChrW(99) & ChrW(32) & ChrW(104) & ChrW(105) & ChrW(7879) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7913) & ChrW(99) & ChrW(32) & ChrW(110) & ChrW(259) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(110) & ChrW(224) & ChrW(121)
                    MessageBoxW Me.hwnd, StrPtr(s), StrPtr("Thông báo"), vbOKOnly

                    Me.MousePointer = 0
                    Exit Sub
                End If
                rs_ktra.MoveNext
            Loop

        End If
        ExecuteSQL5 ("DELETE FROM SOLOTHUOC1")
        ExecuteSQL5 ("INSERT INTO SOLOTHUOC1  SELECT MAVATTU,SOLO,HANDUNG,CONLAI AS SOLUONG  FROM DANHSAchvattu")
        ExecuteSQL5 ("DELETE FROM SOLOTHUOC")
        ExecuteSQL5 ("INSERT INTO SOLOTHUOC  SELECT MAVATTU,SOLO,HANDUNG,SOLUONG  FROM SoLoThuoc1")
        '            If KtraMKAdmin Then
        '                If MsgBox("B¹n ch¾c ch¾n kÕt thóc n¨m " + CStr(pNamTC) + " vµ chuyÓn sang n¨m míi ?" _
                         '                    , vbYesNo + vbExclamation, App.ProductName) <> vbYes Then GoTo KT
        '
        '                HienThongBao "ChuyÓn sè d­ cuèi kú ...  Xin vui lßng chê !", 1
        '                ChuyenNamMoi
        '                lbCty(7).Caption = CStr(pNamTC)
        '                LietKeNam
        '            End If
        If (boolean_kiemtra() = False) Then
            Dim tongsodong As String
            tongsodong = SelectSQL("SELECT count(*) as F1 FROM ChungTu ")
            ExecuteSQL5 "UPDATE license SET sodongId =" + CStr(Int_StrToCodes(tongsodong)) + " , sodong = sodong+ " + CStr(tongsodong)


        End If

        If KtraMKAdmin Then
            Dim sMsg As String
            Dim sCap As String

            sMsg = ChrW(66) & ChrW(7841) & ChrW(110) & " " & _
                   ChrW(99) & ChrW(104) & ChrW(7855) & ChrW(99) & " " & _
                   ChrW(99) & ChrW(104) & ChrW(7855) & ChrW(110) & " " & _
                   ChrW(107) & ChrW(7871) & ChrW(116) & " " & _
                   ChrW(116) & ChrW(104) & ChrW(250) & ChrW(99) & " " & _
                   ChrW(110) & ChrW(259) & ChrW(109) & " " & _
                   CStr(pNamTC) & " " & _
                   ChrW(118) & ChrW(224) & " " & _
                   ChrW(99) & ChrW(104) & ChrW(117) & ChrW(121) & ChrW(7875) & ChrW(110) & " " & _
                   ChrW(115) & ChrW(97) & ChrW(110) & ChrW(103) & " " & _
                   ChrW(110) & ChrW(259) & ChrW(109) & " " & _
                   ChrW(109) & ChrW(7899) & ChrW(105) & ChrW(63)                  ' m?i ?

            sCap = App.ProductName   ' gi? nguyên tiêu d? là tên ?ng d?ng

            If MessageBoxW(0&, StrPtr(sMsg), StrPtr(sCap), vbYesNo + vbExclamation) <> vbYes Then GoTo KT

            HienThongBao "ChuyÓn sè d­ cuèi kú ...  Xin vui lßng chê !", 1
            ChuyenNamMoi
            LbCty(7).Caption = CStr(pNamTC)
            LietKeNam
        End If
        '            Else
        '             MsgBox ("B¹n ph¶i active tr­¬c khi kÕt chuyªn")
        '            End If

    Case 9: FrmKC.Show vbModal
    Case 10: FrmPBCP.Show vbModal
    Case 11: FrmThKC.Show vbModal
    Case 14:
        Form3.chuyen_so_du_dau_ky
        'Form3.Show vbModal ' FrmCTGS.Show vbModal
    Case 16:
        sql = GetSetting(IniPath, "LastYear", "IncTax" + CStr(pNamTC), 0)
        sql = InputBox("Sè ®iÒu chØnh", "ThuÕ thu nhËp doanh nghiÖp " + CStr(pNamTC - 1), sql)
        If IsNumeric(sql) Then SaveSetting IniPath, "LastYear", "IncTax" + CStr(pNamTC), sql
    Case 17:
        sql = ChonTenTep("Chän tÖp d÷ liÖu cña n¨m TC tr­íc (L­u ý cÇn ch¹y kiÓm tra sè liÖu cña n¨m cò)", &H4&, "*.MDB", 1)
        If Len(sql) = 0 Then GoTo KT
        LaySoDauNam sql
    Case 19: If KtraMKAdmin Then FrmE.Show 1
    Case 21:
        If KtraMKAdmin Then
            sql = FrmDB.ChonTepLuu(frmMain.LbCty(8).Caption, pNamTC)
            If Len(sql) > 0 Then
                CloseUp 1
                OpenDB sql
            End If
        End If
    End Select
KT:
    HienThongBao "", 1
    Me.MousePointer = 0
End Sub

Private Sub mnDLdt_Click(Index As Integer)

    AuToNhapTP
    'Nhap kho theo hoa don
    If Index = 1 Then
        
    End If
    
    'Nhap kho theo ngay
    If Index = 2 Then
        
    End If
    
    'Nhap kho theo thang
    If Index = 3 Then
        
    End If
    
End Sub

Public Sub mnHT_Click(Index As Integer)
    Dim psw As String, st As Integer, fn As String

    If Index = 5 Or Index = 6 Or Index = 10 Then
        If Not KtraMKAdmin Then Exit Sub
    End If
    Me.MousePointer = 11
    Select Case Index
    Case 0:    ' Mo tep
a:
        psw = ChonTenTep("Chän tÖp d÷ liÖud÷ liÖu", &H4&, "*.MDB", 1)
MoTep:
        If Len(psw) = 0 Then GoTo KT
        HienThongBao "Më tÖp d÷ liÖu...", 1
        If st = 0 Then CloseUp 1
        If OpenDB(psw, 1) = 0 Then
            GetLicense
            frmLicenseUser.Hide
            If pDataPath <> GetSetting(IniPath, "Environment", "Path") Then
                pProcessMode = 0
            Else
                Select Case UCase(App.ExeName)
                Case "SERVER": pProcessMode = 2
                Case "CLIENT": pProcessMode = 1
                Case Else: pProcessMode = 0
                End Select
            End If
            On Error Resume Next

            FrmMatkhau.Show 1
            If Err.number = 364 Then

            ElseIf Err.number <> 0 Then

            End If
            On Error GoTo 0

            Set FrmMatkhau = Nothing
            SetUserRight

            LietKeTep
        Else
            st = 1
            GoTo a
        End If
    Case 1:    ' Sao chep
        DelTemp
        psw = ChonTenTep("Sao chÐp tÖp d÷ liÖu", &H4&, "*.MDB", 2)
        If Len(psw) = 0 Then GoTo KT
        Me.MousePointer = 11
        HienThongBao "Sao chÐp tÖp d÷ liÖu ...", 1
        CloseUp 1
        On Error Resume Next
        DBEngine.CompactDatabase pDataPath, psw, , , ";pwd=" + pPSW
        On Error GoTo 0
        OpenDB pDataPath
    Case 2:    ' Tep mac dinh
        mnHT_Click 0
        SaveSetting IniPath, "Environment", "Path", pDataPath
    Case 3:    ' Nen tep du lieu
        DelTemp
        psw = ChonTenTep("NÐn tÖp d÷ liÖu", &H4&, "*.SAS", 2)
        If Len(psw) = 0 Then GoTo KT
        Me.MousePointer = 11
        HienThongBao "NÐn tÖp d÷ liÖu ...", 1
        CloseUp 1
        Recycle pCurDir + "TEMP.MDB"
        On Error Resume Next
        DBEngine.CompactDatabase pDataPath, pCurDir + "TEMP.MDB", , , ";pwd=" + pPSW
        On Error GoTo 0
        If Len(Dir(pCurDir + "TEMP.MDB")) > 0 Then
            NenTep pCurDir + "TEMP.MDB", psw
            Recycle pCurDir + "TEMP.MDB"
        Else
            NenTep pDataPath, psw
        End If
X1:
        OpenDB pDataPath
    Case 4:
        psw = ChonTenTep("Chän tÖp d÷ liÖu nÐn", &H4&, "*.SAS", 1)
        If Len(psw) = 0 Then GoTo KT
        fn = ChonTenTep("Chän tªn tÖp d÷ liÖu", &H4&, "*.MDB", 2)
        If Len(fn) = 0 Then GoTo KT
        GianTepNen psw, fn

        CloseUp 1

        OpenDB fn, 1
        GetLicense

        FrmMatkhau.Show 1
        Set FrmMatkhau = Nothing
        SetUserRight
    Case 5:
        EMailDB
    Case 6:
        psw = ChonTenTep("Tªn tÖp d÷ liÖu", &H4&, "*.MDB", 2)
        If Len(psw) = 0 Then GoTo KT
        CloseUp 1
        On Error GoTo KT
        DBEngine.CompactDatabase pDataPath, psw, , , ";pwd=" + pPSW
        On Error GoTo 0
        OpenDB psw
        ExecuteSQL5 "UPDATE License SET LoaiTien=" + IIf(pTien = 0, "1", "0")
        GetLicense
        DoiTyGiaDB
    Case 8:    ' Dat may in
        ChonTenTep "", 0, "", 3
    Case 9:    ' Dat may in
        ChonTenTep "", cdlCFBoth, "", 4
        If Len(dlgCommonDialog.FontName) > 1 And (LoaiFont(dlgCommonDialog.FontName) = FontFlag Or KiemTraMaSoThue(LbCty(8).Caption, "03")) Then
            pFontName = dlgCommonDialog.FontName
            pFontSize = dlgCommonDialog.FontSize
            ExecuteSQL5 "UPDATE License SET FontName='" + pFontName + "', FontSize=" + CStr(pFontSize)
            LbCty(0).FontName = pFontName
            LbCty(1).FontName = pFontName
            mnHT(10).Caption = IIf(FontFlag <> 2, "ChuyÓn ®æi CSDL sang font ABC", "ChuyÓn ®æi CSDL sang font VNI")
            SetFont Me
        End If
    Case 10:
        If MsgBox("B¹n ch¾c ch¾n cÇn ®æi font ? (Chó ý chän font ch÷ tr­íc khi ®æi)", vbCritical + vbYesNo, App.ProductName) = vbNo Then GoTo KT
        Me.MousePointer = 11
        ChuyenDoiFont FontFlag = 2
        GetLicense
    Case 11:    ' Thong so
        If User_Right = 0 Then
            FrmOptions.Show 1
            GetLicense
        Else
            NoRight 0
        End If
    Case 13:    ' Danh sach user
        If User_Right = 0 Then
            FrmUser.Show 1
        Else
            NoRight 0
        End If
    Case 14:    ' Dat mat khau
        'Load FrmMatkhau
        FrmMatkhau.tag = 1
        FrmMatkhau.Show 1
    Case 16:
        If (Not IsNumeric(Left(LbCty(8).Caption, 2))) Then GoTo KT
        If CInt(Left(LbCty(8).Caption, 3)) = 0 Then GoTo KT
        If (Len(pMST) > 0 And Left(LbCty(8).Caption, Len(pMST)) = pMST) Then GoTo B
        If FrmGetStr.GetMK(LbCty(8).Caption) Then
B:
            UpDateDB
            GetLicense
        End If
    Case 18 To 22:
        psw = mnHT(Index).Caption
        GoTo MoTep
    Case 24:
        FrmMatkhau.Show 1
        SetUserRight
    Case 25:
        Unload Me
        Exit Sub
    End Select
KT:
    HienThongBao "", 1
    Me.MousePointer = 0
End Sub

Private Sub mnK_Click(Index As Integer)
    Dim k As Integer
    
    If User_Right <> 0 Then
        NoRight 0
        Exit Sub
    End If
    
    k = SelectSQL("SELECT Lock" + CStr(mnk(Index).tag) + " Mod 10 AS F1 FROM License")
    If MsgBox("CÇn " + IIf(k = 0, "kho¸", "cho nhËp") + IIf(mnk(Index).tag > 0, " ph¸t sinh th¸ng " + CStr(mnk(Index).tag), " sè d­ ®Çu n¨m") + " ?", vbYesNo + vbExclamation, App.ProductName) <> vbYes Then Exit Sub
    ExecuteSQL5 "UPDATE License SET Lock" + CStr(mnk(Index).tag) + "=10*(Lock" + CStr(mnk(Index).tag) + " \ 10)+" + CStr(1 - k)
    mnk(Index).Caption = IIf(1 - k > 0, Trim(mnk(Index).Caption) + "          x", Left(mnk(Index).Caption, Len(mnk(Index).Caption) - 1))
End Sub

Private Sub mnkt_Click(Index As Integer)

End Sub

Private Sub mnNam_Click(Index As Integer)
    Dim i As Integer, path As String
    
    Me.MousePointer = 11
    CloseUp 1
    If Index = 4 Then
        path = GetSetting(IniPath, "Environment", "Path", pCurDir + "DATA\KETOAN.MDB")
    Else
        path = GetSetting(IniPath, "LastYear", mnNam(Index).Caption, pCurDir + "DATA\KETOAN.MDB")
    End If
    If OpenDB(path) <> 0 Then mnHT_Click 0
    For i = 0 To 4
        mnNam(i).CHECKED = (i = Index)
    Next
    pNamTC = CInt5(mnNam(Index).Caption)
    
    LbCty(7).Caption = CStr(pNamTC)
    Me.MousePointer = 0
End Sub

Private Sub mnnh_Click(Index As Integer)

End Sub

Public Sub mnTS_Click(Index As Integer)
    If (Not FADetail) Or User_Right = 2 Then
        NoRight 2
        Exit Sub
    End If
    Me.MousePointer = 11

    Select Case Index
    Case 0:    ' Phan loai TS
        'Load frmPhanLoai
        frmPhanLoai.tag = 1
        frmPhanLoai.Show 1
    Case 11:    ' Phan loai TS
        frmDSTaiSan.Show 1

    Case 1:    ' Phan loai ctu
        'Load frmPhanLoai
        frmPhanLoai.tag = 2
        frmPhanLoai.Show 1
    Case 2:
        frmDSTaiSan.Show 1
    Case 4:    ' Nuoc sx
        'Load FrmKho
        FrmKho.tag = 2
        FrmKho.Show 1
    Case 5:    ' Tinh trang SD
        'Load FrmKho
        FrmKho.tag = 3
        FrmKho.Show 1
    Case 6:    ' DTQL
        'Load FrmKho
        FrmKho.tag = 4
        FrmKho.Show 1
    Case 8:
        If ChoDieuChinhDauKy Then
            pNghiepVu = NV_TANG
            'Load frmTaiSan
            frmTaiSan.tag = 1
            frmTaiSan.Show 1
        End If
    Case 9:
        frmDSTaiSan.Show 1
    Case 10:
        'frmDSTaiSan.Show 1
        If KtraMKAdmin Then DatTKTS
    End Select
    HienThongBao "", 1
    Me.MousePointer = 0
End Sub

Private Sub mnuHLP_Click(Index As Integer)

    Dim nRet As Integer

    Select Case Index
    Case 0:    ' Noi dung
        '         frmTonDauKhachHang.Show
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then MsgBox Err.Description
        On Error GoTo 0
    Case 1:
frmgioithieu.Show vbModal, Me
        'Formimport.Show 1

        ' Tra cuu
        '   frmTonDauSanPham.Sh'ow
        '      On Error Resume Next
        '     nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        '     If Err Then MsgBox Err.Description
        'e    On Error GoTo 0
    Case 3:
        '  frmTonDauDaTaBaSE.Show ' Ban quyen
        '   frmAbout.Show vbModal, Me
        
        frmTaiLieu.Show 1
    Case 4:
        Dim fso As New FileSystemObject
        '  MsgBox pCurDir + "DATA"
        Dim duong_dan As String
        ' duong_dan = Mid(pCurDir, 1, Len(pCurDir) - 1) + CStr(Minute(Now)) + CStr(Second(Now))

        duong_dan = Mid(Mid(pCurDir, 1, Len(pCurDir) - 1), 1, InStrRev(Mid(pCurDir, 1, Len(pCurDir) - 1), "\")) + "VietStar_" + CStr(Minute(Now)) + CStr(Second(Now))

        MkDir duong_dan
        MkDir duong_dan + "\data"
        If Len(Dir(pCurDir + "REPORTS\QD48.MDB")) = 0 Then
            fso.CopyFile pCurDir + "REPORTS\QD15.MDB", duong_dan + "\Data\QD15.mdb", True
        Else
            fso.CopyFile pCurDir + "REPORTS\QD48.MDB", duong_dan + "\Data\QD48.mdb", True
        End If

        fso.CopyFolder pCurDir + "REPORTS", duong_dan + "\REPORTS"
        fso.CopyFolder pCurDir + "Tailieu", duong_dan + "\Tailieu"
        fso.CopyFile pCurDir + "Dummy.xml", duong_dan + "\Dummy.xml"
        fso.CopyFile pCurDir + "UnicodeVowels.xml", duong_dan + "\UnicodeVowels.xml"
        fso.CopyFile pCurDir + "VNIVowelMap.txt", duong_dan + "\VNIVowelMap.txt"
        '  fso.CopyFile pCurDir + "VietStar.exe", duong_dan + "\VietStar.exe"

        fso.CopyFile duong_dan + "\REPORTS\vietstar.exe", duong_dan + "\VietStar.exe"
        CreateShortCut duong_dan + "\VietStar.exe", "VietStar_" + CStr(Minute(Now)) + CStr(Second(Now))
        'MsgBox "B¹n ®· t¹o míi thµnh c«ng: " + duong_dan
        MsgBox "B¹n ®· t¹o míi thµnh c«ng, icon ®· cã ngoµi mµn h×nh:" & vbNewLine & duong_dan
        Shell "EXPLORER.EXE " & duong_dan + "\VietStar.exe"
    End Select

End Sub
Sub CreateShortCut(duonglink As String, tenshortcut As String)
    Dim objShell, strDesktopPath, objLink
    Set objShell = CreateObject("WScript.Shell")
    strDesktopPath = objShell.SpecialFolders("Desktop")
    Set objLink = objShell.CreateShortCut(strDesktopPath & "\" + tenshortcut + ".lnk") '"\saoviet.lnk"
    objLink.Arguments = duonglink
    objLink.Description = "VietStar Accounting"
    objLink.TargetPath = duonglink
    objLink.WindowStyle = 1
    objLink.WorkingDirectory = "c:\windows"
    objLink.Save
  End Sub

Private Sub mnunh_Click()

End Sub

Private Sub mnviet_Click()

End Sub


Private Sub mnversionchild_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Dim fileNumber As Integer
    Dim FilePath As String
    Dim content As String

    'Cap nhat file version
    FilePath = App.path & "\\HoaDon\\\version.txt"
    fileNumber = FreeFile
    On Error Resume Next
    Open FilePath For Output As #fileNumber
    If Err.number = 0 Then
        Print #fileNumber, FindItemByID(Index);
        Close #fileNumber
        'MsgBox "Ðã ghi dè file version.txt thành công!", vbInformation
    Else
        MsgBox "L?i khi ghi dè file!", vbExclamation
    End If
    KiemtraversionGitHub3
End Sub
Private Sub KiemtraversionGitHub3()
    Dim exeDir As String
    Dim extractFolder As String
    Dim jsonPath As String
    Dim json As String
    Dim versions As Collection
    Dim i As Long
    Dim j As Integer

    ' 1. Ðu?ng d?n luu file JSON
    exeDir = App.path
    extractFolder = exeDir & "\Tools\Debug"

    ' T?o thu m?c n?u chua có
    If Dir(extractFolder, vbDirectory) = "" Then
        MkDir extractFolder
    End If

    jsonPath = extractFolder & "\configVersion.json"

    ' 2. Xóa file cu n?u có
    If Dir(jsonPath) <> "" Then Kill jsonPath

    ' 3. T?i file JSON t? GitHub
    Dim jsonContent As String
    jsonContent = DownloadJSONFromGitHub()

    If jsonContent = "" Then
        MsgBox "Không th? t?i file c?u hình t? GitHub!", vbCritical
        Exit Sub
    End If

    ' 4. Luu file JSON xu?ng máy
    Open jsonPath For Output As #1
    Print #1, jsonContent
    Close #1

    ' 5. Load menu t? file JSON v?a t?i
    mnItemCount = 0
    ReDim arrItems(1 To 1)
    json = ReadJsonFile(jsonPath)

    Set versions = GetVersions(json)
    j = 2


    ' 6. L?y version hi?n t?i
    Dim txtPath As String
    txtPath = App.path & "\HoaDon\version.txt"
    Dim content As String
    content = ReadTxt(txtPath)

    ' 7. L?y version m?i nh?t t? JSON
    Dim latestVer As String
    latestVer = GetLatestVersionFromJSON(jsonContent)

    Dim versionData As versionInfo
    Docfilejson "1.1", versionData
    LastVersion = Replace(versionData.fileName, """", "")

    'Kiem tra hien tai co khong
    Dim versioncurrentData As versionInfo
    Docfilejson content, versioncurrentData
    ' 8. So sánh và x? lý c?p nh?t
    'Or versioncurrentData.Version = ""

    Dim downloaDriver As String
    Dim savePathDriver As String
    downloaDriver = "https://github.com/n93studioenter/SaovietGit/releases/download/latest-config/VietstarDriver.zip"
    savePathDriver = App.path & "\Tools\Debug" & "\" & "VietstarDriver" & ".zip"
    Dim savepathexe As String
    savepathexe = App.path & "\Tools\Debug" & "\" & "VietstarDriver" & ".exe"
    If Dir(savePathDriver) <> "" Then Kill savePathDriver
    If Dir(savepathexe) <> "" Then Kill savepathexe
    If DownloadFileWithWinHTTP(downloaDriver, savePathDriver, 10) Then
        UnzipFile savePathDriver, extractFolder
    Else
        Screen.MousePointer = vbDefault    'ho?c 0
    
        Exit Sub
    End If

    ' L?y thông tin file c?n t?i
    Dim fileName As String
    fileName = GetFileNameByVersion(jsonContent, latestVer)

    If fileName <> "" Then
        ' T?o link t?i t? GitHub Release
        Dim downloadUrl As String
        Dim savePath As String
        'downloadUrl = "https://github.com/n93studioenter/SaovietGit/releases/download/latest-config/1.1.zip"
        downloadUrl = "https://github.com/n93studioenter/SaovietGit/releases/download/latest-config/" & content & ".zip"

        savePath = extractFolder & "\" & content & ".zip"

        ' Xóa file cu
        If Dir(savePath) <> "" Then Kill savePath

        ' T?i file
        If DownloadFileWithWinHTTP(downloadUrl, savePath, 60) Then
            Dim s As String
            's = ChrW(67) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7853) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(224) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(99) & ChrW(244) & ChrW(110) & ChrW(103)
            'MessageBoxW Me.hwnd, StrPtr(s), StrPtr("Thông báo"), vbOKOnly

            'UpdateApplication savePath

            ' C?p nh?t version.txt
            If content = "1.1" Then
                WriteTxt txtPath, LastVersion
            End If

            ' Gi?i nén n?u là zip
            UnzipFile savePath, extractFolder

            Dim exePath As String
            exePath = App.path & "\Tools\Debug" & "\" & "VietstarDriver" & ".exe"
            'Dong chuong trinh va mo file update
            If Dir(exePath) <> "" Then
                ' Shell exePath, vbNormalFocus
                ShellExecute 0, "open", exePath, "", App.path & "\Tools\Debug", SW_SHOW
                Unload Me
                End
            Else
                MsgBox "Không tìm th?y file EXE!", vbExclamation
            End If


        Else
            MsgBox "T?i file th?t b?i!", vbCritical
        End If
    Else
        ' MsgBox "B?n dang dùng phiên b?n m?i nh?t: " & content, vbInformation
    End If
End Sub
Private Sub CopyAndRun()
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim src As String
    Dim dest As String
    Dim exe As String
    Dim cmd As String

    src = "\\192.168.1.90\Ke toan 2025 New\1 Copi vao dung 1\Tools\Debug\AutoUpdate"
    dest = App.path & "\Tools\Debug\"
    exe = dest & "AutoUpdate\AutoUpdate.exe"

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Copy
    If Not fso.FolderExists(dest) Then
        fso.CreateFolder dest
    End If

    fso.CopyFolder src, dest, True

    ' Ki?m tra file
    If Not fso.FileExists(exe) Then
        MsgBox "Không tìm th?y AutoUpdate.exe!", vbExclamation
        Exit Sub
    End If

    ' Ch?y AutoUpdate b?ng CMD, không ch? chuong trình k?t thúc
    cmd = "cmd.exe /c start """" """ & exe & """"

    Shell cmd, vbHide

    ' Ðóng form ngay
    Unload Me

    Exit Sub

ErrorHandler:
    MsgBox "L?i: " & Err.Description, vbCritical
    Unload Me
End Sub

Public Sub mnVT_Click(Index As Integer)

    If Index = 22 Then
        Dim response As Integer

        response = MessageBoxW(0&, _
                               StrPtr(ChrW(66) & ChrW(7841) & ChrW(110) & " " & _
                                      ChrW(99) & ChrW(243) & " " & _
                                      ChrW(109) & ChrW(117) & ChrW(7889) & ChrW(110) & " " & _
                                      ChrW(120) & ChrW(243) & ChrW(97) & " " & _
                                      ChrW(116) & ChrW(7845) & ChrW(116) & " " & _
                                      ChrW(99) & ChrW(7843) & " " & _
                                      ChrW(118) & ChrW(7853) & ChrW(116) & " " & _
                                      ChrW(116) & ChrW(432) & " " & _
                                      ChrW(107) & ChrW(104) & ChrW(244) & ChrW(110) & ChrW(103) & " " & _
                                      ChrW(99) & ChrW(243) & " " & _
                                      ChrW(112) & ChrW(104) & ChrW(225) & ChrW(116) & " " & _
                                      ChrW(115) & ChrW(105) & ChrW(110) & ChrW(104) & ChrW(63)), _
                                      StrPtr(ChrW(88) & ChrW(225) & ChrW(99) & " " & _
                                             ChrW(110) & ChrW(104) & ChrW(7853) & ChrW(110) & " " & _
                                             ChrW(120) & ChrW(243) & ChrW(97)), _
                                             vbYesNo + vbQuestion)
        If response = vbYes Then
            Dim sql As String
            sql = "DELETE FROM Vattu WHERE MaSo NOT IN (SELECT MaVatTu FROM TonKho)"
            ExecuteSQL5 sql


            sql = "DELETE  FROM TonKho " & _
                  "WHERE Tien_1 = 0 AND Tien_2 = 0 AND Tien_3 = 0 AND Tien_4 = 0 AND Tien_5 = 0 AND Tien_6 = 0 AND Tien_7 = 0 AND Tien_8 = 0 AND Tien_9 = 0 AND Tien_10 = 0 AND Tien_11 = 0 AND Tien_12 = 0 " & _
                  "AND Tien_Nhap_1 = 0 AND Tien_Nhap_2 = 0 AND Tien_Nhap_3 = 0 AND Tien_Nhap_4 = 0 AND Tien_Nhap_5 = 0 AND Tien_Nhap_6 = 0 AND Tien_Nhap_7 = 0 AND Tien_Nhap_8 = 0 AND Tien_Nhap_9 = 0 AND Tien_Nhap_10 = 0 AND Tien_Nhap_11 = 0 AND Tien_Nhap_12 = 0 " & _
                  "AND Tien_Xuat_1 = 0 AND Tien_Xuat_2 = 0 AND Tien_Xuat_3 = 0 AND Tien_Xuat_4 = 0 AND Tien_Xuat_5 = 0 AND Tien_Xuat_6 = 0 AND Tien_Xuat_7 = 0 AND Tien_Xuat_8 = 0 AND Tien_Xuat_9 = 0 AND Tien_Xuat_10 = 0 AND Tien_Xuat_11 = 0 AND Tien_Xuat_12 = 0 " & _
                  "AND Luong_1 = 0 AND Luong_2 = 0 AND Luong_3 = 0 AND Luong_4 = 0 AND Luong_5 = 0 AND Luong_6 = 0 AND Luong_7 = 0 AND Luong_8 = 0 AND Luong_9 = 0 AND Luong_10 = 0 AND Luong_11 = 0 AND Luong_12 = 0"
            ExecuteSQL5 sql

            sql = "DELETE FROM Vattu WHERE MaSo IN (" & _
                  "SELECT MaVatTu FROM TonKho " & _
                  "WHERE Tien_1 = 0 AND Tien_2 = 0 AND Tien_3 = 0 AND Tien_4 = 0 AND Tien_5 = 0 AND Tien_6 = 0 AND Tien_7 = 0 AND Tien_8 = 0 AND Tien_9 = 0 AND Tien_10 = 0 AND Tien_11 = 0 AND Tien_12 = 0 " & _
                  "AND Tien_Nhap_1 = 0 AND Tien_Nhap_2 = 0 AND Tien_Nhap_3 = 0 AND Tien_Nhap_4 = 0 AND Tien_Nhap_5 = 0 AND Tien_Nhap_6 = 0 AND Tien_Nhap_7 = 0 AND Tien_Nhap_8 = 0 AND Tien_Nhap_9 = 0 AND Tien_Nhap_10 = 0 AND Tien_Nhap_11 = 0 AND Tien_Nhap_12 = 0 " & _
                  "AND Tien_Xuat_1 = 0 AND Tien_Xuat_2 = 0 AND Tien_Xuat_3 = 0 AND Tien_Xuat_4 = 0 AND Tien_Xuat_5 = 0 AND Tien_Xuat_6 = 0 AND Tien_Xuat_7 = 0 AND Tien_Xuat_8 = 0 AND Tien_Xuat_9 = 0 AND Tien_Xuat_10 = 0 AND Tien_Xuat_11 = 0 AND Tien_Xuat_12 = 0 " & _
                  "AND Luong_1 = 0 AND Luong_2 = 0 AND Luong_3 = 0 AND Luong_4 = 0 AND Luong_5 = 0 AND Luong_6 = 0 AND Luong_7 = 0 AND Luong_8 = 0 AND Luong_9 = 0 AND Luong_10 = 0 AND Luong_11 = 0 AND Luong_12 = 0)"
            ExecuteSQL5 sql
            'MsgBox "Xo¸ thµnh c«ng !"
            Dim s As String
            s = ChrW(88) & ChrW(243) & ChrW(97) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(224) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(99) & ChrW(244) & ChrW(110) & ChrW(103)
            MessageBoxW Me.hwnd, StrPtr(s), StrPtr("Thông báo"), vbOKOnly

        End If
        Exit Sub
    End If

    Dim st As String, i As Integer, TK As String, d1 As Date, d2 As Date, j As Integer, k As Integer, mv As Long

    If User_Right = 2 Then
        NoRight 0
        Exit Sub
    End If

    If Not STDetail Then
        NoRight 1
        Exit Sub
    End If

    If Index = 3 Or Index = 4 Or Index = 5 Or Index = 10 Or Index = 12 Then
        If Not KtraMKAdmin Then Exit Sub
    End If

    Me.MousePointer = 11


    Select Case Index
    Case 0:
        frmPhanLoaiVT.tag = 1
        frmPhanLoaiVT.Show 1
    Case 6:
        FrmKho.tag = 1
        FrmKho.Show 1
    Case 1:
        FrmVattu.Show 1

    Case 3:
        FrmLuuChuyen.Show 1
    Case 4:
        FrmThanhPham.Show 1
    Case 2:
        FrmNguon.Show 1
    Case 3:
        If ChoDieuChinhDauKy Then FVTDauKy.Show 1
    Case 8:
        If OutCost <> 2 Then
            If frmMain.Tudongtinhgiavon = False Then
                st = FrmGetStr.GetString("Th¸ng cÇn tÝnh l¹i:", "TÝnh gi¸ xuÊt kho")
            Else
                Dim currentMonth As Integer

                currentMonth = FrmChungtu.Getmonthbynl
                st = currentMonth
            End If
            If IsNumeric(st) Then
                i = CInt5(st)
                j = i
            Else
                i = InStr(st, "-")
                If i > 0 Then
                    j = CInt5(Right(st, Len(st) - i))
                    i = CInt5(Left(st, i - 1))
                Else
                    i = CInt5(st)
                    j = i
                End If
            End If
        Else
            i = 1
            j = 12
        End If
        If i > 0 And i < 13 And j > 0 And j < 13 Then
            st = ""
            If frmMain.Tudongtinhgiavon = False Then
                st = FrmGetStr.GetString("Sè hiÖu vËt t­ cÇn tÝnh l¹i (®Ó trèng nÕu tÝnh l¹i toµn bé):", "TÝnh gi¸ xuÊt kho")
            Else
                st = ""
            End If

            Do While Len(st) > 0
                mv = SoHieu2MaSo(st, "Vattu")
                If mv > 0 Then Exit Do
                st = FrmGetStr.GetString("Sè hiÖu vËt t­ cÇn tÝnh l¹i (®Ó trèng nÕu tÝnh l¹i toµn bé):", "TÝnh gi¸ xuÊt kho")
            Loop

            If OutCost <> 2 Then
                If frmMain.Tudongtinhgiavon = False Then
                    TK = FrmGetStr.GetString("Sè hiÖu tµi kho¶n ghi nî khi xuÊt kho cÇn tÝnh l¹i (®Ó trèng nÕu tÝnh l¹i toµn bé):", "TÝnh gi¸ xuÊt kho", "")
                Else
                    TK = ""
                End If
            End If
            Me.MousePointer = 11
            If OutCost = 0 Then
                If frmMain.Tudongtinhgiavon = False Then
                    ' Nh?p s? t? ngu?i dùng
                    k = CInt5(FrmGetStr.GetString("NhËp sè 1 ®Ó tÝnh b×nh qu©n cuèi kú, sè 2 ®Ó tÝnh b×nh qu©n di ®éng ", "TÝnh l¹i gi¸ xuÊt kho"))
                Else
                    k = 2
                End If    ' Ð?m b?o có End If cho If th? hai
            End If    ' Ð?m b?o có End If cho If d?u tiên
            If k < 1 And k > 2 Then GoTo KT
            If k = 1 Then TinhGXK i, j, st, TK
            If k = 2 Then
                If Tudongtinhgiavon = False Then
                    TinhGXKBQ i, j, st, TK
                Else
                    TinhGXKBQ2 i, j, st, TK
                    MsgBox "Da tinh xong gia von"
                End If
            End If
            If OutCost = 1 Then TinhGVBH NgayDauThang(pNamTC, pThangDauKy), NgayCuoiNam(), 1, mv
            If OutCost = 2 Then TinhGXKFIFO i, j, st, TK
        End If
    Case 9:
        If OutCost = 2 Then
            d1 = NgayDauThang(pNamTC, pThangDauKy)
            d2 = NgayCuoiNam
        Else
            If Not GetDate2.GetDate("TÝnh gi¸ vèn b¸n hµng", d1, d2) Then Exit Sub
        End If

        s = ChrW(76) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(108) & ChrW(7841) & ChrW(105) & ChrW(32) & ChrW(99) & ChrW(225) & ChrW(99) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7913) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(103) & ChrW(105) & ChrW(225) & ChrW(32) & ChrW(118) & ChrW(7889) & ChrW(110) & ChrW(32) & ChrW(273) & ChrW(227) & ChrW(32) & ChrW(116) & ChrW(237) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(63) & ChrW(32) & ChrW(40)
        s = s & ChrW(78) & ChrW(7871) & ChrW(117) & ChrW(32) & ChrW(107) & ChrW(104) & ChrW(244) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(236) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(432) & ChrW(417) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(114) & ChrW(236) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7881)
        s = s & ChrW(32) & ChrW(108) & ChrW(7853) & ChrW(112) & ChrW(32) & ChrW(99) & ChrW(225) & ChrW(99) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7913) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(32) & ChrW(103) & ChrW(105) & ChrW(225) & ChrW(32) & ChrW(118) & ChrW(7889) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(242) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(105) & ChrW(7871) & ChrW(117) & ChrW(41)
        Dim xn As String
        xn = ChrW(88) & ChrW(225) & ChrW(99) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7853) & ChrW(110)
        If MessageBoxW(Me.hwnd, StrPtr(s), StrPtr(xn), vbYesNo + vbExclamation) = vbYes Then
            i = 6
        Else
            i = 7
        End If

        'i = MsgBox("LËp l¹i c¸c chøng tõ gi¸ vèn ®· tÝnh ? (NÕu kh«ng th× ch­¬ng tr×nh chØ lËp c¸c chøng tõ gi¸ vèn cßn thiÕu)", vbCritical + vbYesNo, App.ProductName)
        st = FrmGetStr.GetString("Sè hiÖu vËt t­ cÇn tÝnh l¹i (®Ó trèng nÕu tÝnh l¹i toµn bé):", "TÝnh gi¸ vèn")
        Do While Len(st) > 0
            mv = SoHieu2MaSo(st, "Vattu")
            If mv > 0 Then Exit Do
            st = FrmGetStr.GetString("Sè hiÖu vËt t­ cÇn tÝnh l¹i (®Ó trèng nÕu tÝnh l¹i toµn bé):", "TÝnh gi¸ vèn")
        Loop
        Me.MousePointer = 11
        If OutCost = 0 Then
            k = CInt5(FrmGetStr.GetString("NhËp sè 1 ®Ó tÝnh b×nh qu©n cuèi kú (tÝnh theo th¸ng), sè 2 ®Ó tÝnh b×nh qu©n di ®éng ", "TÝnh l¹i gi¸ vèn"))
            If k < 1 And k > 2 Then GoTo KT
            TinhGVBHBQ month(d1), month(d2), i, mv, k
        Else
            TinhGVBH d1, d2, i, mv
        End If
    Case 6:
        FVTDauKy.tag = 1
        FVTDauKy.Show 1
    Case 7:
        If ChoDieuChinhDauKy Then FVTDauKy.Show 1
    Case 19:
        'Load frmPhanLoaiVT
        frmPhanLoaiVT.tag = 3
        frmPhanLoaiVT.Show 1
    Case 10:
        FrmTP.Show 1
    Case 11:
        If KtraMKAdmin Then DatTKDTTP
    Case 12:
        frmPhanLoaiVT.tag = 3
        frmPhanLoaiVT.Show 1
    Case 13:
        FrmTP.Show 1
    Case 14:
        If KtraMKAdmin Then DatTKDTTP
    Case 16:
        If KtraMKAdmin Then DatTKVT
    Case 18:
        Dim fileNumber As Integer
        Dim FilePath As String
        Dim content As String

        'Cap nhat file version
        FilePath = App.path & "\\HoaDon\\\status.txt"
        fileNumber = FreeFile
        content = "15"
        On Error Resume Next
        Open FilePath For Output As #fileNumber
        If Err.number = 0 Then
            Print #fileNumber, content;
            Close #fileNumber
            'MsgBox "Ðã ghi dè file version.txt thành công!", vbInformation
        Else
            MsgBox "L?i khi ghi dè file!", vbExclamation
        End If

        Dim exePath As String
        exePath = App.path & "\Tools\Debug" & "\" & "Saoviettax" & ".exe"
        'Dong chuong trinh va mo file update
        If Dir(exePath) <> "" Then
            ' Shell exePath, vbNormalFocus
            ShellExecute 0, "open", exePath, "", App.path & "\Tools\Debug", SW_SHOW
        Else
            MsgBox "Không tìm th?y file EXE!", vbExclamation
        End If
    Case 15    ', 16:
        FrmVattu.Show 1
        ' CPGV.tag = Index - 15
        'CPGV.Show 1

    End Select
KT:
    HienThongBao "", 1
    Me.MousePointer = 0
End Sub
Private Sub XKTheoNgay()
    'Lay danh sach chung tu dau thang den hien tai

    Dim ngayCTList As String
    ngayCTList = ""


    Dim rs_chungtu As Recordset
    Dim Query As String

    Dim firstDate As String
    Dim currentDate As String

    ' L?y ngày d?u tháng
    firstDate = Format(DateSerial(Year(Date), month(Date), 1), "yyyy-mm-dd")

    ' L?y ngày hi?n t?i
    currentDate = Format(Date, "yyyy-mm-dd")
    Query = "SELECT NgayCT, COUNT(*) AS SoHoaDon " & _
            "FROM ChungTu " & _
            "WHERE NgayCT >= #" & firstDate & "# AND NgayCT <= #" & currentDate & "# " & _
            "AND MaLoai = 8 " & _
            "GROUP BY NgayCT"
    Set rs_chungtu = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)

    If Not rs_chungtu.EOF Then
        ' B?t d?u vòng l?p While
        While Not rs_chungtu.EOF
            If ngayCTList = "" Then
                ngayCTList = Format(rs_chungtu!NgayCT, "dd/MM/yyyy")
            Else
                ngayCTList = ngayCTList & "," & Format(rs_chungtu!NgayCT, "dd/MM/yyyy")

            End If
            rs_chungtu.MoveNext
        Wend    ' K?t thúc vòng l?p While
    End If

    Dim ngayCTArray() As String
    ngayCTArray = Split(ngayCTList, ",")

    ' Duy?t qua m?ng và hi?n th? t?ng ph?n t?
    Dim item As Variant
    If UBound(ngayCTArray) >= 0 Then
        For Each item In ngayCTArray
            If item = "" Then
                Exit Sub
            End If
            '   Luu tb  tbNhapkhotp
            Dim d As Date
            d = CDate(item)
            Dim currentDateTime As Date
            currentDateTime = Now
            Dim KyHieu As String
            Dim KyHieu2 As String
            KyHieu = "NKTP" & Format(d, "ddMMyyyy")    ' Nên b? d?u / d? tránh l?i khi dùng làm s? hi?u
            KyHieu2 = "XKNL" & Format(d, "ddMMyyyy")
            Dim noiDung As String
            Dim noidung2 As String
            noiDung = "Nh?p kho thành ph?m ngày " & Format(d, "dd/MM/yyyy")
            noidung2 = "Xuaát kho nguyeân lieäu ngày " & Format(d, "dd/MM/yyyy")
            ExecuteSQL5 "INSERT INTO tbNhapkhotp (NgayLap, NgayTao, SoHieu, Ghichu, SoHieu2, Ghichu2, Status) " & _
                        "VALUES (#" & Format(d, "mm/dd/yyyy") & "#, #" & Format(currentDateTime, "mm/dd/yyyy hh:nn:ss") & "#, '" & _
                        KyHieu & "', '" & noiDung & "', '" & KyHieu2 & "', '" & noidung2 & "', 0)"

            'Lay ma moi nhat sau khi luu
            Dim maxMact As String
            maxMact = SelectSQL("SELECT MAX(ID) as f1 FROM tbNhapkhotp")


            'Lay danh sach chung tu den ngay hien tai
            Query = "SELECT * " & _
                    "FROM ChungTu " & _
                    "WHERE NgayCT = #" & Format(item, "yyyy-mm-dd") & "# " & _
                    "AND MaVattu <> 0 " & _
                    "AND SoPS2Co <>0" & _
                    "AND sohieu NOT LIKE '*GV*'"
            Set rs_chungtu = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)

            If Not rs_chungtu.EOF Then

                Dim dongia As Double

                'duyet tung row cua  hoa don
                While Not rs_chungtu.EOF

                    Dim getSHVT As Integer
                    getSHVT = CInt(SelectSQL("SELECT MaPhanLoai as f1 from Vattu where MaSo=" & rs_chungtu!MaVattu))

                    Dim getsh As String
                    getsh = SelectSQL("SELECT SoHieu as f1 from Vattu where MaSo=" & rs_chungtu!MaVattu)

                    Dim getTenVT As String
                    getTenVT = SelectSQL("SELECT TenVattu as f1 from Vattu where MaSo=" & rs_chungtu!MaVattu)

                    'Loc ra neu la thanh pham
                    If IsThanhPham(rs_chungtu!MaVattu, getSHVT) = True Then
                        'MsgBox getTenVT
                        dongia = (val(rs_chungtu!sops) / val(rs_chungtu!SoPS2Co)) * 0.8

                        'Thuc hien luu tbNhapkhotpChitiet

                        'Kiem tra sohieu nay da dc luu chua
                        Dim rs_check As Recordset
                        Dim sql As String
                        Query = "SELECT * FROM tbNhapkhotpChitiet WHERE SoHieu = '" & getsh & "' AND ParentID = " & CInt(maxMact)
                        Set rs_check = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
                        If rs_check.EOF Then
                            sql = "INSERT INTO tbNhapkhotpChitiet (SoHieu, SoLuong, DonGia, ParentID) " & _
                                  "VALUES ('" & getsh & "', " & rs_chungtu!SoPS2Co & ", " & dongia & ", '" & CInt(maxMact) & "')"
                            ExecuteSQL5 sql
                        Else
                            sql = "UPDATE tbNhapkhotpChitiet SET SoLuong = SoLuong + " & rs_chungtu!SoPS2Co & " WHERE SoHieu = '" & getsh & "' AND ParentID = " & CInt(maxMact)
                            ExecuteSQL5 sql
                        End If

                        'Lay ra dong thanh pham vua luu

                        Query = "SELECT * FROM tbNhapkhotpChitiet WHERE SoHieu = '" & getsh & "' AND ParentID = " & CInt(maxMact)
                        Set rs_check = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
                        Dim tt As Double
                        If Not rs_check.EOF Then
                            'Tinh thanh tien thanh pham
                            tt = rs_check!dongia * rs_check!SoLuong
                        End If

                        'Nhap kho nguyen lieu
                        'Loc danh sach   nguyen lieu co trong thanh pham
                        Query = "SELECT * FROM tbNguyenLieuThanhPham WHERE TPSoHieu = '" & getsh & "'"

                        Set rs_check = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
                        If Not rs_check.EOF Then
                            While Not rs_check.EOF
                                'MsgBox rs_check!SoHieuNguyenLieu
                                'Tinh tong tien
                                Dim TongTien As Double
                                Dim SoLuong As Double
                                TongTien = tt * rs_check!TiLe / 100
                                'Tim don gia
                                Dim dg As Double
                                Dim rs_ct As Recordset
                                Query = "SELECT TOP 1 * FROM ChungTu WHERE MaVattu = " & rs_check!IDNguyenLieu & " AND SoPS2No <> 0 AND MaLoai = 1  ORDER BY MaSo DESC"
                                Set rs_ct = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
                                If Not rs_ct.EOF Then

                                    If rs_check!SoHieuNguyenLieu = "RAU.03" Then
                                        MsgBox rs_ct!sohieu
                                    End If
                                    dg = val(rs_ct!sops) / val(rs_ct!SoPS2No)
                                    SoLuong = Round(TongTien / dg, 1)
                                End If

                                sql = "INSERT INTO tbNhapkhonguyenlieu (ParentId, TTien, SL, SoHieu,SoHieuTP) " & _
                                      "VALUES (" & CInt(maxMact) & "," & dg & ", " & SoLuong & ", '" & rs_check!SoHieuNguyenLieu & "', '" & rs_check!TPSoHieu & "')"
                                ExecuteSQL5 sql

                                rs_check.MoveNext
                            Wend
                        End If
                    End If

                    rs_chungtu.MoveNext
                Wend    ' K?t thúc vòng l?p While
            End If
        Next item
    End If
    Command_Click 0
End Sub
Private Sub XKTheoThang()
    Dim ngayCTList As String
    ngayCTList = ""


    Dim rs_chungtu As Recordset
    Dim Query As String

    Dim firstDate As String
    Dim currentDate As String

    ' L?y ngày d?u tháng
    firstDate = Format(DateSerial(Year(Date), month(Date), 1), "yyyy-mm-dd")

    ' L?y ngày hi?n t?i
    currentDate = Format(Date, "yyyy-mm-dd")
    Query = "SELECT * " & _
            "FROM ChungTu " & _
            "WHERE NgayCT >= #" & firstDate & "# AND NgayCT <= #" & currentDate & "# " & _
            "AND MaLoai = 8 "
    Set rs_chungtu = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)

    If Not rs_chungtu.EOF Then
        While Not rs_chungtu.EOF
            rs_chungtu.MoveNext
        Wend
    End If
End Sub

Private Sub mnVTdt_Click(Index As Integer)
    If Index = 1 Then
        XKTheoNgay
    End If
    If Index = 2 Then
        XKTheoThang
    End If
End Sub
Function IsThanhPham(ByVal number As Integer, ByVal MaSo As Integer) As Boolean
'Lay ma phan loai truoc
    Dim getMPL As Integer
    getMPL = CInt(SelectSQL("SELECT MaSo as f1 FROM PhanLoaiVattu WHERE SoHieu='TP'"))
    If getMPL = MaSo Then
        IsThanhPham = True
    Else
        IsThanhPham = False
    End If
End Function

Private Sub mnXoa_Click(Index As Integer)

    If User_Right <> 0 Then
        NoRight 0
        Exit Sub
    End If


    If mnXoa(Index).tag > 0 Then
        Dim sMsg As String
        sMsg = _
        ChrW(&H42) & ChrW(&H1EA1) & ChrW(&H6E) & " " & _
               ChrW(&H63) & ChrW(&H68) & ChrW(&H1EAF) & ChrW(&H63) & " " & _
               ChrW(&H63) & ChrW(&H68) & ChrW(&H1EAF) & ChrW(&H6E) & " " & _
               ChrW(&H63) & ChrW(&H1EA7) & ChrW(&H6E) & " " & _
               ChrW(&H78) & ChrW(&HF3) & ChrW(&H61) & " " & _
               ChrW(&H70) & ChrW(&H68) & ChrW(&HE1) & ChrW(&H74) & " " & _
               ChrW(&H73) & ChrW(&H69) & ChrW(&H6E) & ChrW(&H68) & " " & _
               ChrW(&H74) & ChrW(&H68) & ChrW(&HE1) & ChrW(&H6E) & ChrW(&H67) & " " & _
               CStr(mnXoa(Index).tag) & ChrW(&H20) & ChrW(&H3F)

        ' tháng
        Dim sCap As String
        sCap = ChrW(88) & ChrW(225) & ChrW(99) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(7853) & ChrW(110)
        If MessageBoxW(Me.hwnd, StrPtr(sMsg), _
                       StrPtr(sCap), _
                       vbYesNo + vbExclamation) = vbYes Then
            Me.MousePointer = 11
            sMsg = _
            ChrW(&H58) & ChrW(&HF3) & ChrW(&H61) & " " & _
                   ChrW(&H70) & ChrW(&H68) & ChrW(&HE1) & ChrW(&H74) & " " & _
                   ChrW(&H73) & ChrW(&H69) & ChrW(&H6E) & ChrW(&H68) & " " & _
                   ChrW(&H74) & ChrW(&H68) & ChrW(&HE1) & ChrW(&H6E) & ChrW(&H67) & " " & _
                   CStr(mnXoa(Index).tag) & " ...  " & _
                   ChrW(&H78) & ChrW(&H69) & ChrW(&H6E) & " " & _
                   ChrW(&H76) & ChrW(&H75) & ChrW(&H69) & " " & _
                   ChrW(&H6C) & ChrW(&HF2) & ChrW(&H6E) & ChrW(&H67) & " " & _
                   ChrW(&H63) & ChrW(&H68) & ChrW(&H1EDD)
            HienThongBao sMsg, 1
            XoaPSThang Index
        End If
    Else
        If MsgBox("B¹n ch¾c ch¾n cÇn xãa sè d­ ®Çu n¨m?", vbYesNo + vbExclamation, App.ProductName) = vbYes Then
            Me.MousePointer = 11
            HienThongBao "Xãa sè d­ ®Çu n¨m, xin vui lßng chê !", 1
            XoaDK
        End If
    End If
    HienThongBao "", 1
    Me.MousePointer = 0
End Sub

Private Sub OptNN_Click(Index As Integer)
    CloseItemList
    pNN = Index
'    Img.Visible = (pNN = 0)
    SetFont Me, 1
End Sub

Private Sub sbStatusBar_PanelClick(ByVal Panel As ComctlLib.Panel)
'    Dim path As String
'
'    Select Case Panel.Index
'        Case 2:             Panel.Text = IIf(Panel.tag = 0, "Data File Size: " + Format(FileLen(pDataPath) / 1048576, Mask_2) + " MB, Version: " + IIf(DBKetoan.Version < 4, "97", "2000"), pDataPath)
'                                    Panel.tag = 1 - Panel.tag
'                                    StationList
'        Case 3:             Panel.Text = IIf(Panel.tag = 0, Panel.ToolTipText, UserName)
'                                    Panel.tag = 1 - Panel.tag
'        Case 4:             path = GetSetting(IniPath, "Environment", "BackUpPath")
'                                    path = FrmGetStr.GetString("Th­ môc l­u d÷ liÖu", App.ProductName, path)
'                                    SaveSetting IniPath, "Environment", "BackUpPath", path
'    End Select
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)

End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As ComctlLib.Button)
    Me.MousePointer = 11
    
    Select Case Button.key
        Case "TaiKhoan"
            FrmTaikhoan.tag = 1
            FrmTaikhoan.Show 1
        Case "NgoaiTe"
            FrmNguyente.Show 1
        Case "Kho"
            If STDetail Then
                'Load FrmKho
                FrmKho.tag = 1
                FrmKho.Show 1
            Else
                NoRight 1
            End If
        Case "VatTu"
            If STDetail Then
                FrmVattu.Show 1
            Else
                NoRight 1
            End If
        Case "LuuChuyen"
            If STDetail Then
                FrmLuuChuyen.Show 1
            Else
                NoRight 1
            End If
        Case "DuPhong"
            If STDetail Then
                FrmDuphong.Show 1
            Else
                NoRight 1
            End If
        Case "TaiSan"
            If FADetail Then
                pNghiepVu = NV_KHONG
                frmDSTaiSan.Show 1
            Else
                NoRight 2
            End If
        Case "CN"
            If KHDetail Then FrmKhachHang.Show vbModal
        Case "TongHop"
            FrmTongHop.Show 1
        Case "Help"
            mnuHLP_Click 0
        Case "KetThuc"
         'FVAT.tag = 1
         ' FVAT.Show 1
            Unload Me
            Exit Sub
        Case "ThanhPham"
            If STDetail Then
                FrmThanhPham.Show 1
            Else
                NoRight 1
            End If
    End Select
    HienThongBao "", 1
    Me.MousePointer = 0
End Sub

'======================================================================================
' Function GetLicense : Thí tñc lÃy tËn càng ty v¡ chi nh¤nh
'======================================================================================
Private Sub GetLicense()
    Dim rs_license As Recordset, i As Integer, k As Integer, sh As String

    CloseItemList
    DEMO = 1
    i = er_DBFile
    On Error Resume Next
    ' M? co s? d? li?u
    Set rs_license = DBKetoan.OpenRecordset("License", dbOpenSnapshot)

    If rs_license.EOF Then
        MsgBox "License  DB not working ", vbInformation, "Result"
        'End
    End If
    i = 0
    i = rs_license!Flag1 Mod 100
    If i > 0 Then
        If StationList() > i Then
            i = er_Connection
            Unload Me
            Exit Sub
        End If
    End If

    pTenCty = rs_license!tencty
    pTenCn = rs_license!tencn
        
    LbCty(2).Caption = rs_license!DiaChi
    LbCty(3).Caption = rs_license!Tel
    LbCty(4).Caption = rs_license!Fax
    LbCty(5).Caption = VniToUnicode(rs_license!TaiKhoanVN)
    LbCty(6).Caption = rs_license!TaiKhoanNT
    pNamTC = rs_license!NamTC
    pThangDauKy = rs_license!thang
    LbCty(7).Caption = CStr(pNamTC)
    LbCty(8).Caption = rs_license!masothue
    LbCty(13).Caption = rs_license!email
    LbCty(14).Caption = rs_license!sofax
    pBaoGia = (rs_license!Flag1 Mod 1000) \ 100
    pNVBH = (rs_license!Flag1 Mod 10000) \ 1000

    For i = 5 To 7
        mnCN(i).Visible = (pNVBH > 0)
    Next
    lb(0).tag = "Model"
    SetFont Me
    i = (rs_license!Flag1 Mod 1000000000) \ 100000000
    lb(0).tag = i
    If (i < 3 Or i = 5) And pVersion = 0 Then ExecuteSQL5 "UPDATE License SET Flag1=400000000+Flag1 Mod 100000000", False
    Select Case i
    Case 1: lb(1).Caption = "Doanh nghiÖp Nhµ n­íc"
        lb(0).Caption = "10.1."
    Case 2: lb(1).Caption = "Cæ phÇn - Liªn doanh"
        lb(0).Caption = "10.1."
    Case 3: lb(1).Caption = "C«ng ty TNHH"
        lb(0).Caption = "10.1"
    Case 4: lb(1).Caption = "Doanh nghiÖp t­ nh©n"
        lb(0).Caption = "10.1"
    Case 5: lb(1).Caption = "C¬ së ®µo t¹o"
        lb(0).Caption = "10.1"
    Case 6:
        lb(1).Caption = "Hµnh chÝnh sù nghiÖp"
        lb(0).Caption = "10.1"
        Label(24).Visible = False
        Label(25).Visible = False
        Frame(1).Visible = False
    Case Else
        lb(0).Caption = "10.1"
    End Select
    If pVersion <> 3 Then lb(0).Caption = lb(0).Caption    ' + IIf((rs_license!Flag1 Mod 100000000) \ 10000000 > 0, "1", "0") + IIf((rs_license!Flag1 Mod 10000000) \ 1000000 > 0, "1", "0") + IIf((rs_license!Flag1 Mod 1000000) \ 100000 > 0, "1", "0") + IIf((rs_license!Flag1 Mod 100000) \ 10000 > 0, "1", "0")
    Chk(0).Value = (rs_license!Flag1 Mod 100000000) \ 10000000
    Chk(1).Value = (rs_license!Flag1 Mod 10000000) \ 1000000
    Chk(2).Value = (rs_license!Flag1 Mod 1000000) \ 100000
    Chk(3).Value = (rs_license!Flag1 Mod 100000) \ 10000

    Command(6).Visible = ((rs_license!Flag1 Mod 1000000) \ 100000 > 0)

    Command(4).Visible = (rs_license!Flag1 \ 1000000000 > 0)

    pTygia = IIf(rs_license!tygia > 0, 1, 0)
    pHachToan = 1 - (rs_license!RptOrder Mod 10)
    pRpt = (rs_license!RptOrder Mod 100) \ 10
    OutCost = rs_license!OutCost
    FCost = rs_license!FixedOutCost
    STDetail = rs_license!STDetail
    FADetail = rs_license!FADetail
    KHDetail = rs_license!HDV
    pGiaHT = rs_license!GiaHT
    pGiaVon = (rs_license!STDetail Mod 100) \ 10
    pDTTP = (rs_license!STDetail Mod 1000) \ 100
    pDinhmuc = (rs_license!STDetail Mod 10000) \ 1000

    Command(5).Visible = ((rs_license!Lock0 Mod 100) \ 10 > 0)
    pCongNoHD = (rs_license!Lock0 Mod 1000) \ 100
    pPQTK = (rs_license!Lock0 Mod 10000) \ 1000
    pGiaUSD = (rs_license!Lock0 Mod 100000) \ 10000
    pChietKhau = (rs_license!Lock1 Mod 100) \ 10
    pKiemKeNgay = (rs_license!Lock1 Mod 1000) \ 100
    pNoiBo = (rs_license!Lock1 Mod 10000) \ 1000
    pSoVV = (rs_license!Lock1 Mod 100000) \ 10000
    pNhapKhau = (rs_license!Lock2 Mod 100) \ 10
    pBarCode = (rs_license!Lock2 Mod 1000) \ 100
    pNhapDoiTuong = (rs_license!Lock2 Mod 10000) \ 1000
    pTrungSoHieuKhacThang = (rs_license!Lock2 Mod 100000) \ 10000

    'mnVT(14).Visible = (pNhapKhau > 0)
    'mnVT(15).Visible = (pNhapKhau > 0)

    pTien = 0
    pTien = rs_license!loaitien
    If pTien > 0 Then
        Mask_0 = Mask_2
        pTienStr = TenNT(pTien)
    Else
        Mask_0 = GetSetting(IniPath, "Environment", "IntMask", "###,###,###,###")
        pTienStr = "VND"
    End If
    CTGS_GV = rs_license!CTGS_GV
    pFontName = rs_license!FontName
    pFontSize = rs_license!FontSize
    LbCty(0).FontName = pFontName
    LbCty(1).FontName = pFontName
    LbCty(10).Caption = rs_license!Quan
    LbCty(11).Caption = rs_license!ThanhPho
    frmMain.LbCty(9).Caption = rs_license!email
    pSoKT = rs_license!SoKT
    mnDL(13).Visible = (pSoKT Mod 100 >= 10)
    '    mnDL(14).Visible = (pSoKT Mod 100 >= 10)
    tbToolBar.Buttons("ThanhPham").Visible = (rs_license!RptOrder Mod 10000 >= 1000)
    tbToolBar.Buttons("ThanhPham2").Visible = (rs_license!RptOrder Mod 10000 >= 1000)
    pSongNgu = False
    pSongNgu = (pSoKT Mod 100000 >= 10000)
    pMaVach = 0
    pMaVach = rs_license!mv Mod 10
    pTyGiaBQ = 0
    pTyGiaBQ = IIf(rs_license!mv Mod 10000 >= 1000, 1, 0)
    tbToolBar.Buttons("TongHop").Visible = False
    tbToolBar.Buttons("TongHop").Visible = (rs_license!mv Mod 1000 >= 100)
    DEMO = IIf((rs_license!mv Mod 100000 >= 10000) And (rs_license!MKUP = pRev), 0, 1)
    NgayDauThangMoi = rs_license!NgayDauThang
    FontFlag = LoaiFont(pFontName)

    If (Not pSongNgu) And OptNN(1).Value Then OptNN(0).Value = True
    i = pNN
    pNN = 0

    pNN = i
    mnVT(4).Visible = (OutCost = 0 Or OutCost = 1 Or OutCost = 2)
    mnVT(10).Visible = (pKiemKeNgay > 0)
    For i = 8 To 11
        mnVT(i).Visible = (pDTTP <> 0)
    Next
    mnDL(19).Visible = pSongNgu

    sh = SelectSQL("SELECT App1Path AS F1 FROM License")
    Command(3).Visible = Len(Dir(sh)) > 0

    mnHT(6).Visible = (pTygia > 0)
    mnHT(10).Caption = IIf(FontFlag <> 2, "ChuyÓn ®æi CSDL sang font ABC", "ChuyÓn ®æi CSDL sang font VNI")
    mnHT(10).Visible = (rs_license!RptOrder Mod 1000 >= 100)

    mnCongno.Visible = KHDetail

    Me.Caption = "VietStar Accounting Software - "
    sh = LaySH(rs_license!TKVattu, 1, "-")
    If DEMO = 0 And pVersion <> 2 Then
        Me.Caption = Me.Caption + "12"    '+ sh

        If ((Int_StrToCode(rs_license!masothue) <> rs_license!MST_ID) Or (Int_StrToCode(pTenCty) <> rs_license!TenCty_ID) Or (Int_StrToCode(pTenCn) <> rs_license!tencn_id)) Then
            'If (1 > 2) Then
            pTenCty = ABCtoVNI("Sao chÐp kh«ng b¶n quyÒn")
            pTenCn = ABCtoVNI("Sao chÐp kh«ng b¶n quyÒn")
            ExecuteSQL5 "UPDATE License SET MST_ID=-1"
            pSTOP = 1
        End If
    Else
        Me.Caption = Me.Caption + sh + IIf(pVersion < 2, " - Training Version", " - Ch­¬ng tr×nh phèi hîp ®µo t¹o")
    End If
    If (boolean_kiemtra() = False) Then
        'Label5.Visible = False
        'Label5.Caption = ""
        'Label5.Caption = "  (" & ABCtoVNI("Phiªn b¶n dïng thö") & ")"
        'frmMain.txtdungthu.Caption = (" & ABCtoVNI("Phiªn b¶n dïng thö") & ")"
        frmMain.txtdungthu.Caption = "PhÇn mÒm hÕt h¹n dïng, vui lßng liªn hÖ víi nhµ cung cÊp!"
        ' dung khoa nut thay doi cau hinh doanh nghiep

        'If (SelectSQL("SELECT count(*) as F1 FROM ChungTu ") + rs_license!sodong > 300 Or SelectSQL("SELECT  DateDiff('d',min(NgayCT ), max(NgayCT ))  as F1 from chungtu") > 90) Then
        ' If (SelectSQL("SELECT count(*) as F1 FROM HoaDon ") >= 200 Or SelectSQL("SELECT Sum(ThanhTien) as F1 FROM HoaDon ") > 2000000000) Then

        Dim demslchungtu As Double
        demslchungtu = SelectSQL( _
                       "SELECT COUNT(*) AS F1 " & _
                       "FROM (" & _
                     " SELECT MaCT FROM ChungTu " & _
                     " WHERE SoHieu NOT LIKE '*GV*' " & _
                     " GROUP BY MaCT" & _
                       ") AS T" _
                     )
        Dim demdoanhthu As Double
        demdoanhthu = SelectSQL("SELECT sum(duco_12) as F1 from hethongtk where sohieu ='511' ")
        If (demslchungtu > 200 Or demdoanhthu > 2000000000#) Then
            ' pTenCty = ABCtoVNI("PhÇn mÒm hÕt h¹n dïng thö")
            '  pTenCn = ABCtoVNI("PhÇn mÒm hÕt h¹n dïng thö")
            FrmOptions.Text(0).Enabled = False
            FrmOptions.Text(1).Enabled = False
            FrmOptions.Text(7).Enabled = False
            FrmChungtu.Command(0).Enabled = False
            FrmChungtu.Command(1).Enabled = False
            FrmOptions.Combo(0).Enabled = False
        Else
            FrmOptions.Text(1).Enabled = False
            FrmOptions.Text(0).Enabled = True
            FrmOptions.Text(7).Enabled = True
            FrmChungtu.Command(0).Enabled = True
            FrmChungtu.Command(1).Enabled = True
            FrmOptions.Combo(0).Enabled = True
            frmMain.txtdungthu.Caption = ""

        End If
    Else
        frmMain.txtdungthu.Caption = ""
    End If
    If pVersion = 3 Then
        Me.Caption = Me.Caption + " - HCSN"
        pVATV = "3113"
        pSHPT = "3111"
    Else
        pVATV = "133"
        pSHPT = "131"
    End If

    LbCty(0).tag = rs_license!TenCty_ID
    LbCty(0).Caption = pTenCty
    LbCty(1).Caption = pTenCn
    Frame(0).Visible = pSongNgu

    mnXoa(0).tag = 0
    mnk(0).tag = 0
    mnk(0).Caption = mnk(0).Caption + IIf(rs_license.Fields("Lock0") Mod 10 > 0, "          x", "")
    For i = 1 To 12
        k = CThangFR(i)
        sh = IIf(rs_license.Fields("Lock" + CStr(i)) Mod 10 > 0, "          x", "")
        mnXoa(i).Caption = CStr(k) + "/" + CStr(pNamTC)
        mnk(i).Caption = CStr(k) + "/" + CStr(pNamTC) + sh
        mnXoa(i).tag = k
        mnk(i).tag = k
    Next

    rs_license.Close
    Set rs_license = Nothing

    LietKeNam
    mnVT(15).Visible = True
    'mnCongno.Visible = False

    On Error GoTo 0
End Sub

Private Sub NoRight(id As Integer)
    Select Case id
        Case 0: HienThongBao "Kh«ng cã quyÒn truy cËp!", 1
        Case 1: HienThongBao "Kh«ng ®¨ng ký theo dâi chi tiÕt vËt t­!", 1
        Case 2: HienThongBao "Kh«ng ®¨ng ký theo dâi chi tiÕt TSC§!", 1
    End Select
    Beep
End Sub

Private Sub LietKeNam()
    Dim rs As Recordset, i As Integer
        
    mnNam(MaxNamTC).Caption = CStr(pNamTC)
    mnNam(MaxNamTC).CHECKED = True
    If Not BangDaCo("NamTC") Then Exit Sub
    Set rs = DBKetoan.OpenRecordset("SELECT * FROM NamTC WHERE Nam<" + CStr(pNamTC) + " ORDER BY Nam DESC")
    i = MaxNamTC
    Do While (i > 0) And (Not rs.EOF)
        i = i - 1
        mnNam(i).Caption = CStr(rs!nam)
        mnNam(i).Visible = True
        mnNam(i).tag = rs!path
        rs.MoveNext
    Loop
    Do While (i > 0)
        i = i - 1
        mnNam(i).Visible = False
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Public Sub SetUserRight()
    Dim i As Integer
    
    Command(0).Enabled = (User_Right <> 2)
    Command(4).Enabled = (User_Right <> 2)
    
    For i = 1 To 11
        tbToolBar.Buttons(i).Enabled = (User_Right <> 2)
    Next
    
    For i = 2 To 4
        mnHT(i).Enabled = (User_Right = 0)
    Next
    
    For i = 10 To 11
        mnHT(i).Enabled = (User_Right = 0)
    Next
    
    mnHT(13).Enabled = (User_Right = 0)
    mnDL(0).Enabled = (User_Right = 0)
    mnDL(1).Enabled = (User_Right = 0)
    mnDL(3).Enabled = (User_Right = 0)
    mnDL(4).Enabled = (User_Right = 0)
    mnDL(7).Enabled = (User_Right = 0)
    
    For i = 9 To 12
        mnDL(i).Enabled = (Me.tag Mod 10 >= 1) Or (User_Right = 0)
    Next
    'mnKC(4).Enabled = (User_Right = 0)
    
    mnVatTu.Enabled = (Me.tag Mod 100 >= 10) Or (Me.tag Mod 1000 >= 100)
    'mnTSCD.Enabled = (Me.tag Mod 10000 >= 1000)
    mnCongno.Enabled = (Me.tag Mod 100000 >= 10000)
    Command(2).Enabled = (User_Right <> 1) And (Me.tag Mod 10 >= 1)
    Command(6).Enabled = (User_Right <> 1) And (Me.tag Mod 10 >= 1)
End Sub

Private Sub DatTKCN()
    Dim shtk As String, TK As New ClsTaikhoan
    
    FrmGetStr.tag = 2
    shtk = FrmGetStr.GetString("Sè hiÖu TK", "§Æt/Bá TK theo dâi chi tiÕt")
    If Len(shtk) = 0 Then GoTo KT
    TK.InitTaikhoanSohieu shtk
    If TK.MaSo = 0 Then GoTo KT
    If TK.tk_id = TKVT_ID Or TK.tk_id = TSCD_ID Or TK.tk_id = KHTSCD_ID Or TK.tk_id = TKThue_ID Or TK.tk_id = TKDT_ID Then Exit Sub
    If TK.TkCoPS(0, 0) Or TK.NoDauKy <> 0 Or TK.CoDauKy <> 0 Then
        Me.MousePointer = 11
        If TK.ChuyenChiTietSangDoiTuong Then
            MsgBox "C¸c chi tiÕt tµi kho¶n ®· ®­îc m· ho¸ thµnh ®èi t­îng c«ng nî!", vbCritical, App.ProductName
        Else
            MsgBox "Tµi kho¶n kh«ng chuyÓn ®æi ®­îc!", vbCritical, App.ProductName
        End If
        Me.MousePointer = 0
        GoTo KT
    End If
    If TK.tk_id = TKCNKH_ID Or TK.tk_id = TKCNPT_ID Then ExecuteSQL5 "DELETE SoDuKhachHang.* FROM SoDuKhachHang INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo WHERE HethongTK.SoHieu LIKE '" + TK.sohieu + "*'"
    If TK.loai < 3 Then ExecuteSQL5 "UPDATE HethongTK SET TK_ID=" + IIf(TK.tk_id = TKCNKH_ID, "0", CStr(TKCNKH_ID)) + " WHERE SoHieu LIKE '" + TK.sohieu + "*'"
    If TK.loai > 2 Then ExecuteSQL5 "UPDATE HethongTK SET TK_ID=" + IIf(TK.tk_id = TKCNPT_ID, "0", CStr(TKCNPT_ID)) + " WHERE SoHieu LIKE '" + TK.sohieu + "*'"
KT:
    Set TK = Nothing
End Sub

Private Sub DatTKVT()
    Dim shtk As String, TK As New ClsTaikhoan
    
    FrmGetStr.tag = 1
    shtk = FrmGetStr.GetString("Sè hiÖu TK", "§Æt/Bá TK theo dâi chi tiÕt")
    If Len(shtk) = 0 Then Exit Sub
    TK.InitTaikhoanSohieu shtk
    If TK.MaSo = 0 Then GoTo KT
    If TK.tk_id = TKCNKH_ID Or TK.tk_id = TKCNPT_ID Or TK.tk_id = TSCD_ID Or TK.tk_id = KHTSCD_ID Or TK.tk_id = TKThue_ID Or TK.tk_id = TKDT_ID Then Exit Sub
    If TK.TkCoPS(0, 0) Or TK.NoDauKy <> 0 Or TK.CoDauKy <> 0 Then
        MsgBox "Tµi kho¶n cã ph¸t sinh hoÆc ®Çu kú, kh«ng chuyÓn ®æi ®­îc!", vbCritical, App.ProductName
        GoTo KT
    End If
    If TK.tk_id = TKVT_ID Then ExecuteSQL5 "DELETE TonKho.* FROM TonKho INNER JOIN HethongTK ON TonKho.MaTaiKhoan=HethongTK.MaSo WHERE HethongTK.SoHieu LIKE '" + TK.sohieu + "*'"
    ExecuteSQL5 "UPDATE HethongTK SET TK_ID=" + IIf(TK.tk_id = TKVT_ID, "0", CStr(TKVT_ID)) + " WHERE SoHieu LIKE '" + TK.sohieu + "*'"
KT:
    Set TK = Nothing
End Sub

Private Sub DatTKDTTP()
    Dim shtk As String, TK As New ClsTaikhoan
    
    FrmGetStr.tag = 4
    shtk = FrmGetStr.GetString("Sè hiÖu TK", "§Æt/Bá TK h¹ch to¸n doanh thu")
    If Len(shtk) = 0 Then GoTo KT
    TK.InitTaikhoanSohieu shtk
    If TK.MaSo = 0 Or Left(TK.sohieu, 2) <> "51" Then GoTo KT
    If TK.TkCoPS(0, 0) Then
        MsgBox "Tµi kho¶n cã ph¸t sinh, kh«ng chuyÓn ®æi ®­îc!", vbCritical, App.ProductName
        GoTo KT
    End If
    ExecuteSQL5 "UPDATE HethongTK SET TK_ID2=" + IIf(TK.tk_id2 = TKDT_ID, "0", CStr(TKDT_ID)) + " WHERE SoHieu LIKE '" + TK.sohieu + "*'"
KT:
    Set TK = Nothing
End Sub

Private Sub DatTKTS()
    Dim shtk As String, TK As New ClsTaikhoan
    
    FrmGetStr.tag = 3
    shtk = FrmGetStr.GetString("Sè hiÖu TK", "§Æt/Bá TK theo dâi chi tiÕt")
    If Len(shtk) = 0 Then Exit Sub
    TK.InitTaikhoanSohieu shtk
    If TK.MaSo = 0 Then GoTo KT
    ExecuteSQL5 "UPDATE HethongTK SET TK_ID2=" + IIf(TK.tk_id2 = TKCPSX_ID, "0", CStr(TKCPSX_ID)) + " WHERE SoHieu LIKE '" + TK.sohieu + "*'"
KT:
    Set TK = Nothing
End Sub

Private Sub RunCT()
    Dim pctpath As String
    
    pctpath = SelectSQL("SELECT App1Path AS F1 FROM License")
    If Len(Dir(pctpath)) > 0 Then Shell pctpath, vbNormalFocus
End Sub

Public Function ChonTenTep(Title As String, f As Long, mask As String, act As Integer) As String
    With dlgCommonDialog
        .InitDir = pCurDir + "data\"
        .DialogTitle = Title
        .Flags = f
        .fileName = mask
        .DefaultExt = mask
        .Filter = "TÖp d÷ liÖu (" + mask + ")|" + mask + "|TÊt c¶ (*.*)|*.*"
        On Error GoTo Xong
        Select Case act
        Case 1: .ShowOpen
        Case 2: .ShowSave
        Case 3: .ShowPrinter
        Case 4: .ShowFont
        End Select
        On Error GoTo 0
        If Len(.fileName) = 0 Or Left(.fileName, 1) = "*" Then GoTo Xong

        If act = 2 Then
            If Len(Dir(.fileName)) > 0 Then
                If .fileName = pDataPath Then
                    MsgBox "TÖp d÷ liÖu ®ang më !", vbCritical, App.ProductName
                    GoTo Xong
                End If
                If MsgBox("TÖp " + .fileName + " ®· tån t¹i, tiÕp tôc ? !", vbQuestion + vbYesNo, App.ProductName) = vbNo Then GoTo Xong
                If Recycle(.fileName) <> 0 Then
                    MsgBox "Kh«ng xo¸ ®­îc tÖp " + dlgCommonDialog.fileName + " !", vbExclamation, App.ProductName
                    GoTo Xong
                End If
            End If
        End If
        ChonTenTep = .fileName

        Dim folderPath As String
        Dim FilePath As String

        folderPath = App.path & "\Hoadon"
        FilePath = folderPath & "\vbPath.txt"

        Open FilePath For Output As #1
        Print #1, ChonTenTep;
        Close #1

    End With
Xong:
End Function

Private Function StationList() As Integer
    ' It is important that ReDim be used to define the array as the DLL,
' because the DLL depends on being able to redimension the array.
    ReDim msString(1) As String
' The array is 1-based rather than 0-based, regardless if Option Base 1
' is specified in the declarations section.
    Dim miLoop As Integer, i As Integer, LDBName As String, sql As String, U As String, X As String
 
    LDBName = Left(pDataPath, Len(pDataPath) - 3) + "LDB"
'    miLoop = LDBUser_GetUsers(msString, LDBName, OptLDBLoggedUsers)
' The function calls cannot be combined and must be used individually.
' Get the first user in the selected .LDB file.
    For i = 0 To miLoop - 1
        If i >= LBound(msString, 1) And i <= UBound(msString, 1) Then
            U = SelectSQL("SELECT TenNSD AS F1 FROM Users WHERE WS='" + msString(i) + "' AND TenNSD<>'" + X + "'")
            If U <> "0" Then
                sql = sql + Chr(13) + msString(i) + " : " + U
                X = U
            End If
        End If
    Next
    If miLoop > 1 Then
        LbCty(12).Caption = "C¸c m¸y tr¹m: " + sql
    Else
        LbCty(12).Caption = ""
    End If
    
    StationList = miLoop
End Function

Private Sub LietKeTep()
    Dim i As Integer, fn As String, k As Integer
    
    For i = 1 To 5
        fn = GetSetting(IniPath, "RecentFiles", "File" + CStr(i))
        If Len(fn) > 0 And fn <> pDataPath Then
            'mnHT(17 + i).Caption = fn
           ' mnHT(17 + i).Visible = True
            k = k + 1
        Else
           ' mnHT(17 + i).Visible = False
        End If
    Next
    mnHT(23).Visible = (k > 0)
End Sub

Private Sub XoaQuery()
    Dim q As String
    
    q = InputBox("Tªn query cÇn xo¸: ", App.ProductName)
    If Len(q) > 0 Then
        If QueryDaCo(q) Then DBKetoan.QueryDefs.Delete q
    End If
End Sub

Private Sub FontSetUp()
    Add32Font "VNTIME.TTF"
    Add32Font "VNTIMEB.TTF"
    Add32Font "VNTIMEBI.TTF"
    Add32Font "VNTIMEI.TTF"
    
    Add32Font "VHTIME.TTF"
    Add32Font "VHTIMEB.TTF"
    Add32Font "VHTIMEBI.TTF"
    Add32Font "VHTIMEI.TTF"
    
    Add32Font "VTIMESN.TTF"
    Add32Font "VTIMESB.TTF"
    Add32Font "VTIMESBI.TTF"
    Add32Font "VTIMESI.TTF"
End Sub


Private Sub timerBackup_Timer()
    Dim backupPath As String
    Dim fso As Object
    backupPath = pCurDir + "DATA\backup3.MDB"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(pDataPath) Then
        fso.CopyFile pDataPath, backupPath
        'MsgBox "Sao luu thành công: " & backupPath, vbInformation
    Else
        MsgBox "File co s? d? li?u không t?n t?i!", vbExclamation
    End If

    ' Gi?i phóng d?i tu?ng
    Set fso = Nothing
End Sub

Private Sub timerChaytb_Timer()
    timerChaytb.Enabled = False

    Dim fileNumber As Integer
    Dim content As String
    Dim FilePath As String
    FilePath = App.path & "\\HoaDon\\status.txt"
    content = "11"
    fileNumber = FreeFile
    On Error Resume Next
    Open FilePath For Output As #fileNumber
    If Err.number = 0 Then
        Print #fileNumber, content;
        Close #fileNumber
        'MsgBox "Ðã ghi dè file version.txt thành công!", vbInformation
    Else
        MsgBox "L?i khi ghi dè file!", vbExclamation
    End If
    Dim exePath2 As String
    Dim cmd2 As String
    Dim taskID As Double

    DoEvents  ' Ð? d?m b?o ?ng d?ng có th?i gian kh?i d?ng
    exePath2 = App.path & "\Tools\Debug\SaovietTax.exe"
    ' Dùng runas v?i trust level th?p hon
    cmd2 = "runas /trustlevel:0x20000 """ & exePath2 & """"
    ' L?y Process ID
    taskID = Shell(cmd2, vbNormalFocus)    ' Ho?c vbNormalFocus d? hi?n th? và focus

    ' Ch? ?ng d?ng kh?i t?o
    Sleep 500

    ' Kích ho?t ?ng d?ng
    AppActivate taskID

    'Reset ve lai 1
    content = "1"
    fileNumber = FreeFile
    On Error Resume Next
    Open FilePath For Output As #fileNumber
    If Err.number = 0 Then
        Print #fileNumber, content;
        Close #fileNumber
        'MsgBox "Ðã ghi dè file version.txt thành công!", vbInformation
    Else
        MsgBox "L?i khi ghi dè file!", vbExclamation
    End If
End Sub
