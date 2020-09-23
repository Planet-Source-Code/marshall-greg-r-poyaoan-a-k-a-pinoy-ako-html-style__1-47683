VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "HTML styles"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optHover 
      Caption         =   "Hover "
      Height          =   375
      Left            =   1680
      TabIndex        =   105
      Top             =   6360
      Width           =   1215
   End
   Begin VB.OptionButton optScroll 
      Caption         =   "Scroll Bar"
      Height          =   375
      Left            =   240
      TabIndex        =   40
      Top             =   6720
      Width           =   1335
   End
   Begin VB.OptionButton optButton 
      Caption         =   "Button style"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   6360
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox txtGenCode1 
      Height          =   2775
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   360
      Width           =   4695
   End
   Begin VB.CommandButton cmdPreviewIE 
      Caption         =   "Preview IE"
      Height          =   525
      Left            =   7920
      TabIndex        =   3
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdGenerateCode 
      Caption         =   "Generate Code"
      Height          =   525
      Left            =   5520
      TabIndex        =   1
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdPreviewhtml 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Preview"
      Height          =   525
      Left            =   6720
      TabIndex        =   0
      Top             =   6600
      Width           =   1095
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2655
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   4695
      ExtentX         =   8281
      ExtentY         =   4683
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Frame fraInvisible 
      Caption         =   "Invisibles"
      Height          =   2175
      Left            =   2880
      TabIndex        =   5
      Top             =   4680
      Width           =   2295
      Begin VB.TextBox txtAMainFont 
         Height          =   375
         Left            =   1200
         TabIndex        =   131
         Text            =   "16px  Arial"
         Top             =   120
         Width           =   615
      End
      Begin VB.Timer tmrTempFile3 
         Interval        =   1
         Left            =   1080
         Top             =   360
      End
      Begin VB.TextBox txtHoverFont 
         Height          =   375
         Left            =   1200
         TabIndex        =   124
         Text            =   "16px  Arial"
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtScrollColor8 
         Height          =   375
         Left            =   720
         MaxLength       =   6
         TabIndex        =   102
         Text            =   "FFFFFF"
         Top             =   1440
         Width           =   735
      End
      Begin MSComDlg.CommonDialog dlgImage2 
         Left            =   360
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog dlgImage1 
         Left            =   360
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtScrollColor7 
         Height          =   375
         Left            =   720
         MaxLength       =   6
         TabIndex        =   62
         Text            =   "FFFFFF"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtScrollColor6 
         Height          =   375
         Left            =   720
         MaxLength       =   6
         TabIndex        =   61
         Text            =   "FF6600"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtScrollColor2 
         Height          =   405
         Left            =   600
         MaxLength       =   6
         TabIndex        =   56
         Text            =   "FF7C25"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtScrollColor3 
         Height          =   405
         Left            =   720
         MaxLength       =   6
         TabIndex        =   55
         Text            =   "F8CDB0"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtScrollColor4 
         Height          =   405
         Left            =   840
         MaxLength       =   6
         TabIndex        =   54
         Text            =   "FFEECC"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtScrollColor5 
         Height          =   405
         Left            =   720
         MaxLength       =   6
         TabIndex        =   53
         Text            =   "FF6600"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtScrollColor1 
         Height          =   405
         Left            =   720
         MaxLength       =   6
         TabIndex        =   52
         Text            =   "FFD0B0"
         Top             =   1560
         Width           =   735
      End
      Begin VB.Timer tmrTempFile2 
         Interval        =   1
         Left            =   1560
         Top             =   480
      End
      Begin VB.Timer tmrMain 
         Interval        =   1
         Left            =   720
         Top             =   960
      End
      Begin MSComDlg.CommonDialog dlgButton1 
         Left            =   120
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog dlgButton2 
         Left            =   120
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtCursor 
         Height          =   405
         Left            =   720
         TabIndex        =   13
         Text            =   "hand"
         Top             =   480
         Width           =   495
      End
      Begin VB.Timer tmrTempFile1 
         Interval        =   1
         Left            =   1680
         Top             =   960
      End
      Begin VB.Timer tmrButton 
         Interval        =   1
         Left            =   1200
         Top             =   960
      End
      Begin VB.TextBox txtButtonColor1 
         Height          =   405
         Left            =   0
         MaxLength       =   6
         TabIndex        =   9
         Text            =   "FF0000"
         Top             =   960
         Width           =   735
      End
      Begin MSComDlg.CommonDialog dlgButton3 
         Left            =   120
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin RichTextLib.RichTextBox txtGenCode2 
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393217
         TextRTF         =   $"frmMain.frx":0000
      End
      Begin VB.TextBox txtButtonColor3 
         Height          =   435
         Left            =   0
         MaxLength       =   6
         TabIndex        =   8
         Text            =   "AA0000"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtButtonColor2 
         Height          =   405
         Left            =   0
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "FFFFFF"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtFont 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Text            =   "10px  Arial"
         Top             =   480
         Width           =   615
      End
      Begin MSComDlg.CommonDialog dlgColorAll 
         Left            =   1560
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame fraHover 
      Caption         =   "Hover"
      Height          =   5895
      Left            =   5280
      TabIndex        =   106
      Top             =   240
      Width           =   3855
      Begin VB.CommandButton cmdHoverBorder 
         Caption         =   "Border"
         Height          =   375
         Left            =   480
         TabIndex        =   132
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2280
         TabIndex        =   129
         Text            =   "Pinoy Ako!"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdAMainFont 
         Caption         =   "Fonts"
         Height          =   375
         Left            =   480
         TabIndex        =   128
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton cmdHoverFont 
         Caption         =   "Fonts"
         Height          =   375
         Left            =   2160
         TabIndex        =   123
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txtVC 
         Height          =   375
         Left            =   480
         TabIndex        =   122
         Text            =   "FFFFFF"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtAAC 
         Height          =   375
         Left            =   2280
         TabIndex        =   120
         Text            =   "FFFFFF"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtLC 
         Height          =   375
         Left            =   480
         TabIndex        =   118
         Text            =   "FFFFFF"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtAC 
         Height          =   375
         Left            =   2280
         TabIndex        =   116
         Text            =   "000000"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtABC 
         Height          =   375
         Left            =   480
         TabIndex        =   115
         Text            =   "FF00AA"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtHC 
         Height          =   375
         Left            =   2280
         TabIndex        =   112
         Text            =   "FFFFFF"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtHBC 
         Height          =   375
         Left            =   480
         TabIndex        =   111
         Text            =   "0000FF"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdAColor 
         Caption         =   "Generate  HTML color"
         Height          =   495
         Left            =   2040
         TabIndex        =   108
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox txtAcolor 
         Height          =   495
         Left            =   480
         MaxLength       =   6
         TabIndex        =   107
         Text            =   "FF6600"
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label45 
         Caption         =   "Border:"
         Height          =   255
         Left            =   480
         TabIndex        =   133
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label43 
         Caption         =   "Author:"
         Height          =   255
         Left            =   2280
         TabIndex        =   130
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label41 
         Caption         =   "Edit main font:"
         Height          =   255
         Left            =   480
         TabIndex        =   127
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label48 
         Caption         =   "Edit Hover font:"
         Height          =   255
         Left            =   2160
         TabIndex        =   126
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label47 
         Caption         =   "Generate HTML color:"
         Height          =   255
         Left            =   1200
         TabIndex        =   125
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label46 
         Caption         =   "Visited color:"
         Height          =   255
         Left            =   480
         TabIndex        =   121
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label44 
         Caption         =   "Active color:"
         Height          =   255
         Left            =   2280
         TabIndex        =   119
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label42 
         Caption         =   "Link color:"
         Height          =   255
         Left            =   480
         TabIndex        =   117
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label40 
         Caption         =   "Main BG color:"
         Height          =   255
         Left            =   480
         TabIndex        =   114
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label39 
         Caption         =   "Main Color"
         Height          =   375
         Left            =   2280
         TabIndex        =   113
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label38 
         Caption         =   "Hover color:"
         Height          =   255
         Left            =   2280
         TabIndex        =   110
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label37 
         Caption         =   "Hover BG color:"
         Height          =   375
         Left            =   480
         TabIndex        =   109
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame fraScroll 
      Caption         =   "Scroll Bar"
      Height          =   4575
      Left            =   5280
      TabIndex        =   41
      Top             =   240
      Width           =   3855
      Begin VB.CommandButton cmdScroolFC 
         Caption         =   "Select"
         Height          =   375
         Left            =   2400
         TabIndex        =   104
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdScrollAC 
         Caption         =   "Select"
         Height          =   375
         Left            =   2400
         TabIndex        =   58
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton cmdScrollBC 
         Caption         =   "Select"
         Height          =   375
         Left            =   2400
         TabIndex        =   57
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdScrollDC 
         Caption         =   "Select"
         Height          =   375
         Left            =   2400
         TabIndex        =   46
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdScrollTC 
         Caption         =   "Select"
         Height          =   375
         Left            =   2400
         TabIndex        =   45
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdScroll3DC 
         Caption         =   "Select"
         Height          =   375
         Left            =   2400
         TabIndex        =   44
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdScrollSC 
         Caption         =   "Select"
         Height          =   375
         Left            =   2400
         TabIndex        =   43
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdScrollHC 
         Caption         =   "Select"
         Height          =   375
         Left            =   2400
         TabIndex        =   42
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "Scrollbar-face-color:"
         Height          =   255
         Left            =   360
         TabIndex        =   103
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label28 
         Caption         =   "Scrollbar-arrow-color:"
         Height          =   255
         Left            =   360
         TabIndex        =   60
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label27 
         Caption         =   "Scrollbar-base-color:"
         Height          =   255
         Left            =   360
         TabIndex        =   59
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label label26 
         Caption         =   "Scrollbar-darkshadow-color:"
         Height          =   255
         Left            =   360
         TabIndex        =   51
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label label25 
         Caption         =   "Scrollbar-track-color:"
         Height          =   255
         Left            =   360
         TabIndex        =   50
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label23 
         Caption         =   "Scrollbar-3Dlight-color:"
         Height          =   255
         Left            =   360
         TabIndex        =   49
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label19 
         Caption         =   "Scrollbar-shadow-color:"
         Height          =   255
         Left            =   360
         TabIndex        =   48
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "Scrollbar-highlight-color:"
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame fraButton 
      Caption         =   "Button Style"
      Height          =   6015
      Left            =   5280
      TabIndex        =   15
      Top             =   240
      Width           =   3855
      Begin VB.Frame fraImage 
         Caption         =   "Image"
         Height          =   4575
         Left            =   240
         TabIndex        =   87
         Top             =   240
         Width           =   3375
         Begin VB.CommandButton cmdImage2 
            Caption         =   "..."
            Height          =   375
            Left            =   2880
            TabIndex        =   101
            Top             =   1320
            Width           =   375
         End
         Begin VB.CommandButton cmdImage1 
            Caption         =   "..."
            Height          =   375
            Left            =   2880
            TabIndex        =   100
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtImage2 
            Height          =   375
            Left            =   120
            TabIndex        =   99
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txtImage1 
            Height          =   375
            Left            =   120
            TabIndex        =   98
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtImageWidth 
            Height          =   375
            Left            =   120
            TabIndex        =   91
            Text            =   "150"
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox txtImageHeight 
            Height          =   375
            Left            =   1560
            TabIndex        =   90
            Text            =   "30"
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox txtUrl2 
            Height          =   375
            Left            =   120
            TabIndex        =   89
            Text            =   "http://yahoo.com"
            Top             =   2880
            Width           =   2655
         End
         Begin VB.TextBox txtImageCur 
            Height          =   375
            Left            =   120
            TabIndex        =   88
            Text            =   "hand"
            Top             =   3960
            Width           =   2775
         End
         Begin VB.Label Label30 
            Caption         =   "Main Image (Change "" \ "" to "" / ""' )"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label31 
            Caption         =   "Mouseover Image (Change "" \ "" to "" / ""' )"
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   1080
            Width           =   3015
         End
         Begin VB.Label Label32 
            Caption         =   "Image Width"
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label33 
            Caption         =   "Image Height"
            Height          =   255
            Left            =   1560
            TabIndex        =   94
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label34 
            Caption         =   "URL:"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   2640
            Width           =   375
         End
         Begin VB.Label Label35 
            Caption         =   "Cursor: Example: auto | crosshair | default | move | text | wait | help ,etc.."
            Height          =   375
            Left            =   120
            TabIndex        =   92
            Top             =   3480
            Width           =   2775
         End
      End
      Begin VB.OptionButton optImage 
         Caption         =   "Image"
         Height          =   375
         Left            =   1680
         TabIndex        =   86
         Top             =   5400
         Width           =   735
      End
      Begin VB.TextBox txtUrl 
         Height          =   375
         Left            =   2040
         TabIndex        =   73
         Text            =   "index.htm"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton cmdButtonCur 
         Caption         =   "Edit"
         Height          =   375
         Left            =   1320
         TabIndex        =   72
         Top             =   2640
         Width           =   615
      End
      Begin VB.CommandButton cmdButtonBorder 
         Caption         =   "Edit"
         Height          =   375
         Left            =   600
         TabIndex        =   71
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox txtButLeft 
         Height          =   375
         Left            =   1320
         TabIndex        =   70
         Text            =   "10"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtButTop 
         Height          =   375
         Left            =   600
         TabIndex        =   69
         Text            =   "50"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtButWidth 
         Height          =   375
         Left            =   2040
         TabIndex        =   68
         Text            =   "120"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtButHeight 
         Height          =   375
         Left            =   2760
         TabIndex        =   67
         Text            =   "30"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdButtonBGMO 
         Caption         =   "Select"
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdButtobBG 
         Caption         =   "Selcet"
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdTextColor 
         Caption         =   "Select"
         Height          =   375
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdFont 
         Caption         =   "Select"
         Height          =   375
         Left            =   600
         TabIndex        =   63
         Top             =   1800
         Width           =   615
      End
      Begin VB.OptionButton optAlpha 
         Caption         =   "Alpha"
         Height          =   375
         Left            =   600
         TabIndex        =   20
         Top             =   4920
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optWave 
         Caption         =   "Wave"
         Height          =   375
         Left            =   1680
         TabIndex        =   19
         Top             =   4920
         Width           =   855
      End
      Begin VB.OptionButton optBlur 
         Caption         =   "Blur"
         Height          =   375
         Left            =   2760
         TabIndex        =   18
         Top             =   4920
         Width           =   855
      End
      Begin VB.OptionButton optNormal 
         Caption         =   "Normal"
         Height          =   375
         Left            =   600
         TabIndex        =   17
         Top             =   5400
         Width           =   855
      End
      Begin VB.Frame fraButtonAlpha 
         Caption         =   "Button CSS filter :  Alpha "
         Height          =   1455
         Left            =   600
         TabIndex        =   21
         Top             =   3360
         Width           =   2775
         Begin VB.TextBox txtAlphaStyle 
            Height          =   405
            Left            =   1920
            TabIndex        =   24
            Text            =   "2"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtAlphaFinOp 
            Height          =   405
            Left            =   1080
            TabIndex        =   23
            Text            =   "0"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtAlphaOp 
            Height          =   405
            Left            =   240
            TabIndex        =   22
            Text            =   "100"
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Style: 0-3"
            Height          =   255
            Left            =   1800
            TabIndex        =   27
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Finish Opacity: 0-100"
            Height          =   615
            Left            =   1080
            TabIndex        =   26
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Opacity: 0-100"
            Height          =   495
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame fraButtonNormal 
         Caption         =   "Normal"
         Height          =   1455
         Left            =   600
         TabIndex        =   16
         Top             =   3360
         Width           =   2775
         Begin VB.Line Line1 
            X1              =   240
            X2              =   120
            Y1              =   600
            Y2              =   960
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   480
            Y1              =   480
            Y2              =   960
         End
         Begin VB.Line Line3 
            X1              =   600
            X2              =   360
            Y1              =   600
            Y2              =   960
         End
         Begin VB.Line Line4 
            X1              =   960
            X2              =   720
            Y1              =   600
            Y2              =   840
         End
         Begin VB.Line Line5 
            X1              =   840
            X2              =   1200
            Y1              =   600
            Y2              =   840
         End
         Begin VB.Line Line6 
            X1              =   720
            X2              =   960
            Y1              =   720
            Y2              =   1080
         End
         Begin VB.Line Line7 
            X1              =   1200
            X2              =   840
            Y1              =   720
            Y2              =   1080
         End
         Begin VB.Line Line8 
            X1              =   1320
            X2              =   1320
            Y1              =   600
            Y2              =   1080
         End
         Begin VB.Line Line9 
            X1              =   1200
            X2              =   1560
            Y1              =   600
            Y2              =   720
         End
         Begin VB.Line Line10 
            X1              =   1200
            X2              =   1560
            Y1              =   960
            Y2              =   600
         End
         Begin VB.Line Line11 
            X1              =   1320
            X2              =   1680
            Y1              =   840
            Y2              =   1080
         End
         Begin VB.Line Line12 
            X1              =   1800
            X2              =   1800
            Y1              =   600
            Y2              =   1080
         End
         Begin VB.Line Line13 
            X1              =   1800
            X2              =   1920
            Y1              =   600
            Y2              =   840
         End
         Begin VB.Line Line14 
            X1              =   2040
            X2              =   1800
            Y1              =   600
            Y2              =   840
         End
         Begin VB.Line Line15 
            X1              =   2040
            X2              =   2040
            Y1              =   480
            Y2              =   960
         End
         Begin VB.Line Line16 
            X1              =   2280
            X2              =   1920
            Y1              =   600
            Y2              =   1320
         End
         Begin VB.Line Line17 
            X1              =   2160
            X2              =   2280
            Y1              =   600
            Y2              =   1320
         End
         Begin VB.Line Line18 
            X1              =   1920
            X2              =   2280
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line19 
            X1              =   2520
            X2              =   2520
            Y1              =   480
            Y2              =   1200
         End
         Begin VB.Line Line20 
            X1              =   2400
            X2              =   2640
            Y1              =   1080
            Y2              =   1080
         End
      End
      Begin VB.Frame fraButtonWave 
         Caption         =   "Button CSS filter :  Wave"
         Height          =   1455
         Left            =   600
         TabIndex        =   33
         Top             =   3360
         Width           =   2775
         Begin VB.TextBox txtWaveStr 
            Height          =   405
            Left            =   240
            TabIndex        =   36
            Text            =   "1"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtWaveFreq 
            Height          =   405
            Left            =   1080
            TabIndex        =   35
            Text            =   "3"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtWavelStr 
            Height          =   405
            Left            =   1920
            TabIndex        =   34
            Text            =   "55"
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "Light- strength: 0-100"
            Height          =   615
            Left            =   1920
            TabIndex        =   39
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "Freq: 0-100"
            Height          =   375
            Left            =   1080
            TabIndex        =   38
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "Strength: 0-100"
            Height          =   495
            Left            =   240
            TabIndex        =   37
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame fraButtonBlur 
         Caption         =   "Buttonb CSS filter: Blur"
         Height          =   1455
         Left            =   1080
         TabIndex        =   28
         Top             =   3360
         Width           =   2055
         Begin VB.TextBox txtBlurstr 
            Height          =   375
            Left            =   1200
            TabIndex        =   30
            Text            =   "90"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtBlurDir 
            Height          =   375
            Left            =   240
            TabIndex        =   29
            Text            =   "180"
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Direction (Angle)"
            Height          =   495
            Left            =   240
            TabIndex        =   31
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Strenght 0-1000"
            Height          =   615
            Left            =   1200
            TabIndex        =   32
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Label Label29 
         Caption         =   "URL:"
         Height          =   255
         Left            =   2040
         TabIndex        =   85
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "BG Color"
         Height          =   375
         Left            =   2040
         TabIndex        =   84
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Mouseover BG Color"
         Height          =   375
         Left            =   2760
         TabIndex        =   83
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Text Color"
         Height          =   375
         Left            =   1320
         TabIndex        =   82
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Fonts"
         Height          =   255
         Left            =   600
         TabIndex        =   81
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label24 
         Caption         =   "URL:"
         Height          =   255
         Left            =   2040
         TabIndex        =   80
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "Cursor"
         Height          =   255
         Left            =   1320
         TabIndex        =   79
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Border"
         Height          =   255
         Left            =   600
         TabIndex        =   78
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Left"
         Height          =   255
         Left            =   1440
         TabIndex        =   77
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Top"
         Height          =   255
         Left            =   600
         TabIndex        =   76
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Width"
         Height          =   255
         Left            =   2040
         TabIndex        =   75
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Height"
         Height          =   255
         Left            =   2760
         TabIndex        =   74
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Label Label21 
      Caption         =   "Generated Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label20 
      Caption         =   "Preview:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'constant
Const TempFile1 = "c:\windows\temp\1.html"
Const TempFile2 = "c:\windows\temp\2.html"
Const TempFile3 = "c:\windows\temp\3.html"
Const TempFile4 = "c:\windows\temp\4.html"
Const TempFile5 = "c:\windows\temp\5.html"

Option Explicit
Dim prevwalp As Boolean
Dim cur As String
'html color generator ;hover part
Private Sub cmdAColor_Click()
On Error Resume Next
dlgColorAll.ShowColor
txtAcolor.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub
'<A></A> fonts
Private Sub cmdAMainFont_Click()
Dim itl As String
Dim bld As String
On Error Resume Next
dlgButton1.Flags = 3
dlgButton1.FontName = "Arial"
dlgButton1.ShowFont
If dlgButton1.FontItalic = True Then
itl = "italic"
Else
itl = ""
End If
If dlgButton1.FontBold = True Then
bld = "bold"
Else
bld = ""
End If
txtAMainFont.Text = itl & " " & bld & " " & dlgButton1.FontSize & "px" & " " & dlgButton1.FontName
End Sub

'button border
Private Sub cmdButtonBorder_Click()
Form2.Show
Form1.Enabled = False
End Sub
'button cursor
Private Sub cmdButtonCur_Click()
cur = InputBox("Cursor: ( auto | crosshair | default | move | e-resize | ne-resize | nw-resize | n-resize | se-resize | sw-resize | w-resize | text | wait | help )", "Cursor Editor", "hand")
txtCursor.Text = cur
End Sub

Private Sub cmdHoverBorder_Click()
Form2.Show
Form1.Enabled = False
End Sub

'image choose file location
Private Sub cmdImage1_Click()
On Error Resume Next
dlgImage1.Filter = "Image Files (*.gif)|*.gif|All Files (*.*)|*.*"
dlgImage1.ShowOpen
MsgBox "change ' \ ' to  ' / ' For example c:\1.bmp change to c:/1.bmp. And a folder must not have a space in there name. For exmaple 'My Documents' it should be 'Mydocu~1' like in dos prompt. ", vbInformation + vbOKOnly, "Chnage ' / '"
txtImage1.Text = dlgImage1.FileName
End Sub

'image choose file location
Private Sub cmdImage2_Click()
On Error Resume Next
dlgImage2.Filter = "Image Files (*.gif)|*.gif|All Files (*.*)|*.*"
dlgImage2.ShowOpen
MsgBox "change ' \ ' to  ' / ' For example c:\1.bmp change to c:/1.bmp. And a folder must not have a space in there name. For exmaple 'My Documents' it should be 'Mydocu~1' like in dos prompt. ", vbInformation + vbOKOnly, "Chnage ' / '"
txtImage2.Text = dlgImage2.FileName

End Sub

'scroll 3d color
Private Sub cmdScroll3DC_Click()
On Error Resume Next
dlgColorAll.ShowColor
txtScrollColor1.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub

'scroll arrow color
Private Sub cmdScrollAC_Click()
On Error Resume Next
dlgColorAll.ShowColor
txtScrollColor7.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub

'scroll base color
Private Sub cmdScrollBC_Click()
On Error Resume Next
dlgColorAll.ShowColor
txtScrollColor6.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub

'scroll Dark color
Private Sub cmdScrollDC_Click()
On Error Resume Next
dlgColorAll.ShowColor
txtScrollColor5.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub

'scroll highlight color
Private Sub cmdScrollHC_Click()
On Error Resume Next
dlgColorAll.ShowColor
txtScrollColor1.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub

'scroll shadow color
Private Sub cmdScrollSC_Click()
On Error Resume Next
dlgColorAll.ShowColor
txtScrollColor2.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub

'scroll track color
Private Sub cmdScrollTC_Click()
On Error Resume Next
dlgColorAll.ShowColor
txtScrollColor4.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub

'scroll face color
Private Sub cmdScroolFC_Click()
On Error Resume Next
dlgColorAll.ShowColor
txtScrollColor8.Text = Right(StrReverse(Hex(dlgColorAll.Color)), Len(Hex(dlgColorAll.Color)) - 1) & "000000"
End Sub

'button text color
Private Sub cmdTextColor_Click()
On Error Resume Next
dlgButton3.ShowColor
txtButtonColor2.Text = Right(StrReverse(Hex(dlgButton3.Color)), Len(Hex(dlgButton3.Color)) - 1) & "000000"
End Sub

'button BG color
Private Sub cmdButtobBG_Click()
On Error Resume Next
dlgButton2.ShowColor
txtButtonColor3.Text = Right(StrReverse(Hex(dlgButton2.Color)), Len(Hex(dlgButton2.Color)) - 1) & "000000"
End Sub
'button Mouseover BG color
Private Sub cmdButtonBGMO_Click()
On Error Resume Next
dlgButton1.ShowColor
txtButtonColor1.Text = Right(StrReverse(Hex(dlgButton1.Color)), Len(Hex(dlgButton1.Color)) - 1) & "000000"
End Sub

'button fonts
Private Sub cmdFont_Click()
Dim itl As String
Dim bld As String
On Error Resume Next
dlgButton1.Flags = 3
dlgButton1.FontName = "Arial"
dlgButton1.ShowFont
If dlgButton1.FontItalic = True Then
itl = "italic"
Else
itl = ""
End If
If dlgButton1.FontBold = True Then
bld = "bold"
Else
bld = ""
End If
txtFont.Text = itl & " " & bld & " " & dlgButton1.FontSize & "px" & " " & dlgButton1.FontName
End Sub



'hover fonts
Private Sub cmdHoverFont_Click()
Dim itl As String
Dim bld As String
On Error Resume Next
dlgButton1.Flags = 3
dlgButton1.FontName = "Arial"
dlgButton1.ShowFont
If dlgButton1.FontItalic = True Then
itl = "italic"
Else
itl = ""
End If
If dlgButton1.FontBold = True Then
bld = "bold"
Else
bld = ""
End If
txtHoverFont.Text = itl & " " & bld & " " & dlgButton1.FontSize & "px" & " " & dlgButton1.FontName
End Sub

'What else?
Private Sub Form_Load()
prevwalp = True
Form2.bordercolortxt.MaxLength = 6
fraInvisible.Visible = False

End Sub

' temp file delete and close
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload Form2
Kill TempFile1
Kill TempFile2
Kill TempFile3
Kill TempFile4
End Sub



'main timer loop for button,scroll and hover
Private Sub tmrMain_Timer()
If optButton.Value = True Then
    fraButton.Visible = True
    tmrTempFile1.Enabled = True
Else
    fraButton.Visible = False
    tmrTempFile1.Enabled = False
End If
If optScroll.Value = True Then
    fraScroll.Visible = True
    tmrTempFile2.Enabled = True
Else
    fraScroll.Visible = False
    tmrTempFile2.Enabled = False
End If
If optHover.Value = True Then
    fraHover.Visible = True
    tmrTempFile3.Enabled = True
Else
    fraHover.Visible = False
    tmrTempFile3.Enabled = False
End If


End Sub

'button temp files timer loop
'write/create temp file
Private Sub tmrTempFile1_Timer()

Open TempFile1 For Output As #1
Print #1, "<html>"
Print #1, "<body>"
Print #1, "<button id=button1 style='position:absolute;left:" & txtButLeft.Text & ";top:" & txtButTop.Text & ";background:#" & txtButtonColor3.Text & ";width:" & txtButWidth.Text & ";height:" & txtButHeight.Text & ";cursor:" & txtCursor.Text & ";"
Print #1, "filter:Alpha(opacity=" & txtAlphaOp.Text & ",finishopacity=" & txtAlphaFinOp.Text & ",style=" & txtAlphaStyle.Text & ");font:" & txtFont.Text & ";color:#" & txtButtonColor2.Text & ";border:" & Form2.txtborderwidth.Text & " " & Form2.txtborderstyle.Text & " " & "#" & Form2.bordercolortxt.Text & "'"
Print #1, "onmouseover=button1.style.background='#" & txtButtonColor1.Text & "' onmouseout=button1.style.background='#" & txtButtonColor3.Text & "' onclick=window.open('" & txtUrl.Text & "')>"
Print #1, "Button Style</button>"
Print #1, "</body>"
Print #1, "</html>"
 Close #1
 
 Open TempFile2 For Output As #1
Print #1, "<html>"
Print #1, "<body>"
Print #1, "<button id=button1 style='position:absolute;left:" & txtButLeft.Text & ";top:" & txtButTop.Text & ";background:#" & txtButtonColor3.Text & ";width:" & txtButWidth.Text & ";height:" & txtButHeight.Text & ";cursor:" & txtCursor.Text & ";"
Print #1, "filter:wave(strength=" & txtWaveStr.Text & ",freq=" & txtWaveFreq.Text & ",lightstrength=" & txtWavelStr.Text & ");font:" & txtFont.Text & ";color:#" & txtButtonColor2.Text & ";border:" & Form2.txtborderwidth.Text & " " & Form2.txtborderstyle.Text & " " & "#" & Form2.bordercolortxt.Text & "'"
Print #1, "onmouseover=button1.style.background='#" & txtButtonColor1.Text & "' onmouseout=button1.style.background='#" & txtButtonColor3.Text & "' onclick=window.open('" & txtUrl.Text & "')>"
Print #1, "Button Style</button>"
Print #1, "</body>"
Print #1, "</html>"
 Close #1
 
 
 Open TempFile3 For Output As #1
Print #1, "<html>"
Print #1, "<body>"
Print #1, "<button id=button1 style='position:absolute;left:" & txtButLeft.Text & ";top:" & txtButTop.Text & ";background:#" & txtButtonColor3.Text & ";width:" & txtButWidth.Text & ";height:" & txtButHeight.Text & ";cursor:" & txtCursor.Text & ";"
Print #1, "filter:blur(direction=" & txtBlurDir.Text & ",strength=" & txtBlurstr.Text & ");font:" & txtFont.Text & ";color:#" & txtButtonColor2.Text & ";border:" & Form2.txtborderwidth.Text & " " & Form2.txtborderstyle.Text & " " & "#" & Form2.bordercolortxt.Text & "'"
Print #1, "onmouseover=button1.style.background='#" & txtButtonColor1.Text & "' onmouseout=button1.style.background='#" & txtButtonColor3.Text & "' onclick=window.open('" & txtUrl.Text & "')>"
Print #1, "Button Style</button>"
Print #1, "</body>"
Print #1, "</html>"
 Close #1
 
  Open TempFile4 For Output As #1
Print #1, "<html>"
Print #1, "<body>"
Print #1, "<button id=button1 style='position:absolute;left:" & txtButLeft.Text & ";top:" & txtButTop.Text & ";background:#" & txtButtonColor3.Text & ";width:" & txtButWidth.Text & ";height:" & txtButHeight.Text & ";cursor:" & txtCursor.Text
Print #1, ";font:" & txtFont.Text & ";color:#" & txtButtonColor2.Text & ";border:" & Form2.txtborderwidth.Text & " " & Form2.txtborderstyle.Text & " " & "#" & Form2.bordercolortxt.Text & "'"
Print #1, "onmouseover=button1.style.background='#" & txtButtonColor1.Text & "' onmouseout=button1.style.background='#" & txtButtonColor3.Text & "' onclick=window.open('" & txtUrl.Text & "')>"
Print #1, "Button Style</button>"
Print #1, "</body>"
Print #1, "</html>"
 Close #1
 
 On Error Resume Next
  Open TempFile5 For Output As #1
Print #1, "<html>"
Print #1, "<body >"
Print #1, "<img  name=Im src='" & txtImage1.Text & "' width=" & txtImageWidth.Text & " height=" & txtImageHeight.Text
Print #1, " style='cursor:" & txtImageCur.Text & "' onClick=window.open('" & txtUrl2.Text & "') "
Print #1, " onmouseover=Im.src='" & txtImage2.Text & "' onmouseout=Im.src='" & txtImage1.Text & "' >"
Print #1, "</body>"
Print #1, "</html>"
 Close #1
   
  
If prevwalp = True Then
  WebBrowser1.Navigate TempFile1
  txtGenCode2.FileName = TempFile1
  txtGenCode1.Text = txtGenCode2.Text
       prevwalp = False
    Exit Sub
 End If
   
 End Sub

'button timer loop; option: alpha,wave,normal,image and blur.
Private Sub tmrButton_Timer()
If optAlpha.Value = True Then
    fraButtonAlpha.Visible = True
Else
    fraButtonAlpha.Visible = False
End If
If optWave.Value = True Then
    fraButtonWave.Visible = True
Else
    fraButtonWave.Visible = False
End If
If optBlur.Value = True Then
    fraButtonBlur.Visible = True
Else
    fraButtonBlur.Visible = False
End If
If optNormal.Value = True Then
    fraButtonNormal.Visible = True
Else
    fraButtonNormal.Visible = False
End If
If optImage.Value = True Then
    fraImage.Visible = True

Else
    fraImage.Visible = False
End If
End Sub

'preview from internet explorer
Private Sub cmdPreviewIE_Click()
If optButton.Value = True Then
    If optAlpha.Value = True Then
        Shell "explorer.exe " & TempFile1, 0
    End If
    If optBlur.Value = True Then
        Shell "explorer.exe " & TempFile3, 0
    End If
    If optWave.Value = True Then
        Shell "explorer.exe " & TempFile2, 0
    End If
    If optNormal.Value = True Then
        Shell "explorer.exe " & TempFile4, 0
    End If
    If optImage.Value = True Then
        Shell "explorer.exe " & TempFile5, 0
    End If
End If
If optHover.Value = True Then
Shell "explorer.exe " & TempFile3, 0
End If
If optScroll.Value = True Then
Shell "explorer.exe " & TempFile2, 0
End If
End Sub

'generated code
Private Sub cmdGenerateCode_Click()
If optButton.Value = True Then
    If optAlpha.Value = True Then
        txtGenCode2.FileName = TempFile1
        txtGenCode1.Text = txtGenCode2.Text
    End If
    If optWave.Value = True Then
        txtGenCode2.FileName = TempFile2
        txtGenCode1.Text = txtGenCode2.Text
    End If
    If optBlur.Value = True Then
        txtGenCode2.FileName = TempFile3
        txtGenCode1.Text = txtGenCode2.Text
    End If
    If optNormal.Value = True Then
        txtGenCode2.FileName = TempFile4
        txtGenCode1.Text = txtGenCode2.Text
    End If
    If optImage.Value = True Then
        txtGenCode2.FileName = TempFile5
        txtGenCode1.Text = txtGenCode2.Text
    End If
End If
If optScroll.Value = True Then
        txtGenCode2.FileName = TempFile2
        txtGenCode1.Text = txtGenCode2.Text
End If
If optHover.Value = True Then
        txtGenCode2.FileName = TempFile3
        txtGenCode1.Text = txtGenCode2.Text
End If
End Sub

'preview from form

Private Sub cmdPreviewhtml_Click()
If optButton.Value = True Then
    If optAlpha.Value = True Then
        WebBrowser1.Navigate TempFile1
    End If
    If optWave.Value = True Then
        WebBrowser1.Navigate TempFile2
    End If
    If optBlur.Value = True Then
        WebBrowser1.Navigate TempFile3
    End If
    If optNormal.Value = True Then
        WebBrowser1.Navigate TempFile4
    End If
    If optImage.Value = True Then
        WebBrowser1.Navigate TempFile5
    End If
End If
If optHover.Value = True Then
     WebBrowser1.Navigate TempFile3
End If
If optScroll.Value = True Then
     WebBrowser1.Navigate TempFile2
End If
End Sub


'scroll temp file timer loop
'write/create temp file
Private Sub tmrTempFile2_Timer()

Open TempFile2 For Output As #1
Print #1, "<html>"
Print #1, "<style>"
Print #1, "body"
Print #1, "{SCROLLBAR-HIGHLIGHT-COLOR:#" & txtScrollColor1.Text & ";"
Print #1, "SCROLLBAR-SHADOW-COLOR:#" & txtScrollColor2.Text & ";"
Print #1, "SCROLLBAR-3DLIGHT-COLOR#:"; txtScrollColor3.Text & ";"
Print #1, "SCROLLBAR-TRACK-COLOR#:" & txtScrollColor4.Text & ";"
Print #1, "SCROLLBAR-DARKSHADOW-COLOR:#" & txtScrollColor5.Text & ";"
Print #1, "SCROLLBAR-BASE-COLOR:#" & txtScrollColor6.Text & ";"
Print #1, "SCROLLBAR-ARROW-COLOR:#" & txtScrollColor7.Text & ";"
Print #1, "SCROLLBAR-FACE-COLOR:#" & txtScrollColor8.Text & ";}</style>"
Print #1, "<body>"
Print #1, "Spaces below, so you can preview it well. "
Print #1, "<br><br><br><br><br><br>"
Print #1, "<br><br><br><br><br><br>"
Print #1, "<br><br><br><br><br><br>"
Print #1, "<br><br><br><br><br><br>"
Print #1, "</body>"
Print #1, "</html>"
 Close #1

End Sub

'hover temp file timer loop
'write/create temp file
Private Sub tmrTempFile3_Timer()
Open TempFile3 For Output As #1
Print #1, "<html><head>"
Print #1, "<style type=text/css>"
Print #1, "a {border:" & Form2.txtborderwidth.Text & " " & Form2.txtborderstyle.Text & " " & "#" & Form2.bordercolortxt.Text & "; background:#" & txtABC.Text & ";color:#" & txtAC.Text & ";font:" & txtAMainFont.Text & "};"
Print #1, "a:hover {background:#" & txtHBC.Text & ";color:#" & txtHC.Text & ";font:" & txtHoverFont.Text & "};"
Print #1, "a:link {color:" & txtLC.Text & "};"
Print #1, "a:ative {color:" & txtAAC.Text & "};"
Print #1, "a:visited {color:" & txtVC.Text & "};"
Print #1, "</style></head>"
Print #1, "<body>"
Print #1, "<a href=c:\>Press Preview</a>"
Print #1, "</body>"
Print #1, "</html>"
 Close #1
 
 
End Sub
