VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "border edit"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox bordercolortxt 
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Text            =   "FF6600"
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton back2main 
      Caption         =   "Done"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame7 
      Caption         =   "Border-Color"
      Height          =   1575
      Left            =   3720
      TabIndex        =   2
      Top             =   240
      Width           =   1695
      Begin VB.CommandButton bordercolor 
         Caption         =   "Select"
         Height          =   615
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Border-Style"
      Height          =   1575
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   1695
      Begin VB.TextBox txtborderstyle 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Text            =   "double"
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Style name ( none, double,inset,outset, solid,groove, ridge)"
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog4 
      Left            =   3240
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame8 
      Caption         =   "Border-Width"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1695
      Begin VB.TextBox txtborderwidth 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Text            =   "thin"
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Size(px,in,cm,pt); thin,medium,thick.  example:12pt or thin"
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'border editor form

Private Sub back2main_Click()
Form2.Visible = False
Form1.Enabled = True
Form1.Show
End Sub


Private Sub bordercolor_Click()
On Error Resume Next
CommonDialog4.ShowColor
bordercolortxt.Text = Right(StrReverse(Hex(CommonDialog4.Color)), Len(Hex(CommonDialog4.Color)) - 1) & "000000"
bordercolor.BackColor = CommonDialog4.Color
End Sub


Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
End Sub
