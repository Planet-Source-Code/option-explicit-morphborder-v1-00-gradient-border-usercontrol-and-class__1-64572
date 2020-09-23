VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "MorphBorder Demo - Matthew R. Usner"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   StartUpPosition =   2  'CenterScreen
   Begin prjMorphBorder.MorphListBox MorphListBox1 
      Height          =   2895
      Left            =   4920
      TabIndex        =   5
      Top             =   480
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5106
      BorderWidth     =   10
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjMorphBorder.MorphBorder MorphBorder3 
      Height          =   600
      Left            =   6480
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1058
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   600
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin prjMorphBorder.MorphBorder MorphBorder1 
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1058
      BorderWidth     =   20
      Color1          =   8388608
      Color2          =   16761024
      MiddleOut       =   0   'False
   End
   Begin prjMorphBorder.MorphBorder MorphBorder2 
      Height          =   600
      Left            =   1920
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1058
      BorderWidth     =   16
      Color1          =   16384
      Color2          =   8454016
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":CAD1
      Height          =   1215
      Left            =   600
      TabIndex        =   3
      Top             =   3480
      Width           =   8415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

   Dim i As Long

   MorphBorder2.TargetControlName = "Picture1" ' cat picbox

   MorphListBox1.RedrawFlag = False
   For i = 1 To 30
      MorphListBox1.AddItem String(20, Chr(i + 32))
   Next
   MorphListBox1.RedrawFlag = True

   MorphBorder3.TargetControlName = "MorphListBox1"

End Sub

Private Sub Form_Resize()
   Me.Cls
   MorphBorder1.DisplayBorder ' redraw the form's border.
End Sub
