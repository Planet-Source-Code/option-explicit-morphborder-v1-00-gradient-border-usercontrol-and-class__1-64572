VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "cMorphBorder Class Demo - Matthew R. Usner"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   617
   StartUpPosition =   2  'CenterScreen
   Begin prjMorphBorder.MorphListBox MorphListBox1 
      Height          =   2895
      Left            =   4800
      TabIndex        =   2
      Top             =   480
      Width           =   3855
      _ExtentX        =   6800
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   600
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":CAD1
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   3480
      Width           =   8055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cmb As New cMorphBorder

Private Sub Form_Load()

   Dim i As Long

   With Me
      cmb.DisplayBorder .hdc, .ScaleWidth, .ScaleHeight, 21, &H800000, &HFFC0C0, False
   End With

   With Picture1
      cmb.DisplayBorder .hdc, .ScaleWidth, .ScaleHeight, 18, &H4000&, &HC0FFC0, True
   End With

   MorphListBox1.RedrawFlag = False
   For i = 1 To 30
      MorphListBox1.AddItem String(20, Chr(i + 32))
   Next
   MorphListBox1.RedrawFlag = True

   With MorphListBox1
      cmb.DisplayBorder .hdc, .ScaleWidth, .ScaleHeight, 10, vbBlack, &HE0E0E0, True
   End With

End Sub

Private Sub Form_Resize()

   Me.Cls

   With Me
      cmb.DisplayBorder .hdc, .ScaleWidth, .ScaleHeight, 20, &H800000, &HFFC0C0, False
   End With

End Sub
