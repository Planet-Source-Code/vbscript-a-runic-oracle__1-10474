VERSION 5.00
Begin VB.Form frmRune 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rune Information"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmRune.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   3975
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox txtAnalysis 
      Height          =   1215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3240
      Width           =   3735
   End
   Begin VB.TextBox txtMagic 
      Alignment       =   2  'Center
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Width           =   3735
   End
   Begin VB.TextBox txtDivination 
      Alignment       =   2  'Center
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox txtBasic 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtPhonetic 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblAnalysis 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Analysis"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label lblMagical 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Magical Uses"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label lblDivinatory 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Divinatory Meaning"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label lblBasic 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Basic Meaning"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblPhonetic 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Phonetic"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmRune"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Form_Load()
    Dim I As Integer, Y As Integer

    Me.AutoRedraw = True
    Me.DrawStyle = 6
    Me.DrawMode = 13
    Me.DrawWidth = 13
    Me.ScaleMode = 3
    Me.ScaleHeight = 256

    For I = 0 To 510
        Me.Line (0, Y)-(Me.Width, Y + 1), RGB(0, 0, I), BF
        Y = Y + 1
    Next I
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub
