VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Runes Generator"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdRunes 
      Caption         =   "&Get Runes"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblLableFuture 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Future"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblLablePresent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Present"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblLablePast 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Past"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblFuture 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "future"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblPresent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "present"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblPast 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "past"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   735
   End
   Begin VB.Image imgFuture 
      Height          =   1095
      Left            =   3000
      Top             =   360
      Width           =   735
   End
   Begin VB.Image imgPresent 
      Height          =   1095
      Left            =   1680
      Top             =   360
      Width           =   735
   End
   Begin VB.Image imgPast 
      Height          =   1095
      Left            =   360
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim db As Database
    Dim rs As Recordset
    Dim Rune1 As Integer, Rune2 As Integer, Rune3 As Integer

Private Sub Form_Load()
    Dim I As Integer, Y As Integer

    Set db = OpenDatabase(App.Path & "\runes.mdb")
    Set rs = db.OpenRecordset("tblRunes")
    
    cmdClear_Click
    
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

Private Sub cmdRunes_Click()
    Randomize
    Rune1 = Int(Rnd * rs.RecordCount) + 1
    Rune2 = Int(Rnd * rs.RecordCount) + 1
    Rune3 = Int(Rnd * rs.RecordCount) + 1
    
    rs.MoveFirst
    Do Until rs.Fields("ID") = Rune1
        imgPast.Picture = LoadResPicture(rs.Fields("ID"), 0)
        lblPast.Caption = rs.Fields("Name")
        lblLablePast.Caption = "Past"
        rs.MoveNext
    Loop
    
    rs.MoveFirst
    Do Until rs.Fields("ID") = Rune2
        imgPresent.Picture = LoadResPicture(rs.Fields("ID"), 0)
        lblPresent.Caption = rs.Fields("Name")
        lblLablePresent.Caption = "Present"
        rs.MoveNext
    Loop

    rs.MoveFirst
    Do Until rs.Fields("ID") = Rune3
        imgFuture.Picture = LoadResPicture(rs.Fields("ID"), 0)
        lblFuture.Caption = rs.Fields("Name")
        lblLableFuture.Caption = "Future"
        rs.MoveNext
    Loop
End Sub

Private Sub cmdClear_Click()
    lblLablePast.Caption = ""
    lblLablePresent.Caption = ""
    lblLableFuture.Caption = ""
    imgPast.Picture = LoadPicture()
    imgPresent.Picture = LoadPicture()
    imgFuture.Picture = LoadPicture()
    lblPast.Caption = ""
    lblPresent.Caption = ""
    lblFuture.Caption = ""
End Sub

Private Sub cmdClose_Click()
    End
End Sub

Private Sub imgPast_Click()
    rs.MoveFirst
    
    Do Until rs.Fields("ID") = Rune1 - 1
        rs.MoveNext
    Loop
    
    frmRune.txtName = rs.Fields("Name")
    frmRune.txtPhonetic = rs.Fields("Phonetic")
    frmRune.txtBasic = rs.Fields("BasicMeaning")
    frmRune.txtDivination = rs.Fields("DivinatoryMeaning")
    frmRune.txtMagic = rs.Fields("MagicalUses")
    frmRune.txtAnalysis = rs.Fields("Analysis")
    frmRune.Show

End Sub

Private Sub imgPresent_Click()
    rs.MoveFirst
    
    Do Until rs.Fields("ID") = Rune2 - 1
        rs.MoveNext
    Loop
    
    frmRune.txtName = rs.Fields("Name")
    frmRune.txtPhonetic = rs.Fields("Phonetic")
    frmRune.txtBasic = rs.Fields("BasicMeaning")
    frmRune.txtDivination = rs.Fields("DivinatoryMeaning")
    frmRune.txtMagic = rs.Fields("MagicalUses")
    frmRune.txtAnalysis = rs.Fields("Analysis")
    frmRune.Show
End Sub

Private Sub imgFuture_Click()
    rs.MoveFirst
    
    Do Until rs.Fields("ID") = Rune3 - 1
        rs.MoveNext
    Loop
    
    frmRune.txtName = rs.Fields("Name")
    frmRune.txtPhonetic = rs.Fields("Phonetic")
    frmRune.txtBasic = rs.Fields("BasicMeaning")
    frmRune.txtDivination = rs.Fields("DivinatoryMeaning")
    frmRune.txtMagic = rs.Fields("MagicalUses")
    frmRune.txtAnalysis = rs.Fields("Analysis")
    frmRune.Show
End Sub

