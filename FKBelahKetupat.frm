VERSION 5.00
Begin VB.Form FKBelahKetupat 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keliling Belah Ketupat"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   Icon            =   "FKBelahKetupat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TKBelahKetupat 
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton CKBKHitung 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hitung Keliling"
      Height          =   495
      Left            =   2520
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox TKBKSisi 
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton CKBKTutup 
      BackColor       =   &H000000FF&
      Caption         =   "Tutup"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Keliling Belah Ketupat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "4 * Sisi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   4200
      Picture         =   "FKBelahKetupat.frx":048A
      Top             =   -120
      Width           =   4500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan Sisi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "FKBelahKetupat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CKBKHitung_Click()
If (TKBKSisi.Text = "") Then
    MsgBox "Sisi Tidak Boleh Kosong"
Else
    If IsNumeric(TKBKSisi.Text) Then
        Dim hitung As New BangunDatar
        TKBelahKetupat.Text = hitung.kBelahK(Val(TKBKSisi))
    Else
        MsgBox "bukan angka"
        TKBKSisi.Text = ""
End If
End If
End Sub

Private Sub CKBKTutup_Click()
FKBelahKetupat.Visible = False
Unload Me
End Sub

