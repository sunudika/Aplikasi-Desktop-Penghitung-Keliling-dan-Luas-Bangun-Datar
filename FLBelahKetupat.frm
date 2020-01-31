VERSION 5.00
Begin VB.Form FLBelahKetupat 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Luas Belah Ketupat"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   Icon            =   "FLBelahKetupat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TLBKd2 
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton CLBKTutup 
      BackColor       =   &H000000FF&
      Caption         =   "Tutup"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox TLBKd1 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton CLBKHitung 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hitung Luas"
      Height          =   495
      Left            =   2280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox TLBelahKetupat 
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan d2"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan d1"
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
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   3960
      Picture         =   "FLBelahKetupat.frx":048A
      Top             =   240
      Width           =   4500
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1/2 * d1 * d2"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Luas Belah Ketupat"
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
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
   End
End
Attribute VB_Name = "FLBelahKetupat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CLBKHitung_Click()
If (Trim$(TLBKd1.Text) = "") Then
MsgBox "Diagonal 1 Tidak Boleh Kosong"
ElseIf (Trim$(TLBKd2.Text) = "") Then
MsgBox "Diagonal 2 Tidak Boleh Kosong"
Else
If IsNumeric(TLBKd1.Text) And IsNumeric(TLBKd2.Text) Then
d1 = Val(TLBKd1.Text)
d2 = Val(TLBKd2.Text)
TLBelahKetupat = 1 / 2 * d1 * d2
Else
MsgBox "Bukan angka"
TLBKd1.Text = ""
TLBKd2.Text = ""
End If
End If
End Sub

Private Sub CLBKTutup_Click()
FLBelahKetupat.Visible = False
Unload Me
End Sub

