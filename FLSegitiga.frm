VERSION 5.00
Begin VB.Form FLSegitiga 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Luas Segitiga"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   Icon            =   "FLSegitiga.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TLSt 
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton CLSTutup 
      BackColor       =   &H000000FF&
      Caption         =   "Tutup"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox TLSa 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton CLSHitung 
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
   Begin VB.TextBox TLSegitiga 
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan b"
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
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan a"
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
      TabIndex        =   7
      Top             =   480
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   3960
      Picture         =   "FLSegitiga.frx":048A
      Top             =   0
      Width           =   4500
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1/2 * a * t"
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
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Luas Segitiga"
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
      TabIndex        =   5
      Top             =   2640
      Width           =   2175
   End
End
Attribute VB_Name = "FLSegitiga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TKSSisi1_Change()

End Sub

Private Sub CLSHitung_Click()
If (Trim$(TLSa.Text) = "") Then
MsgBox "Alas Tidak Boleh Kosong"
ElseIf (Trim$(TLSt.Text) = "") Then
MsgBox "Tinggi Tidak Boleh Kosong"
Else
If IsNumeric(TLSa.Text) And IsNumeric(TLSt.Text) Then
a = Val(TLSa.Text)
t = Val(TLSt.Text)
TLSegitiga = 1 / 2 * a * t
Else
MsgBox "Bukan angka"
TLSa.Text = ""
TLSt.Text = ""
End If
End If
End Sub

Private Sub CLSTutup_Click()
FLSegitiga.Visible = False
Unload Me
End Sub
