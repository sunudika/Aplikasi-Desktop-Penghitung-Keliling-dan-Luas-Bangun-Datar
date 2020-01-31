VERSION 5.00
Begin VB.Form FLJajarGenjang 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Luas Jajar Genjang"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   Icon            =   "FLJajarGenjang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TLJajarGenjang 
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton CLJGHitung 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hitung Luas"
      Height          =   495
      Left            =   2280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox TLJGa 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton CLJGTutup 
      BackColor       =   &H000000FF&
      Caption         =   "Tutup"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox TLJGt 
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Luas Jajar Genjang"
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
      TabIndex        =   8
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "a * t"
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
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   3960
      Picture         =   "FLJajarGenjang.frx":048A
      Top             =   0
      Width           =   4500
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
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan t"
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
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "FLJajarGenjang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TKJGSisi1_Change()

End Sub

Private Sub CLJGHitung_Click()
If (Trim$(TLJGa.Text) = "") Then
MsgBox "Alas Tidak Boleh Kosong"
ElseIf (Trim$(TLJGt.Text) = "") Then
MsgBox "Tinggi Tidak Boleh Kosong"
Else
If IsNumeric(TLJGa.Text) And IsNumeric(TLJGt.Text) Then
a = Val(TLJGa.Text)
t = Val(TLJGt.Text)
TLJajarGenjang = a * t
Else
MsgBox "Bukan angka"
TLJGa.Text = ""
TLJGt.Text = ""
End If
End If
End Sub

Private Sub CLJGTutup_Click()
FLJajarGenjang.Visible = False
Unload Me
End Sub
