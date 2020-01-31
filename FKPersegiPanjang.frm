VERSION 5.00
Begin VB.Form FKPersegiPanjang 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keliling Persegi Panjang"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   Icon            =   "FKPersegiPanjang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TKPersegiPanjang 
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton CKPPHitung 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hitung Keliling"
      Height          =   495
      Left            =   2280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox TKPPPanjang 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton CKPPTutup 
      BackColor       =   &H000000FF&
      Caption         =   "Tutup"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox TKPPLebar 
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Keliling Persegi Panjang"
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
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2 * (panjang + lebar)"
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
      Picture         =   "FKPersegiPanjang.frx":048A
      Top             =   0
      Width           =   4500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan Panjang"
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
      Caption         =   "Masukkan Lebar"
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
Attribute VB_Name = "FKPersegiPanjang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CKJGHitung_Click()

End Sub

Private Sub CKPPHitung_Click()
If (Trim$(TKPPPanjang.Text) = "") Then
MsgBox "Panjang Tidak Boleh Kosong"
ElseIf (Trim$(TKPPLebar.Text) = "") Then
MsgBox "Lebar Tidak Boleh Kosong"
Else
If IsNumeric(TKPPPanjang.Text) And IsNumeric(TKPPLebar) Then
p = Val(TKPPPanjang.Text)
l = Val(TKPPLebar.Text)
TKPersegiPanjang.Text = 2 * (p + l)
Else
MsgBox "Bukan Angka"
TKPPPanjang.Text = ""
TKPPLebar.Text = ""
End If
End If
End Sub

Private Sub CKPPTutup_Click()
FKPersegiPanjang.Visible = False
Unload Me
End Sub

