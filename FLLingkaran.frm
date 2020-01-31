VERSION 5.00
Begin VB.Form FLLingkaran 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Luas Lingkaran"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   Icon            =   "FLLingkaran.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CLLingkaranTutup 
      BackColor       =   &H000000FF&
      Caption         =   "Tutup"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox TLLJari 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton CLLingkarangHitung 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hitung Luas"
      Height          =   495
      Left            =   2280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox TLLingkaran 
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan Jari-Jari (r)"
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
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   3960
      Picture         =   "FLLingkaran.frx":048A
      Top             =   0
      Width           =   4500
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "3.14 *  r *  r"
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
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Luas Lingkaran"
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
      Top             =   1920
      Width           =   1935
   End
End
Attribute VB_Name = "FLLingkaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TKLJari_Change()

End Sub

Private Sub CLLingkarangHitung_Click()
If (Trim$(TLLJari.Text) = "") Then
MsgBox "Jari - Jari Tidak Boleh Kosong"
Else
If IsNumeric(TLLJari.Text) Then
r = Val(TLLJari.Text)
TLLingkaran = 3.14 * r * r
Else
MsgBox "Bukan angka"
TLLJari.Text = ""
End If
End If
End Sub

Private Sub CLLingkaranTutup_Click()
FLLingkaran.Visible = False
Unload Me
End Sub
