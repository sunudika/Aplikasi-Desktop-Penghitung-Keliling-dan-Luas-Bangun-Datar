VERSION 5.00
Begin VB.Form FLPersegi 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Luas Persegi"
   ClientHeight    =   4410
   ClientLeft      =   7755
   ClientTop       =   9510
   ClientWidth     =   8760
   Icon            =   "FLPersegi.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8760
   Begin VB.TextBox TLPersegi 
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton CLPersegi 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hitung Luas"
      Height          =   495
      Left            =   2280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox TLPSisi 
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton CLPersegiTutup 
      BackColor       =   &H000000FF&
      Caption         =   "Tutup"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Luas Persegi"
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
      TabIndex        =   6
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sisi * Sisi"
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
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   3960
      Picture         =   "FLPersegi.frx":048A
      Top             =   0
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
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "FLPersegi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TKPSisi_Change()

End Sub

Private Sub CLPersegi_Click()
If (Trim$(TLPSisi.Text) = "") Then
MsgBox "Sisi Tidak Boleh Kosong"
Else
If IsNumeric(TLPSisi.Text) Then
s = Val(TLPSisi.Text)
TLPersegi = s * s
Else
MsgBox "Bukan angka"
TLPSisi.Text = ""
End If
End If
End Sub

Private Sub CLPersegiTutup_Click()
FLPersegi.Visible = False
Unload Me
End Sub
