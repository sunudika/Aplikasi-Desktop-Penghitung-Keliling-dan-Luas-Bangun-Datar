VERSION 5.00
Begin VB.Form FKPersegi 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keliling Persegi"
   ClientHeight    =   4410
   ClientLeft      =   4920
   ClientTop       =   3840
   ClientWidth     =   8760
   Icon            =   "FKPersegi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8760
   Begin VB.CommandButton CKPersegiTutup 
      BackColor       =   &H000000FF&
      Caption         =   "Tutup"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox TKPSisi 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton CKPersegi 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hitung Keliling"
      Height          =   495
      Left            =   2280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox TKPersegi 
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
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
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   3960
      Picture         =   "FKPersegi.frx":048A
      Top             =   0
      Width           =   4500
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
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Keliling Persegi"
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
Attribute VB_Name = "FKPersegi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CKPersegi_Click()
If (Trim$(TKPSisi.Text) = "") Then
MsgBox "Sisi Tidak Boleh Kosong"
Else
If IsNumeric(TKPSisi.Text) Then
TKPersegi = 4 * TKPSisi
Else
MsgBox "Bukan Angka"
TKPSisi.Text = ""
End If
End If
End Sub

Private Sub CKPersegiTutup_Click()
FKPersegi.Visible = False
Unload Me
End Sub

Private Sub TKBKSisi_Change()

End Sub

