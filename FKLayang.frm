VERSION 5.00
Begin VB.Form FKLayang 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keliling Layang - Layang"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   Icon            =   "FKLayang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TKLayang 
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton CKLHitung 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hitung Keliling"
      Height          =   495
      Left            =   2400
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox TKLSisi1 
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox TKLSisi2 
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton CKLayangTutup 
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
      Caption         =   "Keliling Jajar Genjang"
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
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2 * (sisi a + sisi b)"
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
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan Sisi a"
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
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan Sisi b"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   3960
      Picture         =   "FKLayang.frx":048A
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "FKLayang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CKLayangTutup_Click()
FKLayang.Visible = False
Unload Me
End Sub

Private Sub CKLHitung_Click()
If (Trim$(TKLSisi1.Text) = "") Then
MsgBox "Sisi a Tidak Boleh Kosong"
ElseIf (Trim$(TKLSisi2.Text) = "") Then
MsgBox "Sisi b Tidak Boleh Kosong"
Else
If IsNumeric(TKLSisi1.Text) And IsNumeric(TKLSisi2.Text) Then
a = Val(TKLSisi1.Text)
b = Val(TKLSisi2.Text)
TKLayang.Text = 2 * (a + b)
Else
MsgBox "Bukan Angka"
TKLSisi1.Text = ""
TKLSisi2.Text = ""
End If
End If
End Sub

Private Sub TKJGSisi1_Change()

End Sub
