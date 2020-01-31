VERSION 5.00
Begin VB.Form FLLayang 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Luas Layang - Layang"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   Icon            =   "FLLayang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CLLayangTutup 
      BackColor       =   &H000000FF&
      Caption         =   "Tutup"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox TLLd2 
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox TLLd1 
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton CLLHitung 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hitung Luas"
      Height          =   495
      Left            =   2400
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox TLLayang 
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   3960
      Picture         =   "FLLayang.frx":048A
      Top             =   0
      Width           =   4500
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
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
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
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   1815
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
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
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
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   1935
   End
End
Attribute VB_Name = "FLLayang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TKLSisi1_Change()

End Sub

Private Sub CLLayangTutup_Click()
FLLayang.Visible = False
Unload Me
End Sub

Private Sub CLLHitung_Click()
If (Trim$(TLLd1.Text) = "") Then
MsgBox "Diagonal 1 Tidak Boleh Kosong"
ElseIf (Trim$(TLLd2.Text) = "") Then
MsgBox "Diagonal 2 Tidak Boleh Kosong"
Else
If IsNumeric(TLLd1.Text) And IsNumeric(TLLd2.Text) Then
d1 = Val(TLLd1.Text)
d2 = Val(TLLd2.Text)
TLLayang = 1 / 2 * d1 * d2
Else
MsgBox "Bukan angka"
TLLd1.Text = ""
TLLd2.Text = ""
End If
End If
End Sub
