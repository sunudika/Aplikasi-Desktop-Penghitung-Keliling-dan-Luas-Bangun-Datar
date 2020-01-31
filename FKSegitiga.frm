VERSION 5.00
Begin VB.Form FKSegitiga 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keliling Segitiga"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   Icon            =   "FKSegitiga.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TKSSisi3 
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox TKSegitiga 
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
   Begin VB.TextBox TKSSisi1 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CKSTutup 
      BackColor       =   &H000000FF&
      Caption         =   "Tutup"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox TKSSisi2 
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan c"
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
      TabIndex        =   10
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Keliling Segitiga"
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
      Caption         =   "a + b + c"
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
      Picture         =   "FKSegitiga.frx":048A
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
      Top             =   240
      Width           =   1815
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
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "FKSegitiga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CKJGHitung_Click()

End Sub

Private Sub CKPPHitung_Click()
If (Trim$(TKSSisi1.Text) = "") Then
MsgBox "Sisi a Tidak Boleh Kosong"
ElseIf (Trim$(TKSSisi2.Text) = "") Then
MsgBox "Sisi b Tidak Boleh Kosong"
ElseIf (Trim$(TKSSisi3.Text) = "") Then
MsgBox "Sisi c Tidak Boleh Kosong"
Else
If IsNumeric(TKSSisi1.Text) And IsNumeric(TKSSisi2.Text) And IsNumeric(TKSSisi3.Text) Then
a = Val(TKSSisi1.Text)
b = Val(TKSSisi2.Text)
c = Val(TKSSisi3.Text)
TKSegitiga.Text = a + b + c
Else
MsgBox "Bukan Angka"
TKSSisi1.Text = ""
TKSSisi2.Text = ""
TKSSisi3.Text = ""
End If
End If
End Sub

Private Sub CKPPTutup_Click()

End Sub

Private Sub CKSTutup_Click()
FKSegitiga.Visible = False
Unload Me
End Sub
