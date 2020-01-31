VERSION 5.00
Begin VB.Form FKTrapesium 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keliling Trapesium"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   Icon            =   "FKTrapesium.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TKTSisi4 
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox TKTSisi3 
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox TKTrapesium 
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton CKTHitung 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hitung Keliling"
      Height          =   495
      Left            =   2280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox TKTSisi1 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CKTTutup 
      BackColor       =   &H000000FF&
      Caption         =   "Tutup"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox TKTSisi2 
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan d"
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
      Left            =   360
      TabIndex        =   12
      Top             =   2040
      Width           =   1815
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
      Left            =   360
      TabIndex        =   10
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Keliling Trapesium"
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
      Left            =   360
      TabIndex        =   8
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "a + b + c + d"
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
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   3960
      Picture         =   "FKTrapesium.frx":048A
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
      Left            =   360
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
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "FKTrapesium"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CKJGHitung_Click()

End Sub

Private Sub CKTHitung_Click()
If (Trim$(TKTSisi1.Text) = "") Then
MsgBox "Sisi a Tidak Boleh Kosong"
ElseIf (Trim$(TKTSisi2.Text) = "") Then
MsgBox "Sisi b Tidak Boleh Kosong"
ElseIf (Trim$(TKTSisi3.Text) = "") Then
MsgBox "Sisi c Tidak Boleh Kosong"
ElseIf (Trim$(TKTSisi4.Text) = "") Then
MsgBox "Sisi cdTidak Boleh Kosong"
Else
If IsNumeric(TKTSisi1.Text) And IsNumeric(TKTSisi2.Text) And IsNumeric(TKTSisi4.Text) And IsNumeric(TKTSisi4.Text) Then
a = Val(TKTSisi1.Text)
b = Val(TKTSisi2.Text)
c = Val(TKTSisi3.Text)
d = Val(TKTSisi4.Text)
TKTrapesium.Text = a + b + c + d
Else
MsgBox "Bukan Angka"
TKTSisi1.Text = ""
TKTSisi2.Text = ""
TKTSisi3.Text = ""
TKTSisi4.Text = ""
End If
End If
End Sub

Private Sub CKTTutup_Click()
FKTrapesium.Visible = False
Unload Me
End Sub

