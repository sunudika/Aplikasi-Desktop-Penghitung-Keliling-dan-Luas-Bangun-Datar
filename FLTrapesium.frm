VERSION 5.00
Begin VB.Form FLTrapesium 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Luas Trapesium"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   Icon            =   "FLTrapesium.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TLTt 
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton CKTTutup 
      BackColor       =   &H000000FF&
      Caption         =   "Tutup"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox TLTa 
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CLTHitung 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hitung Luas"
      Height          =   495
      Left            =   2280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox TLTrapesium 
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox TLTc 
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
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
      Left            =   360
      TabIndex        =   10
      Top             =   840
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
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   3960
      Picture         =   "FLTrapesium.frx":048A
      Top             =   0
      Width           =   4500
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1/2 * (a+c) * t"
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
      TabIndex        =   8
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Luas Trapesium"
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
      TabIndex        =   7
      Top             =   3360
      Width           =   2175
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
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "FLTrapesium"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TKTSisi1_Change()

End Sub

Private Sub CKTTutup_Click()
FLTrapesium.Visible = False
Unload Me
End Sub

Private Sub CLTHitung_Click()
If (Trim$(TLTa.Text) = "") Then
MsgBox "Alas Tidak Boleh Kosong"
ElseIf (Trim$(TLTt.Text) = "") Then
MsgBox "Tinggi Tidak Boleh Kosong"
ElseIf (Trim$(TLTc.Text) = "") Then
MsgBox "Sisi c Tidak Boleh Kosong"
Else
If IsNumeric(TLTa.Text) And IsNumeric(TLTt.Text) And IsNumeric(TLTc.Text) Then
a = Val(TLTa.Text)
t = Val(TLTt.Text)
c = Val(TLTc.Text)
TLTrapesium = 1 / 2 * (a + c) * t
Else
MsgBox "Bukan angka"
TLTa.Text = ""
TLTt.Text = ""
TLTc.Text = ""
End If
End If
End Sub
