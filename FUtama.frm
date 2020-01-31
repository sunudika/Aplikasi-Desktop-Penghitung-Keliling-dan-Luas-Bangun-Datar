VERSION 5.00
Begin VB.Form FUtama 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Luas Dan Keliling Bangun Datar"
   ClientHeight    =   4110
   ClientLeft      =   1665
   ClientTop       =   6975
   ClientWidth     =   4755
   Icon            =   "FUtama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FUtama.frx":048A
   ScaleHeight     =   4110
   ScaleWidth      =   4755
   Begin VB.CommandButton CFUtama 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      Height          =   615
      Left            =   1680
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   615
      Left            =   2160
      Top             =   1800
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Perhitungan Keliling dan Luas Bangun Datar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Menu MnLuas 
      Caption         =   "Perhitungan Luas"
      Begin VB.Menu MnLPersegi 
         Caption         =   "Persegi"
      End
      Begin VB.Menu MnLPersegiPanjang 
         Caption         =   "Persegi Panjang"
      End
      Begin VB.Menu MnLSegitiga 
         Caption         =   "Segitiga"
      End
      Begin VB.Menu MnLJajarGenjang 
         Caption         =   "Jajar Genjang"
      End
      Begin VB.Menu MnLLayang 
         Caption         =   "Layang - Layang"
      End
      Begin VB.Menu MnLBelahKetupat 
         Caption         =   "Belah Ketupat"
      End
      Begin VB.Menu MnLTrapesium 
         Caption         =   "Trapesium"
      End
      Begin VB.Menu MnLLingkaran 
         Caption         =   "Lingkaran"
      End
   End
   Begin VB.Menu MnKeliling 
      Caption         =   "Perhitungan Keliling"
      Begin VB.Menu MnKPersegi 
         Caption         =   "Persegi"
      End
      Begin VB.Menu MnKPersegiPanjang 
         Caption         =   "Persegi Panjang"
      End
      Begin VB.Menu MnKSegitiga 
         Caption         =   "Segitiga"
      End
      Begin VB.Menu MnKJajarGenjang 
         Caption         =   "Jajar Genjang"
      End
      Begin VB.Menu MnKLayang 
         Caption         =   "Layang - Layang"
      End
      Begin VB.Menu MnKBelahKetupat 
         Caption         =   "Belah Ketupat"
      End
      Begin VB.Menu MnKTrapesium 
         Caption         =   "Trapesium"
      End
      Begin VB.Menu MnKLingkaran 
         Caption         =   "Lingkaran"
      End
   End
   Begin VB.Menu MnTentang 
      Caption         =   "Tentang"
      Begin VB.Menu MnPenulis 
         Caption         =   "Penulis"
      End
      Begin VB.Menu MnTugas 
         Caption         =   "Tugas"
      End
   End
End
Attribute VB_Name = "FUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MnPersegi_Click()

End Sub

Private Sub CFUtama_Click()
End
End Sub

Private Sub MnKBelahKetupat_Click()
FKBelahKetupat.Show
End Sub

Private Sub MnKJajarGenjang_Click()
FKJajarGenjang.Show
End Sub

Private Sub MnKLayang_Click()
FKLayang.Show
End Sub

Private Sub MnKLingkaran_Click()
FKLingkaran.Show
End Sub

Private Sub MnKPersegi_Click()
FKPersegi.Show
End Sub

Private Sub MnKPersegiPanjang_Click()
FKPersegiPanjang.Show
End Sub

Private Sub MnKSegitiga_Click()
FKSegitiga.Show
End Sub

Private Sub MnKTrapesium_Click()
FKTrapesium.Show
End Sub

Private Sub MnLBelahKetupat_Click()
FLBelahKetupat.Show
End Sub

Private Sub MnLJajarGenjang_Click()
FLJajarGenjang.Show
End Sub

Private Sub MnLLayang_Click()
FLLayang.Show
End Sub

Private Sub MnLLingkaran_Click()
FLLingkaran.Show
End Sub

Private Sub MnLPersegi_Click()
FLPersegi.Show
End Sub

Private Sub MnLPersegiPanjang_Click()
FLPersegiPanjang.Show
End Sub

Private Sub MnLSegitiga_Click()
FLSegitiga.Show
End Sub

Private Sub MnLTrapesium_Click()
FLTrapesium.Show
End Sub

Private Sub MnPenulis_Click()
FPenulis.Show
End Sub

Private Sub MnTugas_Click()
FTugas.Show
End Sub
