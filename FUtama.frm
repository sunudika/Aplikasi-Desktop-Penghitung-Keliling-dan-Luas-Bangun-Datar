VERSION 5.00
Begin VB.Form FUtama 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Luas Dan Keliling Bangun Datar"
   ClientHeight    =   3810
   ClientLeft      =   1740
   ClientTop       =   7350
   ClientWidth     =   3570
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   3570
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
End
Attribute VB_Name = "FUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MnPersegi_Click()

End Sub

Private Sub MnKBelahKetupat_Click()
FKBelahKetupat.Show
End Sub

Private Sub MnKJajarGenjang_Click()
FKJajarGenjang.Show
End Sub

Private Sub MnKLayang_Click()
FKJajarGenjang.Show
End Sub

Private Sub MnKLingkaran_Click()
FKLingkaran.Show
End Sub

Private Sub MnKPersegi_Click()
FKPersegi.Show
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

