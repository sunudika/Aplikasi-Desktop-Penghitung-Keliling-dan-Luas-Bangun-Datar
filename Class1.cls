VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BangunDatar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function kPersegi(ByVal s As Double) As Double
    kPersegi = s * 4
End Function

Public Function lPersegi(ByVal s As Double) As Double
    lPersegi = s * s
End Function

Public Function kPersegiP(ByVal p As Double, l As Double) As Double
    kPersegiP = 2 * (p + l)
End Function

Public Function lPersegiP(ByVal p As Double, l As Double) As Double
    kPersegiP = p * l
End Function

Public Function lSegitiga(ByVal a As Double, t As Double) As Double
    lSegitiga = 0.5 * a * t
End Function

Public Function kSegitiga(ByVal a As Double, b As Double, c As Double) As Double
    kSegitiga = a + b + c
End Function

Public Function kJajarG(ByVal a As Double, b As Double) As Double
    kJajarG = 2 * (a + b)
End Function

Public Function lJajarG(ByVal a As Double, t As Double) As Double
    lJajarG = a * t
End Function

Public Function kLayangL(ByVal a As Double, b As Double) As Double
    kLayangL = 2 * (a + b)
End Function

Public Function lLayangL(ByVal d1 As Double, d2 As Double) As Double
    kLayangL = 0.5 * d1 * d2
End Function

Public Function kBelahK(ByVal s As Double) As Double
    kBelahK = s * 4
End Function

Public Function lBelahK(ByVal d1 As Double, d2 As Double) As Double
    lBelahK = 0.5 * d1 * d2
End Function

Public Function kTrapesium(ByVal a As Double, b As Double, c As Double, d As Double) As Double
    kTrapesium = a + b + c + d
End Function

Public Function lTrapesium(ByVal a As Double, c As Double, t As Double) As Double
    lTrapesium = 0.5 * (a + c) * t
End Function

Public Function kLingkaran(ByVal r As Double) As Double
    kLingkaran = 2 * 3.14 * r
End Function

Public Function lLingkaran(ByVal r As Double) As Double
    lLingkaran = 3.14 * r * r
End Function
