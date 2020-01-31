VERSION 5.00
Begin VB.Form FTugas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tentang Tugas"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "FTugas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FTugas.frx":048A
   ScaleHeight     =   4410
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TUGAS 1"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mata Kuliah Pemrograman API"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   1680
      Width           =   3015
   End
End
Attribute VB_Name = "FTugas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
