VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5340
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9105
   LinkTopic       =   "Form3"
   ScaleHeight     =   5340
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Master 
      Caption         =   "Entry"
      Begin VB.Menu Epas 
         Caption         =   "Entry Data Pasien"
         Index           =   0
      End
      Begin VB.Menu Ediag 
         Caption         =   "Entry Data Diagnosa"
         Index           =   1
      End
      Begin VB.Menu Edok 
         Caption         =   "Entry Data Dokter"
         Index           =   2
      End
      Begin VB.Menu Eobat 
         Caption         =   "Entry Data Obat"
         Index           =   3
      End
      Begin VB.Menu Etin 
         Caption         =   "Entry Data Tindakan"
         Index           =   4
      End
      Begin VB.Menu Epem 
         Caption         =   "Entry Data Pemeriksaan"
         Index           =   5
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Ediag_Click(Index As Integer)
Form2.Show

End Sub

Private Sub Edok_Click(Index As Integer)
Form4.Show
End Sub

Private Sub Eobat_Click(Index As Integer)
Form5.Show

End Sub

Private Sub Epas_Click(Index As Integer)
Form1.Show

End Sub

Private Sub Epem_Click(Index As Integer)
Form7.Show
End Sub

Private Sub Etin_Click(Index As Integer)
Form6.Show

End Sub
