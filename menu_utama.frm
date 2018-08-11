VERSION 5.00
Begin VB.MDIForm menu_utama 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8640
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10785
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mn_mas 
      Caption         =   "&Master"
      Begin VB.Menu mn_mass 
         Caption         =   "&Input Inventori"
         Index           =   0
      End
      Begin VB.Menu mn_mass 
         Caption         =   "&Penyesuaian Stock"
         Index           =   1
      End
      Begin VB.Menu mn_mass 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mn_mass 
         Caption         =   "&Historical Inventori"
         Index           =   3
      End
   End
End
Attribute VB_Name = "menu_utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
     buka_koneksi
End Sub

Private Sub mn_mass_Click(Index As Integer)
    If Index = 0 Then
        frm_input_inventori.Show
       ElseIf Index = 1 Then
        frm_penyesuaian_stock.Show
       ElseIf Index = 3 Then
        frm_histo_invent.Show
    End If
         
    
    
End Sub
