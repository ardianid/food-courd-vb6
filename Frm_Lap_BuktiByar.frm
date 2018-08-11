VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form Frm_Lap_BuktiByar 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   4230
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3525
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   0   'False
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "Frm_Lap_BuktiByar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As Report_BuktiBayar

Private Sub Form_Load()

On Error GoTo err_handler

Screen.MousePointer = vbHourglas

With Me
    .Width = 4320
    .Height = 7605
    .ScaleWidth = 4230
    .ScaleHeight = 7245
    .Left = Screen.Width / 2 - .Width / 2
    .Top = Screen.Height / 2 - .Height / 2
End With

Dim sql As String
Dim rs As Recordset
    sql = "select * from qr_penjualan_sebenarnya where no_faktur='" & noff & "'"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, cn
        
        Set Report = New Report_BuktiBayar
            Report.Database.SetDataSource rs, , 1
            
        Report.TxtBayar.SetText Format(byyr, "###,###,###")
        Report.TxtKembali.SetText Format(kemm, "###,###,###")

CRViewer1.Zoom 1

CRViewer1.ReportSource = Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault

On Error GoTo 0
Exit Sub

err_handler:

Screen.MousePointer = vbDefault
    
    Dim konfirm As Integer
        konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
            Err.Clear

End Sub

Private Sub Form_Resize()

CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If htu = True Then frm_jual.baru_lagi
    
End Sub
