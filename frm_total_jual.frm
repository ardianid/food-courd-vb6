VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_total_jual 
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   120
      ScaleHeight     =   6585
      ScaleWidth      =   3825
      TabIndex        =   8
      Top             =   1920
      Width           =   3855
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   4080
      ScaleHeight     =   8385
      ScaleWidth      =   11025
      TabIndex        =   7
      Top             =   120
      Width           =   11055
      Begin CRVIEWERLibCtl.CRViewer CRViewer1 
         Height          =   8295
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   11055
         DisplayGroupTree=   -1  'True
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   -1  'True
         EnableNavigationControls=   -1  'True
         EnableStopButton=   -1  'True
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   -1  'True
         EnableProgressControl=   -1  'True
         EnableSearchControl=   -1  'True
         EnableRefreshButton=   -1  'True
         EnableDrillDown =   -1  'True
         EnableAnimationControl=   -1  'True
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   -1  'True
         DisplayTabs     =   -1  'True
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1665
      ScaleWidth      =   3825
      TabIndex        =   4
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton cmd_set 
         Caption         =   "Printer Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmd_tampil 
         Caption         =   "Tampil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
      End
      Begin MSMask.MaskEdBox msk_tgl1 
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_tgl2 
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S/d"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1800
         TabIndex        =   6
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   840
      End
   End
End
Attribute VB_Name = "frm_total_jual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New lap_total_jual
Dim sql As String
Dim rs As New ADODB.Recordset
Option Explicit

Private Sub cmd_set_Click()
    Report.PrinterSetup Me.hWnd
End Sub


Private Sub cmd_tampil_Click()
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    If msk_tgl1.Text <> "__/__/____" And msk_tgl2.Text <> "__/__/____" Then
        
        sql = "select * from qr_semua_penjualan where"
        sql = sql & " tgl >= datevalue('" & Trim(msk_tgl1.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl2.Text) & "')"
        
        sql = sql & " order by kode_counter"
        
        rs.Open sql, cn
        
        
      '  Report.FormulaFields(0) = "tgl_awal='" & Trim(msk_tgl1.Text) & "'"
      '  Report.FormulaFields(1) = "tgl_akhir='" & Trim(msk_tgl2.Text) & "'"
        Report.SQLQueryString = sql
        Report.Database.SetDataSource rs
        Report.DiscardSavedData
        tampil
        
    Else
           
        tampil_semua
        tampil
           
    End If
    
     
    
        
    
End Sub

Private Sub Form_Load()
    
     CRViewer1.DisplayGroupTree = False
     CRViewer1.EnableExportButton = True
           
     
    ' Report.FormulaFields.Item(0) = "tgl_awal='Semua'"
    ' Report.FormulaFields(1) = "tgl_akhir='Semua'"
     
        
    tampil_semua
    
    tampil
    
End Sub

Private Sub tampil()
Screen.MousePointer = vbHourglass
CRViewer1.ReportSource = Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Sub
    Private Sub tampil_semua()
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    sql = "select * from qr_semua_penjualan order by kode_counter"
    rs.Open sql, cn
    Report.SQLQueryString = sql
    Report.Database.SetDataSource rs


End Sub
