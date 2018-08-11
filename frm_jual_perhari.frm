VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_jual_perhari 
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   3720
      ScaleHeight     =   8265
      ScaleWidth      =   11505
      TabIndex        =   7
      Top             =   120
      Width           =   11535
      Begin CRVIEWERLibCtl.CRViewer CRViewer1 
         Height          =   8175
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   11415
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
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1905
      ScaleWidth      =   3465
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton cmd_set 
         Caption         =   "Printer Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txt_kode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin MSMask.MaskEdBox msk_tgl 
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.CommandButton cmd_tampil 
         Caption         =   "Tampil"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Counter"
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
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Shape Shape1 
      DrawMode        =   5  'Not Copy Pen
      Height          =   6255
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   3495
   End
End
Attribute VB_Name = "frm_jual_perhari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New lap_penjualan_hari
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
        
        If msk_tgl.Text <> "__/__/____" Or txt_kode.Text <> "" Then

            sql = "select * from qr_semua_penjualan where"

            If msk_tgl.Text <> "__/__/____" Then
                sql = sql & " tgl= datevalue('" & Trim(msk_tgl.Text) & "')"
            End If

            If txt_kode.Text <> "" And msk_tgl.Text = "__/__/____" Then
                sql = sql & " kode_counter='" & Trim(txt_kode.Text) & "'"
            End If

            If txt_kode.Text <> "" And msk_tgl.Text <> "__/__/____" Then
                sql = sql & " and kode_counter='" & Trim(txt_kode.Text) & "'"
            End If
        sql = sql & " order by kode_counter"
        
        rs.Open sql, cn
        
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
     
    tampil_semua
    
    tampil

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

Private Sub tampil()
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
    Screen.MousePointer = vbDefault
End Sub

