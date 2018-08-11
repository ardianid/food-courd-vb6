VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_tot_jual1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleksi Total Berdasarkan Disc"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2265
      ScaleWidth      =   5265
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin Crystal.CrystalReport rpt 
         Left            =   840
         Top             =   2400
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.CommandButton cmd_tampil 
         Caption         =   "Tampil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
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
         Left            =   2280
         TabIndex        =   3
         Top             =   2400
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox cek_tanggal 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total  Pertanggal"
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
         Left            =   600
         TabIndex        =   5
         Top             =   2520
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSMask.MaskEdBox msk_tgl1 
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         Left            =   3480
         TabIndex        =   2
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kriteria Pencetakan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   2220
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4920
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   7
         Top             =   720
         Width           =   270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Jenis Barang"
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
         TabIndex        =   6
         Top             =   2280
         Visible         =   0   'False
         Width           =   2010
      End
   End
End
Attribute VB_Name = "frm_tot_jual1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Tampil_Click()
  
On Error GoTo er_tampil
    
    Me.MousePointer = vbHourglass
    utama.MousePointer = vbHourglass
    
    Dim sql As String
'    Dim rs As New ADODB.Recordset
'
'    If rs.State = adStateOpen Then
'        rs.Close
'    End If
    
' If cek_tanggal.Value = vbUnchecked Then
    
'    sql = "select kode_counter,nama_counter,kode_barang,"
'    sql = sql & "nama_barang,no_faktur,tgl,qty,"
    
    sql = ""

    sql = sql & "select * from qr_penjualan_sebenarnya"
    
'End If
    
'If cek_tanggal.Value = vbChecked Then
'
'    sql = "select kode_counter,nama_counter,tgl,harga_sebenarnya,total_harga from qr_penjualan_sebenarnya"
'
'End If
    
    If msk_tgl1.Text <> "__/__/____" And msk_tgl2.Text <> "__/__/____" Then
        sql = sql & " where tgl >= datevalue('" & Trim(msk_tgl1.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl2.Text) & "')"
        
'            If txt_kode.Text <> "" Then
'                sql = sql & " and kode_counter='" & Trim(txt_kode.Text) & "'"
'            End If
            
    End If
    
    sql = sql & " Order by kode_counter"
    
    sqlku = sql
    
'    rs.Open sql, cn
'If cek_tanggal.Value = vbUnchecked Then
'    rpt.ReportFileName = path_lap & "\lap_total_per.rpt"
'End If
'
'If cek_tanggal.Value = vbChecked Then
'    rpt.ReportFileName = path_lap & "\lap_tot_tgl.rpt"
'End If

'    rpt.Connect = cn
'    rpt.RetrieveSQLQuery
'    rpt.SQLQuery = sql
'
'    If msk_tgl1.Text <> "__/__/____" And msk_tgl2.Text <> "__/__/____" Then
'        rpt.Formulas(0) = "tgl_awal='" & Trim(msk_tgl1.Text) & "'"
'        rpt.Formulas(1) = "tgl_akhir='" & Trim(msk_tgl2.Text) & "'"
'    Else
'
'        rpt.Formulas(0) = "tgl_awal='Semua'"
'        rpt.Formulas(1) = "tgl_akhir='Semua'"
'    End If
'
'    rpt.Formulas(2) = "pemakai='" & id_user & "'"
'    rpt.DiscardSavedData = True
'    rpt.WindowState = crptMaximized
'    rpt.Action = 1
    
    If Me.MousePointer = vbHourglass Then
        Me.MousePointer = vbDefault
        utama.MousePointer = vbDefault
    End If
    
    Load Frm_Lap_Tot_Setelah_Discppn
        Frm_Lap_Tot_Setelah_Discppn.Show
    
    Exit Sub
    
er_tampil:
    
    If Me.MousePointer = vbHourglass Then
        Me.MousePointer = vbDefault
        utama.MousePointer = vbDefault
    End If
    
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub


Private Sub Form_Activate()
On Error Resume Next
    msk_tgl1.SetFocus
End Sub

Private Sub Form_Load()
    
    With Me
        .Left = Screen.Width / 2 - .Width / 2
        .Top = 350
    End With
    
End Sub

Private Sub msk_tgl1_GotFocus()
    msk_tgl1.SelStart = 0
    msk_tgl1.SelLength = Len(msk_tgl1)
End Sub

Private Sub msk_tgl1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then msk_tgl2.SetFocus
End Sub

Private Sub msk_tgl2_GotFocus()
    msk_tgl2.SelStart = 0
    msk_tgl2.SelLength = Len(msk_tgl2)
End Sub

Private Sub msk_tgl2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmd_tampil.SetFocus
End Sub

Private Sub txt_kode_GotFocus()
    txt_kode.SelStart = 0
    txt_kode.SelLength = Len(txt_kode)
End Sub
