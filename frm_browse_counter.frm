VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_browse_counter 
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   13200
      ScaleHeight     =   705
      ScaleWidth      =   1905
      TabIndex        =   18
      Top             =   1920
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   13200
      ScaleHeight     =   5385
      ScaleWidth      =   1905
      TabIndex        =   9
      Top             =   2760
      Width           =   1935
      Begin VB.CommandButton cmd_export 
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   4560
         Width           =   1695
      End
      Begin VB.CommandButton cmd_cetak 
         Caption         =   "Cetak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton cmd_setup 
         Caption         =   "Page Setup"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton cmd_hapus 
         Caption         =   "Hapus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton cmd_edit 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   1920
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   1920
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   1920
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   1920
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   120
      ScaleHeight     =   6225
      ScaleWidth      =   12945
      TabIndex        =   8
      Top             =   1920
      Width           =   12975
      Begin TrueDBGrid60.TDBGrid grd_counter 
         Height          =   5895
         Left            =   120
         OleObjectBlob   =   "frm_browse_counter.frx":0000
         TabIndex        =   15
         Top             =   120
         Width           =   12735
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   3840
         Top             =   1680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1665
      ScaleWidth      =   14985
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin VB.TextBox txt_alamat 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txt_nama 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txt_kode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmd_tampil 
         Caption         =   "Tampil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13560
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txt_presentasi 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9600
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txt_counter 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   1080
         Width           =   5175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Pemilik"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pemilik"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Counter"
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
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10680
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Presentasi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   3
         Top             =   1080
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Counter"
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
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1425
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   7680
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pencarian"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frm_browse_counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_counter As New XArrayDB
Dim sql_counter As String
Dim id_ctr As String

Private Sub cmd_cetak_Click()
On Error GoTo er_printer

    With grd_counter.PrintInfo
        
        .PageHeaderFont.Bold = True
        .PageHeaderFont.Italic = True
        .PageHeaderFont.Size = 10
        .PageHeader = "Laporan Data Jenis Barang"
        .RepeatColumnHeaders = True
        .PageFooter = "\tPage: \p" & "..." & id_user
        .PrintPreview
    End With
    Exit Sub
    
er_printer:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub cmd_edit_Click()
    If arr_counter.UpperBound(1) > 0 Then
        mdl_counter = False
        id_counter = id_ctr
        frm_counter.Show
    End If
End Sub

Private Sub cmd_export_Click()
    
    On Error Resume Next

    cd.ShowSave
    grd_counter.ExportToFile cd.FileName, False

End Sub

Private Sub cmd_hapus_Click()

On Error GoTo er_hps

Dim sql, sql1 As String
Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    
If arr_counter.UpperBound(1) > 0 Then
    If MsgBox("Yakin akan dihapus data counter " & arr_counter(grd_counter.Bookmark, 5), vbYesNo + vbQuestion, "Pesan") = vbYes Then
        sql = "select id from tbl_counter where id=" & id_ctr
        rs.Open sql, cn
            If Not rs.EOF Then
            
                sql1 = "delete from tbl_counter where id=" & id_ctr
                rs1.Open sql1, cn
        
            Else
                MsgBox ("Data yang akan dihapus tidak ditemukan")
            End If
        rs.Close
        isi_counter
   Else
        Exit Sub
   End If
End If

Exit Sub
        
er_hps:
    
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
    
End Sub

Private Sub cmd_setup_Click()
On Error GoTo er_set
    
    With grd_counter.PrintInfo
        .PageSetup
    End With
    Exit Sub
    
er_set:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
    
End Sub

Private Sub Cmd_Tampil_Click()
    Dim rs_counter As New ADODB.Recordset
    
On Error GoTo er_tampil
  
If txt_nama.Text = "" And txt_alamat.Text = "" And txt_kode.Text = "" And txt_counter.Text = "" And txt_presentasi.Text = "" Then
    isi_counter
    Exit Sub
End If
    
If txt_nama.Text <> "" Or txt_alamat.Text <> "" Or txt_kode.Text <> "" Or txt_counter.Text <> "" Or txt_presentasi.Text <> "" Then
    
    sql_counter = "select * from tbl_counter where"
    
    If txt_nama.Text <> "" Then
        sql_counter = sql_counter & " nama_pemilik like '%" & Trim(txt_nama.Text) & "%'"
    End If
    
    If txt_alamat.Text <> "" And txt_nama.Text = "" Then
        sql_counter = sql_counter & " alamat like '%" & Trim(txt_alamat.Text) & "%'"
    End If
    
    If txt_alamat.Text <> "" And txt_nama.Text <> "" Then
        sql_counter = sql_counter & " and alamat like '%" & Trim(txt_alamat.Text) & "%'"
    End If
    
    If txt_kode.Text <> "" And txt_nama.Text = "" And txt_alamat.Text = "" Then
        sql_counter = sql_counter & " kode='" & Trim(txt_kode.Text) & "'"
    End If
    
    If txt_kode.Text <> "" And (txt_nama.Text <> "" Or txt_alamat.Text <> "") Then
        sql_counter = sql_counter & " and kode='" & Trim(txt_kode.Text) & "'"
    End If
    
    If txt_counter.Text <> "" And txt_kode.Text = "" And txt_nama.Text = "" And txt_alamat.Text = "" Then
        sql_counter = sql_counter & " nama_counter like '%" & Trim(txt_counter.Text) & "%'"
    End If
    
    If txt_counter.Text <> "" And (txt_kode.Text <> "" Or txt_nama.Text <> "" Or txt_alamat.Text <> "") Then
        sql_counter = sql_counter & " and nama_counter like '%" & Trim(txt_counter.Text) & "%'"
    End If
    
    If txt_presentasi.Text <> "" And txt_counter.Text = "" And txt_kode.Text = "" And txt_nama.Text = "" And txt_alamat.Text = "" Then
        sql_counter = sql_counter & " presentasi_p like '%" & Trim(txt_presentasi.Text) & "%'"
    End If
    
    If txt_presentasi.Text <> "" And (txt_counter.Text <> "" Or txt_kode.Text <> "" Or txt_nama.Text <> "" Or txt_alamat.Text <> "") Then
        sql_counter = sql_counter & " and presentasi_p like '%" & Trim(txt_presentasi.Text) & "%'"
    End If
    
    
    rs_counter.Open sql_counter, cn, adOpenKeyset
    If Not rs_counter.EOF Then
        
        rs_counter.MoveLast
        rs_counter.MoveFirst
        
        l_isi rs_counter
    End If
    rs_counter.Close
End If

Exit Sub

er_tampil:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear

End Sub

Private Sub Form_Load()

    grd_counter.Array = arr_counter
    
    isi_counter
    
    Call cari_wewenang("Form Data Counter")
      
        If edit_form = True Then
            cmd_edit.Enabled = True
        Else
            cmd_edit.Enabled = False
        End If
        
        If hapus_form = True Then
            Cmd_Hapus.Enabled = True
        Else
            Cmd_Hapus.Enabled = False
        End If
        
        If lap_form = True Then
            cmd_setup.Enabled = True
            cmd_cetak.Enabled = True
            cmd_export.Enabled = True
        Else
            cmd_setup.Enabled = False
            cmd_cetak.Enabled = False
            cmd_export.Enabled = False
        End If
    
    'Me.Left = utama.Width \ 2 - Me.Width \ 2
    'Me.Top = utama.Height \ 2 - Me.Height \ 2 - 1500
    
End Sub

Private Sub kosong()
    arr_counter.ReDim 0, 0, 0, 0
    grd_counter.ReBind
    grd_counter.Refresh
End Sub

Public Sub isi_counter()

On Error GoTo er_counter

    Dim rs_counter As New ADODB.Recordset
        
        kosong
        
        sql_counter = "select * from tbl_counter"
        rs_counter.Open sql_counter, cn, adOpenKeyset
            If Not rs_counter.EOF Then
                
                rs_counter.MoveLast
                rs_counter.MoveFirst
                
                l_isi rs_counter
            End If
        rs_counter.Close
        
        Exit Sub
er_counter:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub

Private Sub l_isi(rs_counter As Recordset)
    Dim id_c, kd_c, nm_c, pr_c As String
    Dim nama, alamat, telp As String
    Dim a As Long
        
        
        a = 1
            Do While Not rs_counter.EOF
                arr_counter.ReDim 1, a, 0, 9
                grd_counter.ReBind
                grd_counter.Refresh
                    
                    id_c = rs_counter("id")
                      
                    If Not IsNull(rs_counter("nama_pemilik")) Then
                      nama = rs_counter("nama_pemilik")
                    Else
                      nama = ""
                    End If
                      
                    If Not IsNull(rs_counter("alamat")) Then
                      alamat = rs_counter("alamat")
                    Else
                        alamat = ""
                    End If
                      
                    If Not IsNull(rs_counter("telp")) Then
                      telp = rs_counter("telp")
                    Else
                        telp = ""
                    End If
                      
                    If Not IsNull(rs_counter("kode")) Then
                      kd_c = rs_counter("kode")
                    Else
                        kd_c = ""
                    End If
                      
                    If Not IsNull(rs_counter("nama_counter")) Then
                        nm_c = rs_counter("nama_counter")
                    Else
                        nm_c = ""
                    End If
                    
                    If Not IsNull(rs_counter("presentasi_p")) Then
                        pr_c = rs_counter("presentasi_p")
                    Else
                        pr_c = ""
                    End If
                    
                arr_counter(a, 0) = id_c
                arr_counter(a, 1) = a
                arr_counter(a, 2) = nama
                arr_counter(a, 3) = alamat
                arr_counter(a, 4) = telp
                arr_counter(a, 5) = kd_c
                arr_counter(a, 6) = nm_c
                arr_counter(a, 7) = pr_c
                    
           a = a + 1
           rs_counter.MoveNext
           Loop
           grd_counter.ReBind
           grd_counter.Refresh
           
End Sub

Private Sub grd_counter_Click()
    On Error Resume Next
        If arr_counter.UpperBound(1) > 0 Then
            id_ctr = arr_counter(grd_counter.Bookmark, 0)
        End If
End Sub

Private Sub grd_counter_HeadClick(ByVal ColIndex As Integer)
    Dim sql As String
    Dim rs_counter As New ADODB.Recordset
          
On Error GoTo er_head
          
If arr_counter.UpperBound(1) > 0 Then

    If sql_counter = "" Then
        Exit Sub
    End If
    
    sql = sql_counter
        
Select Case ColIndex
    Case 2
        sql = sql & " order by nama_pemilik"
    Case 3
        sql = sql & " order by alamat"
    Case 4
        sql = sql & " order by telp"
    Case 5
        sql = sql & " order by kode"
    Case 6
        sql = sql & " order by nama_counter"
    Case 7
        sql = sql & " order by presentasi_p"
End Select

rs_counter.Open sql, cn, adOpenKeyset
    If Not rs_counter.EOF Then
        
        rs_counter.MoveLast
        rs_counter.MoveFirst
        
        l_isi rs_counter
    End If
rs_counter.Close
    
End If

Exit Sub

er_head:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear

End Sub

Private Sub grd_counter_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_counter_Click
End Sub

Private Sub txt_counter_GotFocus()
    txt_counter.SelStart = 0
    txt_counter.SelLength = Len(txt_counter)
End Sub

Private Sub txt_presentasi_GotFocus()
    txt_presentasi.SelStart = 0
    txt_presentasi.SelLength = Len(txt_presentasi)
End Sub

Private Sub txt_presentasi_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",")) Then
        Beep
        KeyAscii = 0
    End If
End Sub
