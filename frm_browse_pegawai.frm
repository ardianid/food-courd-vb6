VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_browse_pegawai 
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cd 
      Left            =   3480
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TrueDBGrid60.TDBGrid grd_karyawan 
      Height          =   8055
      Left            =   2160
      OleObjectBlob   =   "frm_browse_pegawai.frx":0000
      TabIndex        =   12
      Top             =   120
      Width           =   12975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   120
      ScaleHeight     =   8025
      ScaleWidth      =   1905
      TabIndex        =   0
      Top             =   120
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
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   7320
         Width           =   1455
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
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   6720
         Width           =   1455
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
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   6120
         Width           =   1455
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
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   5040
         Width           =   1455
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
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   4440
         Width           =   1455
      End
      Begin VB.CommandButton cmd_cari 
         Caption         =   "Cari"
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
         Left            =   840
         TabIndex        =   6
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txt_alamat 
         Appearance      =   0  'Flat
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
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txt_nama 
         Appearance      =   0  'Flat
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
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   1800
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   1800
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   585
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   1800
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pencarian"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frm_browse_pegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_karyawan As New XArrayDB
Dim sql_k As String, id_yawan As String

Private Sub cmd_cari_Click()

On Error GoTo er_cari

    Dim rs_kr As New ADODB.Recordset
 If txt_nama.Text <> "" Or Txt_Alamat.Text <> "" Then
 
        sql_k = "select * from tbl_karyawan where"
    If txt_nama.Text <> "" And Txt_Alamat.Text = "" Then
        sql_k = sql_k & " nama_karyawan like '%" & Trim(txt_nama.Text) & "%'"
    End If
    
    If Txt_Alamat.Text <> "" And txt_nama.Text = "" Then
        sql_k = sql_k & " alamat like '%" & Trim(Txt_Alamat.Text) & "%'"
    End If
    
    If txt_nama.Text <> "" And Txt_Alamat.Text <> "" Then
        sql_k = sql_k & " nama_karyawan like '%" & Trim(txt_nama.Text) & "%' and alamat like '%" & Trim(Txt_Alamat.Text) & "%'"
    End If
    
        rs_kr.Open sql_k, cn, adOpenKeyset
            If Not rs_kr.EOF Then
                
                rs_kr.MoveLast
                rs_kr.MoveFirst
                    
                    lanjut_isi rs_kr
            End If
        rs_kr.Close
 End If
 Exit Sub
            
            
er_cari:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub cmd_cetak_Click()

On Error GoTo er_printer

    With grd_karyawan.PrintInfo
        
        .PageHeaderFont.Bold = True
        .PageHeaderFont.Italic = True
        .PageHeaderFont.Size = 10
        .PageHeader = "Laporan Data Karyawan"
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
    If arr_karyawan.UpperBound(1) > 0 Then
        id_kar = id_yawan
        mdl_karyawan = False
        frm_karyawan_lain.Show
    End If
End Sub

Private Sub cmd_export_Click()

On Error Resume Next

    cd.ShowSave
    grd_karyawan.ExportToFile cd.FileName, False

End Sub

Private Sub cmd_hapus_Click()

On Error GoTo er_h

    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    
If arr_karyawan.UpperBound(1) > 0 Then
    If MsgBox("Yakin akan dihapus nama karyawan " & arr_karyawan(grd_karyawan.Bookmark, 2), vbYesNo + vbQuestion, "Pesan") = vbYes Then
        sql = "select id from tbl_karyawan where id=" & id_yawan
        rs.Open sql, cn
            If Not rs.EOF Then
                sql1 = "delete from tbl_karyawan where id=" & id_yawan
                rs1.Open sql1, cn
            Else
                MsgBox ("Data yang akan dihapus tidak ditemukan")
            End If
        rs.Close
            
        isi_k
   Else
        Exit Sub
   End If
End If
    Exit Sub
    
er_h:
    Dim w
        w = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub

Private Sub cmd_setup_Click()

On Error GoTo er_set

    With grd_karyawan.PrintInfo
        .PageSetup
    End With
    Exit Sub
    
er_set:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub Form_Load()

    grd_karyawan.Array = arr_karyawan
    
    isi_k
    
     Call cari_wewenang("Form Karyawan")
      
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
      
'    Me.Left = Screen.Width \ 2 - Me.Width \ 2
'    Me.Top = Screen.Height \ 2 - Me.Height \ 2
    
End Sub

Private Sub kosong_grid()
    arr_karyawan.ReDim 0, 0, 0, 0
    grd_karyawan.ReBind
    grd_karyawan.Refresh
End Sub

Public Sub isi_k()

On Error GoTo er_ii

    Dim rs_kr As New ADODB.Recordset
    
    sql_k = "select * from tbl_karyawan"
    rs_kr.Open sql_k, cn, adOpenKeyset
        If Not rs_kr.EOF Then
            
            rs_kr.MoveLast
            rs_kr.MoveFirst
                
                lanjut_isi rs_kr
        End If
    rs_kr.Close
    
    Exit Sub
    
er_ii:
     Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub lanjut_isi(rs_kr As Recordset)
    Dim id_k, nm_k, al_k, tg_k, t_r, t_h, g_p As String
    Dim a As Long
        
    kosong_grid
    a = 1
        Do While Not rs_kr.EOF
            arr_karyawan.ReDim 1, a, 0, 9
            grd_karyawan.ReBind
            grd_karyawan.Refresh
                
            If Not IsNull(rs_kr("id")) Then
                id_k = rs_kr("id")
            Else
                id_k = ""
            End If
            If Not IsNull(rs_kr("nama_karyawan")) Then
                nm_k = rs_kr("nama_karyawan")
            Else
                nm_k = ""
            End If
            If Not IsNull(rs_kr("alamat")) Then
                al_k = rs_kr("alamat")
            Else
                al_k = ""
            End If
            
                tg_k = rs_kr("tempat_lhr") & " , " & rs_kr("tgl_lhr")
            If Not IsNull(rs_kr("telp_rumah")) Then
                t_r = rs_kr("telp_rumah")
            Else
                t_r = ""
            End If
            If Not IsNull(rs_kr("telp_hp")) Then
                t_h = rs_kr("telp_hp")
            Else
                t_h = ""
            End If
                
            If Not IsNull(rs_kr("gaji_pokok")) Then
                g_p = rs_kr("gaji_pokok")
            Else
                g_p = ""
            End If
                
            arr_karyawan(a, 0) = id_k
            arr_karyawan(a, 1) = a
            arr_karyawan(a, 2) = nm_k
            arr_karyawan(a, 3) = al_k
            arr_karyawan(a, 4) = tg_k
            arr_karyawan(a, 5) = t_r
            arr_karyawan(a, 6) = t_h
            arr_karyawan(a, 7) = g_p
            
        a = a + 1
        rs_kr.MoveNext
        Loop
        
        grd_karyawan.ReBind
        grd_karyawan.Refresh
End Sub

Private Sub grd_karyawan_Click()
    On Error Resume Next
        If arr_karyawan.UpperBound(1) > 0 Then
            id_yawan = arr_karyawan(grd_karyawan.Bookmark, 0)
        End If
End Sub

Private Sub grd_karyawan_HeadClick(ByVal ColIndex As Integer)
    Dim sql As String
    
On Error GoTo er_head

    If sql_k = "" Then
        Exit Sub
    End If
    
    sql = ""
    sql = sql & sql_k
    
    Select Case ColIndex
        Case 2
            sql = sql & " order by nama_karyawan"
        Case 3
            sql = sql & " order by alamat"
        Case 4
            sql = sql & " order by tempat_lhr,tgl_lhr"
        Case 5
            sql = sql & " order by telp_rumah"
        Case 6
            sql = sql & " order by telp_hp"
        Case 7
            sql = sql & " order by gaji_pokok"
    End Select
    
    Dim rs_kr As New ADODB.Recordset
        rs_kr.Open sql, cn, adOpenKeyset
            If Not rs_kr.EOF Then
                
                rs_kr.MoveLast
                rs_kr.MoveFirst
                
                lanjut_isi rs_kr
            End If
        rs_kr.Close
        
er_head:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub

Private Sub grd_karyawan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_karyawan_Click
End Sub

Private Sub txt_alamat_GotFocus()
    Txt_Alamat.SelStart = 0
    Txt_Alamat.SelLength = Len(Txt_Alamat)
End Sub

Private Sub txt_nama_GotFocus()
    txt_nama.SelStart = 0
    txt_nama.SelLength = Len(txt_nama)
End Sub
