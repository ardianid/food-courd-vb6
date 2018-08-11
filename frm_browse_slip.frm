VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_browse_slip 
   Caption         =   "Browse Slip Pembayaran"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   120
      ScaleHeight     =   8265
      ScaleWidth      =   14985
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin MSComDlg.CommonDialog cd 
         Left            =   1800
         Top             =   4320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmd_export 
         Caption         =   "Export"
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
         Left            =   3240
         TabIndex        =   12
         Top             =   7680
         Width           =   1455
      End
      Begin VB.CommandButton cmd_cetak 
         Caption         =   "Cetak"
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
         Left            =   1680
         TabIndex        =   11
         Top             =   7680
         Width           =   1455
      End
      Begin VB.CommandButton cmd_pagesetup 
         Caption         =   "Page Setup"
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
         TabIndex        =   10
         Top             =   7680
         Width           =   1455
      End
      Begin VB.CommandButton cmd_hapus 
         Caption         =   "Hapus"
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
         Left            =   13440
         TabIndex        =   9
         Top             =   7680
         Width           =   1455
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
         Left            =   13440
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin TrueDBGrid60.TDBGrid grd_daftar 
         Height          =   6135
         Left            =   120
         OleObjectBlob   =   "frm_browse_slip.frx":0000
         TabIndex        =   8
         Top             =   1440
         Width           =   14775
      End
      Begin MSMask.MaskEdBox msk_tgl1 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
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
      Begin VB.TextBox txt_kode 
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
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin MSMask.MaskEdBox msk_tgl2 
         Height          =   375
         Left            =   4800
         TabIndex        =   3
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         TabIndex        =   7
         Top             =   840
         Width           =   345
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frm_browse_slip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB
Dim id_daftar As String
Dim sql1 As String

Private Sub cmd_cetak_Click()

On Error GoTo er_printer

    With grd_daftar.PrintInfo
        
        .PageHeaderFont.Bold = True
        .PageHeaderFont.Italic = True
        .PageHeaderFont.Size = 11
        .PageHeader = "Laporan Transaksi Slip Pembayaran"
        .RepeatColumnHeaders = True
        .PageFooter = "\tPage: \p" & "..." & id_user
        .PrintPreview
    End With
    
    On Error GoTo 0
    
    Exit Sub
    
er_printer:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub cmd_export_Click()
On Error Resume Next
    cd.ShowSave
    grd_daftar.ExportToFile cd.FileName, False
End Sub

Private Sub cmd_hapus_Click()

On Error GoTo er_hapus

Dim sql, sql2 As String
Dim rs As New ADODB.Recordset, rs2 As New ADODB.Recordset
    
    If arr_daftar.UpperBound(1) = 0 Then
        Exit Sub
    End If
    
    If MsgBox("Yakin akan dihapus.........?", vbYesNo + vbQuestion, "Hapus Data") = vbNo Then
        Exit Sub
    End If
    
    sql = "select id from tr_slip_pembayaran where id=" & id_daftar
    rs.Open sql, cn
        If Not rs.EOF Then
            
            sql2 = "delete from tr_slip_pembayaran where id=" & id_daftar
            rs2.Open sql2, cn
            
        Else
            MsgBox ("Data yang akan dihapus tidak ditemukan")
        End If
    rs.Close
    
    cmd_tampil_Click
    
    On Error GoTo 0
    
    Exit Sub
    
er_hapus:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub cmd_pagesetup_Click()
On Error GoTo er_p
    
    With grd_daftar.PrintInfo
        .PageSetup
    End With
    
    On Error GoTo 0
    
    Exit Sub
    
er_p:
        Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
    
End Sub

Private Sub cmd_tampil_Click()

On Error GoTo er_tampil

Dim rs1 As New ADODB.Recordset
    
    kosong_daftar
        
        sql1 = "select * from tr_slip_pembayaran"
    
    If txt_kode.Text <> "" Or (msk_tgl1.Text <> "__/__/____" And msk_tgl2.Text <> "__/__/____") Then
        
        sql1 = sql1 & " where"
        
        If txt_kode.Text <> "" Then
            sql1 = sql1 & " kode_counter='" & Trim(txt_kode.Text) & "'"
        End If
        
        If msk_tgl1.Text <> "__/__/____" And msk_tgl2.Text <> "__/__/____" And txt_kode.Text = "" Then
            sql1 = sql1 & " periode1 >= datevalue('" & Trim(msk_tgl1.Text) & "') and periode2 <= datevalue('" & Trim(msk_tgl2.Text) & "')"
        End If
        
        If msk_tgl1.Text <> "__/__/____" And msk_tgl2.Text <> "__/__/____" And txt_kode.Text <> "" Then
            sql1 = sql1 & " and periode1 >= datevalue('" & Trim(msk_tgl1.Text) & "') and periode2 <= datevalue('" & Trim(msk_tgl2.Text) & "')"
        End If
        
    End If
    
    rs1.Open sql1, cn, adOpenKeyset
    If Not rs1.EOF Then
        
        rs1.MoveLast
        rs1.MoveFirst
        
        lanjut_isi rs1
     End If
     rs1.Close
     
     On Error GoTo 0
     
     Exit Sub
    
er_tampil:
        Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub

Private Sub Form_Load()
    
    Me.Show
    
    grd_daftar.Array = arr_daftar
    
    isi_daftar
    
End Sub

Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub

Private Sub grd_daftar_Click()
On Error Resume Next
    If arr_daftar.UpperBound(1) > 0 Then
        id_daftar = arr_daftar(grd_daftar.Bookmark, 0)
    End If
End Sub

Private Sub grd_daftar_HeadClick(ByVal ColIndex As Integer)

On Error Resume Next

Dim rs1 As New ADODB.Recordset
Dim sql As String

If sql1 = "" Then
    Exit Sub
End If

If arr_daftar.UpperBound(1) = 0 Then
    Exit Sub
End If
    
sql = ""
    
sql = sql1
    
Select Case ColIndex
    
    Case 1
        sql = sql & " order by tgl_bayar"
    Case 2
        sql = sql & " order by periode1"
    Case 3
        sql = sql & " order by periode2"
    Case 4
        sql = sql & " order by kode_counter"
    Case 5
        sql = sql & " order by tot_jual"
    Case 6
        sql = sql & " order by persentase"
    Case 7
        sql = sql & " order by nilai"
    Case 8
        sql = sql & " order by ppn"
    Case 9
        sql = sql & " order by jumlah"
    Case 10
        sql = sql & " order by pot_air"
    Case 11
        sql = sql & " order by pot_listrik"
    Case 12
        sql = sql & " order by pot_lain"
    Case 13
        sql = sql & " order by total"
    Case 15
        sql = sql & " order by nama_user"
End Select
    
    rs1.Open sql, cn, adOpenKeyset
        If Not rs1.EOF Then
            
            rs1.MoveLast
            rs1.MoveFirst
            
            lanjut_isi rs1
        End If
    rs1.Close
    
End Sub

Private Sub grd_daftar_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_daftar_Click
End Sub

Private Sub msk_tgl1_GotFocus()
    msk_tgl1.SelStart = 0
    msk_tgl1.SelLength = Len(msk_tgl1)
End Sub
Private Sub msk_tgl2_GotFocus()
    msk_tgl2.SelStart = 0
    msk_tgl2.SelLength = Len(msk_tgl2)
End Sub

Private Sub txt_kode_GotFocus()
    txt_kode.SelStart = 0
    txt_kode.SelLength = Len(txt_kode)
End Sub

Private Sub isi_daftar()

On Error GoTo er_daftar

    Dim rs1 As New ADODB.Recordset
        
        kosong_daftar
        
        sql1 = "select * from tr_slip_pembayaran"
        rs1.Open sql1, cn, adOpenKeyset
            If Not rs1.EOF Then
                rs1.MoveLast
                rs1.MoveFirst
                
                lanjut_isi rs1
            End If
        rs1.Close
    
    On Error GoTo 0
    
    Exit Sub
    
er_daftar:
    Dim psn
        psn = MsgBox(Err.Number & vbCrLf & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Sub lanjut_isi(rs1 As Recordset)
    Dim tgl, periode1, periode2, kode_counter, tot_jual, persentase, nilai, ppn, jumlah, pot_listrik, pot_air, pot_lain, total, ket, nama_user As String
    Dim a As Long
    Dim idnya As String
        
        a = 1
            Do While Not rs1.EOF
                arr_daftar.ReDim 1, a, 0, 17
                grd_daftar.ReBind
                grd_daftar.Refresh
                                
                    idnya = rs1!id
                                
                    If Not IsNull(rs1!tgl_bayar) Then
                        tgl = rs1!tgl_bayar
                    Else
                        tgl = ""
                    End If
                    
                    If Not IsNull(rs1!periode1) Then
                        periode1 = rs1!periode1
                    Else
                        periode1 = ""
                    End If
                    
                    If Not IsNull(rs1!periode2) Then
                        periode2 = rs1!periode2
                    Else
                        periode2 = ""
                    End If
                    
                    If Not IsNull(rs1!kode_counter) Then
                        kode_counter = rs1!kode_counter
                    Else
                        kode_counter = ""
                    End If
                    
                    If Not IsNull(rs1!tot_jual) Then
                        tot_jual = Format(rs1!tot_jual, "###,###,###")
                    Else
                        tot_jual = 0
                    End If
                    
                    If Not IsNull(rs1!persentase) Then
                        persentase = rs1!persentase
                    Else
                        persentase = 0
                    End If
                    
                    If Not IsNull(rs1!nilai) Then
                        nilai = Format(rs1!nilai, "###,###,###")
                    Else
                        nilai = 0
                    End If
                    
                    If Not IsNull(rs1!ppn) Then
                        ppn = Format(rs1!ppn, "###,###,###")
                    Else
                        ppn = 0
                    End If
                    
                    If Not IsNull(rs1!jumlah) Then
                        jumlah = Format(rs1!jumlah, "###,###,###")
                    Else
                        jumlah = 0
                    End If
                    
                    If Not IsNull(rs1!pot_listrik) Then
                        pot_listrik = Format(rs1!pot_listrik, "###,###,###")
                    Else
                        pot_listrik = 0
                    End If
                    
                    If Not IsNull(rs1!pot_air) Then
                        pot_air = Format(rs1!pot_air, "###,###,###")
                    Else
                        pot_air = 0
                    End If
                    
                    If Not IsNull(rs1!pot_lain) Then
                        pot_lain = Format(rs1!pot_lain, "###,###,###")
                    Else
                        pot_lain = 0
                    End If
                    
                    If Not IsNull(rs1!total) Then
                        total = Format(rs1!total, "###,###,###")
                    Else
                        total = 0
                    End If
                    
                    If Not IsNull(rs1!ket) Then
                        ket = rs1!ket
                    Else
                        ket = 0
                    End If
                    
                    If Not IsNull(rs1!nama_user) Then
                        nama_user = rs1!nama_user
                    Else
                        nama_user = ""
                    End If
                    
                    arr_daftar(a, 0) = idnya
                    arr_daftar(a, 1) = tgl
                    arr_daftar(a, 2) = periode1
                    arr_daftar(a, 3) = periode2
                    arr_daftar(a, 4) = kode_counter
                    arr_daftar(a, 5) = tot_jual
                    arr_daftar(a, 6) = persentase
                    arr_daftar(a, 7) = nilai
                    arr_daftar(a, 8) = ppn
                    arr_daftar(a, 9) = jumlah
                    arr_daftar(a, 10) = pot_air
                    arr_daftar(a, 11) = pot_listrik
                    arr_daftar(a, 12) = pot_lain
                    arr_daftar(a, 13) = total
                    If ket = 0 Then
                        arr_daftar(a, 14) = vbUnchecked
                    Else
                        arr_daftar(a, 14) = vbChecked
                    End If
                    arr_daftar(a, 15) = nama_user
                a = a + 1
                rs1.MoveNext
                Loop
                grd_daftar.ReBind
                grd_daftar.Refresh
End Sub
