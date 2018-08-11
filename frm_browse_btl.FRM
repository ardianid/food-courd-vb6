VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_browse_btl 
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8370
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cd 
      Left            =   3360
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   10320
      ScaleHeight     =   705
      ScaleWidth      =   4785
      TabIndex        =   10
      Top             =   7320
      Width           =   4815
      Begin VB.CommandButton cmd_setup 
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
         TabIndex        =   17
         Top             =   120
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
         TabIndex        =   16
         Top             =   120
         Width           =   1455
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
         TabIndex        =   15
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   9945
      TabIndex        =   9
      Top             =   7320
      Width           =   9975
      Begin VB.CommandButton cmd_batal 
         Caption         =   "Kembalikan Data"
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
         TabIndex        =   19
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1905
      ScaleWidth      =   14985
      TabIndex        =   0
      Top             =   240
      Width           =   15015
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
         Left            =   13080
         TabIndex        =   8
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txt_pegawai 
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
         Left            =   6840
         TabIndex        =   7
         Top             =   1320
         Width           =   4335
      End
      Begin VB.TextBox txt_faktur 
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
         Left            =   6840
         TabIndex        =   5
         Top             =   840
         Width           =   3255
      End
      Begin MSMask.MaskEdBox msk_tgl 
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSMask.MaskEdBox msk_jam 
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_jam1 
         Height          =   375
         Left            =   3120
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   14760
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pencarian"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   1485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pegawai"
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
         Left            =   5040
         TabIndex        =   6
         Top             =   1320
         Width           =   1620
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Faktur"
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
         Left            =   5160
         TabIndex        =   4
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label3 
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
         Left            =   2520
         TabIndex        =   3
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam"
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
         Left            =   480
         TabIndex        =   2
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl."
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
         Left            =   480
         TabIndex        =   1
         Top             =   840
         Width           =   375
      End
   End
   Begin TrueDBGrid60.TDBGrid grd_daftar 
      Height          =   4935
      Left            =   120
      OleObjectBlob   =   "frm_browse_btl.frx":0000
      TabIndex        =   11
      Top             =   2280
      Width           =   15015
   End
End
Attribute VB_Name = "frm_browse_btl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB


Private Sub cmd_batal_Click()

On Error GoTo er_batal

Dim sql, sql1, sql2 As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
    
    If arr_daftar.UpperBound(1) = 0 Then
        Exit Sub
    End If
    
    If MsgBox("Yakinkan data pembatalan yang akan dikembalikan sudah benar" & Chr(13) & "Yakin data yang akan dikembalikan sudah benar", vbYesNo + vbQuestion, "Pesan") = vbNo Then
        Exit Sub
    End If
    
    Dim a As Long
    
    cn.BeginTrans
    
    For a = 1 To arr_daftar.UpperBound(1)
        If arr_daftar(a, 1) <> "" And arr_daftar(a, 1) <> Empty Then
        
            sql = "update tr_faktur_penjualan set ket=0,user_pembatal='-' where no_faktur='" & Trim(arr_daftar(a, 1)) & "' and tgl= datevalue('" & arr_daftar(a, 2) & "') and nama_user='" & Trim(arr_daftar(a, 13)) & "'"
            rs.Open sql, cn
                
             sql2 = "select id from tr_stock where id_barang=" & arr_daftar(a, 0) & " and tgl=datevalue('" & Trim(arr_daftar(a, 2)) & "') and brg_out=" & Trim(arr_daftar(a, 8)) & " and nama_user='" & Trim(arr_daftar(a, 13)) & "' and ket=3"
             rs2.Open sql2, cn
               If Not rs2.EOF Then
                    
                    sql = "update tr_stock set ket=0 where id=" & rs2!id
                    rs.Open sql, cn
                    
                    sql1 = "select id_barang,jml_stock,stock_min from qr_jml_stock where id_barang=" & arr_daftar(a, 0)
                    rs1.Open sql1, cn
                        If Not rs1.EOF Then
                            Dim jml_sekarang As Double
                            jml_sekarang = 0
                            
                            If CDbl(rs1!jml_stock) <= CDbl(rs1!stock_min) Then
                                Dim pn
                                pn = MsgBox("Stock Tidak mencukupi untuk melakukan pembatalan" & Chr(13) & "Stock Min " & rs1!stock_min & Chr(13) & "Jml Stock " & rs1!jml_stock & Chr(13) & "Pembatalan gagal dilakukan")
                                cn.RollbackTrans
                                Exit Sub
                            End If
                            
                            If Not IsNull(rs1!jml_stock) Then
                                jml_sekarang = CDbl(rs1!jml_stock) - CDbl(arr_daftar(a, 8))
                                
                                If CDbl(jml_sekarang) < CDbl(rs1!stock_min) Then
                                    pn = MsgBox("Stock Tidak mencukupi untuk melakukan pembatalan" & Chr(13) & "Stock Min " & rs1!stock_min & Chr(13) & "Jml Stock " & rs1!jml_stock & Chr(13) & "Jumlah yang akan dibatalkan " & arr_daftar(a, 8) & Chr(13) & "Pembatalan gagal dilakukan")
                                    cn.RollbackTrans
                                    Exit Sub
                                End If
                                
                            End If
                            
                            sql = "update tr_jml_stock set jml_stock=" & jml_sekarang & " where id_barang=" & arr_daftar(a, 0)
                            rs.Open sql, cn
                        End If
                    rs1.Close
               End If
             rs2.Close
        End If
    Next a
        
    MsgBox ("Proses pengembalian berhasil dilakukan")
    cn.CommitTrans
    Cmd_Tampil_Click
    Exit Sub
    
er_batal:
        cn.RollbackTrans
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
    
End Sub

Private Sub cmd_cetak_Click()
    On Error GoTo er_printer

    With grd_daftar.PrintInfo
        
        .PageHeaderFont.Bold = True
        .PageHeaderFont.Italic = True
        .PageHeaderFont.Size = 10
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

Private Sub cmd_export_Click()
    On Error Resume Next

    cd.ShowSave
    grd_daftar.ExportToFile cd.FileName, False
End Sub

Private Sub cmd_setup_Click()
On Error GoTo er_setup
    
    With grd_daftar.PrintInfo
        .PageSetup
    End With
    Exit Sub
    
er_setup:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub

Private Sub Cmd_Tampil_Click()

On Error GoTo er_tampil

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        kosong_daftar
        
        sql = "select * from qr_semua_penjualan where ket = 1"
        
        If msk_jam.Text <> "__:__:__" And msk_jam1.Text = "__:__:__" Then
            sql = sql & " and jam = timevalue('" & Trim(msk_jam.Text) & "')"
        End If
        
        If msk_jam1.Text <> "__:__:__" And msk_jam.Text = "__:__:__" Then
            sql = sql & " and jam = timevalue('" & Trim(msk_jam1.Text) & "')"
        End If
        
        If msk_jam.Text <> "__:__:__" And msk_jam1.Text <> "__:__:__" Then
            sql = sql & " and jam >= timevalue('" & Trim(msk_jam.Text) & "') and jam <= timevalue('" & Trim(msk_jam1.Text) & "')"
        End If
        
        If msk_tgl.Text <> "__/__/____" Then
            sql = sql & " and tgl= datevalue('" & Trim(msk_tgl.Text) & "')"
        End If
        
        If txt_faktur.Text <> "" Then
            sql = sql & " and no_faktur='" & Trim(txt_faktur.Text) & "'"
        End If
        
        If txt_pegawai.Text <> "" Then
            sql = sql & " and nama_user like '%" & Trim(txt_pegawai.Text) & "%'"
        End If
        
        sql = sql & " order by tgl,jam,no_faktur"
        rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
                    
                    lanjut rs
            End If
        rs.Close
        
        Exit Sub
        
er_tampil:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub

Private Sub lanjut(rs As Recordset)
    Dim no_faktur, tgl, jam, k_counter, n_counter, k_barang, n_barang, qty, harga, disc, cash, total, user As String
    Dim a As Long
    Dim id_nya, user_pembatal As String
    
    Dim tot_qty, tot_harga, tot_barang As Double
    
        
        a = 1
        tot_qty = 0
        tot_harga = 0
        tot_barang = 0
            Do While Not rs.EOF
                arr_daftar.ReDim 1, a, 0, 15
                grd_daftar.ReBind
                grd_daftar.Refresh
                    
                    id_nya = rs("id_barang")
                    
                    If Not IsNull(rs("no_faktur")) Then
                        no_faktur = rs("no_faktur")
                    Else
                        no_faktur = ""
                    End If
                    
                    If Not IsNull(rs("tgl")) Then
                        tgl = rs("tgl")
                    Else
                        tgl = ""
                    End If
                    
                    If Not IsNull(rs("jam")) Then
                        jam = Format(rs("jam"), "hh:mm:ss")
                    Else
                        jam = ""
                    End If
                    
                    If Not IsNull(rs("kode_counter")) Then
                        k_counter = rs("kode_counter")
                    Else
                        k_counter = ""
                    End If
                    
                    If Not IsNull(rs("nama_counter")) Then
                        n_counter = rs("nama_counter")
                    Else
                        n_counter = ""
                    End If
                    
                    If Not IsNull(rs("kode_barang")) Then
                        k_barang = rs("kode_barang")
                    Else
                        k_barang = ""
                    End If
                    
                    If Not IsNull(rs("nama_barang")) Then
                        n_barang = rs("nama_barang")
                    Else
                        n_barang = ""
                    End If
                    
                    If Not IsNull(rs("qty")) Then
                       qty = rs("qty")
                    Else
                       qty = 0
                    End If
                    
                    If Not IsNull(rs("harga_satuan")) Then
                       harga = rs("harga_satuan")
                    Else
                       harga = 0
                    End If
                    
                    If Not IsNull(rs("disc")) Then
                       disc = rs("disc")
                    Else
                       disc = ""
                    End If
                    
                    If Not IsNull(rs("cash")) Then
                       cash = rs("cash")
                    Else
                       cash = ""
                    End If
                    
                    If Not IsNull(rs("total_harga")) Then
                       total = rs("total_harga")
                    Else
                       total = 0
                    End If
                    
                    If Not IsNull(rs("nama_user")) Then
                       user = rs("nama_user")
                    Else
                       user = ""
                    End If
                    
                    If Not IsNull(rs("user_pembatal")) Then
                       user_pembatal = rs("user_pembatal")
                    Else
                       user_pembatal = ""
                    End If
                    
                    tot_qty = tot_qty + CDbl(qty)
                    tot_harga = tot_harga + CDbl(harga)
                    tot_barang = tot_barang + CDbl(total)
                    
               arr_daftar(a, 0) = id_nya
               arr_daftar(a, 1) = no_faktur
               arr_daftar(a, 2) = tgl
               arr_daftar(a, 3) = jam
               arr_daftar(a, 4) = k_counter
               arr_daftar(a, 5) = n_counter
               arr_daftar(a, 6) = k_barang
               arr_daftar(a, 7) = n_barang
               arr_daftar(a, 8) = qty
               arr_daftar(a, 9) = Format(harga, "###,###,###")
               arr_daftar(a, 10) = disc
               arr_daftar(a, 11) = cash
               arr_daftar(a, 12) = Format(total, "###,###,###")
               arr_daftar(a, 13) = user
               arr_daftar(a, 14) = user_pembatal
               
           a = a + 1
           rs.MoveNext
           Loop
           
           grd_daftar.Columns(7).FooterText = "TOTAL"
           grd_daftar.Columns(8).FooterText = tot_qty
           grd_daftar.Columns(9).FooterText = Format(tot_harga, "###,###,###")
           grd_daftar.Columns(12).FooterText = Format(tot_barang, "###,###,###")
           
           grd_daftar.ReBind
           grd_daftar.Refresh
                    
End Sub

Private Sub Form_Load()

    grd_daftar.Array = arr_daftar
    
    kosong_daftar
    
End Sub

Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub
