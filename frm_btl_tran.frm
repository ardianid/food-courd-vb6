VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_btl_tran 
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   12600
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   14
      Top             =   7560
      Width           =   2415
      Begin VB.CommandButton cmd_ok 
         Caption         =   "OK"
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
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   825
      ScaleWidth      =   12345
      TabIndex        =   13
      Top             =   7560
      Width           =   12375
   End
   Begin TrueDBGrid60.TDBGrid grd_daftar 
      Height          =   5415
      Left            =   120
      OleObjectBlob   =   "frm_btl_tran.frx":0000
      TabIndex        =   12
      Top             =   2040
      Width           =   14895
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1695
      ScaleWidth      =   14895
      TabIndex        =   0
      Top             =   240
      Width           =   14895
      Begin VB.Frame Frame1 
         Height          =   135
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   14415
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
         Left            =   13320
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txt_pegawai 
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
         Left            =   6840
         TabIndex        =   10
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox txt_faktur 
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
         Left            =   6840
         TabIndex        =   8
         Top             =   600
         Width           =   3255
      End
      Begin MSMask.MaskEdBox msk_tgl 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
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
      Begin MSMask.MaskEdBox msk_jam 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         Left            =   3240
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
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
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pegawai"
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
         Left            =   5040
         TabIndex        =   9
         Top             =   1080
         Width           =   1440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Faktur"
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
         Left            =   5160
         TabIndex        =   7
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S/d"
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
         Left            =   2520
         TabIndex        =   5
         Top             =   1080
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam"
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
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl."
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
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   375
      End
   End
End
Attribute VB_Name = "frm_btl_tran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB

Private Sub cmd_ok_Click()
On Error GoTo er_ok
    If arr_daftar.UpperBound(1) > 0 Then
        Dim sql, sql1, sql2 As String
        Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset
        Dim a As Long
            
            If MsgBox("Yakinkan data yang akan diproses sudah benar" & Chr(13) & "Yakin data sudah benar", vbYesNo + vbQuestion, "Pesan") = vbNo Then
                Exit Sub
            End If
            
            cn.BeginTrans
            For a = 1 To arr_daftar.UpperBound(1)
            
                sql = "update tr_faktur_penjualan set ket=1,user_pembatal='" & Trim(utama.lbl_user.Caption) & "' where no_faktur='" & Trim(arr_daftar(a, 1)) & "'"
                rs.Open sql, cn
                
                sql1 = "select id from tr_stock where id_barang=" & arr_daftar(a, 0) & " and tgl=datevalue('" & Trim(arr_daftar(a, 2)) & "') and brg_out=" & Trim(arr_daftar(a, 8)) & " and nama_user='" & Trim(arr_daftar(a, 13)) & "'"
                rs1.Open sql1, cn
                    If Not rs1.EOF Then
                        sql = "update tr_stock set ket=3 where id=" & rs1("id")
                        rs.Open sql, cn
                        
                        sql2 = "select * from tr_jml_stock where id_barang=" & arr_daftar(a, 0)
                        rs2.Open sql2, cn
                            If Not rs2.EOF Then
                                Dim jml As Double
                                    jml = CDbl(rs2("jml_stock")) + CDbl(arr_daftar(a, 8))
                                    
                                    sql = "update tr_jml_stock set jml_stock=" & jml & " where id_barang=" & arr_daftar(a, 0)
                                    rs.Open sql, cn
                            Else
                                    MsgBox ("Stock tidak ditemukan")
                                    cn.RollbackTrans
                                    Exit Sub
                            End If
                        rs2.Close
                        
                    End If
                rs1.Close
           Next a
           MsgBox ("Transaksi pembatan Berhasil dilakukan")
           cn.CommitTrans
           msk_tgl.Text = "__/__/____"
           msk_jam.Text = "__:__:__"
           msk_jam1.Text = "__:__:__"
           txt_faktur.Text = ""
           txt_pegawai.Text = ""
           kosong_daftar
           msk_tgl.SetFocus
    End If
    Exit Sub
    
er_ok:
    cn.RollbackTrans
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub Cmd_Tampil_Click()

On Error GoTo er_tampil

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        kosong_daftar
        
        sql = "select * from qr_semua_penjualan where ket <> 1"
        
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
    Dim id_nya As String
        
    Dim tot_qty, tot_harga, tot_total As Double
    
        tot_qty = 0
        tot_harga = 0
        tot_total = 0
        a = 1
            Do While Not rs.EOF
                arr_daftar.ReDim 1, a, 0, 14
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
                    
                    tot_qty = tot_qty + CDbl(qty)
                    tot_harga = tot_harga + CDbl(harga)
                    tot_total = tot_total + CDbl(total)
                    
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
               
           a = a + 1
           rs.MoveNext
           Loop
           
           grd_daftar.Columns(7).FooterText = "TOTAL"
           grd_daftar.Columns(8).FooterText = tot_qty
           grd_daftar.Columns(9).FooterText = Format(tot_harga, "###,###,###")
           grd_daftar.Columns(12).FooterText = Format(tot_total, "###,###,###")
           
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
