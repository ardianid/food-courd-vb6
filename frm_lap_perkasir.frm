VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_lap_perkasir 
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   120
      ScaleHeight     =   8385
      ScaleWidth      =   14985
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin MSComDlg.CommonDialog cd 
         Left            =   8160
         Top             =   5040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
         Top             =   1320
         Width           =   1455
      End
      Begin VB.PictureBox pic_karyawan 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7215
         Left            =   2400
         ScaleHeight     =   7185
         ScaleWidth      =   6465
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   6495
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   0
            ScaleHeight     =   345
            ScaleWidth      =   6465
            TabIndex        =   17
            Top             =   0
            Width           =   6495
            Begin VB.CommandButton cmd_x 
               Caption         =   "x"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   6000
               TabIndex        =   16
               Top             =   0
               Width           =   495
            End
         End
         Begin VB.TextBox txt_cari 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   6255
         End
         Begin TrueDBGrid60.TDBGrid grd_karyawan 
            Height          =   5655
            Left            =   120
            OleObjectBlob   =   "frm_lap_perkasir.frx":0000
            TabIndex        =   15
            Top             =   1320
            Width           =   6255
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Nama Karyawan"
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
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   6255
         End
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
         Left            =   11880
         TabIndex        =   7
         Top             =   7800
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
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
         Left            =   10320
         TabIndex        =   6
         Top             =   7800
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
         Left            =   13440
         TabIndex        =   8
         Top             =   7800
         Width           =   1455
      End
      Begin TrueDBGrid60.TDBGrid grd_daftar 
         Height          =   5775
         Left            =   120
         OleObjectBlob   =   "frm_lap_perkasir.frx":26EE
         TabIndex        =   5
         Top             =   1920
         Width           =   14775
      End
      Begin VB.TextBox txt_nama 
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
         Left            =   1920
         TabIndex        =   3
         Top             =   1320
         Width           =   4695
      End
      Begin MSMask.MaskEdBox msk_tgl1 
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
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
         Left            =   4680
         TabIndex        =   2
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Kasir"
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
         TabIndex        =   12
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
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
         Left            =   4080
         TabIndex        =   11
         Top             =   840
         Width           =   315
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         X1              =   360
         X2              =   14640
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kriteria Pencarian"
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
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   2970
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
         TabIndex        =   9
         Top             =   840
         Width           =   375
      End
   End
End
Attribute VB_Name = "frm_lap_perkasir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB
Dim arr_karyawan As New XArrayDB
Dim iid As String


Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub

Private Sub kosong_karyawan()
    arr_karyawan.ReDim 0, 0, 0, 0
    grd_karyawan.ReBind
    grd_karyawan.Refresh
End Sub

Private Sub isi_karyawan()

On Error GoTo er_k

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        kosong_karyawan
        
        sql = "select id,nama_karyawan from tbl_karyawan order by nama_karyawan"
        rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                rs.MoveLast
                rs.MoveFirst
                    
                    lanjut_k rs
            End If
       rs.Close
       
       Exit Sub
       
er_k:
       Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub

Private Sub lanjut_k(rs As Recordset)
    Dim a As Long
    Dim id_k, nama As String
        
        a = 1
            Do While Not rs.EOF
                arr_karyawan.ReDim 1, a, 0, 2
                grd_karyawan.ReBind
                grd_karyawan.Refresh
                    
                    id_k = rs!id
                    
                    If Not IsNull(rs!nama_karyawan) Then
                        nama = rs!nama_karyawan
                    End If
                    
               arr_karyawan(a, 0) = id_k
               arr_karyawan(a, 1) = nama
            a = a + 1
            rs.MoveNext
            Loop
            
            grd_karyawan.ReBind
            grd_karyawan.Refresh
            
            
End Sub

Private Sub cmd_cetak_Click()

    On Error GoTo er_printer

    With grd_daftar.PrintInfo
        
        .PageFooterFont.Name = "Arial"
        .PageHeaderFont.Size = 12
        .PageHeader = "LAPORAN PENJUALAN PERKASIR \t\t Periode  : " & Trim(msk_tgl1.Text) & " s/d " & Trim(msk_tgl2.Text)
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

Private Sub cmd_tampil_Click()

On Error GoTo err_tampil

Dim sql As String
Dim rs As New ADODB.Recordset
Dim a, b As Long
Dim tgl, jam, faktur, kode_counter, nama_counter, kode_barang, nama_barang, qty, harga_satuan, cash, disc, total_harga As String
Dim jml_qty, jml_harga_satuan, jml_disc, jml_cash, jml_total As Double
    
    If msk_tgl1.Text = "__/__/____" Or msk_tgl2.Text = "__/__/____" Or txt_nama.Text = "" Then
        MsgBox ("Semua data hrs diisi")
        Exit Sub
    End If
    
    kosong_daftar
    
    grd_daftar.Caption = "NAMA KASIR : " & UCase(Trim(txt_nama.Text))
    
    sql = "select kode_counter,nama_counter,kode_barang,nama_barang,no_faktur,tgl,jam,qty,harga_satuan,disc,cash,total_harga,nama_user from qr_penjualan_sebenarnya"
     
    If msk_tgl1.Text <> "__/__/____" And msk_tgl2.Text <> "__/__/____" Then
        
        sql = sql & " where"
        
        sql = sql & " tgl >= datevalue('" & Trim(msk_tgl1.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl2.Text) & "')"
        
            If txt_nama.Text <> "" Then
                sql = sql & " and nama_user='" & Trim(txt_nama.Text) & "'"
            End If
    End If
    
    sql = sql & " order by tgl,jam"
    rs.Open sql, cn, adOpenKeyset
        If Not rs.EOF Then
            
            rs.MoveLast
            rs.MoveFirst
                
                a = 1
                b = 1
                jml_qty = 0
                jml_harga_satuan = 0
                jml_disc = 0
                jml_cash = 0
                jml_total = 0
                
                Do While Not rs.EOF
                    arr_daftar.ReDim 1, a, 0, 15
                    grd_daftar.ReBind
                    grd_daftar.Refresh
                         
                    If Not IsNull(rs!tgl) Then
                        tgl = rs!tgl
                    Else
                        tgl = ""
                    End If
                        
                    If Not IsNull(rs!jam) Then
                        jam = rs!jam
                    Else
                        jam = ""
                    End If
                    
                    If Not IsNull(rs!no_faktur) Then
                        faktur = rs!no_faktur
                    Else
                        faktur = ""
                    End If
                    
                    If Not IsNull(rs!kode_counter) Then
                        kode_counter = rs!kode_counter
                    Else
                        kode_counter = ""
                    End If
                    
                    If Not IsNull(rs!nama_counter) Then
                        nama_counter = rs!nama_counter
                    Else
                        nama_counter = ""
                    End If
                    
                    If Not IsNull(rs!kode_barang) Then
                        kode_barang = rs!kode_barang
                    Else
                        kode_barang = ""
                    End If
                    
                    If Not IsNull(rs!nama_barang) Then
                        nama_barang = rs!nama_barang
                    Else
                        nama_barang = ""
                    End If
                    
                    If Not IsNull(rs!qty) Then
                        qty = rs!qty
                    Else
                        qty = 0
                    End If
                    
                    If Not IsNull(rs!harga_satuan) Then
                        harga_satuan = rs!harga_satuan
                    Else
                        harga_satuan = 0
                    End If
                    
                    If Not IsNull(rs!disc) Then
                        disc = rs!disc
                    Else
                        disc = 0
                    End If
                    
                    If Not IsNull(rs!cash) Then
                        cash = rs!cash
                    Else
                        cash = 0
                    End If
                    
                    If Not IsNull(rs!total_harga) Then
                        total_harga = rs!total_harga
                    Else
                        total_harga = 0
                    End If
                    
                    If a > 1 Then
                        If faktur <> arr_daftar(a, 3) Then
                            b = b + 1
                        End If
                    End If
                Dim disc_b, cash_b
                    disc_b = Mid(disc, 1, Len(disc) - 1)
                    cash_b = Mid(cash, 1, Len(cash) - 1)
                    
                jml_qty = CDbl(jml_qty) + CDbl(qty)
                jml_harga_satuan = CDbl(jml_harga_satuan) + CDbl(harga_satuan)
                jml_disc = CDbl(jml_disc) + CDbl(disc_b)
                jml_cash = CDbl(jml_cash) + CDbl(cash_b)
                jml_total = CDbl(jml_total) + CDbl(total_harga)
                    
                arr_daftar(a, 0) = b
                arr_daftar(a, 1) = tgl
                arr_daftar(a, 2) = jam
                arr_daftar(a, 3) = faktur
                arr_daftar(a, 4) = kode_counter
                arr_daftar(a, 5) = nama_counter
                arr_daftar(a, 6) = kode_barang
                arr_daftar(a, 7) = nama_barang
                arr_daftar(a, 8) = qty
                arr_daftar(a, 9) = Format(harga_satuan, "###,###,###")
                arr_daftar(a, 10) = disc
                arr_daftar(a, 11) = cash
                arr_daftar(a, 12) = Format(total_harga, "###,###,###")
           a = a + 1
           rs.MoveNext
           Loop
           
           grd_daftar.Columns(7).FooterText = "TOTAL"
           grd_daftar.Columns(8).FooterText = jml_qty
           grd_daftar.Columns(9).FooterText = Format(jml_harga_satuan, "###,###,###")
           grd_daftar.Columns(10).FooterText = jml_disc & "%"
           grd_daftar.Columns(11).FooterText = jml_cash & "%"
           grd_daftar.Columns(12).FooterText = Format(jml_total, "###,###,###")
           
           grd_daftar.ReBind
           grd_daftar.Refresh
       End If
     rs.Close
Exit Sub

err_tampil:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
                
End Sub

Private Sub cmd_x_Click()
    pic_karyawan.Visible = False
    txt_nama.SetFocus
End Sub

Private Sub Command1_Click()
On Error GoTo err_page
    
    With grd_daftar.PrintInfo
        .PageSetup
    End With
    Exit Sub
    
err_page:
        
        Dim p
            p = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub

Private Sub Form_Activate()
    msk_tgl1.SetFocus
End Sub

Private Sub Form_Load()
    grd_daftar.Array = arr_daftar
    
    kosong_daftar
    
    grd_karyawan.Array = arr_karyawan
    
    isi_karyawan
    
End Sub

Private Sub grd_karyawan_Click()
On Error Resume Next
    If arr_karyawan.UpperBound(1) > 0 Then
        iid = arr_karyawan(grd_karyawan.Bookmark, 0)
    End If
End Sub

Private Sub grd_karyawan_DblClick()
    If arr_karyawan.UpperBound(1) > 0 Then
        txt_nama.Text = arr_karyawan(grd_karyawan.Bookmark, 1)
        pic_karyawan.Visible = False
        txt_nama.SetFocus
    End If
End Sub

Private Sub grd_karyawan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_karyawan.Visible = False
        txt_nama.SetFocus
    End If
    
    If KeyCode = 13 Then
        grd_karyawan_DblClick
    End If
End Sub

Private Sub grd_karyawan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_karyawan_Click
End Sub
Private Sub msk_tgl1_GotFocus()
    msk_tgl1.SelStart = 0
    msk_tgl1.SelLength = Len(msk_tgl1)
End Sub
Private Sub msk_tgl2_GotFocus()
    msk_tgl2.SelStart = 0
    msk_tgl2.SelLength = Len(msk_tgl2)
End Sub

Private Sub pic_karyawan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_karyawan.Visible = False
        txt_nama.SetFocus
    End If
End Sub

Private Sub txt_cari_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        grd_karyawan_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        pic_karyawan.Visible = False
        txt_nama.SetFocus
    End If
    
End Sub

Private Sub txt_cari_KeyUp(KeyCode As Integer, Shift As Integer)
    
On Error GoTo cari
    
    kosong_karyawan
    
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
        sql = "select id,nama_karyawan from tbl_karyawan"
            
            If txt_cari.Text <> "" Then
                sql = sql & " where nama_karyawan like '%" & Trim(txt_cari.Text) & "%'"
            End If
            
       sql = sql & " order by nama_karyawan"
       
       rs.Open sql, cn, adOpenKeyset
        If Not rs.EOF Then
            
            rs.MoveLast
            rs.MoveFirst
                
                lanjut_k rs
        End If
      rs.Close
      
      Exit Sub
      
cari:
      Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
      
End Sub

Private Sub txt_nama_GotFocus()
    txt_nama.SelStart = 0
    txt_nama.SelLength = Len(txt_nama)
End Sub

Private Sub txt_nama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txt_nama.Text = ""
        txt_cari.Text = ""
        pic_karyawan.Visible = True
        txt_cari.SetFocus
    End If
End Sub
