VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_browse_pnj 
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   120
      ScaleHeight     =   8265
      ScaleWidth      =   14985
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin MSComDlg.CommonDialog cd 
         Left            =   840
         Top             =   4440
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
         Left            =   13320
         TabIndex        =   23
         Top             =   2160
         Width           =   1455
      End
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
         Left            =   10200
         TabIndex        =   22
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
         Left            =   11760
         TabIndex        =   21
         Top             =   7680
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
         Left            =   13320
         TabIndex        =   20
         Top             =   7680
         Width           =   1455
      End
      Begin TrueDBGrid60.TDBGrid grd_penjualan 
         Height          =   4815
         Left            =   240
         OleObjectBlob   =   "frm_browse_pnj.frx":0000
         TabIndex        =   19
         Top             =   2760
         Width           =   14535
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   7920
         TabIndex        =   14
         Top             =   1440
         Width           =   6855
         Begin MSMask.MaskEdBox msk_tgl1 
            Height          =   375
            Left            =   960
            TabIndex        =   16
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
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
            Left            =   3240
            TabIndex        =   18
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
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
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S/d"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2760
            TabIndex        =   17
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl."
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
            Left            =   360
            TabIndex        =   15
            Top             =   240
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1215
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   6735
         Begin VB.TextBox txt_kasir 
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
            Height          =   405
            Left            =   1920
            TabIndex        =   25
            Top             =   720
            Width           =   4575
         End
         Begin VB.TextBox txt_no_faktur 
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
            Height          =   405
            Left            =   1920
            TabIndex        =   13
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Kasir"
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
            Left            =   360
            TabIndex        =   24
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Faktur"
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
            Left            =   360
            TabIndex        =   12
            Top             =   360
            Width           =   945
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Barang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   7920
         TabIndex        =   6
         Top             =   120
         Width           =   6855
         Begin VB.TextBox txt_kode_barang 
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
            Height          =   405
            Left            =   1800
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txt_nama_barang 
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
            Height          =   405
            Left            =   1800
            TabIndex        =   7
            Top             =   720
            Width           =   2895
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Barang"
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
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Barang"
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
            Left            =   360
            TabIndex        =   9
            Top             =   840
            Width           =   1350
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Jenis Barang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   6735
         Begin VB.TextBox txt_nama_counter 
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
            Height          =   405
            Left            =   1920
            TabIndex        =   5
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox txt_kode_counter 
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
            Height          =   405
            Left            =   1920
            TabIndex        =   3
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
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
            Left            =   360
            TabIndex        =   4
            Top             =   840
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode"
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
            Left            =   360
            TabIndex        =   2
            Top             =   360
            Width           =   510
         End
      End
   End
End
Attribute VB_Name = "frm_browse_pnj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_penjualan As New XArrayDB

Private Sub cmd_cetak_Click()
On Error GoTo er_printer

    With grd_penjualan.PrintInfo
        
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
    grd_penjualan.ExportToFile cd.FileName, False
    
End Sub

Private Sub cmd_setup_Click()

On Error GoTo er_setup

    With grd_penjualan.PrintInfo
        .PageSetup
    End With
    Exit Sub
    
er_setup:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub Cmd_Tampil_Click()
    isi
End Sub

Private Sub Form_Load()

    grd_penjualan.Array = arr_penjualan
    
    kosong_penjualan
    
End Sub

Private Sub kosong_penjualan()
    arr_penjualan.ReDim 0, 0, 0, 0
    grd_penjualan.ReBind
    grd_penjualan.Refresh
End Sub

Private Sub isi()
On Error GoTo er_handler

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        kosong_penjualan
        
    sql = "select * from qr_semua_penjualan where ket=0"
        
        If txt_kode_counter.Text <> "" Or txt_nama_counter.Text <> "" _
            Or Txt_Kode_Barang.Text <> "" Or txt_nama_barang.Text <> "" Or txt_no_faktur.Text <> "" Or txt_kasir.Text <> "" _
                Or msk_tgl1.Text <> "__/__/____" Or msk_tgl2.Text <> "__/__/____" Then
                
         ' kode counter
            If txt_kode_counter.Text <> "" Then
                sql = sql & " and kode_counter like '%" & Trim(txt_kode_counter.Text) & "%'"
            End If
         ' nama counter
            If txt_nama_counter.Text <> "" And txt_kode_counter.Text = "" Then
                sql = sql & " and nama_counter like '%" & Trim(txt_nama_counter.Text) & "%'"
            End If
                
                If txt_nama_counter.Text <> "" And txt_kode_counter.Text <> "" Then
                    sql = sql & " and nama_counter like '%" & Trim(txt_nama_counter.Text) & "%'"
                End If
         ' kode barang
            If Txt_Kode_Barang.Text <> "" And txt_kode_counter.Text = "" And txt_nama_counter.Text = "" Then
                sql = sql & " and kode_barang like '%" & Trim(Txt_Kode_Barang.Text) & "%'"
            End If
                
                If Txt_Kode_Barang.Text <> "" And (txt_kode_counter.Text <> "" Or txt_nama_counter.Text <> "") Then
                    sql = sql & " and kode_barang like '%" & Trim(Txt_Kode_Barang.Text) & "%'"
                End If
         '  Nama barang
            If txt_nama_barang.Text <> "" And txt_kode_counter.Text = "" And txt_nama_counter.Text = "" And Txt_Kode_Barang.Text = "" Then
                sql = sql & " and nama_barang like '%" & Trim(txt_nama_barang.Text) & "%'"
            End If
                
                If txt_nama_barang.Text <> "" And (txt_kode_counter.Text <> "" Or txt_nama_counter.Text <> "" Or Txt_Kode_Barang.Text <> "") Then
                    sql = sql & " and nama_barang like '%" & Trim(txt_nama_barang.Text) & "%'"
                End If
         '  No faktur
            If txt_no_faktur.Text <> "" And txt_kode_counter.Text = "" And txt_nama_counter.Text = "" And Txt_Kode_Barang.Text = "" And txt_nama_barang.Text = "" Then
                sql = sql & " and no_faktur like '%" & Trim(txt_no_faktur.Text) & "%'"
            End If
                
                If txt_no_faktur.Text <> "" And (txt_kode_counter.Text <> "" Or txt_nama_counter.Text <> "" Or Txt_Kode_Barang.Text <> "" Or txt_nama_barang.Text <> "") Then
                    sql = sql & " and no_faktur like '%" & Trim(txt_no_faktur.Text) & "%'"
                End If
         'Nama Kasir
            If txt_kasir.Text <> "" And txt_no_faktur.Text = "" And txt_kode_counter.Text = "" And txt_nama_counter.Text = "" And Txt_Kode_Barang.Text = "" And txt_nama_barang.Text = "" Then
                sql = sql & " and nama_user like '%" & Trim(txt_kasir.Text) & "%'"
            End If
                
                If txt_kasir.Text <> "" And (txt_no_faktur.Text <> "" Or txt_kode_counter.Text <> "" Or txt_nama_counter.Text <> "" Or Txt_Kode_Barang.Text <> "" Or txt_nama_barang.Text <> "") Then
                    sql = sql & " and nama_user like '%" & Trim(txt_kasir.Text) & "%'"
                End If
                
         '  Tgl awal
            If msk_tgl1.Text <> "__/__/____" And msk_tgl2.Text = "__/__/____" And txt_kode_counter.Text = "" And txt_nama_counter.Text = "" And Txt_Kode_Barang.Text = "" And txt_nama_barang.Text = "" And txt_no_faktur.Text = "" And txt_kasir.Text = "" Then
                sql = sql & " and tgl=datevalue('" & msk_tgl1.Text & "')"
            End If
            
                If msk_tgl1.Text <> "__/__/____" And msk_tgl2.Text = "__/__/____" And (txt_kode_counter.Text <> "" Or txt_nama_counter.Text <> "" Or Txt_Kode_Barang.Text <> "" Or txt_nama_barang.Text <> "" Or txt_no_faktur.Text <> "" Or txt_kasir.Text <> "") Then
                    sql = sql & " and tgl=datevalue('" & msk_tgl1.Text & "')"
                 End If
            
         '  Tgl akhir
            If msk_tgl2.Text <> "__/__/____" And txt_kode_counter.Text = "" And txt_nama_counter.Text = "" And Txt_Kode_Barang.Text = "" And txt_nama_barang.Text = "" And txt_no_faktur.Text = "" And msk_tgl1.Text = "__/__/____" And txt_kasir.Text = "" Then
                sql = sql & " and tgl=DateValue('" & msk_tgl2.Text & "') "
            End If
                
                If msk_tgl2.Text <> "__/__/____" And msk_tgl1.Text = "__/__/____" And (txt_kode_counter.Text <> "" Or txt_nama_counter.Text <> "" Or Txt_Kode_Barang.Text <> "" Or txt_nama_barang.Text <> "" Or txt_no_faktur.Text <> "" Or txt_kasir.Text <> "") Then
                    sql = sql & " and tgl=DateValue('" & msk_tgl2.Text & "') "
                End If
                
         ' Tgl awal & akhir
            If msk_tgl1.Text <> "__/__/____" And msk_tgl2.Text <> "__/__/____" Then
                
                If txt_kode_counter.Text = "" And txt_nama_counter.Text = "" And Txt_Kode_Barang.Text = "" And txt_nama_barang.Text = "" And txt_no_faktur.Text = "" And txt_kasir.Text = "" Then
                    sql = sql & " and tgl >= DateValue('" & msk_tgl1.Text & "')  and tgl <= DateValue('" & msk_tgl2.Text & "')"
                End If
                
                If txt_kode_counter.Text <> "" Or txt_nama_counter.Text <> "" Or Txt_Kode_Barang.Text <> "" Or txt_nama_barang.Text <> "" Or txt_no_faktur.Text <> "" Or txt_kasir.Text <> "" Then
                    sql = sql & " and tgl >= DateValue('" & msk_tgl1.Text & "')  and tgl <=  DateValue('" & msk_tgl2.Text & "' ) "
                End If
            End If
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

er_handler:
        
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub

Private Sub lanjut(rs As Recordset)
    Dim tgl, jam, no_faktur, kode_counter, nama_counter As String
    Dim kode_barang, nama_barang, qty, harga_satuan, harga_sebenarnya As String
    Dim user As String
    Dim disc, cash, total_harga As String
    Dim total_qty, total_harga_satuan, total_disc, total_cash, total_semua, total_sebenarnya As Double
    Dim a, b As String
        
        a = 1
        b = 1
        total_qty = 0
        total_harga_satuan = 0
        total_sebenarnya = 0
        total_disc = 0
        total_cash = 0
        total_semua = 0
        
            Do While Not rs.EOF
                arr_penjualan.ReDim 1, a, 0, 15
                grd_penjualan.ReBind
                grd_penjualan.Refresh
                    
                    If Not IsNull(rs("tgl")) Then
                        tgl = rs("tgl")
                    Else
                        tgl = ""
                    End If
                    
                    If Not IsNull(rs("jam")) Then
                        jam = rs("jam")
                    Else
                        jam = ""
                    End If
                    
                    If Not IsNull(rs("no_faktur")) Then
                        no_faktur = rs("no_faktur")
                    Else
                        no_faktur = ""
                    End If
                    
                    If Not IsNull(rs("kode_counter")) Then
                        kode_counter = rs("kode_counter")
                    Else
                        kode_counter = ""
                    End If
                    
                    If Not IsNull(rs("nama_counter")) Then
                        nama_counter = rs("nama_counter")
                    Else
                        nama_counter = ""
                    End If
                    
                    If Not IsNull(rs("kode_barang")) Then
                        kode_barang = rs("kode_barang")
                    Else
                        kode_barang = ""
                    End If
                    
                    If Not IsNull(rs("nama_barang")) Then
                        nama_barang = rs("nama_barang")
                    Else
                        nama_barang = ""
                    End If
                    
                    If Not IsNull(rs("qty")) Then
                        qty = rs("qty")
                    Else
                        qty = 0
                    End If
                    
                    If Not IsNull(rs("harga_satuan")) Then
                        harga_satuan = rs("harga_satuan")
                    Else
                        harga_satuan = 0
                    End If
                    
                    If Not IsNull(rs("harga_sebenarnya")) Then
                        harga_sebenarnya = rs("harga_sebenarnya")
                    Else
                        harga_sebenarnya = 0
                    End If
                    
                    If Not IsNull(rs("cash")) Then
                        cash = rs("cash")
                    Else
                        cash = 0
                    End If
                    
                    If Not IsNull(rs("disc")) Then
                        disc = rs("disc")
                    Else
                        disc = 0
                    End If
                    
                    If Not IsNull(rs("total_harga")) Then
                        total_harga = rs("total_harga")
                    Else
                        total_harga = 0
                    End If
                    
                    If Not IsNull(rs("nama_user")) Then
                        user = rs("nama_user")
                    Else
                        user = ""
                    End If
                
                If a > 1 Then
                    If no_faktur <> arr_penjualan(a - 1, 3) Then
                        b = b + 1
                    End If
                End If
                
              
                arr_penjualan(a, 0) = b
                arr_penjualan(a, 1) = tgl
                arr_penjualan(a, 2) = jam
                arr_penjualan(a, 3) = no_faktur
                arr_penjualan(a, 4) = kode_counter
                arr_penjualan(a, 5) = nama_counter
                arr_penjualan(a, 6) = kode_barang
                arr_penjualan(a, 7) = nama_barang
                arr_penjualan(a, 8) = Format(harga_satuan, "###,###,###")
                arr_penjualan(a, 9) = qty
                arr_penjualan(a, 10) = Format(harga_sebenarnya, "###,###,###")
                arr_penjualan(a, 11) = disc
                arr_penjualan(a, 12) = cash
                arr_penjualan(a, 13) = Format(total_harga, "###,###,###")
                arr_penjualan(a, 14) = user
                
            total_qty = CDbl(qty) + total_qty
            total_harga_satuan = CDbl(harga_satuan) + total_harga_satuan
            total_sebenarnya = CDbl(harga_sebenarnya) + total_sebenarnya
           If disc <> 0 Then
            total_disc = CDbl(Mid(disc, 1, Len(disc) - 1)) + total_disc
           End If
           If cash <> 0 Then
            total_cash = CDbl(Mid(cash, 1, Len(cash) - 1)) + total_cash
           End If
            total_semua = CDbl(total_harga) + total_semua
                    
            a = a + 1
            rs.MoveNext
            Loop
            
            grd_penjualan.Columns(1).FooterText = "TOTAL"
            grd_penjualan.Columns(8).FooterText = Format(total_harga_satuan, "###,###,###")
            grd_penjualan.Columns(9).FooterText = total_qty
            grd_penjualan.Columns(10).FooterText = Format(total_sebenarnya, "###,###,###")
            grd_penjualan.Columns(11).FooterText = total_disc & "%"
            grd_penjualan.Columns(12).FooterText = total_cash & "%"
            grd_penjualan.Columns(13).FooterText = Format(total_semua, "###,###,###")
            
            grd_penjualan.ReBind
            grd_penjualan.Refresh
                
                
                
End Sub

Private Sub msk_tgl1_GotFocus()
    msk_tgl1.SelStart = 0
    msk_tgl1.SelLength = Len(msk_tgl1)
End Sub
Private Sub msk_tgl2_GotFocus()
    msk_tgl2.SelStart = 0
    msk_tgl2.SelLength = Len(msk_tgl2)
End Sub
Private Sub txt_kasir_GotFocus()
    txt_kasir.SelStart = 0
    txt_kasir.SelLength = Len(txt_kasir)
End Sub

Private Sub txt_kode_barang_GotFocus()
    Txt_Kode_Barang.SelStart = 0
    Txt_Kode_Barang.SelLength = Len(Txt_Kode_Barang)
End Sub

Private Sub txt_kode_counter_GotFocus()
    txt_kode_counter.SelStart = 0
    txt_kode_counter.SelLength = Len(txt_kode_counter)
End Sub

Private Sub txt_nama_barang_GotFocus()
    txt_nama_barang.SelStart = 0
    txt_nama_barang.SelLength = Len(txt_nama_barang)
End Sub

Private Sub txt_nama_counter_GotFocus()
    txt_nama_counter.SelStart = 0
    txt_nama_counter.SelLength = Len(txt_nama_counter)
End Sub

Private Sub txt_no_faktur_GotFocus()
    txt_no_faktur.SelStart = 0
    txt_no_faktur.SelLength = Len(txt_no_faktur)
End Sub
