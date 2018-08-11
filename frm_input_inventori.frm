VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frm_input_inventori 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form Input Inventori"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   9975
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   120
      ScaleHeight     =   7425
      ScaleWidth      =   9705
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.CommandButton Cetak 
         Caption         =   "&Preview"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8400
         TabIndex        =   17
         Top             =   6840
         Width           =   1095
      End
      Begin VB.CommandButton Atur 
         Caption         =   "&Page Setup"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         TabIndex        =   16
         Top             =   6840
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Height          =   2535
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   9495
         Begin VB.Frame Frame1 
            Height          =   975
            Left            =   4920
            TabIndex        =   11
            Top             =   1440
            Width           =   4455
            Begin VB.CommandButton Cmd_Baru 
               Caption         =   "&Baru"
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
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton cmd_keluar 
               Caption         =   "&Keluar"
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
               Left            =   3360
               TabIndex        =   13
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton cmd_ref 
               Caption         =   "&Refresh"
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
               Left            =   2280
               TabIndex        =   12
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton cmd_simpan 
               Caption         =   "&Simpan"
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
               Left            =   1200
               TabIndex        =   14
               Top             =   240
               Width           =   975
            End
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
            MaxLength       =   10
            TabIndex        =   6
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txt_nama 
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
            TabIndex        =   5
            Top             =   840
            Width           =   3495
         End
         Begin VB.TextBox txt_stock 
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
            TabIndex        =   3
            Top             =   1320
            Width           =   1095
         End
         Begin TDBNumber6Ctl.TDBNumber TDB_Harga 
            Height          =   375
            Left            =   2160
            TabIndex        =   4
            Top             =   1800
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   661
            Calculator      =   "frm_input_inventori.frx":0000
            Caption         =   "frm_input_inventori.frx":0020
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frm_input_inventori.frx":008C
            Keys            =   "frm_input_inventori.frx":00AA
            Spin            =   "frm_input_inventori.frx":00F4
            AlignHorizontal =   1
            AlignVertical   =   2
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###;;0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999
            MinValue        =   -999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1
            Value           =   0
            MaxValueVT      =   1028849669
            MinValueVT      =   1598423045
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Barang"
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
            Index           =   2
            Left            =   240
            TabIndex        =   10
            Top             =   1800
            Width           =   1350
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Inventory"
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
            TabIndex        =   9
            Top             =   360
            Width           =   1530
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Inventory"
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
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   840
            Width           =   1620
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stock awal"
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
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   1320
            Width           =   1020
         End
      End
      Begin TrueOleDBGrid60.TDBGrid grd_invent 
         Height          =   4095
         Left            =   240
         OleObjectBlob   =   "frm_input_inventori.frx":011C
         TabIndex        =   1
         Top             =   2640
         Width           =   9255
      End
   End
End
Attribute VB_Name = "frm_input_inventori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql_tmp_invent As String
Dim rs_tmp_invent As New ADODB.Recordset
Dim konfirm As String
Dim arr_invent As New XArrayDB

Private Sub Atur_Click()

On Error GoTo er_page
    
    With grd_invent.PrintInfo
        .PageSetup
    End With
    Exit Sub
    
er_page:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear

End Sub

Private Sub Cetak_Click()

On Error GoTo er_printer

    With grd_invent.PrintInfo
        
        .PageFooterFont.Name = "Arial"
        .PageHeaderFont.Size = 12
        .PageHeader = "LAPORAN DATA INVENTORI" '"TOTAL JUMLAH PENJUALAN \t\tPeriode  : " & Trim(msk_tgl1.Text) & " s/d " & Trim(msk_tgl2.Text)
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

Private Sub cmd_baru_Click()
    txt_kode.Text = ""
    txt_nama.Text = ""
    txt_stock.Text = 0
    TDB_Harga.Value = Null
    txt_kode.Enabled = True
    txt_kode.SetFocus
End Sub

Private Sub cmd_keluar_Click()
    Unload Me
End Sub

Private Sub cmd_ref_Click()
    proc_tmp_invent
End Sub

Private Sub cmd_simpan_Click()

On Error GoTo er_simpan

    Dim sql_smp_invent As String
    Dim rs_smp_invent As New ADODB.Recordset
    Dim id_invent As String
    
    Dim Harg As Double
    
    If TDB_Harga.ValueIsNull Then
        Harg = 0
    Else
        Harg = Replace(Trim(TDB_Harga.Value), ",", "")
    End If
    
    '================================================================================='
    'melakukan penyimanan data inventori'
    '================================================================================='
    
    If txt_kode.Text <> "" Then
        If txt_nama.Text <> "" Then
            sql_smp_invent = "select * from tbl_inventori where kode_invent='" & txt_kode.Text & "'"
            Set rs_smp_invent = cn.Execute(sql_smp_invent)
            '========================================================================='
            'meyimpan data pada tabel inventori jika belum ada insert if yes update
            '========================================================================'
            With rs_smp_invent
                If (.BOF And .EOF) Then
                   sql_smp_invent = "insert into tbl_inventori (kode_invent,nama_invent,Harga_Barang)values('" & txt_kode.Text & "','" & txt_nama.Text & "'," & Harg & ")"
                  Else
                   sql_smp_invent = "update tbl_inventori set nama_invent='" & txt_nama.Text & "',Harga_Barang=" & Harg & " where kode_invent='" & txt_kode.Text & "'"
                End If
            End With
            cn.Execute (sql_smp_invent)
            
            '=========================================================================='
            'mencari id_inventori jika ketemu tempil or not id_invent=0
            '========================================================================='
            sql_smp_invent = "select * from tbl_inventori where  kode_invent='" & txt_kode.Text & "'"
            Set rs_smp_invent = cn.Execute(sql_smp_invent)
            With rs_smp_invent
                If (.BOF And .EOF) Then
                    id_invent = 0
                   Else
                    id_invent = rs_smp_invent("id_Invent")
                End If
            End With
            
            '========================================================================='
            'meyimpan data pada tabel tr_inventori sebagai transasaksi
            '========================================================================='
            
            If id_invent <> 0 Then
                sql_smp_invent = "insert into tbl_tr_inventori (id_invent,invent_in," & _
                "tgl_tr,ket,nama_user)values(" & id_invent & ",'" & Val(txt_stock.Text) & "','" & Format(Date, "dd/mm/yyyy") & "'," & _
                "'-','" & Trim(utama.lbl_user.Caption) & "' )"
                cn.Execute (sql_smp_invent)
            End If
            
            If id_invent <> 0 Then
                
                '====================================================================='
                'mencari id_invent pada tbl_stock_invent jika tidak ada maka insert
                'jika ada maka dilakukan update stock dengan menambah stock lama dengan stock baru'
                '====================================================================='
                
                sql_smp_invent = "select * from tbl_stock_invent  where id_invent=" & id_invent
                Set rs_smp_invent = cn.Execute(sql_smp_invent)
                With rs_smp_invent
                    If (.BOF And .EOF) Then
                        sql_smp_invent = "insert into tbl_stock_invent (id_invent,stock_invent)values(" & id_invent & ",'" & txt_stock.Text & "')"
                        cn.Execute (sql_smp_invent)
                       Else
                        Dim stock As Integer
                        stock = !stock_invent + Val(txt_stock.Text)
                        sql_smp_invent = "update tbl_stock_invent set stock_invent='" & stock & "' where id_invent=" & id_invent
                        cn.Execute (sql_smp_invent)
                    End If
                End With
            End If
             
            konfirm = MsgBox("Data sudah berhasil disimpan", vbInformation + vbOKOnly, "informasi")
           Else
            konfirm = MsgBox("Input nama inventori", vbInformation + vbOKOnly, "Informasi")
            Exit Sub
        End If
       Else
        konfirm = MsgBox("Inputkan kode inventori", vbInformation + vbOKOnly, "Informasi")
        Exit Sub
    End If
    
    proc_tmp_invent
    proc_bersih
                
    Exit Sub
    
er_simpan:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
                   
End Sub

Private Sub cmd_simpan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        proc_bersih
    End If
End Sub

Private Sub Form_Activate()
    txt_kode.SetFocus
End Sub

Private Sub Form_Load()
    '==============================================================================='
    'pengaturan form tampilan'
    '==============================================================================='
    
    grd_invent.Array = arr_invent
    
    kosong_invent
    
'    Me.Height = 7560
'    Me.Width = 7800
    Me.Left = (utama.Width - frm_input_inventori.Width) / 2
    Me.Top = (utama.Height - frm_input_inventori.Height) / 4 - 750
    
End Sub

Private Sub kosong_invent()

    arr_invent.ReDim 0, 0, 0, 0
    grd_invent.ReBind
    grd_invent.Refresh
    
End Sub

Public Sub proc_tmp_invent()
    '================================================================================='
    'mengisi dat pada grid dengan data inventori
    '================================================================================='
    
    On Error GoTo er_temp
    
    Dim nama, kode As String
    Dim Harg As Double
    Dim a As Long
    Dim total_harga As Double
    Dim id_invent As String
    
    kosong_invent
    
    sql_tmp_invent = "select * from tbl_inventori order by id_invent"
    rs_tmp_invent.Open sql_tmp_invent, cn, adOpenKeyset
        If Not rs_tmp_invent.EOF Then
            
            rs_tmp_invent.MoveLast
            rs_tmp_invent.MoveFirst
            
            a = 1
            total_harga = 0
                Do While Not rs_tmp_invent.EOF
                    
                    arr_invent.ReDim 1, a, 0, grd_invent.Columns.Count
                    grd_invent.ReBind
                    grd_invent.Refresh
                    
                    If Not IsNull(rs_tmp_invent!kode_invent) Then
                        kode = rs_tmp_invent!kode_invent
                    Else
                        kode = ""
                    End If
                    
                    If Not IsNull(rs_tmp_invent!nama_invent) Then
                        nama = rs_tmp_invent!nama_invent
                    Else
                        nama = ""
                    End If
                    
                    Harg = IIf(Not IsNull(rs_tmp_invent!Harga_Barang), rs_tmp_invent!Harga_Barang, 0)
                    id_invent = rs_tmp_invent!id_invent
                    
                    total_harga = total_harga + Harg
                    
                arr_invent(a, 0) = kode
                arr_invent(a, 1) = nama
                arr_invent(a, 2) = Harg
                arr_invent(a, 3) = id_invent
                
               a = a + 1
               rs_tmp_invent.MoveNext
               Loop
        End If
   rs_tmp_invent.Close
    
    grd_invent.Columns(1).FooterText = "TOTAL"
    grd_invent.Columns(2).FooterText = Format(total_harga, "###,###,###")
    
    grd_invent.ReBind
    grd_invent.Refresh
    
    grd_invent.MoveFirst
    
    Exit Sub
    
er_temp:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub

Private Sub grd_invent_DblClick()
    
    If arr_invent.UpperBound(1) = 0 Then Exit Sub
    
    txt_kode.Text = arr_invent(grd_invent.Bookmark, 0)
    txt_nama.Text = arr_invent(grd_invent.Bookmark, 1)
    
    If arr_invent(grd_invent.Bookmark, 2) = 0 Then
        TDB_Harga.Value = Null
    Else
        TDB_Harga.Value = arr_invent(grd_invent.Bookmark, 2)
    End If
    
    txt_kode.Enabled = False
    
    
End Sub

Private Sub grd_invent_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo err_handler
    
    If KeyCode = 13 Then
        grd_invent_DblClick
        
        On Error GoTo 0
        Exit Sub
        'cn.CommitTrans
    End If
    
    If KeyCode = vbKeyDelete Then
    
    cn.BeginTrans
    
    Dim sql As String
    Dim rs As Recordset
        
        If MsgBox("Apakah anda akan menghapus inventori " & arr_invent(grd_invent.Bookmark, 1), vbYesNo + vbQuestion, "Konfirmasi") = vbNo Then
            
            cn.RollbackTrans
            
            grd_invent.SetFocus
            
            On Error GoTo 0
            Exit Sub
        End If
        
        sql = "delete from tbl_stock_invent where id_invent=" & arr_invent(grd_invent.Bookmark, 3)
        
        Set rs = New ADODB.Recordset
            rs.Open sql, cn
        
        sql = "delete from tbl_tr_inventori where id_invent=" & arr_invent(grd_invent.Bookmark, 3)
            
        Set rs = New ADODB.Recordset
            rs.Open sql, cn
        
        sql = "delete from tbl_inventori where id_invent=" & arr_invent(grd_invent.Bookmark, 3)
        
        Set rs = New ADODB.Recordset
            rs.Open sql, cn
        
        cn.CommitTrans
        
        Dim Konfirmasi As Integer
        Konfirmasi = CInt(MsgBox("Data inventori " & arr_invent(grd_invent.Bookmark, 1) & " telah dihapus", vbOKOnly + vbInformation, "Informasi"))
        
        cmd_ref_Click
        cmd_baru_Click
        
        grd_invent.SetFocus
        
    On Error GoTo 0
    Exit Sub
        
    End If
    
    On Error GoTo 0
    Exit Sub
    
err_handler:
        
    If KeyCode = vbKeyDelete Then cn.RollbackTrans
        
            Konfirmasi = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
                Err.Clear

End Sub

Private Sub TDB_Harga_GotFocus()
    With TDB_Harga
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub TDB_Harga_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmd_simpan.SetFocus
End Sub

Private Sub TDB_Harga_LostFocus()
    
    With TDB_Harga
        If .Value = Null Then .Value = Null
    End With
    
End Sub

Private Sub txt_kode_GotFocus()
    txt_kode.SelStart = 0
    txt_kode.SelLength = Len(txt_kode)
End Sub

Private Sub txt_kode_KeyDown(KeyCode As Integer, Shift As Integer)
       Dim sql_cr_invent As String
       Dim rs_cr_invent As New ADODB.Recordset
       
       '============================================================================='
       'melakukan proses pencarian berdasarkankode invent jika ada maka akn diedit
       'jika tidak maka akan dilakukan insert baru
       '============================================================================='
       
       On Error GoTo er_kode
       
       If KeyCode = 13 Then
            If txt_kode.Text <> "" Then
                sql_cr_invent = "select * from tbl_inventori where kode_invent='" & txt_kode.Text & "'"
                Set rs_cr_invent = cn.Execute(sql_cr_invent)
                With rs_cr_invent
                    If (.BOF And .EOF) Then
                        txt_nama.Text = ""
                        txt_stock.Text = 0
                        TDB_Harga.Value = Null
                        txt_nama.SetFocus
                       Else
                        konfirm = MsgBox("Kode barang sudah ada apaka anda akan menedit data???", vbQuestion + vbYesNo, "Informasi")
                        If konfirm = vbYes Then
                            txt_kode.Enabled = False
                            txt_nama.Text = !nama_invent
                            TDB_Harga.Value = IIf(Not IsNull(!Harga_Barang), !Harga_Barang, Null)
                            txt_stock.Text = 0
                            txt_nama.SetFocus
                           Else
                            txt_kode.Text = ""
                            txt_kode.SetFocus
                            Exit Sub
                        End If
                    End If
                End With
               Else
                konfirm = MsgBox("Input kode inventori yang akan diproses", vbInformation + vbOKOnly, "Informasi")
                Exit Sub
            End If
        End If
                        
        Exit Sub
        
er_kode:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
                        
End Sub

Private Sub txt_kode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Public Sub proc_bersih()
    If txt_kode.Enabled = False Then
        txt_kode.Enabled = True
    End If
    txt_kode.Text = ""
    txt_nama.Text = ""
    txt_stock.Text = ""
    TDB_Harga.Value = Null
    
    txt_kode.SetFocus
    
    
End Sub

Private Sub txt_nama_GotFocus()
    txt_nama.SelStart = 0
    txt_nama.SelLength = Len(txt_nama)
End Sub

Private Sub txt_nama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If txt_nama.Text <> "" Then
            txt_stock.SetFocus
           Else
            MsgBox ("Input Nama Inventori")
            txt_nama.SetFocus
        End If
    End If
    
    If KeyCode = vbKeyEscape Then
        proc_bersih
    End If
End Sub

Private Sub txt_nama_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txt_stock_GotFocus()
    txt_stock.SelStart = 0
    txt_stock.SelLength = Len(txt_stock)
End Sub

Private Sub txt_stock_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
        If txt_stock.Text <> "" Then
            TDB_Harga.SetFocus
           Else
            MsgBox ("Input jumlah stock")
            txt_stock.SetFocus
        End If
    End If
    
    If KeyCode = vbKeyEscape Then
        proc_bersih
    End If
End Sub

Private Sub txt_stock_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
    Beep
    KeyAscii = 0
End If
End Sub
