VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_lap_perkasir1 
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
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
      Begin MSComCtl2.DTPicker dtp_tgl1 
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   12582912
         CalendarTitleForeColor=   16777215
         Format          =   52101121
         CurrentDate     =   39215
      End
      Begin VB.CommandButton cmd_cetak_faktur 
         Caption         =   "Cetak Ke Faktur"
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
         TabIndex        =   11
         Top             =   7800
         Width           =   1935
      End
      Begin VB.TextBox txt_nama 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         TabIndex        =   6
         Top             =   1320
         Width           =   4695
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   7800
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
         Left            =   11880
         TabIndex        =   2
         Top             =   7800
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
         TabIndex        =   1
         Top             =   1320
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   7080
         Top             =   7680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin TrueDBGrid60.TDBGrid grd_daftar 
         Height          =   5775
         Left            =   120
         OleObjectBlob   =   "frm_lap_perkasir1.frx":0000
         TabIndex        =   5
         Top             =   1920
         Width           =   14775
      End
      Begin MSComCtl2.DTPicker dtp_tgl2 
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   12582912
         CalendarTitleForeColor=   16777215
         Format          =   52101121
         CurrentDate     =   39215
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
         TabIndex        =   10
         Top             =   840
         Width           =   375
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
         TabIndex        =   9
         Top             =   240
         Width           =   2970
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         X1              =   360
         X2              =   14640
         Y1              =   600
         Y2              =   600
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
         TabIndex        =   8
         Top             =   840
         Width           =   315
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
         TabIndex        =   7
         Top             =   1320
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frm_lap_perkasir1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB
Dim sql_tampil As String

Private Sub cmd_cetak_Click()

    On Error GoTo er_printer

    With grd_daftar.PrintInfo
        
        .PageFooterFont.Name = "Arial"
        .PageHeaderFont.Size = 12
        .PageHeader = "LAPORAN PENJUALAN PERKASIR \t\t Periode  : " & Trim(DTP_Tgl1.Value) & " s/d " & Trim(DTP_Tgl2.Value)
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

Private Sub cmd_cetak_faktur_Click()
Dim sql, sql1 As String
Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
Dim a As Long
Dim jml_seluruh As Double
    
    sql = "select distinct(kode_counter)as kd_counter from qr_penjualan_sebenarnya where tgl >= datevalue('" & Trim(DTP_Tgl1.Value) & "') and tgl <= datevalue('" & Trim(DTP_Tgl2.Value) & "') and nama_user='" & Trim(txt_nama.Text) & "'"
    rs.Open sql, cn, adOpenKeyset
        If Not rs.EOF Then
                        
            Printer.Font = "Arial"
            Printer.FontSize = 11
        
            Printer.CurrentX = 0
            Printer.CurrentY = 0
            
            Printer.Print
            Printer.Print "Laporan penjualan Kasir"
            Printer.Print
            Printer.Print "Nama"; Tab(11); ":"; Tab(13); txt_nama.Text
            Printer.Print "Peroide"; Tab(11); ":"; Tab(13); DTP_Tgl1.Value & " s/d " & DTP_Tgl2.Value
            Printer.Print
            Printer.Print "No."; Tab(5); "Kd Counter"; Tab(17); "Jumlah"
            Printer.Print
                        
            rs.MoveLast
            rs.MoveFirst
                
              jml_seluruh = 0
              a = 1
              Do While Not rs.EOF
                sql1 = "select sum(total_harga) as harga from qr_penjualan_sebenarnya where"
                sql1 = sql1 & " kode_counter='" & rs!kd_counter & "' and tgl >= datevalue('" & Trim(DTP_Tgl1.Value) & "') and tgl <= datevalue('" & Trim(DTP_Tgl2.Value) & "') and nama_user='" & Trim(txt_nama.Text) & "'"
                rs1.Open sql1, cn
                    If Not rs1.EOF Then
                        
                        Dim harga
                        harga = Format(rs1!harga, "###,###,###")
                        
                        jml_seluruh = CDbl(jml_seluruh) + CDbl(rs1!harga)
                        
                        Printer.Print a; Tab(5); rs!kd_counter; Tab(17); harga
                        
                        a = a + 1
                    End If
                rs1.Close
            
              rs.MoveNext
              Loop
            
            Dim grs
            
            grs = String$(29, "-")
            
            Printer.Print
            Printer.Print grs
            Printer.Print "TOTAL"; Tab(17); Format(jml_seluruh, "###,###,###")
            Printer.EndDoc
        End If
      rs.Close
End Sub

Private Sub cmd_export_Click()
On Error Resume Next
    cd.ShowSave
    grd_daftar.ExportToFile cd.FileName, False
End Sub

Private Sub Cmd_Tampil_Click()

On Error GoTo err_tampil


Dim rs As New ADODB.Recordset
Dim a, b As Long
Dim tgl, jam, faktur, kode_counter, nama_counter, kode_barang, nama_barang, qty, harga_satuan, cash, disc, total_harga As String
Dim jml_qty, jml_harga_satuan, jml_disc, jml_cash, jml_total As Double
Dim jml_qty1, jml_harga_satuan1, jml_disc1, jml_cash1, jml_total1 As Double
    
    kosong_daftar
    
    grd_daftar.Caption = "NAMA KASIR : " & UCase(Trim(txt_nama.Text))
    
    sql_tampil = "select kode_counter,nama_counter,kode_barang,nama_barang,no_faktur,tgl,jam,qty,harga_satuan,disc,cash,total_harga,nama_user,ppn from qr_penjualan_sebenarnya"
        
        sql_tampil = sql_tampil & " where"
        
        sql_tampil = sql_tampil & " tgl >= datevalue('" & Trim(DTP_Tgl1.Value) & "') and tgl <= datevalue('" & Trim(DTP_Tgl2.Value) & "')"
        
            If txt_nama.Text <> "" Then
                sql_tampil = sql_tampil & " and nama_user='" & Trim(txt_nama.Text) & "'"
            End If
    
    sql_tampil = sql_tampil & " order by kode_counter"
    rs.Open sql_tampil, cn, adOpenKeyset
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
                
                jml_qty1 = 0: jml_harga_satuan1 = 0: jml_disc1 = 0: jml_cash1 = 0: jml_total1 = 0
                
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
                    
                    If Not IsNull(rs!ppn) Then
                        cash = rs!ppn
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
                    
                    If a > 1 Then
                        If kode_counter <> arr_daftar(a - 1, 4) Then
                            
                            
                            
                            arr_daftar(a, 0) = ""
                            arr_daftar(a, 1) = ""
                            arr_daftar(a, 2) = ""
                            arr_daftar(a, 3) = ""
                            arr_daftar(a, 4) = ""
                            arr_daftar(a, 5) = ""
                            arr_daftar(a, 6) = ""
                            arr_daftar(a, 7) = "Grand Total"
                            arr_daftar(a, 8) = jml_qty1
                            arr_daftar(a, 9) = Format(jml_harga_satuan1, "###,###,###")
                            arr_daftar(a, 10) = jml_disc1 & "%"
                            arr_daftar(a, 11) = jml_cash1 & "%"
                            arr_daftar(a, 12) = Format(jml_total1, "###,###,###")
                            
                            a = a + 2
                            
                            arr_daftar.ReDim 1, a, 0, 15
                            grd_daftar.ReBind
                            grd_daftar.Refresh
                                        
                            jml_qty1 = 0: jml_harga_satuan1 = 0: jml_disc1 = 0: jml_cash1 = 0: jml_total1 = 0
                                        
                        End If
                    End If
                        
                Dim disc_b, cash_b
                    disc_b = Mid(disc, 1, Len(disc) - 1)
                    cash_b = Mid(cash, 1, Len(cash) - 1)
                    
                ' untuk seluruh
                    
                jml_qty = CDbl(jml_qty) + CDbl(qty)
                jml_harga_satuan = CDbl(jml_harga_satuan) + CDbl(harga_satuan)
                jml_disc = CDbl(jml_disc) + CDbl(disc_b)
                jml_cash = CDbl(jml_cash) + CDbl(cash_b)
                jml_total = CDbl(jml_total) + CDbl(total_harga)
                    
                ' untuk satu
                    jml_qty1 = CDbl(jml_qty1) + CDbl(qty)
                    jml_harga_satuan1 = CDbl(jml_harga_satuan1) + CDbl(harga_satuan)
                    jml_disc1 = CDbl(jml_disc1) + CDbl(disc_b)
                    jml_cash1 = CDbl(jml_cash1) + CDbl(cash_b)
                    jml_total1 = CDbl(jml_total1) + CDbl(total_harga)
                    
                    
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
           
           
           'untuk terakhir
            
                            
                            
                            arr_daftar.ReDim 1, a, 0, 15
                            grd_daftar.ReBind
                            grd_daftar.Refresh
           
            
                            arr_daftar(a, 0) = ""
                            arr_daftar(a, 1) = ""
                            arr_daftar(a, 2) = ""
                            arr_daftar(a, 3) = ""
                            arr_daftar(a, 4) = ""
                            arr_daftar(a, 5) = ""
                            arr_daftar(a, 6) = ""
                            arr_daftar(a, 7) = "Grand Total"
                            arr_daftar(a, 8) = jml_qty1
                            arr_daftar(a, 9) = Format(jml_harga_satuan1, "###,###,###")
                            arr_daftar(a, 10) = jml_disc1 & "%"
                            arr_daftar(a, 11) = jml_cash1 & "%"
                            arr_daftar(a, 12) = Format(jml_total1, "###,###,###")
                            
            ' Sampai Disini
           
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

Private Sub Form_Load()

grd_daftar.Array = arr_daftar

DTP_Tgl1.Value = Format(Date, "dd,mm,yyyy")
DTP_Tgl2.Value = Format(Date, "dd/mm/yyyy")
txt_nama.Text = utama.lbl_user.Caption

kosong_daftar

Cmd_Tampil_Click

End Sub

Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub

Private Sub grd_daftar_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next

If sql_tampil = "" Then
    Exit Sub
End If

If arr_daftar.UpperBound(1) = 0 Then
    Exit Sub
End If
Dim sql As String
Dim rs As New ADODB.Recordset
Dim a, b As Long
Dim tgl, jam, faktur, kode_counter, nama_counter, kode_barang, nama_barang, qty, harga_satuan, cash, disc, total_harga As String
Dim jml_qty, jml_harga_satuan, jml_disc, jml_cash, jml_total As Double
sql = ""
sql = sql_tampil

Select Case ColIndex
    Case 1
        sql = sql & " order by tgl"
    Case 2
        sql = sql & " order by jam"
    Case 3
        sql = sql & " order by no_faktur"
    Case 4
        sql = sql & " order by kode_counter"
    Case 5
        sql = sql & " order by nama_counter"
    Case 6
        sql = sql & " order by kode_barang"
    Case 7
        sql = sql & " order by nama_barang"
    Case 8
        sql = sql & " order by qty"
    Case 9
        sql = sql & " order by harga_satuan"
    Case 10
        sql = sql & " order by disc"
    Case 11
        sql = sql & " order by cash"
    Case 12
        sql = sql & " order by total_harga"
End Select
    
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
End Sub
