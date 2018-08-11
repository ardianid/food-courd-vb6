VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_persentase_jual 
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   120
      ScaleHeight     =   8385
      ScaleWidth      =   14985
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin MSComDlg.CommonDialog cd 
         Left            =   1680
         Top             =   4680
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
         Left            =   3120
         TabIndex        =   11
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
         Left            =   1680
         TabIndex        =   10
         Top             =   7800
         Width           =   1455
      End
      Begin VB.CommandButton cmd_page 
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
         Left            =   240
         TabIndex        =   9
         Top             =   7800
         Width           =   1455
      End
      Begin TrueDBGrid60.TDBGrid grd_daftar 
         Height          =   6375
         Left            =   120
         OleObjectBlob   =   "frm_persentase_jual.frx":0000
         TabIndex        =   5
         Top             =   1320
         Width           =   14655
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
         Left            =   13200
         TabIndex        =   4
         Top             =   840
         Width           =   1575
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
         Height          =   390
         Left            =   2640
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin MSMask.MaskEdBox msk_tgl1 
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
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
         Left            =   5280
         TabIndex        =   2
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode counter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   4800
         TabIndex        =   7
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   315
      End
   End
End
Attribute VB_Name = "frm_persentase_jual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB

Private Sub cmd_cetak_Click()
On Error GoTo er_printer

    With grd_daftar.PrintInfo
        
        .PageFooterFont.Name = "Arial"
        .PageHeaderFont.Size = 12
        .PageHeader = "TOTAL JUMLAH PENJUALAN \t\tPeriode  : " & Trim(msk_tgl1.Text) & " s/d " & Trim(msk_tgl2.Text)
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

Private Sub cmd_page_Click()
On Error GoTo er_page
    
    With grd_daftar.PrintInfo
        .PageSetup
    End With
    Exit Sub
    
er_page:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub Cmd_Tampil_Click()

On Error GoTo er_tampil

Dim sql, sql1, sql2, sql3 As String
Dim rs As New ADODB.Recordset, rs2 As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset, rs3 As New ADODB.Recordset
Dim sql5 As String
Dim rs5 As New ADODB.Recordset

    
    If msk_tgl1.Text = "__/__/____" Or msk_tgl2.Text = "__/__/____" Then
        MsgBox ("Tanggal harus diisi")
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    utama.MousePointer = vbHourglass
    
    sql5 = "delete from tbl_temp_counter"
    rs5.Open sql5, cn
    
    sql5 = "delete from tbl_temp_barang"
    rs5.Open sql5, cn
    
    sql = "select distinct(kode_counter)as kode_counter from qr_penjualan_sebenarnya where tgl >= datevalue('" & Trim(msk_tgl1.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl2.Text) & "')"
    
    If txt_kode.Text <> "" Then
        sql = sql & " and kode_counter='" & Trim(txt_kode.Text) & "'"
    End If
    
    rs.Open sql, cn, adOpenKeyset
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            
                Do While Not rs.EOF
                    sql1 = "select count(kode_counter) as jml_counter from qr_penjualan_sebenarnya where tgl >= datevalue('" & Trim(msk_tgl1.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl2.Text) & "') and kode_counter='" & rs!kode_counter & "'"
                    rs1.Open sql1, cn
                        
                        If Not rs1.EOF Then
                            sql2 = "select distinct(kode_barang) as kode_barang from qr_penjualan_sebenarnya where tgl >= datevalue('" & Trim(msk_tgl1.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl2.Text) & "') and kode_counter='" & rs!kode_counter & "'"
                            rs2.Open sql2, cn, adOpenKeyset
                                If Not rs2.EOF Then
                                    rs2.MoveLast
                                    rs2.MoveFirst
                                        
                                        sql5 = "insert into tbl_temp_counter (kode_counter,jml) values('" & Trim(rs!kode_counter) & "'," & rs1!jml_counter & ")"
                                        rs5.Open sql5, cn
                                        
                                        Do While Not rs2.EOF
                                            sql3 = "select sum(qty) as jml_barang from qr_penjualan_sebenarnya where tgl >= datevalue('" & Trim(msk_tgl1.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl2.Text) & "') and kode_counter='" & rs!kode_counter & "'"
                                            sql3 = sql3 & " and kode_barang = '" & rs2!kode_barang & "'"
                                            rs3.Open sql3, cn
                                                If Not rs3.EOF Then
                                                    sql5 = "insert into tbl_temp_barang (kode_counter,kode_barang,jml) values('" & Trim(rs!kode_counter) & "','" & Trim(rs2!kode_barang) & "'," & Trim(rs3!jml_barang) & ")"
                                                    rs5.Open sql5, cn
                                                End If
                                            rs3.Close
                                        rs2.MoveNext
                                        Loop
                                End If
                            rs2.Close
                        End If
                    rs1.Close
                rs.MoveNext
                Loop
        End If
    rs.Close
    
    Call isi_grid
    
    If Me.MousePointer = vbHourglass Then
        Me.MousePointer = vbDefault
        utama.MousePointer = vbDefault
    End If
        
    
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

Private Sub isi_grid()
    
    kosong_daftar
         
    Dim kode_counter, nama_counter, kode_barang, nama_barang As String
    Dim jml, jml_counter, a, b As Long
    Dim tot_counter, tot_barang As Double
         
        Dim sql, sql1 As String
        Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
            
                    sql = "select * from qr_temp_counter order by jml desc"
                    rs.Open sql, cn, adOpenKeyset
                        If Not rs.EOF Then
                            
                            rs.MoveLast
                            rs.MoveFirst
                                
                                a = 1
                                b = 1
                                tot_counter = 0
                                tot_barang = 0
                                    Do While Not rs.EOF
                                      sql1 = "select * from qr_temp_barang where kode_counter='" & Trim(rs!kode_counter) & "' order by jml desc"
                                      rs1.Open sql1, cn, adOpenKeyset
                                       If Not rs1.EOF Then
                                        
                                        rs1.MoveLast
                                        rs1.MoveFirst
                                        
                                        If Not IsNull(rs!jml) Then
                                            tot_counter = tot_counter + CDbl(rs!jml)
                                        Else
                                            tot_counter = tot_counter + 0
                                        End If
                                        
                                            Do While Not rs1.EOF
                                                arr_daftar.ReDim 1, a, 0, 8
                                                grd_daftar.ReBind
                                                grd_daftar.Refresh
                                                    
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
                                                    
                                                    If Not IsNull(rs!jml) Then
                                                        jml_counter = rs!jml
                                                    Else
                                                        jml_counter = 0
                                                    End If
                                                    
                                                    If Not IsNull(rs1!kode_barang) Then
                                                        kode_barang = rs1!kode_barang
                                                    Else
                                                        kode_barang = ""
                                                    End If
                                                    
                                                    If Not IsNull(rs1!nama_barang) Then
                                                        nama_barang = rs1!nama_barang
                                                    Else
                                                        nama_barang = ""
                                                    End If
                                                    
                                                    If Not IsNull(rs1!jml) Then
                                                        jml = rs1!jml
                                                    Else
                                                        jml = 0
                                                    End If
                                                    
                                                    If a > 1 Then
                                                        If kode_counter <> arr_daftar(a - 1, 1) Then
                                                            b = b + 1
                                                        End If
                                                    End If
                                                    
                                                    
                                                    tot_barang = tot_barang + CDbl(jml)
                                                    
                                                    arr_daftar(a, 0) = b
                                                    arr_daftar(a, 1) = kode_counter
                                                    arr_daftar(a, 2) = nama_counter
                                                    arr_daftar(a, 3) = jml_counter
                                                    arr_daftar(a, 4) = kode_barang
                                                    arr_daftar(a, 5) = nama_barang
                                                    arr_daftar(a, 6) = jml
                                            a = a + 1
                                            rs1.MoveNext
                                            Loop
                                          End If
                                          rs1.Close
                                          rs.MoveNext
                                       Loop
                                            
                                            grd_daftar.Columns(2).FooterText = "TOTAL"
                                            grd_daftar.Columns(3).FooterText = tot_counter
                                            grd_daftar.Columns(6).FooterText = tot_barang
                                            
                                            grd_daftar.ReBind
                                            grd_daftar.Refresh
                                  End If
                                rs.Close
                            
        
        
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
