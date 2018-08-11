VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_histock 
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   14985
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cd 
      Left            =   960
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   9360
      ScaleHeight     =   735
      ScaleWidth      =   5415
      TabIndex        =   19
      Top             =   7680
      Width           =   5415
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
         TabIndex        =   22
         Top             =   120
         Width           =   1575
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
         Left            =   1920
         TabIndex        =   21
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmd_Export 
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
         Left            =   3600
         TabIndex        =   20
         Top             =   120
         Width           =   1575
      End
   End
   Begin TrueDBGrid60.TDBGrid grd_stock 
      Height          =   5055
      Left            =   360
      OleObjectBlob   =   "frm_histock.frx":0000
      TabIndex        =   18
      Top             =   2520
      Width           =   14415
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   360
      ScaleHeight     =   2145
      ScaleWidth      =   14385
      TabIndex        =   0
      Top             =   240
      Width           =   14415
      Begin VB.TextBox txt_user 
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
         Height          =   405
         Left            =   7680
         TabIndex        =   24
         Top             =   1440
         Width           =   3135
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
         Left            =   12840
         TabIndex        =   10
         Top             =   1560
         Width           =   1455
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   9960
         ScaleHeight     =   465
         ScaleWidth      =   4425
         TabIndex        =   17
         Top             =   0
         Width           =   4455
         Begin VB.OptionButton opt_semua 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Semua"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   9
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton opt_out 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Barang Out"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   8
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton opt_in 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Barang In"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.TextBox txt_nama_barang 
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
         Height          =   405
         Left            =   7680
         TabIndex        =   6
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox txt_kode_barang 
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
         Height          =   405
         Left            =   7680
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txt_kode_counter 
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
         Height          =   405
         Left            =   2280
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txt_nama_counter 
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
         Height          =   405
         Left            =   2280
         TabIndex        =   4
         Top             =   1320
         Width           =   3135
      End
      Begin MSMask.MaskEdBox msk_tgl 
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox msk_tgl1 
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama User"
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
         Left            =   6000
         TabIndex        =   23
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
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
         Left            =   6000
         TabIndex        =   16
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
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
         Left            =   6000
         TabIndex        =   15
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Counter"
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
         TabIndex        =   14
         Top             =   1320
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode counter"
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
         TabIndex        =   13
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
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
         Left            =   3840
         TabIndex        =   12
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl"
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
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   300
      End
   End
End
Attribute VB_Name = "frm_histock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_stock As New XArrayDB
Dim sql As String

Private Sub cmd_cetak_Click()
    On Error GoTo er_printer

    With grd_stock.PrintInfo
        
        .PageHeaderFont.Bold = True
        .PageHeaderFont.Italic = True
        .PageHeaderFont.Size = 10
        .PageHeader = "Laporan Historical Stock"
        .RepeatColumnHeaders = True
        .PageFooter = "\tPage: \p" & "..." & id_user
        .PrintPreview
    End With
    Exit Sub
    
er_printer:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clea
End Sub

Private Sub cmd_export_Click()
    
On Error Resume Next

    cd.ShowSave
    grd_stock.ExportToFile cd.FileName, False
    
End Sub

Private Sub cmd_setup_Click()
   
On Error GoTo er_setup
   
   With grd_stock.PrintInfo
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

    grd_stock.Array = arr_stock
    
    kosong
    
    opt_semua.Value = True
    
End Sub

Private Sub kosong()
    arr_stock.ReDim 0, 0, 0, 0
    grd_stock.ReBind
    grd_stock.Refresh
End Sub

Private Sub isi()

On Error GoTo er_isi

    Dim rs As New ADODB.Recordset
    Dim dimana As Integer
        
        kosong
        
        dimana = 0
        sql = "select * from qr_historical_stock where ket <> 3"
If msk_tgl.Text <> "__/__/____" Or msk_tgl1.Text <> "__/__/____" Or txt_kode_counter.Text <> "" Or txt_nama_counter.Text <> "" Or _
    txt_kode_barang.Text <> "" Or txt_nama_barang.Text <> "" Then
        
        
        dimana = 1
        If msk_tgl.Text <> "__/__/____" And msk_tgl1.Text = "__/__/____" Then
            sql = sql & " and tgl= datevalue('" & Trim(msk_tgl.Text) & "')"
        End If
        
        If msk_tgl1.Text <> "__/__/____" And msk_tgl.Text = "__/__/____" Then
            sql = sql & " and tgl= datevalue('" & Trim(msk_tgl1.Text) & "')"
        End If
        
        If msk_tgl.Text <> "__/__/____" And msk_tgl1.Text <> "__/__/____" Then
            sql = sql & " and tgl >= datevalue('" & Trim(msk_tgl.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl1.Text) & "')"
        End If
        
        If txt_kode_counter.Text <> "" And msk_tgl.Text = "__/__/____" And msk_tgl1.Text = "__/__/____" Then
            sql = sql & " and kode_counter like '%" & Trim(txt_kode_counter.Text) & "%'"
        End If
        
        If txt_kode_counter.Text <> "" And (msk_tgl.Text <> "__/__/____" Or msk_tgl1.Text <> "__/__/____") Then
            sql = sql & " and kode_counter like '%" & Trim(txt_kode_counter.Text) & "%'"
        End If
        
        If txt_nama_counter.Text <> "" And txt_kode_counter.Text = "" And msk_tgl.Text = "__/__/____" And msk_tgl1.Text = "__/__/____" Then
            sql = sql & " and nama_counter like '%" & Trim(txt_nama_counter.Text) & "%'"
        End If
        
        If txt_nama_counter.Text <> "" And (txt_kode_counter.Text <> "" Or msk_tgl.Text <> "__/__/____" Or msk_tgl1.Text <> "__/__/____") Then
            sql = sql & " and nama_counter like '%" & Trim(txt_nama_counter.Text) & "%'"
        End If
        
        If txt_kode_barang.Text <> "" And txt_nama_counter.Text = "" And txt_kode_counter.Text = "" And msk_tgl.Text = "__/__/____" And msk_tgl1.Text = "__/__/____" Then
            sql = sql & " and kode_barang like '%" & Trim(txt_kode_barang.Text) & "%'"
        End If
        
        If txt_kode_barang.Text <> "" And (txt_nama_counter.Text <> "" Or txt_kode_counter.Text <> "" Or msk_tgl.Text = "__/__/____" Or msk_tgl1.Text = "__/__/____") Then
            sql = sql & " and kode_barang like '%" & Trim(txt_kode_barang.Text) & "%'"
        End If
        
        If txt_nama_barang.Text <> "" And txt_kode_barang.Text = "" And txt_nama_counter.Text = "" And txt_kode_counter.Text = "" And msk_tgl.Text = "__/__/____" And msk_tgl1.Text = "__/__/____" Then
            sql = sql & " and nama_barang like '%" & Trim(txt_nama_barang.Text) & "%'"
        End If
                 
        If txt_nama_barang.Text <> "" And (txt_kode_barang.Text <> "" Or txt_nama_counter.Text <> "" Or txt_kode_counter.Text <> "" Or msk_tgl.Text <> "__/__/____" Or msk_tgl1.Text <> "__/__/____") Then
            sql = sql & " and nama_barang like '%" & Trim(txt_nama_barang.Text) & "%'"
        End If
        
        If txt_user.Text <> "" And txt_nama_barang.Text = "" And txt_kode_barang.Text = "" And txt_nama_counter.Text = "" And txt_kode_counter.Text = "" And msk_tgl.Text = "__/__/____" And msk_tgl1.Text = "__/__/____" Then
            sql = sql & " and nama_karyawan like '%" & Trim(txt_user.Text) & "%'"
        End If
        
        If txt_user.Text <> "" And (txt_nama_barang.Text <> "" Or txt_kode_barang.Text <> "" Or txt_nama_counter.Text <> "" Or txt_kode_counter.Text <> "" Or msk_tgl.Text <> "__/__/____" Or msk_tgl1.Text <> "__/__/____") Then
            sql = sql & " and nama_karyawan like '%" & Trim(txt_user.Text) & "%'"
        End If
End If
        
        If opt_in.Value = True Then
    
            sql = sql & " and brg_in <> 0"
         
        End If
        
        If opt_out.Value = True Then
         
            sql = sql & " and brg_out <> 0"
         
        End If
        
        sql = sql & " order by tgl,kode_counter,nama_counter"
        rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
                
                lanjut rs
            End If
        rs.Close
            
Exit Sub

er_isi:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
            
End Sub

Private Sub lanjut(rs As Recordset)
    Dim tgl, kode_counter, nama_counter, kode_barang, nama_barang, brg_in, brg_out, ket, user, kemana As String
    Dim a, b As Long
    Dim jml_in, jml_out As String
    
        a = 1
        b = 1
        jml_in = 0
        jml_out = 0
            Do While Not rs.EOF
                arr_stock.ReDim 1, a, 0, 12
                grd_stock.ReBind
                grd_stock.Refresh
                
                    If Not IsNull(rs("tgl")) Then
                        tgl = rs("tgl")
                    Else
                        tgl = ""
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
                    
                    If Not IsNull(rs("brg_in")) Then
                        brg_in = rs("brg_in")
                    Else
                        brg_in = 0
                    End If
                    
                    If Not IsNull(rs("brg_out")) Then
                        brg_out = rs("brg_out")
                    Else
                        brg_out = 0
                    End If
                    
                    If Not IsNull(rs("ket")) Then
                        ket = rs("ket")
                        
                        If ket = 0 Then
                            ket = "-"
                        End If
                        
                        If ket = "1" Then
                            ket = "Penyesuaian"
                        End If
                        
                    Else
                        ket = ""
                    End If
                    
                    If Not IsNull(rs("nama_karyawan")) Then
                        user = rs("nama_karyawan")
                    Else
                        user = ""
                    End If
                    
                    If Not IsNull(rs("kemana")) Then
                        kemana = rs("kemana")
                    Else
                        kemana = ""
                    End If
                    
                    If a > 1 Then
                      If tgl <> arr_stock(a - 1, 1) Then
                        b = b + 1
                      End If
                    End If
                    
               jml_in = jml_in + CDbl(brg_in)
               jml_out = jml_out + CDbl(brg_out)
                    
               arr_stock(a, 0) = b
               arr_stock(a, 1) = tgl
               arr_stock(a, 2) = kode_counter
               arr_stock(a, 3) = nama_counter
               arr_stock(a, 4) = kode_barang
               arr_stock(a, 5) = nama_barang
               arr_stock(a, 6) = brg_in
               arr_stock(a, 7) = brg_out
               arr_stock(a, 8) = ket
               arr_stock(a, 9) = kemana
               arr_stock(a, 10) = user
               
            a = a + 1
            rs.MoveNext
            Loop
            
            grd_stock.Columns(5).FooterText = "TOTAL"
            grd_stock.Columns(6).FooterText = jml_in
            grd_stock.Columns(7).FooterText = jml_out
            
            grd_stock.ReBind
            grd_stock.Refresh
                                        
End Sub

Private Sub msk_tgl_GotFocus()
    msk_tgl.SelStart = 0
    msk_tgl.SelLength = Len(msk_tgl)
End Sub

Private Sub msk_tgl1_GotFocus()
    msk_tgl1.SelStart = 0
    msk_tgl1.SelLength = Len(msk_tgl1)
End Sub
Private Sub txt_kode_barang_GotFocus()
    txt_kode_barang.SelStart = 0
    txt_kode_barang.SelLength = Len(txt_kode_barang)
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
Private Sub txt_user_GotFocus()
    txt_user.SelStart = 0
    txt_user.SelLength = Len(txt_user)
End Sub
