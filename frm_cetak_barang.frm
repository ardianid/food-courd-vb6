VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_cetak_barang 
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   360
      ScaleHeight     =   705
      ScaleWidth      =   14505
      TabIndex        =   13
      Top             =   7680
      Width           =   14535
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
         Left            =   9120
         TabIndex        =   16
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmd_Cetak 
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
         Left            =   10920
         TabIndex        =   15
         Top             =   120
         Width           =   1695
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
         Left            =   12720
         TabIndex        =   14
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   360
      ScaleHeight     =   7425
      ScaleWidth      =   14505
      TabIndex        =   0
      Top             =   120
      Width           =   14535
      Begin MSComDlg.CommonDialog cd 
         Left            =   3960
         Top             =   4680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin TrueDBGrid60.TDBGrid grd_daftar 
         Height          =   4575
         Left            =   120
         OleObjectBlob   =   "frm_cetak_barang.frx":0000
         TabIndex        =   12
         Top             =   2760
         Width           =   14295
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2505
         ScaleWidth      =   14265
         TabIndex        =   1
         Top             =   120
         Width           =   14295
         Begin VB.TextBox txt_counter 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2760
            TabIndex        =   4
            Top             =   960
            Width           =   6735
         End
         Begin VB.Frame Frame1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   7080
            TabIndex        =   19
            Top             =   240
            Width           =   6735
            Begin VB.OptionButton opt_tidak 
               Caption         =   "Barang &Tdk Aktif"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4440
               TabIndex        =   22
               Top             =   360
               Width           =   1935
            End
            Begin VB.OptionButton opt_aktif 
               Caption         =   "Barang &Aktif"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2400
               TabIndex        =   21
               Top             =   360
               Width           =   1455
            End
            Begin VB.OptionButton opt_semua 
               Caption         =   "&Semua Barang"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   20
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.TextBox txt_kode 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2760
            TabIndex        =   18
            Top             =   480
            Width           =   2415
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
            Left            =   12720
            TabIndex        =   11
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox txt_harga_jual 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2760
            TabIndex        =   10
            Top             =   1920
            Width           =   1935
         End
         Begin VB.TextBox txt_nama_barang 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   6240
            TabIndex        =   8
            Top             =   1440
            Width           =   3255
         End
         Begin VB.TextBox txt_kode_barang 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2760
            TabIndex        =   6
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Counter"
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
            Left            =   1200
            TabIndex        =   17
            Top             =   600
            Width           =   1350
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Jual"
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
            Left            =   1200
            TabIndex        =   9
            Top             =   2040
            Width           =   1035
         End
         Begin VB.Label Label4 
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
            Left            =   4800
            TabIndex        =   7
            Top             =   1560
            Width           =   1350
         End
         Begin VB.Label Label3 
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
            Left            =   1200
            TabIndex        =   5
            Top             =   1560
            Width           =   1275
         End
         Begin VB.Label Label2 
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
            Left            =   1200
            TabIndex        =   3
            Top             =   1080
            Width           =   1425
         End
         Begin VB.Line Line1 
            X1              =   360
            X2              =   13800
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kriteria Pencetakan"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   2850
         End
      End
   End
End
Attribute VB_Name = "frm_cetak_barang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB
Dim sql As String

Private Sub cmd_cetak_Click()
    
    On Error GoTo er_printer

    With grd_daftar.PrintInfo
        
        .PageHeaderFont.Bold = True
        .PageHeaderFont.Italic = True
        .PageHeaderFont.Size = 10
        .PageHeader = "Laporan Data Barang Counter"
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

Private Sub cmd_tampil_Click()

On Error GoTo er_tampil

    Dim rs As New ADODB.Recordset
    Dim masuk As Boolean
    
    kosong_daftar
    
    sql = "select * from qr_barang"
        
        masuk = False
            
            
        
        If txt_kode.Text <> "" Or txt_counter.Text <> "" Or txt_kode_barang.Text <> "" Or txt_nama_barang.Text <> "" Or txt_harga_jual.Text <> "" Then
            
            masuk = True
            
            sql = sql & " where"
            
            If txt_kode.Text <> "" Then
                sql = sql & " kode_counter like '%" & Trim(txt_kode.Text) & "%'"
            End If
            
            If txt_counter.Text <> "" And txt_kode.Text = "" Then
                sql = sql & " nama_counter like '%" & Trim(txt_counter.Text) & "%'"
            End If
            
            If txt_counter.Text <> "" And txt_kode.Text <> "" Then
                sql = sql & " and nama_counter like '%" & Trim(txt_counter.Text) & "%'"
            End If
            
            If txt_kode_barang.Text <> "" And txt_counter.Text = "" And txt_kode.Text = "" Then
                sql = sql & " kode like '%" & Trim(txt_kode_barang.Text) & "%'"
            End If
            
            If txt_kode_barang.Text <> "" And (txt_counter.Text <> "" Or txt_kode.Text <> "") Then
                sql = sql & " and kode like '%" & Trim(txt_kode_barang.Text) & "%'"
            End If
            
            If txt_nama_barang.Text <> "" And txt_counter.Text = "" And txt_kode_barang.Text = "" And txt_kode.Text = "" Then
                sql = sql & " nama_barang like '%" & Trim(txt_nama_barang.Text) & "%'"
            End If
            
            If txt_nama_barang.Text <> "" And (txt_counter.Text <> "" Or txt_kode_barang.Text <> "" Or txt_kode.Text <> "") Then
                sql = sql & " and nama_barang like '%" & Trim(txt_nama_barang.Text) & "%'"
            End If
            
            If txt_harga_jual.Text <> "" And txt_counter.Text = "" And txt_kode_barang.Text = "" And txt_nama_barang.Text = "" And txt_kode.Text = "" Then
                sql = sql & " harga_jual= ccur(" & Trim(txt_harga_jual.Text) & ")"
            End If
            
            If txt_harga_jual.Text <> "" And (txt_kode.Text <> "" Or txt_counter.Text <> "" Or txt_kode_barang.Text <> "" Or txt_nama_barang.Text <> "") Then
                sql = sql & " and harga_jual= ccur(" & Trim(txt_harga_jual.Text) & ")"
            End If
            
        End If
        
        If opt_aktif.Value = True Then
            If masuk = True Then
                sql = sql & " and aktif=1"
            Else
                sql = sql & " where aktif=1"
            End If
        End If
        
        If opt_tidak.Value = True Then
            If masuk = True Then
                sql = sql & " and aktif=0"
            Else
                sql = sql & " where aktif=0"
            End If
        End If
        
        sql = sql & " order by nama_counter"
        
        rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
                
            Dim nama_counter, kode_barang, nama_barang, harga_jual, stock_awal, stock_max As String
            Dim a, b As Long
            a = 1
            b = 1
                Do While Not rs.EOF
                    arr_daftar.ReDim 1, a, 0, 8
                    grd_daftar.ReBind
                    grd_daftar.Refresh
                        
                        If Not IsNull(rs("nama_counter")) Then
                            nama_counter = rs("nama_counter")
                        Else
                            nama_counter = ""
                        End If
                        
                        If Not IsNull(rs("kode")) Then
                            kode_barang = rs("kode")
                        Else
                            kode_barang = ""
                        End If
                        
                        If Not IsNull(rs("nama_barang")) Then
                            nama_barang = rs("nama_barang")
                        Else
                            nama_barang = ""
                        End If
                        
                        If Not IsNull(rs("harga_jual")) Then
                            harga_jual = rs("harga_jual")
                        Else
                            harga_jual = ""
                        End If
                        
                        If Not IsNull(rs("stock_min")) Then
                            stock_awal = rs("stock_min")
                        Else
                            stock_awal = ""
                        End If
                        
                        If Not IsNull(rs("stock_max")) Then
                            stock_max = rs("stock_max")
                        Else
                            stock_max = ""
                        End If
                        
                   If a > 1 Then
                    If nama_counter <> arr_daftar(a - 1, 1) And nama_counter <> "" Then
                        b = b + 1
                    End If
                   End If
                   
                   arr_daftar(a, 0) = b
                   arr_daftar(a, 1) = nama_counter
                   arr_daftar(a, 2) = kode_barang
                   arr_daftar(a, 3) = nama_barang
                   arr_daftar(a, 4) = harga_jual
                   arr_daftar(a, 5) = stock_awal
                   arr_daftar(a, 6) = stock_max
                   
                a = a + 1
                rs.MoveNext
                Loop
                grd_daftar.ReBind
                grd_daftar.Refresh
            End If
        rs.Close
         
        Exit Sub
        
er_tampil:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
         
End Sub

Private Sub Form_Load()

    grd_daftar.Array = arr_daftar
    
    opt_semua.Value = True
    
    kosong_daftar
    
End Sub


Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub
Private Sub txt_counter_GotFocus()
    txt_counter.SelStart = 0
    txt_counter.SelLength = Len(txt_counter)
End Sub

Private Sub txt_harga_jual_GotFocus()
    txt_harga_jual.SelStart = 0
    txt_harga_jual.SelLength = Len(txt_harga_jual)
End Sub

Private Sub txt_harga_jual_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txt_kode_barang_GotFocus()
    txt_kode_barang.SelStart = 0
    txt_kode_barang.SelLength = Len(txt_kode_barang)
End Sub

Private Sub txt_nama_barang_GotFocus()
    txt_nama_barang.SelStart = 0
    txt_nama_barang.SelLength = Len(txt_nama_barang)
End Sub
