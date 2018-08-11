VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Begin VB.Form frm_bc 
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin TDBContainer3D6Ctl.TDBContainer3D pic_tambah 
      Height          =   4335
      Left            =   -4680
      TabIndex        =   22
      Top             =   3600
      Visible         =   0   'False
      Width           =   10455
      _Version        =   65536
      _ExtentX        =   18441
      _ExtentY        =   7646
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "frm_bc.frx":0000
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_bc.frx":001C
      Childs          =   "frm_bc.frx":00C8
      Begin VB.PictureBox pipi 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   240
         ScaleHeight     =   3705
         ScaleWidth      =   9825
         TabIndex        =   23
         Top             =   240
         Width           =   9855
         Begin VB.Frame Frame2 
            Height          =   15
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   9615
         End
         Begin VB.Frame Frame1 
            Height          =   615
            Left            =   4320
            TabIndex        =   44
            Top             =   480
            Width           =   5415
            Begin VB.CheckBox Check_Diskon 
               Caption         =   "Tidak &Disc"
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
               Left            =   2520
               TabIndex        =   26
               Top             =   240
               Width           =   1455
            End
            Begin VB.CheckBox cek_ket 
               Caption         =   "&Perhitungan Stock"
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
               Left            =   120
               TabIndex        =   25
               Top             =   240
               Width           =   2175
            End
            Begin VB.CheckBox cek_aktif 
               Caption         =   "A&ktif"
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
               Left            =   4320
               TabIndex        =   27
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.TextBox txt_kode_barang 
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
            Left            =   1920
            TabIndex        =   24
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox txt_nama_barang 
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
            Left            =   1920
            TabIndex        =   29
            Top             =   1200
            Width           =   5535
         End
         Begin VB.TextBox txt_harga 
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
            Left            =   1920
            TabIndex        =   31
            Top             =   1680
            Width           =   2295
         End
         Begin VB.TextBox txt_stock_min 
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
            Left            =   1920
            TabIndex        =   33
            Top             =   2400
            Width           =   735
         End
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
            Height          =   495
            Left            =   6960
            TabIndex        =   35
            Top             =   3000
            Width           =   1335
         End
         Begin VB.CommandButton cmd_cancel 
            Caption         =   "Cancel"
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
            Left            =   8400
            TabIndex        =   36
            Top             =   3000
            Width           =   1335
         End
         Begin VB.TextBox txt_stock_max 
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
            Left            =   4440
            TabIndex        =   34
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox txt_satuan 
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
            Left            =   5160
            TabIndex        =   32
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Data Barang"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   240
            TabIndex        =   43
            Top             =   120
            Width           =   1305
         End
         Begin VB.Label Label3 
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
            Left            =   360
            TabIndex        =   42
            Top             =   720
            Width           =   1260
         End
         Begin VB.Label Label4 
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
            Left            =   360
            TabIndex        =   41
            Top             =   1200
            Width           =   1350
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Jual"
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
            TabIndex        =   40
            Top             =   1680
            Width           =   1020
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Minimum"
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
            TabIndex        =   39
            Top             =   2400
            Width           =   1470
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Maximum"
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
            Left            =   2760
            TabIndex        =   38
            Top             =   2400
            Width           =   1560
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Counter :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   5640
            TabIndex        =   37
            Top             =   120
            Width           =   990
         End
         Begin VB.Label lbl_kode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_kode"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   6720
            TabIndex        =   30
            Top             =   120
            Width           =   900
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Satuan"
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
            Left            =   4320
            TabIndex        =   28
            Top             =   1680
            Width           =   675
         End
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D pic_counter 
      Height          =   6615
      Left            =   2640
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   11668
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "frm_bc.frx":00E4
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_bc.frx":0100
      Childs          =   "frm_bc.frx":01AC
      Begin VB.TextBox txt_nm 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2040
         TabIndex        =   18
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox txt_nm 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   17
         Top             =   840
         Width           =   1455
      End
      Begin TrueDBGrid60.TDBGrid grd_counter 
         Height          =   4695
         Left            =   240
         OleObjectBlob   =   "frm_bc.frx":01C8
         TabIndex        =   21
         Top             =   1680
         Width           =   5535
      End
      Begin VB.Frame Frame5 
         Height          =   135
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   16
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nama Jenis Barang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Width           =   1635
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Kode Jenis Barang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   1560
      End
   End
   Begin VB.PictureBox mm 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   -5400
      ScaleHeight     =   5625
      ScaleWidth      =   5625
      TabIndex        =   9
      Top             =   8280
      Width           =   5655
      Begin VB.CommandButton cmd_x 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   5625
         TabIndex        =   10
         Top             =   0
         Width           =   5655
      End
   End
   Begin VB.PictureBox pic_barang 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   240
      ScaleHeight     =   6585
      ScaleWidth      =   14745
      TabIndex        =   7
      Top             =   1560
      Width           =   14775
      Begin TrueDBGrid60.TDBGrid grd_barang 
         Height          =   5775
         Left            =   120
         OleObjectBlob   =   "frm_bc.frx":2EB5
         TabIndex        =   3
         Top             =   120
         Width           =   14535
      End
      Begin VB.CommandButton cmd_tambah 
         Caption         =   "Tambah"
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
         Left            =   11880
         TabIndex        =   4
         Top             =   6000
         Width           =   1335
      End
      Begin VB.CommandButton cmd_edit 
         Caption         =   "Edit"
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
         Left            =   13320
         TabIndex        =   5
         Top             =   6000
         Width           =   1335
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
         Left            =   13320
         TabIndex        =   6
         Top             =   6000
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   240
      ScaleHeight     =   1305
      ScaleWidth      =   14745
      TabIndex        =   0
      Top             =   120
      Width           =   14775
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
         Height          =   615
         Left            =   12600
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txt_nama_counter 
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
         Left            =   2520
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label lbl_nama_counter 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2520
         TabIndex        =   13
         Top             =   720
         Width           =   5415
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama  Counter"
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
         TabIndex        =   12
         Top             =   720
         Width           =   1500
      End
      Begin VB.Line Line1 
         X1              =   12240
         X2              =   12240
         Y1              =   0
         Y2              =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Counter"
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
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frm_bc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_counter As New XArrayDB
Dim arr_barang As New XArrayDB
Dim id_cntr As String, sql_barang As String, id_brg As String
Dim arah_simpan As Boolean
Dim Moving As Boolean
Dim yold, xold As Long


Private Sub besar_form()
    Me.Height = 7350
    Me.Width = 9585
    Me.ScaleHeight = 6870
    Me.ScaleWidth = 9495
    pic_tambah.Visible = False
    pic_barang.Visible = True
End Sub

Private Sub setengah_form()
    Me.Height = 4530
    Me.Width = 9585
    Me.ScaleHeight = 4050
    Me.ScaleWidth = 9495
    pic_barang.Visible = False
    pic_tambah.Visible = True
End Sub

Private Sub kosong_counter()
    arr_counter.ReDim 0, 0, 0, 0
    grd_counter.ReBind
    grd_counter.Refresh
End Sub

Private Sub kosong_barang()
    arr_barang.ReDim 0, 0, 0, 0
    grd_barang.ReBind
    grd_barang.Refresh
End Sub

Private Sub isi_counter()

On Error GoTo er_isi

    Dim rs_counter As New ADODB.Recordset
    Dim sql As String
        
        kosong_counter
        
        sql = "select id,kode,nama_counter from tbl_counter"
        rs_counter.Open sql, cn, adOpenKeyset
            If Not rs_counter.EOF Then
                
                rs_counter.MoveLast
                rs_counter.MoveFirst
                    
                    lanjut_counter rs_counter
            End If
        rs_counter.Close
        
        Exit Sub
        
er_isi:
        Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub

Private Sub lanjut_counter(rs_counter As Recordset)
    Dim i_c, k_c, n_c As String
    Dim a As Long
        
        a = 1
            Do While Not rs_counter.EOF
                
                arr_counter.ReDim 1, a, 0, 3
                grd_counter.ReBind
                grd_counter.Refresh
                    
                    i_c = rs_counter("id")
                  If Not IsNull(rs_counter("kode")) Then
                    k_c = rs_counter("kode")
                  Else
                    k_c = ""
                  End If
                  If Not IsNull(rs_counter("nama_counter")) Then
                    n_c = rs_counter("nama_counter")
                  Else
                    n_c = ""
                  End If
                    
                arr_counter(a, 0) = i_c
                arr_counter(a, 1) = k_c
                arr_counter(a, 2) = n_c
                
            a = a + 1
            rs_counter.MoveNext
            Loop
            grd_counter.ReBind
            grd_counter.Refresh
                    
End Sub

Private Sub cek_ket_Click()
    If cek_ket.Value = vbChecked Then
        txt_stock_min.Enabled = True
        txt_stock_max.Enabled = True
    End If
    
    If cek_ket.Value = vbUnchecked Then
        txt_stock_min.Text = 0
        txt_stock_max.Text = 0
        txt_stock_min.Enabled = False
        txt_stock_max.Enabled = False
    End If
    
End Sub

Private Sub cmd_cancel_Click()
    'besar_form
    pic_tambah.Visible = False
    Picture1.Enabled = True
    Cmd_Tampil_Click
End Sub

Private Sub cmd_edit_Click()

On Error GoTo er_edit

    If arr_barang.UpperBound(1) > 0 Then
        arah_simpan = False
        pic_tambah.Visible = True
        'pic_tambah.Left = Me.Width / 2 - pic_tambah.Width / 2
        'pic_tambah.Top = Me.Height / 2 - pic_tambah.Height / 2 - 750
        'setengah_form
        Picture1.Enabled = False
        txt_kode_barang.Text = arr_barang(grd_barang.Bookmark, 2)
        txt_nama_barang.Text = arr_barang(grd_barang.Bookmark, 3)
        txt_satuan.Text = arr_barang(grd_barang.Bookmark, 4)
        txt_harga.Text = arr_barang(grd_barang.Bookmark, 5)
        txt_stock_min.Text = arr_barang(grd_barang.Bookmark, 6)
        txt_stock_max.Text = arr_barang(grd_barang.Bookmark, 7)
        Dim ket_c
            ket_c = arr_barang(grd_barang.Bookmark, 8)
                If ket_c = 1 Then
                    cek_ket.Value = vbChecked
                Else
                    cek_ket.Value = vbUnchecked
                End If
                
       Dim aktif_grid As Integer
            aktif_grid = arr_barang(grd_barang.Bookmark, 9)
                
                If aktif_grid = 1 Then
                    cek_aktif.Value = vbChecked
                Else
                    cek_aktif.Value = vbUnchecked
                End If
                
        Dim Discn As Integer
            Discn = arr_barang(grd_barang.Bookmark, 10)
                
                If Discn = 1 Then
                    Check_Diskon.Value = vbChecked
                Else
                    Check_Diskon.Value = vbUnchecked
                End If
                
        cek_ket_Click
        txt_kode_barang.SetFocus
    End If
    
    Exit Sub
    
er_edit:
        Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub cmd_hapus_Click()

On Error GoTo er_h:

Dim sql, sql1 As String
Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    
If arr_barang.UpperBound(1) > 0 Then
    If MsgBox("Yakin akan hapus barang " & arr_barang(grd_barang.Bookmark, 2), vbYesNo + vbQuestion, "Pesan") = vbYes Then
        
        sql = "select id from tbl_barang where id=" & id_brg
        rs.Open sql, cn
            If Not rs.EOF Then
                
                sql1 = "delete from tbl_barang where id=" & id_brg
                rs1.Open sql1, cn
                
            Else
                MsgBox ("Data yang akan dihapus tidak ditemukan")
            End If
            
        rs.Close
        Cmd_Tampil_Click
        Exit Sub
    Else
        Exit Sub
    End If
End If

Exit Sub
    
er_h:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
    
End Sub

Private Sub cmd_ok_Click()

On Error GoTo er_s

Dim sql, sql1 As String
Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    
If txt_kode_barang.Text = "" Then
    MsgBox ("Kode barang harus diisi")
    Exit Sub
End If

If txt_nama_barang.Text = "" Then
    MsgBox ("Nama barang harus diisi")
    Exit Sub
End If

If MsgBox("Yakin semua data yang dimasukkan sudah benar.....?", vbYesNo + vbQuestion, "Pesan") = vbNo Then
    Exit Sub
End If

If arah_simpan = True Then
    
    sql = "select kode from tbl_barang where kode='" & Trim(txt_kode_barang.Text) & "' and id_counter=" & id_cntr
    rs.Open sql, cn
        If Not rs.EOF Then
            If MsgBox("Kode barang " & txt_kode_barang.Text & " sudah ada", vbOKOnly + vbQuestion, "Pesan") = vbOK Then
                Exit Sub
            End If
        Else
            simpan_aja
        End If
    rs.Close
 MsgBox ("Data berhasil_dismpan")
 Cmd_Tampil_Click
 arah_simpan = True
 kosong_text
 txt_kode_barang.SetFocus
 Exit Sub
 
 ElseIf arah_simpan = False Then
    
    sql = "select id from tbl_barang where id=" & id_brg
    rs.Open sql, cn
        If Not rs.EOF Then
        Dim per
            
            If cek_ket.Value = vbChecked Then
                per = 1
            Else
                per = 0
            End If
            
            Dim aktifnya As Integer
                
                If cek_aktif.Value = vbChecked Then
                    aktifnya = 1
                Else
                    aktifnya = 0
                End If
                
            Dim disc As Integer
                If Check_Diskon.Value = vbChecked Then
                    disc = 2
                Else
                    disc = 1
                End If
                
            sql1 = "update tbl_barang set kode='" & Trim(txt_kode_barang.Text) & "',nama_barang='" & Trim(txt_nama_barang.Text) & "',harga_jual=" & CCur(Trim(txt_harga.Text)) & ",stock_min=" & Trim(txt_stock_min.Text) & ",stock_max=" & Trim(txt_stock_max.Text) & ",ket=" & per & ",satuan= '" & Trim(txt_satuan.Text) & "',aktif=" & aktifnya & ",Per_Disc=" & disc & " where id=" & id_brg
            rs1.Open sql1, cn
            
        Else
            
            MsgBox ("Data yang akan diedit tidak ditemukan")
            
        End If
    rs.Close
    cmd_cancel_Click
    Exit Sub
    
End If
    
er_s:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub simpan_aja()
    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    Dim per
    
        If cek_ket.Value = vbChecked Then
            per = 1
        Else
            per = 0
        End If
 
        
        Dim aktif_gak As Integer
            If cek_aktif.Value = vbChecked Then
                aktif_gak = 1
            Else
                aktif_gak = 0
            End If
               
        Dim disc As Integer
            If Check_Diskon.Value = vbChecked Then
                disc = 2
            Else
                disc = 1
            End If
               
     sql = "insert into tbl_barang (id_counter,kode,nama_barang,harga_jual,stock_min,stock_max,ket,satuan,aktif,Per_Disc)"
     sql = sql & " values(" & id_cntr & ",'" & Trim(txt_kode_barang.Text) & "','" & Trim(txt_nama_barang.Text) & "'," & CCur(Trim(txt_harga.Text)) & "," & Trim(txt_stock_min.Text) & "," & Trim(txt_stock_max.Text) & "," & per & ",'" & Trim(txt_satuan.Text) & "'," & aktif_gak & "," & disc & ")"
     rs.Open sql, cn
     
     
     
End Sub

Private Sub cmd_tambah_Click()
If txt_nama_counter.Text <> "" Then
    arah_simpan = True
    pic_tambah.Visible = True
    pic_tambah.Left = Me.Width / 2 - pic_tambah.Width / 2
    pic_tambah.Top = Me.Height / 2 - pic_tambah.Height / 2 - 750
   ' setengah_form
    kosong_text
    
    cek_aktif.Value = vbChecked
    Check_Diskon.Value = vbUnchecked
    
    Picture1.Enabled = False
    
    txt_kode_barang.SetFocus
    
End If
End Sub

Private Sub kosong_text()
    cek_ket.Value = vbUnchecked
    Check_Diskon.Value = vbUnchecked
    txt_nama_barang.Text = ""
    txt_kode_barang.Text = ""
    txt_harga.Text = ""
    txt_stock_min.Text = ""
    txt_stock_min.Text = 0
    txt_stock_max.Text = ""
    txt_stock_max.Text = 0
End Sub

Private Sub Cmd_Tampil_Click()

On Error GoTo er_tampil

    If txt_nama_counter.Text <> "" And id_cntr <> "" Then
        
        Dim rs_barang As New ADODB.Recordset
            
            kosong_barang
            
            sql_barang = "select id,kode,nama_barang,harga_jual,stock_min,stock_max,ket,satuan,aktif,Per_Disc from tbl_barang where id_counter=" & id_cntr
            rs_barang.Open sql_barang, cn, adOpenKeyset
                If Not rs_barang.EOF Then
                    
                    rs_barang.MoveLast
                    rs_barang.MoveFirst
                        
                    lanjut_barang rs_barang
                End If
             rs_barang.Close
    End If
           
    Exit Sub
    
er_tampil:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
           
End Sub

Private Sub lanjut_barang(rs_barang As Recordset)
    
    Dim id_b, kd_b, nm_b, hg_jl, stk_min, stk_max, cek_ket, sat, akt As String
    Dim perdisc As Integer
    Dim a As Long
        
        a = 1
            Do While Not rs_barang.EOF
                arr_barang.ReDim 1, a, 0, grd_barang.Columns.Count
                grd_barang.ReBind
                grd_barang.Refresh
                    
                    id_b = rs_barang("id")
                    
                    If Not IsNull(rs_barang("kode")) Then
                        kd_b = rs_barang("kode")
                    Else
                        kd_b = ""
                    End If
                    
                    If Not IsNull(rs_barang("nama_barang")) Then
                        nm_b = rs_barang("nama_barang")
                    Else
                        nm_b = ""
                    End If
                    
                    If Not IsNull(rs_barang("harga_jual")) Then
                        hg_jl = rs_barang("harga_jual")
                    Else
                        hg_jl = ""
                    End If
                    
                    If Not IsNull(rs_barang("stock_min")) Then
                        stk_min = rs_barang("stock_min")
                    Else
                        stk_min = 0
                    End If
                    
                    If Not IsNull(rs_barang("stock_max")) Then
                        stk_max = rs_barang("stock_max")
                    Else
                        stk_max = 0
                    End If
                    If Not IsNull(rs_barang("ket")) Then
                        cek_ket = rs_barang("ket")
                    Else
                        cek_ket = ""
                    End If
                    
                    If Not IsNull(rs_barang("satuan")) Then
                        sat = rs_barang("satuan")
                    Else
                        sat = ""
                    End If
                    
                    perdisc = IIf(Not IsNull(rs_barang!Per_disc), rs_barang!Per_disc, 0)
                    
                    akt = IIf(Not IsNull(rs_barang!aktif), rs_barang!aktif, 0)
                    
                arr_barang(a, 0) = id_b
                arr_barang(a, 1) = a
                arr_barang(a, 2) = kd_b
                arr_barang(a, 3) = nm_b
                arr_barang(a, 4) = sat
                arr_barang(a, 5) = hg_jl
                arr_barang(a, 6) = stk_min
                arr_barang(a, 7) = stk_max
                arr_barang(a, 8) = cek_ket
                
                If akt = 1 Then
                    arr_barang(a, 9) = vbChecked
                Else
                    arr_barang(a, 9) = vbUnchecked
                End If
                
                If perdisc <> 0 Then
                
                If perdisc = 1 Then
                    arr_barang(a, 10) = vbUnchecked
                Else
                    arr_barang(a, 10) = vbChecked
                End If
                
                End If
                
            a = a + 1
            rs_barang.MoveNext
            
            Loop
            
            grd_barang.ReBind
            grd_barang.Refresh
        
    
End Sub

Private Sub cmd_x_Click()
    pic_counter.Visible = False
    txt_nama_counter.SetFocus
End Sub



Private Sub Form_Load()

    grd_counter.Array = arr_counter
    
    grd_barang.Array = arr_barang
    
    With pic_counter
        .Left = 3240
        .Top = 360
    End With
    
    isi_counter
    
    txt_stock_min.Enabled = False
    txt_stock_max.Enabled = False
    
    kosong_barang
    
    besar_form
    
    Call cari_wewenang("Form Data Barang Counter")
        
        If tambah_form = True Then
            Cmd_Tambah.Enabled = True
        Else
            Cmd_Tambah.Enabled = False
        End If
        
        If edit_form = True Then
            cmd_edit.Enabled = True
        Else
            cmd_edit.Enabled = False
        End If
        
        If hapus_form = True Then
            Cmd_Hapus.Enabled = True
        Else
            Cmd_Hapus.Enabled = False
        End If
        
       
    
End Sub



Private Sub grd_barang_Click()
On Error Resume Next
    If arr_barang.UpperBound(1) > 0 Then
        id_brg = arr_barang(grd_barang.Bookmark, 0)
    End If
End Sub

Private Sub grd_barang_HeadClick(ByVal ColIndex As Integer)
    
On Error GoTo er_head
    
If arr_barang.UpperBound(1) > 0 Then

    Dim rs_barang As New ADODB.Recordset
    Dim sql As String
    
    If sql_barang = "" Then
        Exit Sub
    End If
    
    sql = sql_barang
    
    Select Case ColIndex
        
        Case 2
            
            sql = sql & " order by kode"
            
        Case 3
            
            sql = sql & " order by nama_barang"
            
        Case 5
            
            sql = sql & " order by harga_jual"
            
        Case 6
            
            sql = sql & " order by stock_min"
            
        Case 7
            
            sql = sql & " order by stock_max"
            
    End Select
    
    rs_barang.Open sql, cn, adOpenKeyset
    If Not rs_barang.EOF Then
        
        rs_barang.MoveLast
        rs_barang.MoveFirst
        
        lanjut_barang rs_barang
    End If
    rs_barang.Close
    
End If

Exit Sub

er_head:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear

End Sub

Private Sub grd_barang_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_barang_Click
End Sub

Private Sub grd_counter_Click()
    On Error Resume Next
        If arr_counter.UpperBound(1) > 0 Then
            id_cntr = arr_counter(grd_counter.Bookmark, 0)
        End If
End Sub

Private Sub grd_counter_DblClick()
If arr_counter.UpperBound(1) > 0 Then
    
    lbl_kode.Caption = arr_counter(grd_counter.Bookmark, 1)
    txt_nama_counter.Text = arr_counter(grd_counter.Bookmark, 1)
    lbl_nama_counter.Caption = arr_counter(grd_counter.Bookmark, 2)
    pic_counter.Visible = False
    txt_nama_counter.SetFocus
End If
    
End Sub

Private Sub grd_counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        grd_counter_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        pic_counter.Visible = False
        txt_nama_counter.SetFocus
    End If
End Sub

Private Sub grd_counter_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_counter_Click
End Sub

Private Sub pic_counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_counter.Visible = False
        txt_nama_counter.SetFocus
    End If
End Sub

Private Sub pic_counter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub pic_counter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   pic_counter.Top = pic_counter.Top - (yold - Y)
   pic_counter.Left = pic_counter.Left - (xold - X)
End If

End Sub

Private Sub pic_counter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub pic_tambah_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub pic_tambah_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   pic_tambah.Top = pic_tambah.Top - (yold - Y)
   pic_tambah.Left = pic_tambah.Left - (xold - X)
End If

End Sub

Private Sub pic_tambah_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub txt_harga_GotFocus()
    txt_harga.SelStart = 0
    txt_harga.SelLength = Len(txt_harga)
End Sub

Private Sub txt_harga_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txt_harga_KeyUp(KeyCode As Integer, Shift As Integer)
    txt_harga.Text = Format(txt_harga.Text, "###,###,###")
    txt_harga.SelStart = Len(txt_harga.Text)
End Sub

Private Sub txt_kode_barang_GotFocus()
    txt_kode_barang.SelStart = 0
    txt_kode_barang.SelLength = Len(txt_kode_barang)
End Sub

Private Sub txt_nama_barang_GotFocus()
    txt_nama_barang.SelStart = 0
    txt_nama_barang.SelLength = Len(txt_nama_barang)
End Sub


Private Sub txt_nama_counter_GotFocus()
    txt_nama_counter.SelStart = 0
    txt_nama_counter.SelLength = Len(txt_nama_counter)
End Sub

Private Sub txt_nama_counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        kosong_counter
        txt_nama_counter.Text = ""
        pic_counter.Visible = True
        txt_nm(0).Text = ""
        txt_nm(1).Text = ""
        txt_nm(1).SetFocus
    End If
End Sub

Private Sub txt_nama_counter_LostFocus()

On Error GoTo er_nama

If txt_nama_counter.Text <> "" Then

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        sql = "select id,kode,nama_counter from tbl_counter where kode='" & Trim(txt_nama_counter.Text) & "'"
        rs.Open sql, cn
            If Not rs.EOF Then
                id_cntr = rs("id")
                lbl_kode.Caption = rs("kode")
                lbl_nama_counter.Caption = rs("nama_counter")
            Else
                MsgBox ("Kode counter yang anda masukkan tidak ditemukan")
                txt_nama_counter.SetFocus
            End If
        rs.Close
        
End If
Exit Sub

er_nama:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear

End Sub

Private Sub txt_nm_GotFocus(Index As Integer)
    Select Case Index
        Case 0
            txt_nm(0).SelStart = 0
            txt_nm(0).SelLength = Len(txt_nm(0))
        Case 1
            txt_nm(1).SelStart = 0
            txt_nm(1).SelLength = Len(txt_nm(1))
    End Select
End Sub

Private Sub txt_nm_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_counter.Visible = False
        txt_nama_counter.SetFocus
    End If
    
    If KeyCode = 13 Then
        lbl_kode.Caption = arr_counter(grd_counter.Bookmark, 1)
        txt_nama_counter.Text = arr_counter(grd_counter.Bookmark, 1)
        lbl_nama_counter.Caption = arr_counter(grd_counter.Bookmark, 2)
        pic_counter.Visible = False
        txt_nama_counter.SetFocus
    End If
End Sub

Private Sub txt_nm_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo er_nm

        Dim sql As String
        Dim rs_counter As New ADODB.Recordset
            
      sql = "select id,kode,nama_counter from tbl_counter"
            
      Select Case Index
      
      Case 0
            sql = sql & " where nama_counter like '%" & Trim(txt_nm(0).Text) & "%'"
      Case 1
            sql = sql & " where kode like '%" & Trim(txt_nm(1).Text) & "%'"
      End Select
      
            rs_counter.Open sql, cn, adOpenKeyset
                If Not rs_counter.EOF Then
                    
                    rs_counter.MoveLast
                    rs_counter.MoveFirst
                    
                    lanjut_counter rs_counter
                End If
            rs_counter.Close
     
     Exit Sub
     
er_nm:
        Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
     
End Sub

Private Sub txt_satuan_GotFocus()
    txt_satuan.SelStart = 0
    txt_satuan.SelLength = Len(txt_satuan)
End Sub

Private Sub txt_stock_max_GotFocus()
    txt_stock_max.SelStart = 0
    txt_stock_max.SelLength = Len(txt_stock_max)
End Sub

Private Sub txt_stock_max_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txt_stock_min_GotFocus()
    txt_stock_min.SelStart = 0
    txt_stock_min.SelLength = Len(txt_stock_min)
End Sub

Private Sub txt_stock_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub
