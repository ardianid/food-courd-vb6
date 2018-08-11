VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_penjualan 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic_counter 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   -3120
      ScaleHeight     =   5625
      ScaleWidth      =   5865
      TabIndex        =   46
      Top             =   3720
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txt_counter 
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
         Index           =   1
         Left            =   2520
         TabIndex        =   51
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txt_counter 
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
         Index           =   0
         Left            =   120
         TabIndex        =   50
         Top             =   840
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
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
         Left            =   5400
         TabIndex        =   53
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   5865
         TabIndex        =   47
         Top             =   0
         Width           =   5895
      End
      Begin TrueDBGrid60.TDBGrid grd_counter 
         Height          =   4335
         Left            =   120
         OleObjectBlob   =   "frm_penjualan.frx":0000
         TabIndex        =   52
         Top             =   1200
         Width           =   5655
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Counter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Counter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   48
         Top             =   480
         Width           =   3735
      End
   End
   Begin VB.TextBox txt_jml_bayar 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   9720
      TabIndex        =   44
      Text            =   "txt_jml_bayar"
      Top             =   9120
      Width           =   5415
   End
   Begin VB.TextBox txt_kode_counter 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5520
      TabIndex        =   20
      Top             =   2400
      Width           =   2055
   End
   Begin VB.PictureBox pic_barang 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   -3000
      ScaleHeight     =   5625
      ScaleWidth      =   5865
      TabIndex        =   9
      Top             =   5520
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton Command2 
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
         Left            =   5400
         TabIndex        =   14
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   5865
         TabIndex        =   15
         Top             =   0
         Width           =   5895
      End
      Begin VB.TextBox txt 
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
         Index           =   1
         Left            =   2040
         TabIndex        =   12
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txt 
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
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1935
      End
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
         Left            =   5400
         TabIndex        =   10
         Top             =   0
         Width           =   495
      End
      Begin TrueDBGrid60.TDBGrid grd_barang 
         Height          =   4335
         Left            =   120
         OleObjectBlob   =   "frm_penjualan.frx":2E15
         TabIndex        =   13
         Top             =   1200
         Width           =   5655
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.TextBox txt_faktur 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5520
      TabIndex        =   6
      Top             =   1920
      Width           =   3375
   End
   Begin VB.PictureBox pic_samping 
      Height          =   7335
      Left            =   0
      ScaleHeight     =   7275
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   840
      Width           =   3015
      Begin VB.PictureBox Picture3 
         Height          =   3015
         Left            =   120
         ScaleHeight     =   2955
         ScaleWidth      =   2595
         TabIndex        =   5
         Top             =   1200
         Width           =   2655
         Begin VB.Image img_foto 
            Height          =   3015
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2655
         End
      End
      Begin VB.Label lbl_user 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   480
         TabIndex        =   4
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lbl_tgl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_tgl"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label lbl_jam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_jam"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   1050
      End
      Begin VB.Image img2 
         Height          =   1935
         Left            =   0
         Stretch         =   -1  'True
         Top             =   -1680
         Width           =   2895
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   240
      Top             =   480
   End
   Begin TrueDBGrid60.TDBGrid grd_daftar 
      Height          =   3015
      Left            =   3120
      OleObjectBlob   =   "frm_penjualan.frx":5AF1
      TabIndex        =   8
      Top             =   5280
      Width           =   12135
   End
   Begin MSComCtl2.DTPicker dtp_tgl 
      Height          =   375
      Left            =   11280
      TabIndex        =   18
      Top             =   1920
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   57737217
      CurrentDate     =   38622
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   13035
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      Begin VB.Image img 
         Height          =   495
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox txt_beli 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4440
      TabIndex        =   28
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox txt_disc 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4440
      TabIndex        =   30
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txt_charge 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4440
      TabIndex        =   33
      Top             =   4680
      Width           =   735
   End
   Begin VB.CheckBox cek_faktur 
      BackColor       =   &H00800000&
      Caption         =   "&c"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   43
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox txt_kode_barang 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5520
      TabIndex        =   21
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lbl_grand_total 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   11280
      TabIndex        =   55
      Top             =   4440
      Width           =   3975
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   8640
      TabIndex        =   54
      Top             =   4560
      Width           =   1845
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Cetak Faktur"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   6240
      TabIndex        =   45
      Top             =   3840
      Width           =   2040
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Kembali"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   675
      Left            =   5160
      TabIndex        =   42
      Top             =   9960
      Width           =   4230
   End
   Begin VB.Label lbl_kembali 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lbl_kembali"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   9720
      TabIndex        =   41
      Top             =   9960
      Width           =   5415
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Bayar"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   675
      Left            =   5760
      TabIndex        =   40
      Top             =   9120
      Width           =   3585
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Harga"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   675
      Left            =   6360
      TabIndex        =   39
      Top             =   8280
      Width           =   3075
   End
   Begin VB.Label lbl_total_bayar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lbl_total_bayar"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   9720
      TabIndex        =   38
      Top             =   8280
      Width           =   5415
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Satuan"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   8640
      TabIndex        =   37
      Top             =   3840
      Width           =   2145
   End
   Begin VB.Label lbl_harga 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   11280
      TabIndex        =   36
      Top             =   3720
      Width           =   3975
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Charge"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   3120
      TabIndex        =   35
      Top             =   4800
      Width           =   1125
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   5280
      TabIndex        =   34
      Top             =   4680
      Width           =   285
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   5280
      TabIndex        =   32
      Top             =   4200
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disc"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   3120
      TabIndex        =   31
      Top             =   4320
      Width           =   690
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   3120
      TabIndex        =   29
      Top             =   3720
      Width           =   540
   End
   Begin VB.Label lbl_nama_counter 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   27
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Counter"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   9120
      TabIndex        =   26
      Top             =   2400
      Width           =   2250
   End
   Begin VB.Label lbl_nama_barang 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   25
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama  Barang"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   9120
      TabIndex        =   24
      Top             =   3000
      Width           =   2220
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode counter"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   3120
      TabIndex        =   23
      Top             =   2400
      Width           =   2130
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Barang"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   3120
      TabIndex        =   22
      Top             =   3000
      Width           =   2040
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   9120
      TabIndex        =   19
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Faktur"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   3120
      TabIndex        =   7
      Top             =   1920
      Width           =   1545
   End
   Begin VB.Image img_dasar 
      Height          =   735
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "frm_penjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB
Dim arr_barang As New XArrayDB
Dim arr_counter As New XArrayDB
Dim kode_barang As String, id_barang As String, st_s, stk As Double, ket_b As Boolean
Dim id_counter As String, uang_disc, uang_charge As Double
Dim sementara, s_disc, s_charge As String
    
    

Private Sub kosong_counter()
    arr_counter.ReDim 0, 0, 0, 0
    grd_counter.ReBind
    grd_counter.Refresh
End Sub

Private Sub isi_counter()
    Dim sql As String
    Dim rs_counter  As New ADODB.Recordset
        
        kosong_counter
        
        sql = "select id,kode,nama_counter from tbl_counter order by kode"
        rs_counter.Open sql, cn, adOpenKeyset
            If Not rs_counter.EOF Then
                
                rs_counter.MoveLast
                rs_counter.MoveFirst
                
                lanjut_counter rs_counter
            End If
        rs_counter.Close
End Sub

Private Sub lanjut_counter(rs_counter As Recordset)
    Dim id_c, kd_c, nm_c As String
    Dim a As Long
        
        a = 1
            Do While Not rs_counter.EOF
                arr_counter.ReDim 1, a, 0, 3
                grd_counter.ReBind
                grd_counter.Refresh
                    
                    id_c = rs_counter("id")
                    kd_c = rs_counter("kode")
                    nm_c = rs_counter("nama_counter")
                    
                arr_counter(a, 0) = id_c
                arr_counter(a, 1) = kd_c
                arr_counter(a, 2) = nm_c
                
            a = a + 1
            rs_counter.MoveNext
            Loop
            grd_counter.ReBind
            grd_counter.Refresh
End Sub

Private Sub cmd_faktur_Click()

On Error GoTo er_print



Dim a As Long
Dim grs
    Printer.Font = "Arial"
    
        
        Printer.CurrentX = 0
        Printer.CurrentY = 0
            
            Printer.Print
            Printer.Print Tab((55 / 2 - Len("Bukti Pembayaran") / 2) - 25); "Bukti Pembayaran"
            Printer.Print Tab((55 / 2 - Len("Simpur Food Center") / 2) - 25); "Simpur Food Center"
            Printer.Print
            
            grs = String$(55, "-")
            
            Printer.Print grs
            
            Printer.Print "Tgl. " & dtp_tgl.Value; Tab(21); "Jam. " & lbl_jam.Caption
            Printer.Print grs
            Printer.Print "No Faktur " & Trim(txt_faktur.Text)
            Printer.Print
            Printer.Print "Qty"; Tab(5); "Nama Barang"; Tab(25); "Harga"
            Printer.Print grs
            
       For a = 1 To arr_daftar.UpperBound(1)
            
            Printer.Print arr_daftar(a, 3); Tab(5); arr_daftar(a, 2); Tab(25); Space(Len(arr_daftar(a, 4)) - Len(arr_daftar(a, 4))) + arr_daftar(a, 4)
            
      Next a
            
            Printer.Print grs
            Printer.Print "Discount"; Tab(11); grd_daftar.Columns(6).FooterText
            Printer.Print "Cash"; Tab(11); grd_daftar.Columns(8).FooterText
            Printer.Print grs
            Printer.Print "Total"; Tab(25); Space(Len(grd_daftar.Columns(10).FooterText) - Len(grd_daftar.Columns(10).FooterText)) + grd_daftar.Columns(10).FooterText
            Printer.Print grs
            Printer.Print "jml Bayar "; Tab(25); Format(Space(Len(txt_bayar.Text) - Len(txt_bayar.Text)) + txt_bayar.Text, "Currency")
            Printer.Print "Kembali"; Tab(25); Space(Len(lbl_kembali.Caption) - Len(lbl_kembali.Caption)) + lbl_kembali.Caption
            Printer.Print
            Printer.Print " *********** Terima Kasih ***********"
            
      
      Printer.EndDoc
      Exit Sub
            
er_print:
           Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
            
End Sub

Private Sub kosong_aja_semua()
       lbl_total_bayar.Caption = ""
       txt_bayar.Text = ""
       lbl_kembali.Caption = ""
       pic_bayar.Visible = False
       kosong1
       kosong_dc
       lbl_total.Caption = ""
       kosong_daftar
       txt_kode_barang.SetFocus
End Sub

Private Sub cmd_hapus_Click()
    If MsgBox("Yakin akan dihapus.....?", vbYesNo + vbQuestion, "Pesan") = vbNo Then
        Exit Sub
    End If
    
If arr_daftar.UpperBound(1) > 1 Then
    lbl_total.Caption = Format(CDbl(lbl_total.Caption) - CDbl(grd_daftar.Columns(10).Text), "Currency")
    grd_daftar.Delete
Else
    lbl_total.Caption = ""
    kosong_daftar
End If
    grd_daftar.ReBind
    grd_daftar.Refresh
    
End Sub

Private Sub cmd_ok_Click()

On Error GoTo er_ok

    Dim sql, sql1, sql2 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset
    Dim a As Long
        
 If txt_faktur.Text = "" Then
    MsgBox ("No Faktur tidak boleh kosong")
    txt_faktur.SetFocus
    Exit Sub
 End If
        
 If CDbl(txt_bayar.Text) < CDbl(lbl_total_bayar.Caption) Then
    MsgBox ("Jumlah uang tidak boleh kurang dari total bayar")
    Exit Sub
 End If
        
 If MsgBox("Yakin semua data yang dimasukkan sudah benar", vbYesNo + vbQuestion, "Pesan") = vbNo Then
    Exit Sub
 End If
        
     cn.BeginTrans
        
    sql = "select no_faktur from tr_penjualan where no_faktur='" & Trim(txt_faktur.Text) & "'"
    rs.Open sql, cn
        If Not rs.EOF Then
            
                MsgBox ("No Faktur " & Trim(txt_faktur.Text) & " Sudah Ada")
                cn.RollbackTrans
                pic_bayar.Visible = False
                Exit Sub
        End If
    rs.Close
    
    sql1 = "insert into tr_faktur_penjualan (no_faktur,tgl,jam)"
    sql1 = sql1 & " values('" & Trim(txt_faktur.Text) & "','" & Trim(dtp_tgl.Value) & "','" & Trim(frm_jam.Caption) & "')"
    rs1.Open sql1, cn
    
    For a = 1 To arr_daftar.UpperBound(1)
           
            stock_sekarang (arr_daftar(a, 11))
            If ket_b = False Then
                Exit Sub
            End If
                
                If st_s <> 0 Then
                    If (stk - CDbl(arr_daftar(a, 3))) < st_s Then
                        Dim jangan
                        jangan = MsgBox("Stock barang tidak mencukupi untuk memenuhi penjualan" & Chr(13) & "Stock Sekarang " & stk & Chr(13) & "Stock Min " & st_s & Chr(13) & "Jml Beli " & arr_daftar(a, 3), vbOKOnly + vbInformation, "Pesan")
                        cn.RollbackTrans
                        Exit Sub
                    End If
                 End If
            
            ' isi transaksi penjualan
            
                sql1 = "insert into tr_penjualan (no_faktur,id_barang,qty,harga_satuan,harga_sebenarnya,disc,harga_disc,cash,harga_cash,total_harga)"
                sql1 = sql1 & " values('" & Trim(txt_faktur.Text) & "'," & arr_daftar(a, 11) & "," & arr_daftar(a, 3) & "," & CCur(arr_daftar(a, 4)) & ", " & CCur(arr_daftar(a, 5)) & ",'" & arr_daftar(a, 6) & "'," & CCur(arr_daftar(a, 7)) & ",'" & arr_daftar(a, 8) & "'," & CCur(arr_daftar(a, 9)) & "," & CCur(arr_daftar(a, 10)) & ")"
                rs1.Open sql1, cn
                
            ' seleksi dari tr_jml_stok
                
                sql2 = "select jml_stock from tr_jml_stock where id_barang=" & arr_daftar(a, 11)
                rs2.Open sql2, cn
                    If Not rs2.EOF Then
                        Dim j_stock_sekarang As Double
                        
            'update tbl_jml_stock ( kalau ada)
            
                            j_stock_sekarang = CDbl(rs2("jml_stock")) - CDbl(arr_daftar(a, 3))
                            sql1 = "update tr_jml_stock set jml_stock=" & j_stock_sekarang & " where id_barang=" & arr_daftar(a, 11)
                            rs1.Open sql1, cn
                            
            ' isi transaksi stok_barang
                
                            sql1 = "insert into tr_stock (id_barang,brg_in,brg_out,tgl,ket)"
                            sql1 = sql1 & " values(" & arr_daftar(a, 11) & ",0," & arr_daftar(a, 3) & ",'" & Trim(dtp_tgl.Value) & "',0)"
                            rs1.Open sql1, cn
                            
                    End If
                rs2.Close
    Next a
        
    MsgBox ("Data berhasil disimpan")
    cn.CommitTrans
    pic_bayar.Visible = False
    cmd_faktur.Enabled = True
    Exit Sub
    
er_ok:
    cn.RollbackTrans
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub

Private Sub stock_sekarang(id_b As String)
    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        sql = "select stock_min,jml_stock from qr_jml_stock where id_barang=" & id_b
        rs.Open sql, cn
            If Not rs.EOF Then
                If CDbl(rs("jml_stock")) = CDbl(rs("stock_min")) Then
                    MsgBox ("Jumlah stok hampir mendekati batas minimum")
                    ket_b = False
                    Exit Sub
                End If
                st_s = rs("stock_min")
                stk = rs("jml_stock")
                ket_b = True
            Else
                ket_b = True
                st_s = 0
            End If
        rs.Close
End Sub

Private Sub cmd_x_Click()
    pic_barang.Visible = False
End Sub

Private Sub Command1_Click()
    pic_counter.Visible = False
End Sub

Private Sub Command2_Click()
    pic_barang.Visible = False
End Sub

Private Sub dtp_tgl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_kode_counter.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    txt_kode_counter.SetFocus
End Sub

Private Sub Form_Load()
    
    buka_path
    
    buka_koneksi
    
    grd_daftar.Array = arr_daftar
        
    grd_counter.Array = arr_counter
        
    grd_barang.Array = arr_barang
        
    kosong_daftar
    
    isi_counter
    
    kosong_barang

    dtp_tgl.Value = Format(Date, "long date")
    img_foto.Picture = LoadPicture("D:\foto\09.jpg")
        
    lbl_harga.Caption = 0
    txt_beli.Text = 0
    txt_disc.Text = 0
    txt_charge.Text = 0
    lbl_grand_total.Caption = 0
    lbl_total_bayar.Caption = 0
    txt_jml_bayar.Text = 0
    lbl_kembali.Caption = 0
        
End Sub

Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub

Sub kosong_barang()
    arr_barang.ReDim 0, 0, 0, 0
    grd_barang.ReBind
    grd_barang.Refresh
End Sub
    
Private Sub isi_barang()
    Dim sql As String
    Dim rs_barang As New ADODB.Recordset
        
        kosong_barang
        
        sql = "select nama_counter,kode,nama_barang from qr_barang where id_counter=" & id_counter & "  order by kode"
        rs_barang.Open sql, cn, adOpenKeyset
            If Not rs_barang.EOF Then
                
                rs_barang.MoveLast
                rs_barang.MoveFirst
                
                lanjut_barang rs_barang
            End If
        rs_barang.Close
        
End Sub

Private Sub lanjut_barang(rs_barang As Recordset)
    Dim counter, kode, barang As String
    Dim a As Long
        
        a = 1
            Do While Not rs_barang.EOF
                arr_barang.ReDim 1, a, 0, 3
                grd_barang.ReBind
                grd_barang.Refresh
                    
                    If Not IsNull(rs_barang("nama_counter")) Then
                        counter = rs_barang("nama_counter")
                    Else
                        counter = ""
                    End If
                    
                    If Not IsNull(rs_barang("kode")) Then
                        kode = rs_barang("kode")
                    Else
                        kode = ""
                    End If
                    
                    If Not IsNull(rs_barang("nama_barang")) Then
                        barang = rs_barang("nama_barang")
                    Else
                        barang = ""
                    End If
                    
                arr_barang(a, 0) = counter
                arr_barang(a, 1) = kode
                arr_barang(a, 2) = barang
                
            a = a + 1
            rs_barang.MoveNext
            Loop
            grd_barang.ReBind
            grd_barang.Refresh
        
End Sub

Private Sub Form_Resize()
    Picture1.Left = Me.Left
    Picture1.Top = Me.Top
    Picture1.Width = Me.Width
    Picture1.Height = 1750
    img.Left = Picture1.Left
    img.Top = Picture1.Top
    img.Width = Picture1.Width
    img.Height = Picture1.Height
    img.Picture = LoadPicture(App.path & "\BANNER.jpg")
    
    pic_samping.Move Me.Left, Picture1.Height - 50, 3000, Me.Height - Picture1.Height
    img2.Move pic_samping.Left, 0, pic_samping.Width, pic_samping.Height
    img2.Picture = LoadPicture(App.path & "\banner3.1.jpg")
    img2.ZOrder 1
    lbl_jam.Move img2.Left + 150, img2.Top + 100, 0, 0
    lbl_tgl.Move img2.Left + 150, lbl_jam.Top + 300, 0, 0
    lbl_tgl.Caption = Format(Date, "long date")
    
    lbl_user.Left = img2.Left + 150
    lbl_user.Top = lbl_tgl.Top + 500
    
    img_dasar.Left = Me.ScaleLeft
    img_dasar.Width = Me.ScaleWidth
    img_dasar.Width = Me.ScaleWidth
    img_dasar.Height = Me.ScaleHeight
    img_dasar.Picture = LoadPicture(App.path & "\3.jpg")
    img_dasar.ZOrder 1
        
  
        
    End Sub

Private Sub grd_barang_Click()
    On Error Resume Next
        If arr_barang.UpperBound(1) > 0 Then
            kode_barang = arr_barang(grd_barang.Bookmark, 1)
        End If
End Sub

Private Sub grd_barang_DblClick()
 If arr_barang.UpperBound(1) > 0 Then
    txt_kode_barang.Text = kode_barang
    kasih_tahu
    pic_barang.Visible = False
    txt_kode_barang.SetFocus
 End If
End Sub

Private Sub kasih_tahu()
    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        sql = "select nama_counter,nama_barang,harga_jual,id_barang from qr_barang where kode='" & Trim(txt_kode_barang.Text) & "' and id_counter=" & id_counter
        rs.Open sql, cn
            If Not rs.EOF Then
                id_barang = rs("id_barang")
                lbl_nama_barang.Caption = rs("nama_barang")
                lbl_harga.Caption = Format(rs("harga_jual"), "Currency")
            Else
                MsgBox ("Kode barang yang anda masukkan tidak ditemukan")
                txt_kode_barang.SetFocus
            End If
        rs.Close
End Sub

Private Sub grd_barang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        grd_barang_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        pic_barang.Visible = False
    End If
    
End Sub

Private Sub grd_barang_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_barang_Click
End Sub

Private Sub grd_counter_Click()
    On Error Resume Next
        If arr_counter.UpperBound(1) > 0 Then
            id_counter = arr_counter(grd_counter.Bookmark, 0)
        End If
End Sub

Private Sub grd_counter_DblClick()
    If arr_counter.UpperBound(1) > 0 Then
        txt_kode_counter.Text = arr_counter(grd_counter.Bookmark, 1)
        lbl_nama_counter.Caption = arr_counter(grd_counter.Bookmark, 2)
        pic_counter.Visible = False
        txt_kode_counter.SetFocus
        isi_barang
     End If
End Sub

Private Sub grd_counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_counter.Visible = False
    End If
    If KeyCode = 13 Then
        grd_counter_DblClick
    End If
End Sub

Private Sub grd_counter_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_counter_Click
End Sub

Private Sub opt_cash_Click()
    If opt_cash.Value = True Then
        frm_pakai.Enabled = True
        frm_pakai.Caption = "Cash"
        Label12.Caption = "Jl Beli x Harga Satuan" & "+ Cash"
        kosong_dc
    End If
End Sub

Private Sub opt_discount_Click()
    If opt_discount.Value = True Then
        frm_pakai.Enabled = True
        frm_pakai.Caption = "Discount"
        Label12.Caption = "Jl Beli x Harga Satuan" & "- Discount"
        kosong_dc
    End If
End Sub

Private Sub opt_tidak_Click()
    If opt_tidak.Value = True Then
        frm_pakai.Enabled = False
        Label12.Caption = "Jl Beli x Harga Satuan"
        frm_pakai.Caption = ""
        kosong_dc
    End If
End Sub

Private Sub grd_daftar_Click()
    On Error Resume Next
        If arr_daftar.UpperBound(1) > 0 Then
            sementara = arr_daftar(grd_daftar.Bookmark, 10)
            s_disc = arr_daftar(grd_daftar.Bookmark, 6)
            s_charge = arr_daftar(grd_daftar.Bookmark, 8)
        End If
End Sub

Private Sub grd_daftar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If arr_daftar.UpperBound(1) > 1 Then
            grd_daftar.Delete
        Else
            arr_daftar.ReDim 0, 0, 0, 0
        End If
    grd_daftar.ReBind
    grd_daftar.Refresh
    Dim jml, jml_d, jml_c As Double
    
        jml = CDbl(grd_daftar.Columns(10).FooterText) - CDbl(sementara)
        grd_daftar.Columns(10).FooterText = Format(jml, "currency")
        lbl_total_bayar.Caption = Format(jml, "currency")
        txt_jml_bayar.Text = Format(jml, "###,###,###")
        lbl_kembali.Caption = "Rp." & Format(jml, "###,###,###")
        
        
        jml_d = Mid(grd_daftar.Columns(6).FooterText, 1, Len(grd_daftar.Columns(6).FooterText) - 1)
        s_disc = Mid(s_disc, 1, Len(s_disc) - 1)
        jml_d = CDbl(jml_d) - CDbl(s_disc)
        grd_daftar.Columns(6).FooterText = jml_d & "%"
        
        jml_c = Mid(grd_daftar.Columns(8).FooterText, 1, Len(grd_daftar.Columns(8).FooterText) - 1)
        s_charge = Mid(s_charge, 1, Len(s_charge) - 1)
        jml_c = CDbl(jml_c) - CDbl(s_charge)
        grd_daftar.Columns(8).FooterText = jml_c & "%"
        
    End If
End Sub

Private Sub grd_daftar_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_daftar_Click
End Sub

Private Sub pic_counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_counter.Visible = False
    End If
End Sub

Private Sub Timer1_Timer()
    
    lbl_jam.Caption = Format(Time, "hh:mm:ss")
End Sub



Private Sub txt_bayar_GotFocus()
    cmd_ok.Default = True
End Sub

Private Sub txt_bayar_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txt_bayar_KeyUp(KeyCode As Integer, Shift As Integer)

lbl_kembali.Caption = 0

    If txt_bayar.Text <> "" Then
                
            
            txt_bayar.Text = Format(txt_bayar.Text, "###,###,###")
            txt_bayar.SelStart = Len(txt_bayar.Text)
                
            Dim cari_kembali As Double
            
                cari_kembali = CDbl(txt_bayar.Text) - CDbl(lbl_total_bayar.Caption)
                
                lbl_kembali.Caption = Format(cari_kembali, "##,###,###")
                
                
                
    End If
End Sub

Private Sub txt_bayar_LostFocus()
    cmd_ok.Default = False
End Sub

Private Sub txt_beli_GotFocus()
    txt_beli.SelStart = 0
    txt_beli.SelLength = Len(txt_beli)
End Sub

Private Sub isi_daftar_belanjaan()
    arr_daftar.ReDim 1, arr_daftar.UpperBound(1) + 1, 0, 12
    grd_daftar.ReBind
    grd_daftar.Refresh
        
        Dim jml_baris As Long
            
            jml_baris = arr_daftar.UpperBound(1)
            
            arr_daftar(jml_baris, 0) = Trim(txt_faktur.Text)
            arr_daftar(jml_baris, 1) = Trim(txt_kode_barang.Text)
            arr_daftar(jml_baris, 2) = Trim(lbl_nama_barang.Caption)
            arr_daftar(jml_baris, 3) = Trim(txt_beli.Text)
            arr_daftar(jml_baris, 4) = Trim(lbl_harga.Caption)
            arr_daftar(jml_baris, 5) = CDbl(lbl_harga.Caption) * CDbl(txt_beli.Text)
            arr_daftar(jml_baris, 6) = Trim(txt_disc.Text) & "%"
            arr_daftar(jml_baris, 7) = uang_disc
            arr_daftar(jml_baris, 8) = Trim(txt_charge.Text) & "%"
            arr_daftar(jml_baris, 9) = uang_charge
            arr_daftar(jml_baris, 10) = Trim(lbl_grand_total.Caption)
            arr_daftar(jml_baris, 11) = id_barang
            
            Dim jml_diskon, jml_cash, jml_biaya As Double
            
            If grd_daftar.Columns(10).FooterText = "" Then
                grd_daftar.Columns(10).FooterText = 0
            End If
            
            jml_biaya = CDbl(lbl_grand_total.Caption) + CDbl(grd_daftar.Columns(10).FooterText)
            grd_daftar.Columns(10).FooterText = Format(jml_biaya, "Currency")
            lbl_total_bayar.Caption = Format(jml_biaya, "currency")
            txt_jml_bayar.Text = Format(jml_biaya, "###,###,###")
            lbl_kembali.Caption = "Rp." & Format(jml_biaya, "###,###,###")
                     
            If grd_daftar.Columns(6).FooterText = "" Then
                grd_daftar.Columns(6).FooterText = 0 & "%"
                jml_diskon = 0
            End If
            
            
                jml_diskon = CDbl(txt_disc.Text) + CDbl(Mid(grd_daftar.Columns(6).FooterText, 1, Len(grd_daftar.Columns(6).FooterText) - 1))
                grd_daftar.Columns(6).FooterText = jml_diskon & "%"
            

            If grd_daftar.Columns(8).FooterText = "" Then
                grd_daftar.Columns(8).FooterText = 0 & "%"
                jml_cash = 0
            End If
            
                jml_cash = CDbl(txt_charge.Text) + CDbl(Mid(grd_daftar.Columns(8).FooterText, 1, Len(grd_daftar.Columns(8).FooterText) - 1))
                grd_daftar.Columns(8).FooterText = jml_cash & "%"
           
            
            
            
            
     grd_daftar.ReBind
     grd_daftar.Refresh
End Sub

Private Sub kosong1()
    txt_kode_barang.Text = ""
    lbl_nama_barang.Caption = ""
    lbl_nama_counter.Caption = ""
    lbl_harga.Caption = ""
    txt_beli.Text = ""
    lbl_jumlah.Caption = ""
End Sub

Private Sub txt_beli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And txt_beli.Text <> "" Then
        txt_disc.SetFocus
    Else
        txt_beli.SetFocus
    End If
End Sub

Private Sub txt_beli_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txt_beli_KeyUp(KeyCode As Integer, Shift As Integer)
    lbl_grand_total.Caption = 0
    If txt_beli.Text <> "" Then
        Dim grand As Double
            grand = CDbl(txt_beli.Text) * CDbl(lbl_harga.Caption)
            grand = grand + CDbl(lbl_grand_total.Caption)
            lbl_grand_total.Caption = Format(grand, "Currency")
    End If
End Sub

Private Sub txt_discount1_GotFocus()
    txt_discount1.SelStart = 0
    txt_discount1.SelLength = Len(txt_discount1)
End Sub

Private Sub txt_discount1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_discount2.SetFocus
    End If
End Sub

Private Sub txt_discount1_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",")) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txt_discount1_KeyUp(KeyCode As Integer, Shift As Integer)
    
 
    
    lbl_jumlah.Caption = ""
    txt_discount2.Text = ""

    If txt_beli.Text <> "" And lbl_harga.Caption <> "" And txt_discount1.Text <> "" Then
        Dim persen, harga As Currency
            harga = Trim(lbl_harga.Caption)
            persen = Trim(txt_discount1.Text)
        txt_discount2.Text = Val(harga) * (Val(persen) / 100)
        
        
        If opt_discount.Value = True Then
            hitung_jumlah (True)
        ElseIf opt_cash.Value = True Then
            hitung_jumlah (False)
        End If
               
    End If
        
 
    
End Sub
Private Sub hitung_jumlah(jml As Boolean)
    
    Dim jumlah, disc As Double
    jumlah = CDbl(lbl_harga.Caption) * CDbl(txt_beli.Text)
    disc = CDbl(txt_discount2.Text) * CDbl(txt_beli.Text)
    
    Select Case jml
        
        Case True
        
            jumlah = jumlah - disc
        
        Case False
            
            jumlah = jumlah + disc
            
        End Select
            
    lbl_jumlah.Caption = Format(jumlah, "Currency")
        
End Sub

Private Sub kosong_dc()
    txt_discount1.Text = ""
    txt_discount2.Text = ""
    lbl_jumlah.Caption = ""
End Sub

Private Sub txt_discount2_GotFocus()
    txt_discount2.SelStart = 0
    txt_discount2.SelLength = Len(txt_discount2)
End Sub

Private Sub txt_discount2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        
        If lbl_total.Caption = "" Then
            lbl_total.Caption = 0
        End If
        
        Dim tot As Double
            tot = CDbl(lbl_total.Caption) + CDbl(lbl_jumlah.Caption)
            lbl_total.Caption = Format(tot, "Currency")
                
            isi_daftar_belanjaan
                
            kosong1
            kosong_dc
            txt_kode_barang.SetFocus
    End If
            
End Sub

Private Sub txt_discount2_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txt_discount2_KeyUp(KeyCode As Integer, Shift As Integer)
    
 
    lbl_jumlah.Caption = ""
    txt_discount1.Text = ""

    If lbl_harga.Caption <> "" And txt_beli.Text <> "" And txt_discount2.Text <> "" Then
        Dim tinggi, rendah As Currency
            rendah = Trim(txt_discount2.Text)
            tinggi = Trim(lbl_harga.Caption)
            
            Dim persen_sementara
            
            persen_sementara = (CDbl(rendah) / CDbl(tinggi)) * 100
            txt_discount1.Text = Round(persen_sementara, 1)
            
            
            txt_discount2.Text = Format(txt_discount2.Text, "###,###,###")
            txt_discount2.SelStart = Len(txt_discount2.Text)
            
            If opt_discount.Value = True Then
                hitung_jumlah (True)
            ElseIf opt_cash.Value = True Then
                hitung_jumlah (False)
            End If
            
    End If
End Sub

Private Sub txt_beli_LostFocus()
    If txt_beli.Text = "" Then
        txt_beli.Text = 0
    End If
End Sub

Private Sub txt_charge_GotFocus()
    txt_charge.SelStart = 0
    txt_charge.SelLength = Len(txt_charge)
End Sub

Private Sub txt_charge_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If txt_charge.Text = "" Then
       txt_charge.Text = 0
    End If
    If KeyCode = 13 And txt_charge.Text <> "" Then
     If txt_kode_counter.Text <> "" And txt_kode_barang.Text <> "" And txt_beli.Text <> "" Then
        isi_daftar_belanjaan
        txt_kode_counter.Text = ""
        lbl_nama_counter.Caption = ""
        txt_kode_barang.Text = ""
        lbl_nama_barang.Caption = ""
        txt_beli.Text = 0
        txt_disc.Text = 0
        txt_charge.Text = 0
        lbl_harga.Caption = 0
        lbl_grand_total.Caption = 0
        txt_kode_counter.SetFocus
     Else
        MsgBox ("Data beli harus diisi")
        Exit Sub
     End If
    End If
        
End Sub

Private Sub txt_charge_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",")) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txt_charge_KeyUp(KeyCode As Integer, Shift As Integer)

    lbl_grand_total.Caption = 0
    uang_charge = 0
        If txt_charge.Text <> "" Then
            
            Dim disc, persen, grand, charge As Double
            
            persen = Trim(txt_disc.Text)
            charge = Trim(txt_charge.Text)
            
            grand = CDbl(txt_beli.Text) * CDbl(lbl_harga.Caption)
            grand = grand + CDbl(lbl_grand_total.Caption)
            disc = Val(grand) * (Val(persen) / 100)
            charge = Val(grand) * (Val(charge) / 100)
            uang_charge = charge
            grand = grand - disc + charge
            
            lbl_grand_total.Caption = Format(grand, "currency")
            
        End If
        
End Sub

Private Sub txt_charge_LostFocus()
    If txt_charge.Text = "" Then
        txt_charge.Text = 0
    End If
End Sub

Private Sub txt_counter_GotFocus(Index As Integer)
    Select Case Index
        Case 0
            txt_counter(0).SelStart = 0
            txt_counter(0).SelLength = Len(txt_counter(0))
        Case 1
            txt_counter(1).SelStart = 0
            txt_counter(1).SelLength = Len(txt_counter(1))
    End Select
End Sub

Private Sub txt_counter_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        pic_counter.Visible = False
    End If
        
    If KeyCode = 13 Then
     If arr_counter.UpperBound(1) > 0 Then
        txt_kode_counter.Text = arr_counter(grd_counter.Bookmark, 1)
        lbl_nama_counter.Caption = arr_counter(grd_counter.Bookmark, 2)
        pic_counter.Visible = False
        txt_kode_counter.SetFocus
        isi_barang
      End If
    End If
End Sub

Private Sub txt_counter_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim sql As String
    Dim rs_counter As New ADODB.Recordset
        
        kosong_counter
        
        sql = "select id,kode,nama_counter from tbl_counter"
            
            
                
                If txt_counter(0).Text <> "" And txt_counter(1).Text = "" Then
                    sql = sql & " where kode like '%" & Trim(txt_counter(0).Text) & "%'"
                End If
                
                If txt_counter(1).Text <> "" And txt_counter(0).Text = "" Then
                    sql = sql & " where nama_counter like '%" & Trim(txt_counter(1).Text) & "%'"
                End If
                
                If txt_counter(0).Text <> "" And txt_counter(1).Text <> "" Then
                    sql = sql & " where kode like '%" & Trim(txt_counter(0).Text) & "%' and nama_counter like '%" & Trim(txt_counter(1).Text) & "%'"
                End If
                                    
            
            
        sql = sql & " order by kode"
        rs_counter.Open sql, cn, adOpenKeyset
            If Not rs_counter.EOF Then
                
                rs_counter.MoveLast
                rs_counter.MoveFirst
                    
                    lanjut_counter rs_counter
            End If
        rs_counter.Close
End Sub

Private Sub txt_disc_GotFocus()
    txt_disc.SelStart = 0
    txt_disc.SelLength = Len(txt_disc)
End Sub

Private Sub txt_disc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And txt_disc.Text <> "" Then
        txt_charge.SetFocus
    Else
        txt_disc.SetFocus
    End If
End Sub

Private Sub txt_disc_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",")) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txt_disc_KeyUp(KeyCode As Integer, Shift As Integer)
    
    lbl_grand_total.Caption = 0
    uang_disc = 0
    If txt_disc.Text <> "" Then
        Dim disc, persen, grand As Double
            
            persen = Trim(txt_disc.Text)
            
            grand = CDbl(txt_beli.Text) * CDbl(lbl_harga.Caption)
            grand = grand + CDbl(lbl_grand_total.Caption)
            disc = Val(grand) * (Val(persen) / 100)
            uang_disc = CDbl(disc)
            grand = grand - disc
            
            lbl_grand_total.Caption = Format(grand, "currency")
    End If
    
End Sub

Private Sub txt_disc_LostFocus()
    If txt_disc.Text = "" Then
        txt_disc.Text = 0
    End If
End Sub

Private Sub txt_faktur_GotFocus()
    
    txt_faktur.SelStart = 0
    txt_faktur.SelLength = Len(txt_faktur)
    
End Sub

Private Sub txt_faktur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        dtp_tgl.SetFocus
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Select Case Index
        Case 0
            txt(0).SelStart = 0
            txt(0).SelLength = Len(txt(0))
        Case 1
            txt(1).SelStart = 0
            txt(1).SelLength = Len(txt(1))
    End Select
End Sub

Private Sub txt_jml_bayar_GotFocus()
    txt_jml_bayar.SelStart = 0
    txt_jml_bayar.SelLength = Len(txt_jml_bayar)
End Sub

Private Sub txt_jml_bayar_KeyUp(KeyCode As Integer, Shift As Integer)
    
    lbl_kembali.Caption = 0
    
    If txt_jml_bayar.Text <> "" Then
        Dim yang_dibayar, kembali As Double
        yang_dibayar = txt_jml_bayar.Text
        txt_jml_bayar.Text = Format(txt_jml_bayar.Text, "###,###,###")
        txt_jml_bayar.SelStart = Len(txt_jml_bayar.Text)
        kembali = CDbl(yang_dibayar) - CDbl(lbl_total_bayar.Caption)
        lbl_kembali.Caption = "Rp." & Format(kembali, "###,###,###")
    End If
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_barang.Visible = False
    End If
    
    If KeyCode = 13 Then
        If arr_barang.UpperBound(1) > 0 Then
            txt_kode_barang.Text = kode_barang
            kasih_tahu
            pic_barang.Visible = False
            txt_kode_barang.SetFocus
        End If
    End If
 
    
End Sub

Private Sub txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim sql1 As String
    Dim rs_barang As New ADODB.Recordset
        
    
        
 If arr_barang.UpperBound(1) > 0 Then
 
        
                
        sql1 = "select nama_counter,kode,nama_barang from qr_barang where id_counter=" & id_counter & ""
        
    Select Case Index
        
        Case 0
         
            sql1 = sql1 & " and kode like '%" & Trim(txt(0).Text) & "%'"
         
        Case 1
         
            sql1 = sql1 & " and nama_barang like '%" & Trim(txt(1).Text) & "%'"
         
    End Select
        
        sql1 = sql1 & " order by kode"
        rs_barang.Open sql1, cn, adOpenKeyset
            If Not rs_barang.EOF Then
                
                rs_barang.MoveLast
                rs_barang.MoveFirst
                
                lanjut_barang rs_barang
            End If
        rs_barang.Close
        
End If
      
            
                       
End Sub

Private Sub txt_kode_barang_GotFocus()
    txt_kode_barang.SelStart = 0
    txt_kode_barang.SelLength = Len(txt_kode_barang)
End Sub

Private Sub txt_kode_barang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txt_kode_barang.Text = ""
        txt(0).Text = ""
        txt(1).Text = ""
        pic_barang.Visible = True
        txt(0).SetFocus
    End If
    If KeyCode = 13 Then
        txt_beli.SetFocus
    End If
        
End Sub

Private Sub txt_kode_barang_LostFocus()
    If txt_kode_barang.Text <> "" Then
        kasih_tahu
    End If
End Sub

Private Sub txt_kode_counter_GotFocus()
    txt_kode_counter.SelStart = 0
    txt_kode_counter.SelLength = Len(txt_kode_counter)
End Sub

Private Sub txt_kode_counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txt_kode_counter.Text = ""
        txt_counter(0).Text = ""
        txt_counter(1).Text = ""
        pic_counter.Visible = True
        txt_counter(0).SetFocus
    End If
    
    If KeyCode = 13 Then
        txt_kode_barang.SetFocus
    End If
End Sub

Private Sub txt_kode_counter_LostFocus()
    If txt_kode_counter.Text <> "" Then
        Dim sql As String
        Dim rs As New ADODB.Recordset
            
            sql = "select id,nama_counter from tbl_counter where kode='" & Trim(txt_kode_counter.Text) & "'"
            rs.Open sql, cn
                If Not rs.EOF Then
                    id_counter = rs("id")
                    lbl_nama_counter.Caption = rs("nama_counter")
                    isi_barang
                Else
                    MsgBox ("Kode Counter yang anda masukkan tidak ditemukan")
                    txt_kode_counter.SetFocus
                End If
            rs.Close
    End If
End Sub
