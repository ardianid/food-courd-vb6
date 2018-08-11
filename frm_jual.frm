VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{EC76FE26-BAFD-4E89-AA40-E748DA83A570}#1.0#0"; "IsButton_Ard.ocx"
Begin VB.Form frm_jual 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frm_jual.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Txt_Tgl_Membr 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   -1200
      Locked          =   -1  'True
      TabIndex        =   67
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDBMember 
      Height          =   5895
      Left            =   2880
      TabIndex        =   52
      Top             =   6600
      Visible         =   0   'False
      Width           =   8775
      _Version        =   65536
      _ExtentX        =   15478
      _ExtentY        =   10398
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "frm_jual.frx":217CF
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_jual.frx":217EB
      Childs          =   "frm_jual.frx":21897
      Begin VB.TextBox Txt_Cr_Member 
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
         Left            =   3840
         TabIndex        =   59
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox Txt_Cr_Member 
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
         Left            =   1200
         TabIndex        =   58
         Top             =   840
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   53
         Top             =   480
         Width           =   8175
      End
      Begin TrueDBGrid60.TDBGrid Grid_Member 
         Height          =   4335
         Left            =   240
         OleObjectBlob   =   "frm_jual.frx":218B3
         TabIndex        =   57
         Top             =   1320
         Width           =   8175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN DATA MEMBER"
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
         Index           =   0
         Left            =   240
         TabIndex        =   56
         Top             =   240
         Width           =   2490
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
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
         Left            =   360
         TabIndex        =   55
         Top             =   840
         Width           =   585
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         Left            =   2880
         TabIndex        =   54
         Top             =   840
         Width           =   645
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D pic_counter 
      Height          =   6975
      Left            =   3120
      TabIndex        =   42
      Top             =   8880
      Visible         =   0   'False
      Width           =   6375
      _Version        =   65536
      _ExtentX        =   11245
      _ExtentY        =   12303
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "frm_jual.frx":2563C
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_jual.frx":25658
      Childs          =   "frm_jual.frx":25704
      Begin VB.TextBox txt_counter 
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
         Index           =   0
         Left            =   960
         TabIndex        =   45
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txt_counter 
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
         Index           =   1
         Left            =   3240
         TabIndex        =   46
         Top             =   720
         Width           =   2895
      End
      Begin VB.Frame Frame5 
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   43
         Top             =   480
         Width           =   5895
      End
      Begin TrueDBGrid60.TDBGrid grd_counter 
         Height          =   5415
         Left            =   240
         OleObjectBlob   =   "frm_jual.frx":25720
         TabIndex        =   49
         Top             =   1200
         Width           =   5895
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         Left            =   2505
         TabIndex        =   48
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
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
         TabIndex        =   47
         Top             =   720
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN COUNTER"
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
         Index           =   1
         Left            =   240
         TabIndex        =   44
         Top             =   240
         Width           =   1965
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D pic_barang 
      Height          =   6855
      Left            =   3120
      TabIndex        =   34
      Top             =   9240
      Visible         =   0   'False
      Width           =   6735
      _Version        =   65536
      _ExtentX        =   11880
      _ExtentY        =   12091
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "frm_jual.frx":2853D
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_jual.frx":28559
      Childs          =   "frm_jual.frx":28605
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   37
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3600
         TabIndex        =   38
         Top             =   720
         Width           =   2655
      End
      Begin VB.Frame Frame5 
         Height          =   135
         Index           =   1
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   6255
      End
      Begin TrueDBGrid60.TDBGrid grd_barang 
         Height          =   5415
         Left            =   240
         OleObjectBlob   =   "frm_jual.frx":28621
         TabIndex        =   41
         Top             =   1080
         Width           =   6255
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
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
         Left            =   390
         TabIndex        =   40
         Top             =   720
         Width           =   525
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         Left            =   2775
         TabIndex        =   39
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN DATA BARANG"
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
         Index           =   22
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   2460
      End
   End
   Begin VB.PictureBox c 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   120
      ScaleHeight     =   6585
      ScaleWidth      =   5865
      TabIndex        =   21
      Top             =   17000
      Visible         =   0   'False
      Width           =   5895
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   0
         Width           =   5895
      End
   End
   Begin VB.PictureBox X 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   5640
      ScaleHeight     =   6465
      ScaleWidth      =   5865
      TabIndex        =   7
      Top             =   17000
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
         TabIndex        =   9
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
         TabIndex        =   10
         Top             =   0
         Width           =   5895
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
         TabIndex        =   8
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.TextBox txt_faktur 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5760
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   2160
      Top             =   960
   End
   Begin TrueDBGrid60.TDBGrid grd_daftar 
      Height          =   2775
      Left            =   2760
      OleObjectBlob   =   "frm_jual.frx":2B8F1
      TabIndex        =   6
      Top             =   2760
      Width           =   12375
   End
   Begin VB.TextBox txt_beli 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      TabIndex        =   14
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txt_disc 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13920
      TabIndex        =   15
      Top             =   1680
      Width           =   615
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
      Left            =   -240
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox pic_samping 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   8535
      Left            =   0
      Picture         =   "frm_jual.frx":32D49
      ScaleHeight     =   8475
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   600
      Width           =   2655
      Begin IsButton_Ard.isButton Cmd_Baru 
         Height          =   975
         Left            =   1320
         TabIndex        =   64
         Top             =   7200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1720
         Icon            =   "frm_jual.frx":54518
         Style           =   10
         Caption         =   "&Baru"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin IsButton_Ard.isButton Cmd_Simpan 
         Height          =   975
         Left            =   120
         TabIndex        =   63
         Top             =   7200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1720
         Icon            =   "frm_jual.frx":54534
         Style           =   10
         Caption         =   "&Simpan"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin VB.CheckBox cek_faktur 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cetak Faktur"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   6360
         Width           =   2415
      End
      Begin VB.PictureBox pic_cepat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillColor       =   &H00C0C0FF&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   120
         ScaleHeight     =   1905
         ScaleWidth      =   2385
         TabIndex        =   25
         Top             =   4200
         Width           =   2415
         Begin VB.ListBox lst_cepat 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   570
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2835
         ScaleWidth      =   2355
         TabIndex        =   4
         Top             =   1200
         Width           =   2415
         Begin VB.Image img_foto 
            Height          =   2895
            Left            =   0
            Picture         =   "frm_jual.frx":54550
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2415
         End
      End
      Begin VB.Label lbl_user 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   480
      End
      Begin VB.Label lbl_tgl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_tgl"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lbl_jam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_jam"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   705
      End
      Begin VB.Image img2 
         Height          =   3135
         Left            =   0
         Stretch         =   -1  'True
         Top             =   -1680
         Visible         =   0   'False
         Width           =   2895
      End
   End
   Begin VB.TextBox txt_jml_bayar 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   11880
      TabIndex        =   20
      Text            =   "0"
      Top             =   7320
      Width           =   3135
   End
   Begin VB.TextBox txt_kode_counter 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   720
      Width           =   1455
   End
   Begin IsButton_Ard.isButton Browse_Faktur 
      Height          =   375
      Left            =   9480
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Icon            =   "frm_jual.frx":650D4
      Style           =   10
      Caption         =   "..."
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin VB.TextBox Txt_Tot_Disc 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   10200
      TabIndex        =   30
      Text            =   "Txt_Tot_Disc"
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox Txt_Disc_PPN 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   14520
      TabIndex        =   31
      Text            =   "Txt_Disc_PPN"
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtp_tgl 
      Height          =   375
      Left            =   11520
      TabIndex        =   28
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   12582912
      CalendarTitleForeColor=   16777215
      Format          =   54001665
      CurrentDate     =   39212
   End
   Begin VB.TextBox Txt_Order 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -120
      TabIndex        =   29
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Txt_Kode_Barang 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5760
      TabIndex        =   51
      Top             =   1200
      Width           =   1455
   End
   Begin IsButton_Ard.isButton Browse_Barang 
      Height          =   300
      Left            =   7320
      TabIndex        =   61
      Top             =   1200
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      Icon            =   "frm_jual.frx":650F0
      Style           =   5
      Caption         =   "..."
      USeCustomColors =   -1  'True
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      UseMaskColor    =   -1  'True
   End
   Begin IsButton_Ard.isButton Browse_Jenis 
      Height          =   300
      Left            =   7320
      TabIndex        =   60
      Top             =   720
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      Icon            =   "frm_jual.frx":6510C
      Style           =   5
      Caption         =   "..."
      USeCustomColors =   -1  'True
      BackColor       =   16777215
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      UseMaskColor    =   -1  'True
   End
   Begin VB.TextBox txt_ppn 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   5040
      TabIndex        =   65
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin IsButton_Ard.isButton Browse_Member 
      Height          =   375
      Left            =   6120
      TabIndex        =   69
      Top             =   5760
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Icon            =   "frm_jual.frx":65128
      Style           =   10
      Caption         =   "..."
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin VB.TextBox Txt_Member 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   5040
      TabIndex        =   66
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox Txt_Nama_Membr 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1035
      Left            =   5040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   68
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   9
      Left            =   6120
      TabIndex        =   102
      Top             =   7440
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   6
      Left            =   4800
      TabIndex        =   101
      Top             =   7440
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax  && Serv"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   8
      Left            =   2760
      TabIndex        =   100
      Top             =   7440
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   0
      Left            =   4800
      TabIndex        =   99
      Top             =   5760
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Member"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   7
      Left            =   2760
      TabIndex        =   98
      Top             =   5760
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jml Kembali"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   6
      Left            =   8160
      TabIndex        =   97
      Top             =   7920
      Width           =   1950
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   5
      Left            =   11640
      TabIndex        =   96
      Top             =   7920
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jml Dibayar"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   5
      Left            =   8160
      TabIndex        =   95
      Top             =   7320
      Width           =   1860
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   4
      Left            =   11640
      TabIndex        =   94
      Top             =   7320
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jml Harus Dibayar"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   4
      Left            =   8160
      TabIndex        =   93
      Top             =   6720
      Width           =   2880
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   3
      Left            =   11640
      TabIndex        =   92
      Top             =   6720
      Width           =   90
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   3
      Left            =   11160
      TabIndex        =   91
      Top             =   6120
      Width           =   285
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disc Total"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   2
      Left            =   8160
      TabIndex        =   90
      Top             =   6120
      Width           =   1590
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   2
      Left            =   11640
      TabIndex        =   89
      Top             =   6120
      Width           =   90
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   1
      Left            =   8160
      TabIndex        =   88
      Top             =   5640
      Width           =   1845
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   1
      Left            =   11640
      TabIndex        =   87
      Top             =   5640
      Width           =   90
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   10
      Left            =   14640
      TabIndex        =   86
      Top             =   1680
      Width           =   285
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disc"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   9
      Left            =   13200
      TabIndex        =   85
      Top             =   1680
      Width           =   570
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   8
      Left            =   11640
      TabIndex        =   84
      Top             =   1680
      Width           =   450
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   7
      Left            =   9720
      TabIndex        =   83
      Top             =   2160
      Width           =   1320
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   6
      Left            =   3000
      TabIndex        =   82
      Top             =   2160
      Width           =   2145
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   5
      Left            =   11280
      TabIndex        =   81
      Top             =   1200
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   5
      Left            =   8400
      TabIndex        =   80
      Top             =   1200
      Width           =   2130
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   4
      Left            =   11280
      TabIndex        =   79
      Top             =   600
      Width           =   90
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
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   4
      Left            =   8400
      TabIndex        =   78
      Top             =   600
      Width           =   2250
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   3
      Left            =   11280
      TabIndex        =   77
      Top             =   120
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   3
      Left            =   8400
      TabIndex        =   76
      Top             =   120
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   2
      Left            =   5520
      TabIndex        =   75
      Top             =   1200
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kd Barang"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   2
      Left            =   2760
      TabIndex        =   74
      Top             =   1200
      Width           =   1650
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   1
      Left            =   5520
      TabIndex        =   73
      Top             =   600
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kd Counter"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   1
      Left            =   2760
      TabIndex        =   72
      Top             =   600
      Width           =   1770
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   0
      Left            =   5520
      TabIndex        =   71
      Top             =   120
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Faktur"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   405
      Index           =   0
      Left            =   2760
      TabIndex        =   70
      Top             =   120
      Width           =   1635
   End
   Begin VB.Label Lbl_Harus 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   11880
      TabIndex        =   62
      Top             =   6720
      Width           =   3135
   End
   Begin VB.Label Lbl_Tot_Disc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   11880
      TabIndex        =   50
      Top             =   6120
      Width           =   3135
   End
   Begin VB.Label Lbl_Tot_DiscP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lbl_Tot_DiscP"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   14280
      TabIndex        =   32
      Top             =   -120
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lbl_kembali 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   11880
      TabIndex        =   19
      Top             =   7920
      Width           =   3135
   End
   Begin VB.Label lbl_total_bayar 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      TabIndex        =   18
      Top             =   5575
      Width           =   3135
   End
   Begin VB.Label lbl_harga 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5640
      TabIndex        =   17
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label lbl_nama_counter 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   13
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label lbl_nama_barang 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   12
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Image img_dasar 
      Height          =   735
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbl_grand_total 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   11640
      TabIndex        =   24
      Top             =   2160
      Width           =   3495
   End
End
Attribute VB_Name = "frm_jual"
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
Dim Satuanku As String
Dim Pot_Disc As Integer
Dim Tot_Tanpa_Disc As Double
Dim Arr_Member As New XArrayDB
Dim InMember As Boolean
Dim Rubah As Boolean
Dim Letakdirubah As Long
Dim hap_detail As Boolean

Dim Moving As Boolean
Dim yold, xold As Long
Dim harusppn As Double
Dim rubah1 As Boolean

Private Function cek_Barang_noorder_sama() As Long
    
    cek_Barang_noorder_sama = 0
    
    If arr_barang.UpperBound(1) = 0 Then Exit Function
    
    Dim a As Double
        For a = 1 To arr_daftar.UpperBound(1)
            If UCase(Txt_Kode_Barang.Text) = UCase(arr_daftar(a, 1)) And UCase(Txt_Order.Text) = UCase(arr_daftar(a, 13)) Then
                cek_Barang_noorder_sama = a
                
                grd_daftar.Col = a
                
                Exit Function
            End If
        Next
    
End Function

Private Sub Ok_Rubah(ByVal baru_lagis As Boolean)

On Error GoTo er_ok

    Dim sql, sql1, sql2 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset
    Dim a As Long

 cn.BeginTrans
 
 If arr_daftar.UpperBound(1) = 0 Then
    
    Dim k As Integer
        k = CInt(MsgBox("Tidak ada barang yang akan disimpan", vbOKOnly + vbInformation, "Informasi"))
        
        cn.RollbackTrans
        On Error GoTo 0
        Exit Sub
 End If


 If txt_faktur.Text = "" Then
    MsgBox ("No Faktur tidak boleh kosong")
    txt_faktur.SetFocus
    
    cn.RollbackTrans
    On Error GoTo 0
    Exit Sub
 End If

' If CDbl(txt_jml_bayar.Text) < CDbl(Lbl_Harus.Caption) Then
'    MsgBox ("Jumlah uang tidak boleh kurang dari total bayar")
'    Exit Sub
' End If

If baru_lagis <> False Then

 If MsgBox("Yakin semua data yang dimasukkan sudah benar", vbYesNo + vbQuestion, "Pesan") = vbNo Then
    
    cn.RollbackTrans
    On Error GoTo 0
    Exit Sub
 End If

End If


    sql = "select no_faktur from tr_faktur_penjualan where no_faktur='" & Trim(txt_faktur.Text) & "'"
    rs.Open sql, cn
        If Not rs.EOF Then
        
            sql1 = "update tr_faktur_penjualan set tgl='" & Trim(dtp_tgl.Value) & "',jam='" & Trim(lbl_jam.Caption) & "',nama_user='" & Trim(lbl_user.Caption) & "',ket=0,user_pembatal='-',No_Member='" & Trim(Txt_Member.Text) & "',Disc=" & Trim(Txt_Tot_Disc.Text) & ",Nilai_Disc=" & CCur(Trim(Lbl_Tot_Disc.Caption)) & ",Tot_Sth_Disc=" & CCur(Lbl_Harus.Caption)
            sql1 = sql1 & ",ppn=" & Trim(txt_ppn.Text) & " where no_faktur='" & Trim(txt_faktur.Text) & "'"
            rs1.Open sql1, cn
        
            
        For a = 1 To arr_daftar.UpperBound(1)
            
            Dim msql As String
            Dim mrs As Recordset
            
            msql = "select no_faktur from tr_penjualan where no_faktur='" & Trim(txt_faktur.Text) & "' and No_order='" & arr_daftar(a, 13) & "' and id_barang= " & arr_daftar(a, 12)
            
            Set mrs = New ADODB.Recordset
                mrs.Open msql, cn, adOpenKeyset
            
            With mrs
            
            If .EOF Then
            

                    stock_sekarang (arr_daftar(a, 12))
        
        
                    If st_s <> 0 Then
                        If (stk - CDbl(arr_daftar(a, 3))) < st_s Then
                            Dim jangan
                            jangan = MsgBox("Stock barang tidak mencukupi untuk memenuhi penjualan" & Chr(13) & "Stock Sekarang " & stk & Chr(13) & "Stock Min " & st_s & Chr(13) & "Jml Beli " & arr_daftar(a, 3), vbOKOnly + vbInformation, "Pesan")
                            cn.RollbackTrans
                            
                            On Error GoTo 0
                            Exit Sub
                        End If
                     End If


                ' isi transaksi penjualan
    
                    sql1 = "insert into tr_penjualan (no_faktur,id_barang,qty,harga_satuan,harga_sebenarnya,disc,harga_disc,cash,harga_cash,total_harga,No_order)"
                    sql1 = sql1 & " values('" & Trim(txt_faktur.Text) & "'," & arr_daftar(a, 12) & "," & arr_daftar(a, 3) & "," & CCur(arr_daftar(a, 5)) & ", " & CCur(arr_daftar(a, 6)) & ",'" & arr_daftar(a, 7) & "'," & CCur(arr_daftar(a, 8)) & ",'" & arr_daftar(a, 9) & "'," & CCur(arr_daftar(a, 10)) & "," & CCur(arr_daftar(a, 11)) & ",'" & arr_daftar(a, 13) & "')"
                    rs1.Open sql1, cn
    
                ' seleksi dari tr_jml_stok
    
                    sql2 = "select jml_stock from tr_jml_stock where id_barang=" & arr_daftar(a, 12)
                    rs2.Open sql2, cn
                        If Not rs2.EOF Then
                            Dim j_stock_sekarang As Double
    
                'update tbl_jml_stock ( kalau ada)
    
                                j_stock_sekarang = CDbl(rs2("jml_stock")) - CDbl(arr_daftar(a, 3))
                                sql1 = "update tr_jml_stock set jml_stock=" & j_stock_sekarang & " where id_barang=" & arr_daftar(a, 12)
                                rs1.Open sql1, cn
    
                ' isi transaksi stok_barang
    
                                sql1 = "insert into tr_stock (id_barang,brg_in,brg_out,tgl,ket,nama_user)"
                                sql1 = sql1 & " values(" & arr_daftar(a, 12) & ",0," & arr_daftar(a, 3) & ",'" & Trim(dtp_tgl.Value) & "',0,'" & Trim(lbl_user.Caption) & "')"
                                rs1.Open sql1, cn
    
                        End If
                    rs2.Close
'
'            Next a
'                    cek_keb_stock (arr_daftar(a, 12))
          
            Else
            
                    
                ' seleksi dari tr_jml_stok
    
                    sql2 = "select jml_stock from tr_jml_stock where id_barang=" & arr_daftar(a, 12)
                    rs2.Open sql2, cn
                        If Not rs2.EOF Then
'                            Dim j_stock_sekarang As Double
                            Dim b_sebelumnya As Double
                            
                            stock_sekarang (arr_daftar(a, 12))
                            b_sebelumnya = Cari_Jml_Barang_Sebelumnya(Trim(txt_faktur.Text), arr_daftar(a, 12), arr_daftar(a, 13))
                
                            If st_s <> 0 Then
                                If ((stk + b_sebelumnya)) - CDbl(arr_daftar(a, 3)) < st_s Then
                                   ' Dim jangan
                                    jangan = MsgBox("Stock barang tidak mencukupi untuk memenuhi penjualan" & Chr(13) & "Stock Sekarang " & stk & Chr(13) & "Stock Min " & st_s & Chr(13) & "Jml Beli " & arr_daftar(a, 3), vbOKOnly + vbInformation, "Pesan")
                                    cn.RollbackTrans
                                    
                                    On Error GoTo 0
                                    Exit Sub
                                End If
                             End If


                'update tbl_jml_stock ( kalau ada)
    
                                j_stock_sekarang = (CDbl(rs2("jml_stock")) + b_sebelumnya) - CDbl(arr_daftar(a, 3))
                                sql1 = "update tr_jml_stock set jml_stock=" & j_stock_sekarang & " where id_barang=" & arr_daftar(a, 12)
                                rs1.Open sql1, cn
    
                ' isi transaksi stok_barang
                                
                                Dim Jml_Stock_Masuk As Double
                                
                                If CDbl(b_sebelumnya) > CDbl(arr_daftar(a, 3)) Then
                                    Jml_Stock_Masuk = CDbl(b_sebelumnya) - CDbl(arr_daftar(a, 3))
                                Else
                                    Jml_Stock_Masuk = CDbl(arr_daftar(a, 3)) - CDbl(b_sebelumnya)
                                End If
                                
                                sql1 = "insert into tr_stock (id_barang,brg_in,brg_out,tgl,ket,nama_user)"
                                sql1 = sql1 & " values(" & arr_daftar(a, 12) & ",0," & Jml_Stock_Masuk & ",'" & Trim(dtp_tgl.Value) & "',0,'" & Trim(lbl_user.Caption) & "')"
                                rs1.Open sql1, cn
    
                        End If
                    rs2.Close

                ' update transaksi penjualan
    
                    sql1 = "update tr_penjualan set qty=" & arr_daftar(a, 3) & ",harga_satuan=" & CCur(arr_daftar(a, 5)) & ",harga_sebenarnya=" & CCur(arr_daftar(a, 6)) & ",disc='" & arr_daftar(a, 7) & "',harga_disc=" & CCur(arr_daftar(a, 8)) & ",cash='" & arr_daftar(a, 9) & "',harga_cash=" & CCur(arr_daftar(a, 10)) & ",total_harga=" & CCur(arr_daftar(a, 11)) & ",No_order='" & arr_daftar(a, 13) & "'"
                    sql1 = sql1 & " where no_faktur='" & Trim(txt_faktur.Text) & "' and id_barang=" & arr_daftar(a, 12) & " and No_order='" & Trim(arr_daftar(a, 13)) & "'"
                    rs1.Open sql1, cn


'                    sql1 = "update tr_faktur_penjualan set tgl='" & Trim(dtp_tgl.Value) & "',jam='" & Trim(lbl_jam.Caption) & "',nama_user='" & Trim(lbl_user.Caption) & "',ket=0,user_pembatal='-',No_Member='" & Trim(Txt_Member.Text) & "',Disc=" & Trim(Txt_Tot_Disc.Text) & ",Nilai_Disc=" & CCur(Trim(Lbl_Harus.Caption))
'                    sql1 = sql1 & " where no_faktur='" & Trim(txt_faktur.Text) & "'"
'                    rs1.Open sql1, cn
                      
            End If
          End With
          
                    cek_keb_stock (arr_daftar(a, 12))
            Next a
          
    cn.CommitTrans
    Dim konfirm As Integer
    konfirm = CInt(MsgBox("Data telah disimpan", vbOKOnly + vbInformation, "Infomasi"))
    

    If cek_faktur.Value = vbChecked Then
'        faktur
    End If
    
    If baru_lagis = True Then
    If cek_faktur.Value = vbChecked Then
        If MsgBox("Apakah anda ingin mencetak bukti pembayaran", vbYesNo + vbQuestion, "Konfirmasi") = vbYes Then
            
            noff = Trim(txt_faktur.Text)
            
            If txt_jml_bayar.Text = "" Then
                byyr = 0
            ElseIf txt_jml_bayar.Text = 0 Then
                byyr = 0
            Else
                byyr = Replace(txt_jml_bayar, ".", "")
            End If
            
            If Trim(lbl_kembali.Caption) = "Rp." Then
                kemm = 0
            Else
                kemm = Replace(Trim(lbl_kembali.Caption), "Rp.", "")
                kemm = Replace(kemm, ",", "")
                kemm = CDbl(kemm)
            End If
            
            htu = True
            
            Load Frm_Lap_BuktiByar
                Frm_Lap_BuktiByar.Show
        
        End If
    End If
    End If
       
'     Call baru_lagi
    'Rubah = False

        Else
                
                konfirm = CInt(MsgBox("No Faktur " & Trim(txt_faktur.Text) & " Tidak ditemukan", vbOKOnly + vbInformation, "Infomasi"))
                cn.RollbackTrans
                
                On Error GoTo 0
                Exit Sub
        End If
    rs.Close
    
    On Error GoTo 0
    Exit Sub
    
    
er_ok:
    cn.RollbackTrans
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear

End Sub

Private Function Cari_Jml_Barang_Sebelumnya(ByVal no_faktur As String, ByVal id_barang As String, ByVal Noorder As String) As Double
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select qty from tr_penjualan where no_faktur='" & no_faktur & "' and id_barang=" & id_barang & " and No_order='" & Noorder & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, cn
            
            With rs
                
                If Not .EOF Then
                    Cari_Jml_Barang_Sebelumnya = IIf(Not IsNull(!qty), !qty, 0)
                Else
                    Cari_Jml_Barang_Sebelumnya = 0
                End If
                
            End With
    
End Function

Private Sub kosong_counter()
    arr_counter.ReDim 0, 0, 0, 0
    grd_counter.ReBind
    grd_counter.Refresh
End Sub

Private Sub isi_counter()

On Error GoTo err_counter

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
        Exit Sub
        
err_counter:

    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub

Private Sub lanjut_counter(rs_counter As Recordset)
    Dim id_c, kd_c, nm_c As String
    Dim a As Long
        
            arr_counter.ReDim 0, 0, 0, 0
                grd_counter.ReBind
                grd_counter.Refresh

        
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

Private Sub faktur()

On Error GoTo er_print



Dim a As Long
Dim grs
    Printer.Font = "Arial"
    Printer.FontSize = 7.5
        
        Printer.CurrentX = 0
        Printer.CurrentY = 0
            
            Printer.Print
            Printer.Print Tab((55 / 2 - Len("Bukti Pembayaran") / 2) - 25); "Bukti Pembayaran"
            Printer.Print Tab((55 / 2 - Len("Istana Kuring") / 2) - 25); "Istana Kuring"
            
            grs = String$(73, "-")
            
            Printer.Print grs
            
            Printer.Print "Tgl. " & dtp_tgl.Value; Tab(21); "Jam. " & lbl_jam.Caption
            Printer.Print grs
            Printer.Print "No Faktur " & Trim(txt_faktur.Text)
            Printer.Print "Qty"; Tab(5); "Nama Barang"; Tab(25); "Harga"; Tab(35); "Disc"; Tab(41); "Charge"
            Printer.Print grs
            
       For a = 1 To arr_daftar.UpperBound(1)
            
            Printer.Print arr_daftar(a, 3); Tab(5); arr_daftar(a, 2); Tab(25); Space(Len(arr_daftar(a, 4)) - Len(arr_daftar(a, 4))) + arr_daftar(a, 4); Tab(35); arr_daftar(a, 6); Tab(41); arr_daftar(a, 8)
            
      Next a
            
            Printer.Print "Total Discount"; Tab(21); grd_daftar.Columns(6).FooterText
            Printer.Print "Total Charge"; Tab(21); grd_daftar.Columns(8).FooterText
            Printer.Print grs
            Printer.Print "Total"; Tab(25); Space(Len(grd_daftar.Columns(10).FooterText) - Len(grd_daftar.Columns(10).FooterText)) + grd_daftar.Columns(10).FooterText
            Printer.Print "jml Bayar "; Tab(25); Format(Space(Len(txt_jml_bayar.Text) - Len(txt_jml_bayar.Text)) + txt_jml_bayar.Text, "Currency")
            Printer.Print "Kembali"; Tab(25); Space(Len(lbl_kembali.Caption) - Len(lbl_kembali.Caption)) + lbl_kembali.Caption
            Printer.Print
            Printer.Print " ******* Terima Kasih Atas Kunjungan Anda *******"
            
      
      Printer.EndDoc
      Exit Sub
            
er_print:
           Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
            
End Sub

Private Sub kosong_aja_semua()
       lbl_total_bayar.Caption = ""
       
       lbl_kembali.Caption = ""
       
       kosong1
       
       
       kosong_daftar
       Txt_Kode_Barang.SetFocus
End Sub

Private Sub ok()

On Error GoTo er_ok

    Dim sql, sql1, sql2 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset
    Dim a As Long

cn.BeginTrans

 If arr_daftar.UpperBound(1) = 0 Then
    
    Dim k As Integer
        k = CInt(MsgBox("Tidak ada barang yang akan disimpan", vbOKOnly + vbInformation, "Informasi"))
        
        cn.RollbackTrans
        On Error GoTo 0
        Exit Sub
End If
    
    
 If txt_faktur.Text = "" Then
    MsgBox ("No Faktur tidak boleh kosong")
    txt_faktur.SetFocus
    
    cn.RollbackTrans
    On Error GoTo 0
    Exit Sub
 End If

' If CDbl(txt_jml_bayar.Text) < CDbl(Lbl_Harus.Caption) Then
'    MsgBox ("Jumlah uang tidak boleh kurang dari total bayar")
'    Exit Sub
' End If

 If MsgBox("Yakin semua data yang dimasukkan sudah benar", vbYesNo + vbQuestion, "Pesan") = vbNo Then
    
    cn.RollbackTrans
    On Error GoTo 0
    Exit Sub
 End If

     

    sql = "select no_faktur from tr_faktur_penjualan where no_faktur='" & Trim(txt_faktur.Text) & "'"
    rs.Open sql, cn
        If Not rs.EOF Then

                MsgBox ("No Faktur " & Trim(txt_faktur.Text) & " Sudah Ada")
                cn.RollbackTrans
                
                On Error GoTo 0
                Exit Sub
        End If
    rs.Close

    sql1 = "insert into tr_faktur_penjualan (no_faktur,tgl,jam,nama_user,ket,user_pembatal,No_Member,Disc,Nilai_Disc,Tot_Sth_Disc,ppn)"
    sql1 = sql1 & " values('" & Trim(txt_faktur.Text) & "','" & Trim(dtp_tgl.Value) & "','" & Trim(lbl_jam.Caption) & "','" & Trim(lbl_user.Caption) & "',0,'-','" & Trim(Txt_Member.Text) & "'," & Trim(Txt_Tot_Disc.Text) & "," & CCur(Trim(Lbl_Tot_Disc.Caption)) & "," & CCur(Lbl_Harus.Caption) & "," & Trim(txt_ppn.Text) & " )"
    rs1.Open sql1, cn

    For a = 1 To arr_daftar.UpperBound(1)

            stock_sekarang (arr_daftar(a, 12))
'            If ket_b = False Then
'                Exit Sub
'            End If

                If st_s <> 0 Then
                    If (stk - CDbl(arr_daftar(a, 3))) < st_s Then
                        Dim jangan
                        jangan = MsgBox("Stock barang tidak mencukupi untuk memenuhi penjualan" & Chr(13) & "Stock Sekarang " & stk & Chr(13) & "Stock Min " & st_s & Chr(13) & "Jml Beli " & arr_daftar(a, 3), vbOKOnly + vbInformation, "Pesan")
                        cn.RollbackTrans
                        
                        On Error GoTo 0
                        Exit Sub
                    End If
                 End If

            ' isi transaksi penjualan

                sql1 = "insert into tr_penjualan (no_faktur,id_barang,qty,harga_satuan,harga_sebenarnya,disc,harga_disc,cash,harga_cash,total_harga,No_order)"
                sql1 = sql1 & " values('" & Trim(txt_faktur.Text) & "'," & arr_daftar(a, 12) & "," & arr_daftar(a, 3) & "," & CCur(arr_daftar(a, 5)) & ", " & CCur(arr_daftar(a, 6)) & ",'" & arr_daftar(a, 7) & "'," & CCur(arr_daftar(a, 8)) & ",'" & arr_daftar(a, 9) & "'," & CCur(arr_daftar(a, 10)) & "," & CCur(arr_daftar(a, 11)) & ",'" & arr_daftar(a, 13) & "')"
                rs1.Open sql1, cn

            ' seleksi dari tr_jml_stok

                sql2 = "select jml_stock from tr_jml_stock where id_barang=" & arr_daftar(a, 12)
                rs2.Open sql2, cn
                    If Not rs2.EOF Then
                        Dim j_stock_sekarang As Double

            'update tbl_jml_stock ( kalau ada)

                            j_stock_sekarang = CDbl(rs2("jml_stock")) - CDbl(arr_daftar(a, 3))
                            sql1 = "update tr_jml_stock set jml_stock=" & j_stock_sekarang & " where id_barang=" & arr_daftar(a, 12)
                            rs1.Open sql1, cn

            ' isi transaksi stok_barang

                            sql1 = "insert into tr_stock (id_barang,brg_in,brg_out,tgl,ket,nama_user)"
                            sql1 = sql1 & " values(" & arr_daftar(a, 12) & ",0," & arr_daftar(a, 3) & ",'" & Trim(dtp_tgl.Value) & "',0,'" & Trim(lbl_user.Caption) & "')"
                            rs1.Open sql1, cn

                    End If
                rs2.Close
                cek_keb_stock (arr_daftar(a, 12))
    Next a
    
    cn.CommitTrans
    MsgBox ("Data berhasil disimpan")
    

    If cek_faktur.Value = vbChecked Then
       If MsgBox("Apakah anda ingin mencetak bukti pembayaran", vbYesNo + vbQuestion, "Konfirmasi") = vbYes Then
            
            noff = Trim(txt_faktur.Text)
            
            If txt_jml_bayar.Text = "" Then
                byyr = 0
            ElseIf txt_jml_bayar.Text = 0 Then
                byyr = 0
            Else
                byyr = Replace(txt_jml_bayar, ",", "")
            End If
            
            If Trim(lbl_kembali.Caption) = "Rp." Then
                kemm = 0
            Else
                kemm = Replace(Trim(lbl_kembali.Caption), "Rp.", "")
                kemm = Replace(kemm, ",", "")
                kemm = CDbl(kemm)
            End If
            
            htu = True
            
            Load Frm_Lap_BuktiByar
                Frm_Lap_BuktiByar.Show
            
            
        Else
            
            Call baru_lagi
        
       End If
    End If
    
  '  Call baru_lagi
    Rubah = False
    
    On Error GoTo 0
    Exit Sub
    
    
er_ok:
    cn.RollbackTrans
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub

Private Sub stock_sekarang(id_b As String)

On Error GoTo er_stock

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        sql = "select stock_min,jml_stock from qr_jml_stock where id_barang=" & id_b
        rs.Open sql, cn
            If Not rs.EOF Then
                st_s = rs("stock_min")
                stk = rs("jml_stock")
                ket_b = True
            Else
                ket_b = False
                st_s = 0
                stk = 0
            End If
        rs.Close
        
        Exit Sub
        
er_stock:
    
    Dim p
        p = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub
Private Sub cek_keb_stock(param As String)

On Error GoTo err_keb

Dim sql As String
Dim rs As New ADODB.Recordset

        sql = "select stock_min,jml_stock from qr_jml_stock where id_barang=" & param
        rs.Open sql, cn
            If Not rs.EOF Then
                If CDbl(rs("jml_stock")) = CDbl(rs("stock_min")) Then
                    Dim sk
                    sk = MsgBox("Jumlah stok hampir mendekati batas minimum" & Chr(13) & "Stock Min " & rs("stock_min") & Chr(13) & "Laporkan segera untuk kembali diisi")
                   ' Exit Sub
                End If
            End If
        rs.Close
        
        On Error GoTo 0
        Exit Sub
        
err_keb:
        
        Dim p
            p = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub Browse_Barang_Click()
        
    If txt_kode_counter.Text = "" Then Exit Sub
        
    With pic_barang
        
        If .Visible = False Then
            
            .Left = (Browse_Barang.Left + Browse_Barang.Width) / 2
            .Top = (Browse_Barang.Top + Browse_Barang.Height) + 15
            
            Txt_Kode_Barang.Text = ""
            txt(0).Text = ""
            txt(1).Text = ""
            pic_barang.Visible = True
            txt(0).SetFocus

        
        Else
            .Visible = False
        End If
        
    End With
    
End Sub

Private Sub Browse_Jenis_Click()
    
    With pic_counter
        
        If .Visible = False Then
            
            .Left = (Browse_Jenis.Left + Browse_Jenis.Width) / 2
            .Top = (Browse_Jenis.Top + Browse_Jenis.Height) + 15
            
            txt_kode_counter.Text = ""
            txt_counter(0).Text = ""
            txt_counter(1).Text = ""
            pic_counter.Visible = True
            txt_counter(0).SetFocus

            
        Else
            .Visible = False
        End If
        
    End With
    
End Sub

Private Sub Browse_Member_Click()

If lbl_total_bayar.Caption = 0 Then Exit Sub

With TDBMember
    
    If .Visible = False Then
        
        Txt_Cr_Member(0).Text = ""
        Txt_Cr_Member(1).Text = ""
        
        Txt_Cr_Member_KeyUp 0, 0, 0
        
        .Visible = True
        
        Txt_Cr_Member(0).SetFocus
        
    Else
        .Visible = False
    End If
    
End With

End Sub

Private Sub cmd_baru_Click()
    baru_lagi
End Sub

Private Sub Cmd_Simpan_Click()
    If Rubah = False Then
        Call ok
    Else
        Call Ok_Rubah(True)
    End If
        
    hap_detail = False
    
End Sub

Private Sub cmd_x_Click()
    pic_barang.Visible = False
End Sub

Private Sub Command1_Click()
    pic_counter.Visible = False
    txt_kode_counter.SetFocus
End Sub

Private Sub Command2_Click()
    pic_barang.Visible = False
    Txt_Kode_Barang.SetFocus
End Sub

Private Sub dtp_tgl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
'        Txt_Order.SetFocus
        txt_kode_counter.SetFocus
    End If
    
    If KeyCode = vbKeyF2 Then
        baru_lagi
    End If
    
    If KeyCode = vbKeyF4 Then
        frm_cfaktur.Show
    End If
        
    If KeyCode = vbKeyF1 Then
        frm_ganti_pwd.Show
    End If
        
End Sub

Private Sub Form_Activate()
    cek_faktur.Value = vbChecked
    txt_faktur.SetFocus
'    txt_kode_counter.SetFocus
End Sub

Private Sub Form_Load()
    
    Me.Picture = LoadPicture(App.path & "\background transaksi.jpg")
    pic_samping.Picture = LoadPicture(App.path & "\background transaksi.jpg")
    
    'Me.PaintPicture utama.Picture, 0, 0
    pic_samping.PaintPicture Me.Picture, 0, 0
    
     
    Rubah = False
    InMember = False
    hap_detail = False
     
    grd_daftar.Array = arr_daftar
        
    grd_counter.Array = arr_counter
        
    grd_barang.Array = arr_barang
        
    kosong_daftar
    
    isi_counter
    
    kosong_barang

    dtp_tgl.Value = Format(Date, "dd/mm/yyyy")
    
    With pic_barang
        .Left = 4440
        .Top = 1875
    End With
    
     With pic_counter
        .Left = 4680
        .Top = 1375
    End With
    
    With TDBMember
        .Left = 6480
        .Top = 2400
    End With
    
    Grid_Member.Array = Arr_Member
    
    
    txt_jml_bayar.Text = 0
    lbl_harga.Caption = 0
    txt_beli.Text = 0
    Txt_Disc.Text = 0
    txt_charge.Text = 0
    lbl_grand_total.Caption = 0
    lbl_total_bayar.Caption = 0
    txt_jml_bayar.Text = 0
    lbl_kembali.Caption = 0
    
    
    Lbl_Harus.Caption = 0
    Txt_Member.Text = ""
    
    txt_ppn.Text = 0
    Txt_Tot_Disc.Text = 0
    Txt_Disc_PPN.Text = 0
    Lbl_Tot_Disc.Caption = 0
    Lbl_Tot_DiscP.Caption = 0
    
    
    Tot_Tanpa_Disc = 0
    
    cepat
    
    cari_foto
    
    dtp_tgl.Value = Date
    
  '  isi_faktur
    
    
End Sub

Private Sub cari_foto()
On Error Resume Next
    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        sql = " select foto from qr_user where id=" & id_user
        rs.Open sql, cn
            
            If Not rs.EOF Then
                
                If Not IsNull(rs("foto")) Then
                    Set img_foto.Picture = LoadPicture(path_foto & "\foto" & rs("foto"))
                Else
                    Set img_foto.Picture = LoadPicture()
                End If
            Else
                Set img_foto.Picture = LoadPicture()
            End If
        rs.Close
        
            
End Sub

Private Sub cepat()
    With lst_cepat
        .AddItem "             Shortcut"
        .AddItem "============================================"
        .AddItem "F2 Baru"
        .AddItem "F3 Bantuan "
        .AddItem "F4 Cetak Faktur Lagi"
        .AddItem "Alt+C Cek Faktur"
        
        .AddItem "-----------------------------------------------------"
    End With
End Sub

Public Sub isi_faktur()

On Error GoTo er_faktur

    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    Dim faktur, f_sementara, ln
    Dim pj
        
        faktur = ""
    
        faktur = DatePart("d", Now)
      If Len(faktur) = 1 Then
        faktur = 0 & faktur
      End If
        faktur = faktur & DatePart("m", Now)
        faktur = faktur & Right(DatePart("yyyy", Now), 2)
        
        pj = 0
        pj = Len(faktur) + 1
        
        faktur = faktur & id_user
        
        
        sql = "select no_faktur from tr_faktur_penjualan where ucase(nama_user)=ucase('" & Trim(lbl_user.Caption) & "') and tgl=datevalue('" & Trim(dtp_tgl.Value) & "')"
        rs.Open sql, cn
            If Not rs.EOF Then
                sql1 = "select no_faktur  from tr_faktur_penjualan where  tgl=datevalue('" & Trim(dtp_tgl.Value) & "') and ucase(nama_user)=ucase('" & Trim(lbl_user.Caption) & "')"
                rs1.Open sql1, cn, adOpenKeyset
                    If Not rs1.EOF Then
                        
                        rs1.MoveLast
                        rs1.MoveFirst
                        
                        Dim sementara As Long
                        Dim f, j
                            
                            sementara = 0
                            ln = Len(id_user)
                            
                            Do While Not rs1.EOF
                                
                                j = Mid(rs1("no_faktur"), pj + ln, Len(rs1("no_faktur")))
                                    
                                If j > sementara Then
                                    sementara = j
                                End If
                                
                           rs1.MoveNext
                           Loop
                            
                            f_sementara = CDbl(sementara) + 1
                            faktur = faktur & f_sementara
                            
                            
                    End If
                rs1.Close
            Else
                faktur = faktur & "1"
            End If
        rs.Close
        txt_faktur.Text = ""
        txt_faktur.Text = faktur
            
        On Error GoTo 0
        Exit Sub
        
er_faktur:
            Dim p
                p = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
                Err.Clear
        
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

On Error GoTo isi_barang

    Dim sql As String
    Dim rs_barang As New ADODB.Recordset
        
        kosong_barang
        
        sql = "select top 100 nama_counter,kode,nama_barang,Satuan from qr_barang where id_counter=" & id_counter & " and aktif=1 order by kode"
        rs_barang.Open sql, cn, adOpenKeyset
            If Not rs_barang.EOF Then
                
                rs_barang.MoveLast
                rs_barang.MoveFirst
                
                lanjut_barang rs_barang
            End If
        rs_barang.Close
        
        Exit Sub
        
isi_barang:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub

Private Sub lanjut_barang(rs_barang As Recordset)
    Dim counter, kode, barang As String
    Dim sat As String
    Dim a As Long
        
        arr_barang.ReDim 0, 0, 0, 0
                grd_barang.ReBind
                grd_barang.Refresh
        
        
        a = 1
            Do While Not rs_barang.EOF
                arr_barang.ReDim 1, a, 0, grd_barang.Columns.Count
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
                    
                    sat = IIf(Not IsNull(rs_barang!satuan), rs_barang!satuan, "")
                    
                arr_barang(a, 0) = counter
                arr_barang(a, 1) = kode
                arr_barang(a, 2) = barang
                arr_barang(a, 3) = sat
                
            a = a + 1
            rs_barang.MoveNext
            Loop
            grd_barang.ReBind
            grd_barang.Refresh
        
End Sub


Private Sub Form_Resize()
On Error Resume Next
    
'    Picture1.Width = Me.Width
'    Picture1.Height = Me.Height
'    Picture1.PaintPicture Me.Picture, 0, 0
    
    pic_samping.Left = Me.Left
    pic_samping.Top = utama.pic_atas.Height - 1850
    pic_samping.Width = 2655
    pic_samping.Height = Me.Height - utama.pic_atas
    
    img2.Move pic_samping.Left, 0, pic_samping.Width, pic_samping.Height
    img2.Picture = LoadPicture(App.path & "\banner3.1.jpg")
    img2.ZOrder 1
    lbl_jam.Move img2.Left + 150, img2.Top + 100, 0, 0
    lbl_tgl.Move img2.Left + 150, lbl_jam.Top + 300, 0, 0
    lbl_tgl.Caption = Format(Date, "long date")
    
    lbl_user.Left = img2.Left + 150
    lbl_user.Top = lbl_tgl.Top + 500
    
    img_dasar.Left = Me.ScaleLeft
    img_dasar.Top = Me.ScaleTop
    img_dasar.Width = Me.ScaleWidth
    img_dasar.Width = Me.ScaleWidth
    img_dasar.Height = Me.ScaleHeight
    img_dasar.Picture = LoadPicture(App.path & "\3.jpg")
    img_dasar.ZOrder 1
        
    pic_cepat.Top = Picture3.Height + 1500
    pic_cepat.Left = img2.Left + 150
    lst_cepat.Width = pic_cepat.Width
    lst_cepat.Height = pic_cepat.Height
    lst_cepat.Width = pic_cepat.Width
    
    cek_faktur.Top = pic_cepat.Top + pic_cepat.Height + 150
    pic_cepat.Left = pic_cepat.Left
    
End Sub

Private Sub grd_barang_Click()
    On Error Resume Next
        If arr_barang.UpperBound(1) > 0 Then
            kode_barang = arr_barang(grd_barang.Bookmark, 1)
        End If
End Sub

Private Sub grd_barang_DblClick()
 If arr_barang.UpperBound(1) > 0 Then
    Txt_Kode_Barang.Text = kode_barang
    kasih_tahu
    pic_barang.Visible = False
    Txt_Kode_Barang.SetFocus
 End If
End Sub

Private Sub kasih_tahu()

On Error GoTo err_kasih

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        Satuanku = ""
        
        sql = "select nama_counter,nama_barang,harga_jual,id_barang,Satuan,Per_Disc from qr_barang where kode='" & Trim(Txt_Kode_Barang.Text) & "' and id_counter=" & id_counter
        rs.Open sql, cn
        
            If Not rs.EOF Then
                id_barang = rs("id_barang")
                lbl_nama_barang.Caption = rs("nama_barang")
                lbl_harga.Caption = Format(rs("harga_jual"), "Currency")
                Satuanku = IIf(Not IsNull(rs("Satuan")), rs("Satuan"), "")
                
                Dim p As Integer
                    p = rs!Per_disc
                    If p = 1 Then
                        Pot_Disc = 1
                        Txt_Disc.Enabled = True
                    ElseIf p = 2 Then
                        Txt_Disc.Enabled = False
                        Pot_Disc = 2
                    End If
                
            Else
                MsgBox ("Kode barang yang anda masukkan tidak ditemukan")
                Txt_Kode_Barang.Text = ""
                Txt_Kode_Barang.SetFocus
            End If
        rs.Close
        
        Exit Sub
        
err_kasih:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub

Private Sub grd_barang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        grd_barang_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        pic_barang.Visible = False
        Txt_Kode_Barang.SetFocus
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
'        isi_barang
     End If
End Sub

Private Sub grd_counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_counter.Visible = False
        txt_kode_counter.SetFocus
    End If
    If KeyCode = 13 Then
        grd_counter_DblClick
    End If
    
End Sub

Private Sub grd_counter_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_counter_Click
End Sub

Private Sub grd_daftar_Click()
 '  On Error Resume Next
        If arr_daftar.UpperBound(1) > 0 Then
            sementara = arr_daftar(grd_daftar.Bookmark, 11)
            s_disc = arr_daftar(grd_daftar.Bookmark, 7)
            s_charge = arr_daftar(grd_daftar.Bookmark, 9)
           ' txt_beli.SetFocus
        End If
End Sub

Private Sub grd_daftar_DblClick()
    
    If arr_daftar.UpperBound(1) = 0 Then Exit Sub
    
    Txt_Order.Text = arr_daftar(grd_daftar.Bookmark, 13)
    Txt_Kode_Barang.Text = arr_daftar(grd_daftar.Bookmark, 1)
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select kode_counter,nama_counter from qr_barang where kode='" & Trim(Txt_Kode_Barang.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, cn
        With rs
            
            If Not .EOF Then
                txt_kode_counter.Text = IIf(Not IsNull(!kode_counter), !kode_counter, "")
                lbl_nama_counter.Caption = IIf(Not IsNull(!nama_counter), !nama_counter, "")
            Else
                txt_kode_counter.Text = ""
                lbl_nama_counter.Caption = ""
            End If
            
        End With
    
    lbl_nama_barang.Caption = arr_daftar(grd_daftar.Bookmark, 2)
    lbl_harga.Caption = Format(arr_daftar(grd_daftar.Bookmark, 5), "currency")
    txt_beli.Text = arr_daftar(grd_daftar.Bookmark, 3)
    
    If Txt_Tot_Disc.Text = 0 Then
        Txt_Disc.Text = Mid(arr_daftar(grd_daftar.Bookmark, 7), 1, Len(arr_daftar(grd_daftar.Bookmark, 7)) - 1)
    Else
        Txt_Disc.Text = 0
    End If
        
    Satuanku = arr_daftar(grd_daftar.Bookmark, 4)
    
    id_barang = arr_daftar(grd_daftar.Bookmark, 12)
    
        lbl_grand_total.Caption = 0
        If txt_beli.Text <> "" Then
            Dim grand As Double
                grand = CDbl(txt_beli.Text) * CDbl(lbl_harga.Caption)
                grand = grand + CDbl(lbl_grand_total.Caption)
                lbl_grand_total.Caption = Format(grand, "Currency")
        End If
    
    Pot_Disc = arr_daftar(grd_daftar.Bookmark, 14)
    
    If Pot_Disc = 1 Then
        Txt_Disc.Enabled = True
    Else
        Txt_Disc.Enabled = False
    End If
    
    Letakdirubah = grd_daftar.Bookmark
    hap_detail = True
    rubah1 = True
    
    txt_kode_counter.SetFocus
    
    
End Sub

Private Sub grd_daftar_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo er_down

    If KeyCode = vbKeyDelete Then
    
    If arr_daftar.UpperBound(1) = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    
    
    sementara = arr_daftar(grd_daftar.Bookmark, 11)
    s_disc = arr_daftar(grd_daftar.Bookmark, 7)
    s_charge = arr_daftar(grd_daftar.Bookmark, 9)

    If Rubah = False Then
    
                    Dim jml, jml_d, jml_c As Double
                
                    jml = CDbl(grd_daftar.Columns(11).FooterText) - CDbl(arr_daftar(grd_daftar.Bookmark, 11))
                    grd_daftar.Columns(11).FooterText = Format(jml, "currency")
                    
                    lbl_total_bayar.Caption = Format(CDbl(lbl_total_bayar) - (CDbl(arr_daftar(grd_daftar.Bookmark, 3)) * CDbl(arr_daftar(grd_daftar.Bookmark, 5))), "currency")
                            
                    
        '                If Pot_Disc <> 2 Then
        '                    Tot_Tanpa_Disc = Tot_Tanpa_Disc - CDbl(arr_daftar(grd_daftar.Bookmark, 11))
        '                End If
                    
                    'If Txt_Tot_Disc.Text <> 0 Then
                        
                        Dim nil_disc As Double
                        If arr_daftar(grd_daftar.Bookmark, 14) <> 2 Then
                            nil_disc = Tot_Tanpa_Disc - (CDbl(arr_daftar(grd_daftar.Bookmark, 3)) * CDbl(arr_daftar(grd_daftar.Bookmark, 5)))
                            Tot_Tanpa_Disc = nil_disc
                        Else
                            nil_disc = Tot_Tanpa_Disc
                        End If
                        
                        If nil_disc <= 0 Then
                            nil_disc = 0
                        Else
                            nil_disc = nil_disc * (CDbl(Txt_Tot_Disc.Text) / 100)
                        End If
                        
                        Lbl_Tot_Disc.Caption = Format(nil_disc, "currency")
                        Lbl_Harus.Caption = Format(CDbl(lbl_total_bayar.Caption) - CDbl(Lbl_Tot_Disc.Caption), "currency")
                        txt_jml_bayar.Text = Format(Lbl_Harus.Caption, "###,###,###")
                        
                   ' End If
                    
                    'txt_jml_bayar.Text = Format(jml, "###,###,###")
                    
                  '  lbl_kembali.Caption = "Rp." & Format(jml, "###,###,###")
                    
                    
                    jml_d = Mid(grd_daftar.Columns(7).FooterText, 1, Len(grd_daftar.Columns(7).FooterText) - 1)
                    s_disc = Mid(arr_daftar(grd_daftar.Bookmark, 7), 1, Len(arr_daftar(grd_daftar.Bookmark, 7)) - 1)
                    jml_d = CDbl(jml_d) - CDbl(s_disc)
                    
                    If jml_d <= 0 Then
                        grd_daftar.Columns(7).FooterText = 0 & "%"
                    Else
                        grd_daftar.Columns(7).FooterText = jml_d & "%"
                    End If
                        
                        If arr_daftar.UpperBound(1) > 1 Then
                            grd_daftar.Delete
                        Else
                            arr_daftar.ReDim 0, 0, 0, 0
                        End If
                            grd_daftar.ReBind
                            grd_daftar.Refresh
                        
                        

    Else
        
                        If arr_daftar.UpperBound(1) = 1 Then
                            Dim m As Integer
                                m = CInt(MsgBox("Anda tidak diperbolehkan menghapus semua barang yang telah dibeli", vbOKOnly + vbInformation, "Informasi"))
                                
                                On Error GoTo 0
                                Exit Sub
                        End If
                        
                        Dim sql As String
                        Dim rs As Recordset
                            
                        Dim sql1 As String
                        Dim rs1 As Recordset
                            
                        cn.BeginTrans
                
                            ' seleksi dari tr_jml_stok
                
                                sql1 = "select jml_stock from tr_jml_stock where id_barang=" & arr_daftar(grd_daftar.Bookmark, 12)
                                
                                Set rs1 = New ADODB.Recordset
                                rs1.Open sql1, cn
                                    If Not rs1.EOF Then
                                        Dim j_stock_sekarang As Double
                
                            'update tbl_jml_stock ( kalau ada)
                
                                            j_stock_sekarang = CDbl(rs1("jml_stock")) + CDbl(arr_daftar(grd_daftar.Bookmark, 3))
                                            sql = "update tr_jml_stock set jml_stock=" & j_stock_sekarang & " where id_barang=" & arr_daftar(grd_daftar.Bookmark, 12)
                                            
                                            Set rs = New ADODB.Recordset
                                            rs.Open sql, cn
                
                            ' isi transaksi stok_barang
                
                                            sql = "insert into tr_stock (id_barang,brg_in,brg_out,tgl,ket,nama_user)"
                                            sql = sql & " values(" & arr_daftar(grd_daftar.Bookmark, 12) & "," & arr_daftar(grd_daftar.Bookmark, 3) & ",0,'" & Trim(dtp_tgl.Value) & "',0,'" & Trim(lbl_user.Caption) & "')"
                                            
                                            Set rs = New ADODB.Recordset
                                            rs.Open sql, cn
                                End If
                                
                                sql = "delete from tr_penjualan where no_faktur='" & arr_daftar(grd_daftar.Bookmark, 0) & "' and id_barang=" & arr_daftar(grd_daftar.Bookmark, 12) & " and No_order='" & arr_daftar(grd_daftar.Bookmark, 13) & "'"
                                
                                Set rs = New ADODB.Recordset
                                    rs.Open sql, cn
                        
                            jml = CDbl(grd_daftar.Columns(11).FooterText) - CDbl(arr_daftar(grd_daftar.Bookmark, 11))
                            grd_daftar.Columns(11).FooterText = Format(jml, "currency")
                            
                            lbl_total_bayar.Caption = Format(CDbl(lbl_total_bayar) - (CDbl(arr_daftar(grd_daftar.Bookmark, 3)) * CDbl(arr_daftar(grd_daftar.Bookmark, 5))), "currency")
                            
                    
        '                If Pot_Disc <> 2 Then
        '                    Tot_Tanpa_Disc = Tot_Tanpa_Disc - CDbl(arr_daftar(grd_daftar.Bookmark, 11))
        '                End If
                    
                    'If Txt_Tot_Disc.Text <> 0 Then
                        
'                        Dim nil_disc As Double
                        If arr_daftar(grd_daftar.Bookmark, 14) <> 2 Then
                            nil_disc = Tot_Tanpa_Disc - (CDbl(arr_daftar(grd_daftar.Bookmark, 3)) * CDbl(arr_daftar(grd_daftar.Bookmark, 5)))
                            Tot_Tanpa_Disc = nil_disc
                        Else
                            nil_disc = Tot_Tanpa_Disc
                        End If
                        
                        If nil_disc <= 0 Then
                            nil_disc = 0
                        Else
                            nil_disc = nil_disc * (CDbl(Txt_Tot_Disc.Text) / 100)
                        End If
                        
                        Lbl_Tot_Disc.Caption = Format(nil_disc, "currency")
                        Lbl_Harus.Caption = Format(CDbl(lbl_total_bayar.Caption) - CDbl(Lbl_Tot_Disc.Caption), "currency")
                        txt_jml_bayar.Text = Format(Lbl_Harus.Caption, "###,###,###")
                        
                   ' End If
                    
                    'txt_jml_bayar.Text = Format(jml, "###,###,###")
                    
                 '   lbl_kembali.Caption = "Rp." & Format(jml, "###,###,###")
                    
                    
                    jml_d = Mid(grd_daftar.Columns(7).FooterText, 1, Len(grd_daftar.Columns(7).FooterText) - 1)
                    s_disc = Mid(arr_daftar(grd_daftar.Bookmark, 7), 1, Len(arr_daftar(grd_daftar.Bookmark, 7)) - 1)
                    jml_d = CDbl(jml_d) - CDbl(s_disc)
                    If jml_d <= 0 Then
                        grd_daftar.Columns(7).FooterText = 0 & "%"
                    Else
                        grd_daftar.Columns(7).FooterText = jml_d & "%"
                    End If

                        If arr_daftar.UpperBound(1) > 1 Then
                            grd_daftar.Delete
                        Else
                            arr_daftar.ReDim 0, 0, 0, 0
                        End If
                            grd_daftar.ReBind
                            grd_daftar.Refresh

                        
                        cn.CommitTrans
                        
                        Ok_Rubah (False)
                        
                        
                        grd_daftar.ReBind
                        grd_daftar.Refresh
                        
                        grd_daftar.MoveFirst
              
End If
            
        
End If
    
    Txt_Tot_Disc_Change
    cmd_simpan.SetFocus
    
    On Error GoTo 0
    Exit Sub
    
er_down:
    
    If Rubah = True Then cn.RollbackTrans
    
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub

Private Sub grd_daftar_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_daftar_Click
End Sub

Private Sub pic_barang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_barang.Visible = False
        Txt_Kode_Barang.SetFocus
    End If
End Sub

Private Sub Grid_Member_DblClick()
On Error Resume Next
    
    If Arr_Member.UpperBound(1) = 1 And Arr_Member(1, 1) = Empty Then Exit Sub
    
    InMember = False
    
    Dim konfirm As Integer
        If Arr_Member(Grid_Member.Bookmark, 3) = vbUnchecked Then
        
        konfirm = CInt(MsgBox("No. member yang anda pilih sudah tida aktif", vbOKOnly + vbInformation, "Informasi"))
        Grid_Member.SetFocus
        Exit Sub
        End If
    
    If CDate(Arr_Member(Grid_Member.Bookmark, 2)) < Date Then
        konfirm = CInt(MsgBox("Kartu member anda sudah habis sejak " & Arr_Member(Grid_Member.Bookmark, 2), vbOKOnly + vbInformation, "Informasi"))
        Exit Sub
    End If
    
    
    Txt_Member.Text = Arr_Member(Grid_Member.Bookmark, 0)
    Txt_Nama_Membr.Text = Arr_Member(Grid_Member.Bookmark, 1)
    Txt_Tgl_Membr.Text = Arr_Member(Grid_Member.Bookmark, 2)

                    Txt_Tot_Disc.Text = Arr_Member(Grid_Member.Bookmark, 4)
                    
                    Dim Nilai_Disc As Double
                     If Txt_Tot_Disc <> 0 Then
                        Nilai_Disc = Tot_Tanpa_Disc * (CDbl(Txt_Tot_Disc.Text) / 100)
                     Else
                        Nilai_Disc = 0
                     End If
                     
                     Lbl_Tot_Disc.Caption = Format(Nilai_Disc, "currency")

                    Dim harus As Double
                        
                        harus = CDbl(lbl_total_bayar.Caption) - CDbl(Lbl_Tot_Disc.Caption)
                            
                        Lbl_Harus.Caption = Format(harus, "currency")
                        txt_jml_bayar.Text = harus
                        lbl_kembali.Caption = 0


    TDBMember.Visible = False
    
    
    Txt_Tot_Disc_Change
    hitung_ppn
    
    Txt_Tot_Disc.SetFocus
    InMember = True
    
End Sub

Private Sub Grid_Member_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Member_DblClick
    If KeyCode = vbKeyEscape Then TDBMember.Visible = False
End Sub

Private Sub pic_barang_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub pic_barang_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   pic_barang.Top = pic_barang.Top - (yold - Y)
   pic_barang.Left = pic_barang.Left - (xold - X)
End If

End Sub

Private Sub pic_barang_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub pic_counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_counter.Visible = False
        txt_kode_counter.SetFocus
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

Private Sub TDBMember_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDBMember_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDBMember.Top = TDBMember.Top - (yold - Y)
   TDBMember.Left = TDBMember.Left - (xold - X)
End If

End Sub

Private Sub TDBMember_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub Timer1_Timer()
    DoEvents
    lbl_jam.Caption = Format(Time, "hh:mm:ss")
End Sub

Private Sub txt_bayar_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txt_beli_GotFocus()
    txt_beli.SelStart = 0
    txt_beli.SelLength = Len(txt_beli)
End Sub

Private Sub isi_daftar_belanjaan()

On Error GoTo er_belanjaan
    
    
    If (Rubah = True And hap_detail = True) Or rubah1 = True Then
    
    grd_daftar.Bookmark = Letakdirubah
    
            rubah1 = False
            
            Dim jml, jml_d, jml_c As Double
        
            jml = CDbl(grd_daftar.Columns(11).FooterText) - CDbl(arr_daftar(Letakdirubah, 11))
            grd_daftar.Columns(11).FooterText = Format(jml, "currency")
            
            lbl_total_bayar.Caption = Format(jml, "currency")
                    
                
                Dim nil_disc As Double
                If arr_daftar(Letakdirubah, 14) <> 2 Then
                    nil_disc = Tot_Tanpa_Disc - (CDbl(arr_daftar(Letakdirubah, 3)) * CDbl(arr_daftar(Letakdirubah, 5)))
                    Tot_Tanpa_Disc = nil_disc
                Else
                    nil_disc = Tot_Tanpa_Disc
                End If
                
                    nil_disc = nil_disc * (CDbl(Txt_Tot_Disc.Text) / 100)
                
                Lbl_Tot_Disc.Caption = Format(nil_disc, "currency")
                Lbl_Harus.Caption = Format(CDbl(lbl_total_bayar.Caption) - CDbl(Lbl_Tot_Disc.Caption), "currency")
                
                
           ' End If
            
            'txt_jml_bayar.Text = Format(jml, "###,###,###")
            
            lbl_kembali.Caption = "Rp." & Format(jml, "###,###,###")
            
            
            jml_d = Mid(grd_daftar.Columns(7).FooterText, 1, Len(grd_daftar.Columns(7).FooterText) - 1)
            s_disc = Mid(arr_daftar(Letakdirubah, 7), 1, Len(arr_daftar(Letakdirubah, 7)) - 1)
            jml_d = CDbl(jml_d) - CDbl(s_disc)
            grd_daftar.Columns(7).FooterText = jml_d & "%"
                
                If arr_daftar.UpperBound(1) > 1 Then
                    grd_daftar.Delete
                Else
                    arr_daftar.ReDim 0, 0, 0, 0
                End If
                    grd_daftar.ReBind
                    grd_daftar.Refresh
    
    End If
    
'    harga_s = CDbl(arr_daftar(a, 3)) * CDbl(arr_daftar(a, 5))
'    harga_disc = harga_s * (CDbl(nil_disc) / 100)
'    harga_disc = harga_s - harga_disc


    arr_daftar.ReDim 1, arr_daftar.UpperBound(1) + 1, 0, grd_daftar.Columns.Count
    grd_daftar.ReBind
    grd_daftar.Refresh
        
        Dim jml_baris As Long
            
            jml_baris = arr_daftar.UpperBound(1)
            
            arr_daftar(jml_baris, 0) = Trim(txt_faktur.Text)
            arr_daftar(jml_baris, 1) = Trim(Txt_Kode_Barang.Text)
            arr_daftar(jml_baris, 2) = Trim(lbl_nama_barang.Caption)
            arr_daftar(jml_baris, 3) = Trim(txt_beli.Text)
            arr_daftar(jml_baris, 4) = Satuanku
            arr_daftar(jml_baris, 5) = Trim(lbl_harga.Caption)
            arr_daftar(jml_baris, 6) = CDbl(lbl_harga.Caption) * CDbl(txt_beli.Text)
         If Txt_Tot_Disc.Text <> 0 And Pot_Disc <> 2 Then
            arr_daftar(jml_baris, 7) = Trim(Txt_Tot_Disc.Text) & "%"
         Else
            arr_daftar(jml_baris, 7) = Trim(Txt_Disc.Text) & "%"
         End If
            
            arr_daftar(jml_baris, 8) = uang_disc
            arr_daftar(jml_baris, 9) = Trim(txt_charge.Text) & "%"
            arr_daftar(jml_baris, 10) = uang_charge
            arr_daftar(jml_baris, 11) = Trim(lbl_grand_total.Caption)
            arr_daftar(jml_baris, 12) = id_barang
            arr_daftar(jml_baris, 13) = Trim(Txt_Order.Text)
            arr_daftar(jml_baris, 14) = Pot_Disc
            
'            If Rubah = False Then
                If Pot_Disc <> 2 Then
                    Tot_Tanpa_Disc = Tot_Tanpa_Disc + CDbl(lbl_grand_total.Caption)
                End If
            
            
            Dim jml_diskon, jml_cash, jml_biaya As Double
            
            If grd_daftar.Columns(11).FooterText = "" Then
                grd_daftar.Columns(11).FooterText = 0
            End If
            
            jml_biaya = CDbl(grd_daftar.Columns(11).FooterText) + CDbl(lbl_grand_total.Caption)
            grd_daftar.Columns(11).FooterText = Format(jml_biaya, "currency")
            lbl_total_bayar.Caption = Format(CDbl(lbl_total_bayar.Caption) + (CDbl(lbl_harga.Caption) * CDbl(txt_beli.Text)), "currency")
            
            Lbl_Harus.Caption = Format(jml_biaya, "currency")
            
            txt_jml_bayar.Text = Format(jml_biaya, "###,###,###")
'            lbl_kembali.Caption = "Rp." & Format(jml_biaya, "###,###,###")
                     
            If grd_daftar.Columns(7).FooterText = "" Then
                grd_daftar.Columns(7).FooterText = 0 & "%"
                jml_diskon = 0
            End If
            
            
                jml_diskon = CDbl(Txt_Disc.Text) + CDbl(Mid(grd_daftar.Columns(7).FooterText, 1, Len(grd_daftar.Columns(7).FooterText) - 1))
                grd_daftar.Columns(7).FooterText = jml_diskon & "%"
            

            If grd_daftar.Columns(9).FooterText = "" Then
                grd_daftar.Columns(9).FooterText = 0 & "%"
                jml_cash = 0
            End If
            
                jml_cash = CDbl(txt_charge.Text) + CDbl(Mid(grd_daftar.Columns(9).FooterText, 1, Len(grd_daftar.Columns(9).FooterText) - 1))
                grd_daftar.Columns(9).FooterText = jml_cash & "%"
     
     grd_daftar.ReBind
     grd_daftar.Refresh
     
     If Txt_Tot_Disc.Text <> 0 And Txt_Member.Text = "" Then
            Txt_Tot_Disc_Change
     Else
     
             If Pot_Disc <> 2 Then
                Tot_Tanpa_Disc = Tot_Tanpa_Disc - CDbl(lbl_grand_total.Caption)
             End If
              
            If Tot_Tanpa_Disc <= 0 Then Tot_Tanpa_Disc = 0
    
            Dim harga_s As Double
            
            harga_s = CDbl(txt_beli.Text) * CDbl(lbl_harga.Caption)
            
            Tot_Tanpa_Disc = Tot_Tanpa_Disc + harga_s
    
            Txt_Member_KeyDown 13, 0
        
     End If
        
     hap_detail = False
     
     On Error GoTo 0
     Exit Sub
     
er_belanjaan:
     Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
     
End Sub

Private Sub kosong1()
    Txt_Kode_Barang.Text = ""
    lbl_nama_barang.Caption = ""
    lbl_nama_counter.Caption = ""
    lbl_harga.Caption = ""
    txt_beli.Text = ""
    
End Sub

Private Sub txt_beli_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If txt_beli.Text = "" Then
        txt_beli.Text = 0
    End If
    
    If KeyCode = 13 And txt_beli.Text <> 0 Then
    
        If cek_Barang_noorder_sama = True Then
            Dim nk As Integer
            nk = CInt(MsgBox("Barang dengan no order yang anda masukkan sudah ada dalam daftar pesanan,silahkan periksa kembali", vbOKOnly + vbInformation, "Informasi"))
            
            Txt_Order.SetFocus
            On Error GoTo 0
            Exit Sub
        End If

    
        If Txt_Disc.Enabled = True Then
            Txt_Disc.SetFocus
        Else
            Txt_Disc.Text = 0
            txt_disc_KeyDown 13, 0
'            Txt_Tot_Disc_Change
'            hitung_ppn
        End If

    Else
        txt_beli.SetFocus
    End If
    
    If KeyCode = vbKeyF2 Then
        baru_lagi
    End If
    
    If KeyCode = vbKeyF4 Then
        frm_cfaktur.Show
    End If
    
    If KeyCode = vbKeyF1 Then
        frm_ganti_pwd.Show
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
    
    If Txt_Tot_Disc.Text <> 0 And Pot_Disc <> 2 Then
        
        uang_disc = 0
        
        Dim harga_s As Double
        Dim harga_disc As Double
        
        harga_s = CDbl(txt_beli.Text) * CDbl(lbl_harga.Caption)
        harga_disc = harga_s * (CDbl(Txt_Tot_Disc.Text) / 100)
        uang_disc = harga_disc
        harga_disc = harga_s - harga_disc
        
        lbl_grand_total.Caption = Format(harga_disc, "currency")
    Else
        Dim grand As Double
            grand = CDbl(txt_beli.Text) * CDbl(lbl_harga.Caption)
            grand = grand + CDbl(lbl_grand_total.Caption)
            lbl_grand_total.Caption = Format(grand, "Currency")
    
    End If
    
    End If
End Sub

Private Sub txt_discount1_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",")) Then
        Beep
        KeyAscii = 0
    End If
End Sub


Private Sub txt_discount2_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
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
     If txt_kode_counter.Text <> "" And Txt_Kode_Barang.Text <> "" And txt_beli.Text <> "" Then
        isi_daftar_belanjaan
        txt_kode_counter.Text = ""
        lbl_nama_counter.Caption = ""
        Txt_Kode_Barang.Text = ""
        lbl_nama_barang.Caption = ""
        txt_beli.Text = 0
        Txt_Disc.Text = 0
        txt_charge.Text = 0
        lbl_harga.Caption = 0
        lbl_grand_total.Caption = 0
        txt_kode_counter.SetFocus
     Else
        MsgBox ("Data beli harus diisi")
        Exit Sub
     End If
    End If
        
    If KeyCode = vbKeyF2 Then
        baru_lagi
    End If
    
    If KeyCode = vbKeyF4 Then
        frm_cfaktur.Show
    End If
    
    If KeyCode = vbKeyF1 Then
        frm_ganti_pwd.Show
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
            
            persen = Trim(Txt_Disc.Text)
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
        txt_kode_counter.SetFocus
    End If
        
    If KeyCode = 13 Then
     If arr_counter.UpperBound(1) > 0 Then
        txt_kode_counter.Text = arr_counter(grd_counter.Bookmark, 1)
        lbl_nama_counter.Caption = arr_counter(grd_counter.Bookmark, 2)
        pic_counter.Visible = False
        txt_kode_counter.SetFocus
'        isi_barang
      End If
    End If
    
    
    
End Sub

Private Sub txt_counter_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo er_counter

    Dim sql As String
    Dim rs_counter As New ADODB.Recordset
        
        kosong_counter
        
        sql = "select top 100 id,kode,nama_counter from tbl_counter"
            
            
                
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
        
        Exit Sub
        
er_counter:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub Txt_Cr_Member_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Member_DblClick
    If KeyCode = vbKeyEscape Then TDBMember.Visible = False
    End Sub

Private Sub Txt_Cr_Member_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next

    Dim sql As String
    Dim rs As Recordset
    
    Dim a As Long
    Dim no_mem, nama, tgl_a, akt As String
    Dim disc As Double
    
    sql = "select top 100 No_Member,Nama,Tanggal2,Aktif,Disc from qr_member where aktif=1 " ' where No_Member='" & Trim(Txt_Member.Text) & "'"
        
    Select Case Index
        Case 0
            sql = sql & " and No_Member like '%" & Trim(Txt_Cr_Member(0).Text) & "%'"
        Case 1
            sql = sql & " and Nama like '%" & Trim(Txt_Cr_Member(1).Text) & "%'"
    End Select
    
    
    sql = sql & " order by No_Member asc"
    
    Set rs = New ADODB.Recordset
        rs.Open sql, cn, adOpenKeyset
        
        a = 1
        Arr_Member.ReDim 0, 0, 0, 0
        Arr_Member.ReDim 1, 1, 1, 1
            Grid_Member.ReBind
            Grid_Member.Refresh
        
        With rs
            
            Do While Not .EOF
                Arr_Member.ReDim 1, a, 0, Grid_Member.Columns.Count
                    Grid_Member.ReBind
                    Grid_Member.Refresh
                
                no_mem = IIf(Not IsNull(!No_Member), !No_Member, "")
                nama = IIf(Not IsNull(!nama), !nama, "")
                tgl_a = IIf(Not IsNull(!Tanggal2), !Tanggal2, "")
                akt = !aktif
                disc = IIf(Not IsNull(!disc), !disc, 0)
                
                Arr_Member(a, 0) = no_mem
                Arr_Member(a, 1) = nama
                Arr_Member(a, 2) = tgl_a
                
                If akt = 1 Then
                    Arr_Member(a, 3) = vbChecked
                Else
                    Arr_Member(a, 3) = vbUnchecked
                End If
                
                Arr_Member(a, 4) = disc
                
            a = a + 1
            .MoveNext
            Loop
            
            Grid_Member.ReBind
            Grid_Member.Refresh
            
            Grid_Member.MoveFirst
            
        End With
        
    
        
End Sub

Private Sub txt_disc_GotFocus()
    Txt_Disc.SelStart = 0
    Txt_Disc.SelLength = Len(Txt_Disc)
End Sub

Private Sub txt_disc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Txt_Disc.Text <> "" Then
    
    If txt_kode_counter.Text <> "" And Txt_Kode_Barang.Text <> "" And txt_beli.Text <> "" Then
        
        Dim no_gr As Long
            no_gr = cek_Barang_noorder_sama
        
        If no_gr <> 0 And hap_detail = False Then
            Dim nk As Integer
            If MsgBox("PERHATIAN" & Chr(13) & "Barang dengan no order yang anda masukkan sudah ada dalam daftar pesanan" & Chr(13) & "data yang akan dirubah hanya harga barang,Qty,serta total" & Chr(13) & "apakah anda ingin menambahkan datanya ...", vbYesNo + vbQuestion, "Konfirmasi") = vbYes Then
            
            arr_daftar(no_gr, 3) = arr_daftar(no_gr, 3) + CDbl(Trim(txt_beli.Text))
            arr_daftar(no_gr, 5) = Trim(lbl_harga.Caption)
            arr_daftar(no_gr, 6) = CDbl(lbl_harga.Caption) * arr_daftar(no_gr, 3)
                        
            Dim disc, persen, grand As Double
            
            If Left(arr_daftar(no_gr, 7), 1) <> "%" Then
                persen = Mid(arr_daftar(no_gr, 7), 1, Len(arr_daftar(no_gr, 7)) - 1)
            Else
                persen = 0
            End If
            
            grand = arr_daftar(no_gr, 6)
            grand = grand
            disc = Val(grand) * (Val(persen) / 100)
            grand = grand - disc
                        
            arr_daftar(no_gr, 11) = grand
            
'            If Txt_Tot_Disc.Text <> 0 And Pot_Disc <> 2 Then
'                arr_daftar(no_gr, 7) = arr_daftar(no_gr, 7) & "%"
'            Else
                If Left(arr_daftar(no_gr, 7), 1) <> "%" Then
                    arr_daftar(no_gr, 7) = Mid(arr_daftar(no_gr, 7), 1, Len(arr_daftar(no_gr, 7)) - 1) & "%"
                Else
                    arr_daftar(no_gr, 7) = 0 & "%"
                End If
                
'            End If
            
            arr_daftar(no_gr, 8) = arr_daftar(no_gr, 8)
            
            If Left(arr_daftar(no_gr, 9), 1) <> "%" Then
                arr_daftar(no_gr, 9) = Mid(arr_daftar(no_gr, 9), 1, Len(arr_daftar(no_gr, 9)) - 1) & "%"
            Else
                arr_daftar(no_gr, 9) = 0 & "%"
            End If
            
            arr_daftar(no_gr, 10) = arr_daftar(no_gr, 10)
'            arr_daftar(jml_baris, 11) = Trim(lbl_grand_total.Caption)
            arr_daftar(no_gr, 12) = arr_daftar(no_gr, 12)
            arr_daftar(no_gr, 13) = arr_daftar(no_gr, 13)
            arr_daftar(no_gr, 14) = arr_daftar(no_gr, 14)
            
            grd_daftar.ReBind
            grd_daftar.Refresh
            
'            If grd_daftar.Columns(11).FooterText = "" Then
                grd_daftar.Columns(11).FooterText = 0
'            End If
            
            Dim a As Long
            Dim harga_se As Double
            
            harga_se = 0
            For a = arr_daftar.LowerBound(1) To arr_daftar.UpperBound(1)
               grd_daftar.Columns(11).FooterText = CDbl(grd_daftar.Columns(11).FooterText) + CDbl(arr_daftar(a, 11))
               
               harga_se = harga_se + CDbl(arr_daftar(a, 6))
               
            Next
            
            
            
            lbl_total_bayar.Caption = Format(harga_se, "currency")
            
            grd_daftar.ReBind
            grd_daftar.Refresh
            
            Dim jml_t_d As Double
                If Left(grd_daftar.Columns(7).FooterText, 1) <> "%" Then
                    jml_t_d = Mid(grd_daftar.Columns(7).FooterText, 1, Len(grd_daftar.Columns(7).FooterText) - 1)
                Else
                    jml_t_d = 0
                End If

                jml_t_d = Val(harga_se) * (Val(jml_t_d) / 100)
                harga_se = harga_se - jml_t_d
            
            Dim tot_disc, nil_dis As Double
                tot_disc = Trim(Txt_Tot_Disc.Text)
                nil_dis = CDbl(lbl_total_bayar.Caption) * (CDbl(tot_disc) / 100)
'                nil_dis = CDbl(lbl_total_bayar.Caption) - nil_dis

            Lbl_Tot_Disc.Caption = Format(nil_dis, "currency")

            Dim harus As Double
            If Txt_Tot_Disc.Text <> 0 And Txt_Tot_Disc.Text <> "" Then
                harus = CDbl(lbl_total_bayar.Caption) - CDbl(nil_dis)
            Else
                harus = grd_daftar.Columns(11).FooterText
            End If

            Lbl_Harus.Caption = Format(harus, "currency")

            txt_jml_bayar.Text = harus
            
            Txt_Tot_Disc_Change
            hitung_ppn
            txt_jml_bayar_Change
            
            
            
            Else
                
                nk = CInt(MsgBox("Maaf anda tidak boleh memasukkan data barang dan order yang sama...", vbOKOnly + vbInformation, "Informasi"))
            
            End If
            
        Else
            
            isi_daftar_belanjaan
            
        End If
    
'        txt_charge.SetFocus
'    Else
        
        hap_detail = False
                
'        If Txt_Tot_Disc.Text <> 0 Then
'            Txt_Tot_Disc_Change
'        End If
            
        txt_kode_counter.Text = ""
        lbl_nama_counter.Caption = ""
        Txt_Kode_Barang.Text = ""
        lbl_nama_barang.Caption = ""
        txt_beli.Text = 0
        Txt_Disc.Text = 0
        txt_charge.Text = 0
        lbl_harga.Caption = 0
        lbl_grand_total.Caption = 0
        txt_kode_counter.SetFocus
        
        Txt_Tot_Disc_Change
        hitung_ppn
        txt_jml_bayar_Change
        
     Else
        MsgBox ("Data beli harus diisi")
        Exit Sub
     End If
    End If
    
    If KeyCode = vbKeyF2 Then
        baru_lagi
    End If
    
    If KeyCode = vbKeyF4 Then
        frm_cfaktur.Show
    End If
    
    If KeyCode = vbKeyF1 Then
        frm_ganti_pwd.Show
    End If
    
End Sub

Private Sub txt_disc_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",")) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txt_disc_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If Txt_Tot_Disc.Text <> 0 Then
        Txt_Disc.Text = 0
        Exit Sub
    End If

    lbl_grand_total.Caption = 0
    uang_disc = 0
    If Txt_Disc.Text <> "" Then
        Dim disc, persen, grand As Double
            
            persen = Trim(Txt_Disc.Text)
            
            grand = CDbl(txt_beli.Text) * CDbl(lbl_harga.Caption)
            grand = grand + CDbl(lbl_grand_total.Caption)
            disc = Val(grand) * (Val(persen) / 100)
            uang_disc = CDbl(disc)
            grand = grand - disc
            
            lbl_grand_total.Caption = Format(grand, "currency")
    End If
    
End Sub

Private Sub txt_disc_LostFocus()
    If Txt_Disc.Text = "" Then
        Txt_Disc.Text = 0
    End If
End Sub

Private Sub Txt_Disc_PPN_Change()
On Error Resume Next

    If lbl_total_bayar.Caption = 0 Then Exit Sub
    
    Dim disc As Double
        If Txt_Disc_PPN.Text = "" Then
            disc = 0
        Else
            disc = Trim(Txt_Disc_PPN.Text)
        End If
    
    Dim Hasil As Double
        Hasil = CDbl(lbl_total_bayar.Caption) * (CDbl(disc) / 100)
        
        Lbl_Tot_DiscP.Caption = Hasil
    
End Sub

Private Sub Txt_Disc_PPN_GotFocus()
    Call Focus_(Txt_Disc_PPN)
End Sub

Private Sub Txt_Disc_PPN_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_jml_bayar.SetFocus
End Sub

Private Sub Txt_Disc_PPN_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",")) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Txt_Disc_PPN_LostFocus()
    If Txt_Disc_PPN.Text = "" Then Txt_Disc_PPN.Text = 0
End Sub

Private Sub txt_faktur_GotFocus()
    
'    txt_faktur.SelStart = 0
'    txt_faktur.SelLength = Len(txt_faktur)
    
End Sub

Private Sub txt_faktur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then

'    Dim konfirm As Integer
'        If txt_faktur.Text = "" Then
'
'             konfirm = CInt(MsgBox("No faktur tu gak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
'
'            Exit Sub
'        End If
'
'        Dim sql As String
'        Dim rs As Recordset
'
'        Dim sql1 As String
'        Dim rs1 As Recordset
'
'        sql1 = "select * from qr_faktur_penjualan_group where no_faktur='" & Trim(txt_faktur.Text) & "'"
'
'        Set rs1 = New ADODB.Recordset
'            rs1.Open sql1, cn
'
'        If Not rs1.EOF Then
'
'
'            Rubah = True
'            hap_detail = False
'
'            dtp_tgl.Value = rs1!tgl
'            Txt_Member.Text = IIf(Not IsNull(rs1!No_Member), rs1!No_Member, "")
'            Txt_Nama_Membr.Text = IIf(Not IsNull(rs1!nama), rs1!nama, "")
'            Txt_Tgl_Membr.Text = IIf(Not IsNull(rs1!Tanggal2), rs1!Tanggal2, "")
'            Txt_Tot_Disc.Text = IIf(Not IsNull(rs1!tot_disc), rs1!tot_disc, 0)
'            Lbl_Tot_Disc.Caption = Format(IIf(Not IsNull(rs1!Tot_Nilai_disc), rs1!Tot_Nilai_disc, 0), "currency")
'
'            Lbl_Harus.Caption = IIf(Not IsNull(rs1!Tot_Sth_Disc), rs1!Tot_Sth_Disc, 0)
'            Lbl_Harus.Caption = Format(Lbl_Harus.Caption, "currency")
'
'            txt_ppn.Text = IIf(Not IsNull(rs1!ppn), rs1!ppn, 0)
'
'            kosong_daftar
'
'            If Txt_Member.Text <> "" Then
'                InMember = True
'            Else
'                InMember = False
'            End If
'
'                    sql = "select * from qr_penjualan_sebenarnya where no_faktur='" & Trim(txt_faktur.Text) & "'"
'
'                    Set rs = New ADODB.Recordset
'                        rs.Open sql, cn
'
'                    With rs
'
'                    If Not .EOF Then
'
'                        Isi_Grid_Transaksi
'
'                        If Txt_Member.Text <> "" Then
'                            Txt_Member_KeyDown 13, 0
'                        ElseIf Txt_Member.Text = "" And Txt_Tot_Disc.Text <> 0 Then
'                            Txt_Tot_Disc_Change
'                        End If
'
'                        dtp_tgl.SetFocus
'                  End If
'                  End With
'
'          Else
'
'            Rubah = False
'            hap_detail = False
'
'             dtp_tgl.Value = Date
'             Txt_Order.Text = ""
'             txt_kode_counter.Text = ""
'             lbl_nama_counter.Caption = ""
'             Txt_Kode_Barang.Text = ""
'             lbl_nama_barang.Caption = ""
'             txt_beli.Text = 0
'             txt_disc.Text = 0
'             lbl_harga.Caption = 0
'             lbl_grand_total.Caption = 0
'             Txt_Member.Text = ""
'             Txt_Nama_Membr.Text = ""
'             Txt_Tgl_Membr.Text = ""
'             Txt_Tot_Disc.Text = 0
'             lbl_total_bayar.Caption = 0
'             Lbl_Tot_Disc.Caption = 0
'             Lbl_Harus.Caption = 0
'             txt_jml_bayar.Text = 0
'             lbl_kembali.Caption = 0
'
'             arr_daftar.ReDim 0, 0, 0, 0
''                arr_daftar.ReDim 1, 1, 1, 1
'                 grd_daftar.ReBind
'                 grd_daftar.Refresh
'
''                If grd_daftar.Columns(11).FooterText = "" Then
'                 grd_daftar.Columns(11).FooterText = 0
''                End If
'
''                If grd_daftar.Columns(7).FooterText = "" Then
'                 grd_daftar.Columns(7).FooterText = 0 & "%"
''                End If
'
'
'          End If

'                    End With

        txt_kode_counter.SetFocus
    End If
    
    If KeyCode = vbKeyF2 Then
        baru_lagi
    End If
    
    If KeyCode = vbKeyF4 Then
        frm_cfaktur.Show
    End If
    
    If KeyCode = vbKeyF1 Then
        frm_ganti_pwd.Show
    End If
    
End Sub

Private Sub Isi_Grid_Transaksi()
    
    Dim sql As String
    Dim rs As Recordset
    Dim a As Long
    Dim no_faktur, kode_barang, nama_barang, qty, satuan, harga, harga_sebenarnya, disc, harga_disc, carge, harga_charge, total, id_barang, no_order, perdisc As String
    Dim jml_tot As Double
    Dim jml_diskon As Double
    Dim woi As Double
    
        sql = "select * from qr_penjualan_sebenarnya where no_faktur='" & Trim(txt_faktur.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenKeyset
            
            a = 1
            jml_tot = 0
            Tot_Tanpa_Disc = 0: woi = 0
            
            arr_daftar.ReDim 0, 0, 0, 0
            arr_daftar.ReDim 1, 1, 1, 1
                grd_daftar.ReBind
                grd_daftar.Refresh
         
            If grd_daftar.Columns(11).FooterText = "" Then
                grd_daftar.Columns(11).FooterText = 0
            End If
        
            If grd_daftar.Columns(7).FooterText = "" Then
                grd_daftar.Columns(7).FooterText = 0 & "%"
                jml_diskon = 0
            End If

         
            With rs
                
                Do While Not .EOF
                    arr_daftar.ReDim 1, a, 0, grd_daftar.Columns.Count
                        grd_daftar.ReBind
                        grd_daftar.Refresh
                    
                    no_faktur = IIf(Not IsNull(!no_faktur), !no_faktur, "")
                    kode_barang = IIf(Not IsNull(!kode_barang), !kode_barang, "")
                    nama_barang = IIf(Not IsNull(!nama_barang), !nama_barang, "")
                    qty = IIf(Not IsNull(!qty), !qty, 0)
                    satuan = IIf(Not IsNull(!satuan), !satuan, "")
                    harga = IIf(Not IsNull(!harga_satuan), !harga_satuan, 0)
                    harga_sebenarnya = IIf(Not IsNull(!harga_sebenarnya), !harga_sebenarnya, 0)
                    disc = IIf(Not IsNull(!disc), !disc, "0%")
                    harga_disc = IIf(Not IsNull(!harga_disc), !harga_disc, 0)
                    carge = IIf(Not IsNull(!cash), !cash, "0%")
                    harga_charge = IIf(Not IsNull(!harga_cash), !harga_cash, 0)
                    total = IIf(Not IsNull(!total_harga), !total_harga, 0)
                    id_barang = IIf(Not IsNull(!id_barang), !id_barang, "")
                    no_order = IIf(Not IsNull(!no_order), !no_order, "")
                    perdisc = IIf(Not IsNull(!Per_disc), !Per_disc, 0)
                    
                    arr_daftar(a, 0) = no_faktur
                    arr_daftar(a, 1) = kode_barang
                    arr_daftar(a, 2) = nama_barang
                    arr_daftar(a, 3) = qty
                    arr_daftar(a, 4) = satuan
                    arr_daftar(a, 5) = harga
                    arr_daftar(a, 6) = harga_sebenarnya
                    arr_daftar(a, 7) = disc
                    arr_daftar(a, 8) = harga_disc
                    arr_daftar(a, 9) = carge
                    arr_daftar(a, 10) = harga_charge
                    arr_daftar(a, 11) = total
                    arr_daftar(a, 12) = id_barang
                    arr_daftar(a, 13) = no_order
                    arr_daftar(a, 14) = perdisc
                    
                    woi = woi + (CDbl(qty) * CDbl(harga))
                    
                    If perdisc <> 2 Then
                        Tot_Tanpa_Disc = CDbl(Tot_Tanpa_Disc) + (CDbl(qty) * CDbl(harga))
                    End If
                    
                    jml_tot = CDbl(jml_tot) + CDbl(total)
                    
                    jml_diskon = CDbl(jml_diskon) + CDbl(Mid(disc, 1, Len(disc) - 1))
                    
                a = a + 1
                .MoveNext
                Loop
                
                grd_daftar.ReBind
                grd_daftar.Refresh
                
                grd_daftar.MoveFirst
                
                
                lbl_total_bayar.Caption = Format(woi, "currency")
                
'             If txt_disc.Text <> 0 Then
'                Lbl_Harus.Caption = Format(Tot_Tanpa_Disc, "currency")
'             Else
                
                Dim discc As Double
                    discc = IIf((Lbl_Tot_Disc.Caption = ""), 0, Lbl_Tot_Disc.Caption)
                
                Dim Hrs_h As Double
                    Hrs_h = jml_tot - discc
                Lbl_Harus.Caption = Format(Hrs_h, "currency")
'             End If
'
                'jml_biaya = CDbl(lbl_grand_total.Caption) + CDbl(grd_daftar.Columns(11).FooterText)
                grd_daftar.Columns(11).FooterText = Format(jml_tot, "Currency")
                                
'                txt_jml_bayar.Text = Format(jml_biaya, "###,###,###")
'                lbl_kembali.Caption = "Rp." & Format(jml_biaya, "###,###,###")
                     
                
                grd_daftar.Columns(7).FooterText = jml_diskon & "%"
                
                
'                If grd_daftar.Columns(9).FooterText = "" Then
'                grd_daftar.Columns(9).FooterText = 0 & "%"
'                jml_cash = 0
'                End If
                
'                jml_cash = CDbl(txt_charge.Text) + CDbl(Mid(grd_daftar.Columns(9).FooterText, 1, Len(grd_daftar.Columns(9).FooterText) - 1))
'                grd_daftar.Columns(9).FooterText = jml_cash & "%"
            
                
                
            End With
        
    
End Sub

Private Sub txt_faktur_LostFocus()
    
    If Len(txt_faktur.Text) < 7 Then
        Dim konfirm As Integer
            konfirm = CInt(MsgBox("No faktur tidak boleh kurang dari 7 huruf", vbOKOnly + vbInformation, "Informasi"))
            
            txt_faktur.SetFocus
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

Private Sub txt_jml_bayar_Change()
'On Error Resume Next

    lbl_kembali.Caption = 0
    
   ' If txt_jml_bayar.Text = "" Then txt_jml_bayar.Text = 0
    
    If txt_jml_bayar.Text <> "" Then
        Dim yang_dibayar, kembali As Double
        yang_dibayar = txt_jml_bayar.Text
        txt_jml_bayar.Text = Format(txt_jml_bayar.Text, "###,###,###")
        txt_jml_bayar.SelStart = Len(txt_jml_bayar.Text)
        kembali = CDbl(yang_dibayar) - CDbl(Lbl_Harus.Caption)
        lbl_kembali.Caption = "Rp." & Format(kembali, "###,###,###")
        
    End If

End Sub

Private Sub txt_jml_bayar_GotFocus()
    txt_jml_bayar.SelStart = 0
    txt_jml_bayar.SelLength = Len(txt_jml_bayar)
End Sub

Private Sub txt_jml_bayar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
    If Rubah = False Then
        Call ok
    Else
        Call Ok_Rubah(True)
    End If
    
'    If cek_faktur.Value = vbChecked Then
'        If MsgBox("Apakah anda ingin mencetak bukti pembayaran", vbYesNo + vbQuestion, "Konfirmasi") = vbYes Then
'
'            noff = Trim(txt_faktur.Text)
'
'            If txt_jml_bayar.Text = "" Then byyr = 0
'            If txt_jml_bayar.Text = 0 Then
'                byyr = 0
'            Else
'                byyr = Replace(txt_jml_bayar, ",", "")
'            End If
'
'            kemm = CDbl(lbl_kembali.Caption)
'
'
'            Load Frm_Lap_BuktiByar
'                Frm_Lap_BuktiByar.Show
'
'
'        Else
'
'            Call baru_lagi
'
'        End If
'    End If
    
    hap_detail = False
    
    End If
    
    If KeyCode = vbKeyF2 Then
        baru_lagi
    End If
    
    If KeyCode = vbKeyF4 Then
        frm_cfaktur.Show
    End If
    
    If KeyCode = vbKeyF1 Then
        frm_ganti_pwd.Show
    End If
    
End Sub

Public Sub baru_lagi()
        
    hap_detail = False
    Rubah = False
    rubah1 = False
        
        Txt_Disc.Enabled = True
        
        dtp_tgl.Value = Date
        txt_faktur.Text = ""
        Call isi_faktur
        txt_kode_counter.Text = ""
        Txt_Kode_Barang.Text = ""
        lbl_nama_counter.Caption = ""
        lbl_nama_barang.Caption = ""
        
        lbl_harga.Caption = 0
        txt_beli.Text = 0
        Txt_Disc.Text = 0
        txt_charge.Text = 0
        lbl_grand_total.Caption = 0
        lbl_total_bayar.Caption = 0
        txt_jml_bayar.Text = 0
        lbl_kembali.Caption = 0
        Txt_Order.Text = ""
        
        Txt_Member.Text = ""
        Txt_Nama_Membr.Text = ""
        Txt_Tgl_Membr.Text = ""
        
        Txt_Tot_Disc.Text = 0
        Txt_Disc_PPN.Text = 0
        Lbl_Tot_Disc.Caption = 0
        Lbl_Tot_DiscP.Caption = 0
        Lbl_Harus.Caption = 0
        txt_ppn.Text = 0
        
        kosong_daftar
        grd_daftar.Columns(7).FooterText = ""
        grd_daftar.Columns(9).FooterText = ""
        grd_daftar.Columns(11).FooterText = ""
        
        txt_faktur.SetFocus
End Sub

Private Sub txt_jml_bayar_KeyPress(KeyAscii As Integer)
   If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
   End If
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_barang.Visible = False
        Txt_Kode_Barang.SetFocus
    End If
    
    If KeyCode = 13 Then
        If arr_barang.UpperBound(1) > 0 Then
            Txt_Kode_Barang.Text = kode_barang
            kasih_tahu
            pic_barang.Visible = False
            Txt_Kode_Barang.SetFocus
        End If
    End If
 
    
End Sub

Private Sub txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo er_ko

    Dim sql1 As String
    Dim rs_barang As New ADODB.Recordset
        
    
        
 If arr_barang.UpperBound(1) > 0 Then
 
        
                
        sql1 = "select top 100 nama_counter,kode,nama_barang,Satuan from qr_barang where id_counter=" & id_counter & " and aktif=1"
        
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
      
Exit Sub

er_ko:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
                       
End Sub

Private Sub txt_kode_barang_GotFocus()
    Txt_Kode_Barang.SelStart = 0
    Txt_Kode_Barang.SelLength = Len(Txt_Kode_Barang)
End Sub

Private Sub txt_kode_barang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Txt_Kode_Barang.Text = ""
        txt(0).Text = ""
        txt(1).Text = ""
        pic_barang.Visible = True
        txt(0).SetFocus
    End If
    
    If KeyCode = 13 Then
        txt_beli.SetFocus
    End If
    
    If KeyCode = vbKeyF2 Then
        baru_lagi
    End If
    
    If KeyCode = vbKeyF4 Then
        frm_cfaktur.Show
    End If
    
    If KeyCode = vbKeyF1 Then
        frm_ganti_pwd.Show
    End If
    
End Sub

Private Sub txt_kode_barang_LostFocus()
    
    If txt_kode_counter.Text = "" Then
        Dim konfirm As Integer
            konfirm = CInt(MsgBox("Kode jenis barang harus diisi", vbOKOnly + vbInformation, "Informasi"))
             txt_kode_counter.SetFocus
             Exit Sub
    End If
            
    
    If Txt_Kode_Barang.Text <> "" Then
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
    
    If KeyCode = 13 And txt_kode_counter.Text <> "" Then
        Txt_Kode_Barang.SetFocus
    ElseIf KeyCode = 13 And txt_kode_counter.Text = "" And arr_daftar.UpperBound(1) > 0 Then
        Txt_Member.SetFocus
        'txt_jml_bayar.SetFocus
    End If
    
    If KeyCode = vbKeyF2 Then
        baru_lagi
    End If
    
    If KeyCode = vbKeyF4 Then
        frm_cfaktur.Show
    End If
    
    If KeyCode = vbKeyF1 Then
        frm_ganti_pwd.Show
    End If
    
End Sub

Private Sub txt_kode_counter_LostFocus()

On Error GoTo er_ls

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
    
    Exit Sub
    
er_ls:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub

Private Sub Txt_Member_GotFocus()
    Call Focus_(Txt_Member)
End Sub

Private Sub Txt_Member_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyF3 Then Browse_Member_Click
If KeyCode = 13 Then
        
    InMember = False
        
    Dim konfirm As Integer
    If lbl_total_bayar.Caption = 0 Then
        konfirm = CInt(MsgBox("Isi dulu barang barang yang dibeli", vbOKOnly + vbInformation, "Informasi"))
        
        Exit Sub
    End If
        
    If Txt_Member.Text <> "" Then
    
    Dim sql As String
    Dim rs As Recordset
        
    
        sql = "select No_Member,Nama,Tanggal2,Aktif,Disc from qr_member where No_Member='" & Trim(Txt_Member.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, cn
            
            With rs
                If Not .EOF Then
                    
                    If !aktif = 0 Then
                        konfirm = CInt(MsgBox("No. member sudah tidak aktif lagi", vbOKOnly + vbInformation, "Informasi"))
                        
                        Exit Sub
                    End If
                    
'                    If Date > !Tanggal2 Then
'                        Konfirm = CInt(MsgBox("No. member anda sudah habis,silahkan registrasi ulang kartu member anda", vbOKOnly + vbInformation, "Informasi"))
'
'                        Exit Sub
'                    End If
                    
                    If CDate(!Tanggal2) < Date Then
                        konfirm = CInt(MsgBox("Kartu member anda sudah habis sejak " & !Tanggal2, vbOKOnly + vbInformation, "Informasi"))
                        Exit Sub
                    End If
                    
                    Txt_Nama_Membr.Text = IIf(Not IsNull(!nama), !nama, "")
                    Txt_Tgl_Membr.Text = !Tanggal2
                    Txt_Tot_Disc.Text = IIf(Not IsNull(!disc), !disc, 0)
                    
                    Dim Nilai_Disc As Double
                        
                        If Txt_Tot_Disc.Text <> 0 Then
                            Nilai_Disc = Tot_Tanpa_Disc * (CDbl(Txt_Tot_Disc.Text) / 100)
                        Else
                            Nilai_Disc = 0
                        End If
                        
                        Lbl_Tot_Disc.Caption = Format(Nilai_Disc, "currency")
                        
                    Dim harus As Double
                        
                        harus = CDbl(lbl_total_bayar.Caption) - CDbl(Lbl_Tot_Disc.Caption)
                            
                        Lbl_Harus.Caption = Format(harus, "currency")
                        txt_jml_bayar.Text = harus
                        lbl_kembali.Caption = 0
                        
                    Txt_Tot_Disc_Change
                    
                    hitung_ppn
                    
                    Txt_Tot_Disc.SetFocus
                    
                    
                    Exit Sub
                Else
                    
                    konfirm = CInt(MsgBox("No. member yang anda masukkanb tidak ditemukan", vbOKOnly + vbInformation, "Informasi"))
                    
                    Exit Sub
                End If
            End With
            
    End If
    
    
    Txt_Tot_Disc_Change
    
    hitung_ppn
    
    Txt_Tot_Disc.SetFocus
    
    End If
    
    
    If KeyCode = vbKeyF2 Then
        baru_lagi
    End If
    
    If KeyCode = vbKeyF4 Then
        frm_cfaktur.Show
    End If
    
    If KeyCode = vbKeyF1 Then
        frm_ganti_pwd.Show
    End If

    
End Sub

Private Sub Txt_Member_LostFocus()
    If Txt_Member.Text <> "" Then
        InMember = True
    Else
        Txt_Nama_Membr.Text = ""
        Txt_Tgl_Membr.Text = ""
    End If
        
End Sub

Private Sub Txt_Order_GotFocus()
    Call Focus_(Txt_Order)
End Sub

Private Sub Txt_Order_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Txt_Order.Text <> "" Then txt_kode_counter.SetFocus
    
    If KeyCode = vbKeyF2 Then
        baru_lagi
    End If
    
    If KeyCode = vbKeyF4 Then
        frm_cfaktur.Show
    End If
    
    If KeyCode = vbKeyF1 Then
        frm_ganti_pwd.Show
    End If

    
End Sub

Private Sub txt_ppn_Change()

    If txt_ppn.Text = "" Then txt_ppn.Text = 0
    
    hitung_ppn

End Sub

Private Sub txt_ppn_GotFocus()
    On Error Resume Next
    With txt_ppn
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txt_ppn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Tot_Disc.SetFocus
    
    If KeyCode = vbKeyF2 Then
        baru_lagi
    End If
    
    If KeyCode = vbKeyF4 Then
        frm_cfaktur.Show
    End If
    
    If KeyCode = vbKeyF1 Then
        frm_ganti_pwd.Show
    End If
    
End Sub

Private Sub txt_ppn_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",")) Then
        KeyAscii = 0
    End If
End Sub
Private Sub hitung_ppn()

    Dim tot_b As Double
        If lbl_total_bayar.Caption = 0 Then
            tot_b = 0
        Else
            tot_b = lbl_total_bayar.Caption
        End If
        
    Dim ppn As Double
        ppn = (CDbl(tot_b) - CDbl(Lbl_Tot_Disc.Caption)) * (CDbl(txt_ppn.Text) / 100)
        
        Dim sebelum_ppn As Double
            sebelum_ppn = CDbl(tot_b) - CDbl(Lbl_Tot_Disc.Caption)
        
        harusppn = CDbl(sebelum_ppn) + CDbl(ppn)
        
        Lbl_Harus.Caption = Format(harusppn, "currency")
        txt_jml_bayar.Text = Format(harusppn, "currency")
        
End Sub

Private Sub Txt_Tot_Disc_Change()
On Error Resume Next

    If lbl_total_bayar.Caption = 0 Then Exit Sub
    If Txt_Tot_Disc.Text = "" Then
        Lbl_Tot_Disc.Caption = 0
    End If
    
    If Txt_Tot_Disc.Text = "" Then
        Lbl_Tot_Disc.Caption = 0
    End If
    
    If InMember = True Then
    
        Txt_Member.Text = ""
        Txt_Nama_Membr.Text = ""
        Txt_Tgl_Membr.Text = ""
        
    End If
    
        
    Dim nil_disc As Double
        If Txt_Tot_Disc.Text = "" Then
            nil_disc = 0
        Else
            nil_disc = Trim(Txt_Tot_Disc.Text)
        End If
        
        If arr_daftar.UpperBound(1) > 0 Then
        
        Dim a As Long
        Dim harga_disc, harga_s As Double
        Dim n As Double
        Dim tot, Harga_Bawah, n_d As Double
            
            harga_disc = 0: harga_s = 0
            n = 0: Tot_Tanpa_Disc = 0
            tot = 0: Harga_Bawah = 0
            n_d = 0
            
            For a = arr_daftar.LowerBound(1) To arr_daftar.UpperBound(1)
            
            If arr_daftar(a, 14) <> 2 Then
                harga_s = CDbl(arr_daftar(a, 3)) * CDbl(arr_daftar(a, 5))
                tot = tot + harga_s
                harga_disc = harga_s * (CDbl(nil_disc) / 100)
                n = harga_disc
                harga_disc = harga_s - harga_disc
                
                Harga_Bawah = Harga_Bawah + harga_disc
                Tot_Tanpa_Disc = Tot_Tanpa_Disc + harga_s
                
                n_d = CDbl(n_d) + CDbl(nil_disc)
                
                arr_daftar(a, 7) = nil_disc & "%"
                arr_daftar(a, 8) = n
                arr_daftar(a, 11) = harga_disc
                
                 
            Else
                
                Harga_Bawah = Harga_Bawah + (CDbl(arr_daftar(a, 3)) * CDbl(arr_daftar(a, 5)))
                tot = tot + CDbl(arr_daftar(a, 11))
                
            End If
            
            Next
            
            grd_daftar.ReBind
            grd_daftar.Refresh
            
            lbl_total_bayar.Caption = Format(tot, "currency")

            If grd_daftar.Columns(7).FooterText = "" Then
                grd_daftar.Columns(7).FooterText = 0
            End If
            
            grd_daftar.Columns(7).FooterText = n_d & "%"

            If grd_daftar.Columns(11).FooterText = "" Then
                grd_daftar.Columns(11).FooterText = 0
            End If
            
            grd_daftar.Columns(11).FooterText = Format(Harga_Bawah, "currency")
    
        End If
        
    Dim Hasil As Double
        Hasil = Tot_Tanpa_Disc * (nil_disc / 100)
        
     lbl_total_bayar.Caption = Format(tot, "currency")
        
     Lbl_Tot_Disc.Caption = Format(Hasil, "currency")
    
    Dim harus As Double
          
          harus = CDbl(lbl_total_bayar.Caption) - CDbl(Lbl_Tot_Disc.Caption)
            
          Lbl_Harus.Caption = Format(harus, "currency")
          txt_jml_bayar.Text = Format(harus, "###,###,###")
          lbl_kembali.Caption = 0
        
          hitung_ppn
            
            
    
End Sub

Private Sub Txt_Tot_Disc_GotFocus()
    Call Focus_(Txt_Tot_Disc)
End Sub

Private Sub Txt_Tot_Disc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_jml_bayar.SetFocus
    
    If KeyCode = vbKeyF2 Then
        baru_lagi
    End If
    
    If KeyCode = vbKeyF4 Then
        frm_cfaktur.Show
    End If
    
    If KeyCode = vbKeyF1 Then
        frm_ganti_pwd.Show
    End If

    
End Sub

Private Sub Txt_Tot_Disc_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",")) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Txt_Tot_Disc_LostFocus()
    If Txt_Tot_Disc.Text = "" Then Txt_Tot_Disc.Text = 0
End Sub
