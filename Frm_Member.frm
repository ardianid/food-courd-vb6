VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{EC76FE26-BAFD-4E89-AA40-E748DA83A570}#1.0#0"; "IsButton_Ard.ocx"
Begin VB.Form Frm_Member 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MEMBER"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   8775
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Info 
      Height          =   4335
      Left            =   480
      TabIndex        =   67
      Top             =   3000
      Visible         =   0   'False
      Width           =   6615
      _Version        =   65536
      _ExtentX        =   11668
      _ExtentY        =   7646
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Member.frx":0000
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Member.frx":001C
      Childs          =   "Frm_Member.frx":00C8
      Begin VB.TextBox Txt_Cr_Info 
         Height          =   315
         Index           =   1
         Left            =   3720
         TabIndex        =   74
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox Txt_Cr_Info 
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   73
         Top             =   840
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   68
         Top             =   480
         Width           =   6015
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Info 
         Height          =   2895
         Left            =   240
         OleObjectBlob   =   "Frm_Member.frx":00E4
         TabIndex        =   69
         Top             =   1200
         Width           =   6015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   210
         Index           =   4
         Left            =   3120
         TabIndex        =   72
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Member"
         Height          =   210
         Index           =   3
         Left            =   360
         TabIndex        =   71
         Top             =   840
         Width           =   1005
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
         Index           =   22
         Left            =   240
         TabIndex        =   70
         Top             =   240
         Width           =   2490
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Hapus 
      Height          =   4335
      Left            =   -240
      TabIndex        =   59
      Top             =   2760
      Visible         =   0   'False
      Width           =   6615
      _Version        =   65536
      _ExtentX        =   11668
      _ExtentY        =   7646
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Member.frx":2EF6
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Member.frx":2F12
      Childs          =   "Frm_Member.frx":2FBE
      Begin VB.TextBox Txt_Cr_Hapus 
         Height          =   315
         Index           =   1
         Left            =   3720
         TabIndex        =   66
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox Txt_Cr_Hapus 
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   65
         Top             =   840
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   60
         Top             =   480
         Width           =   6015
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Hapus 
         Height          =   2895
         Left            =   240
         OleObjectBlob   =   "Frm_Member.frx":2FDA
         TabIndex        =   61
         Top             =   1200
         Width           =   6015
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
         Index           =   21
         Left            =   240
         TabIndex        =   64
         Top             =   240
         Width           =   2490
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Member"
         Height          =   210
         Index           =   2
         Left            =   360
         TabIndex        =   63
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   210
         Index           =   1
         Left            =   3120
         TabIndex        =   62
         Top             =   840
         Width           =   450
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Rubah 
      Height          =   4335
      Left            =   120
      TabIndex        =   51
      Top             =   1920
      Visible         =   0   'False
      Width           =   6615
      _Version        =   65536
      _ExtentX        =   11668
      _ExtentY        =   7646
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Member.frx":5DED
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Member.frx":5E09
      Childs          =   "Frm_Member.frx":5EB5
      Begin VB.TextBox Txt_Cr_Rubah 
         Height          =   315
         Index           =   1
         Left            =   3720
         TabIndex        =   58
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox Txt_Cr_Rubah 
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   57
         Top             =   840
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   52
         Top             =   480
         Width           =   6015
      End
      Begin TrueOleDBGrid60.TDBGrid Grid_Rubah 
         Height          =   2895
         Left            =   240
         OleObjectBlob   =   "Frm_Member.frx":5ED1
         TabIndex        =   53
         Top             =   1200
         Width           =   6015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   210
         Index           =   0
         Left            =   3120
         TabIndex        =   56
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Member"
         Height          =   210
         Index           =   14
         Left            =   360
         TabIndex        =   55
         Top             =   840
         Width           =   1005
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
         Index           =   20
         Left            =   240
         TabIndex        =   54
         Top             =   240
         Width           =   2490
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDB_Aktivasi 
      Height          =   2055
      Left            =   -5400
      TabIndex        =   39
      Top             =   240
      Visible         =   0   'False
      Width           =   5775
      _Version        =   65536
      _ExtentX        =   10186
      _ExtentY        =   3625
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "Frm_Member.frx":8CE4
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "Frm_Member.frx":8D00
      Childs          =   "Frm_Member.frx":8DAC
      Begin MSComCtl2.DTPicker Dtp_Tgl1 
         Height          =   315
         Left            =   1440
         TabIndex        =   46
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   12582912
         CalendarTitleForeColor=   16777215
         Format          =   20578305
         CurrentDate     =   39211
      End
      Begin VB.TextBox Txt_Disc 
         Height          =   315
         Left            =   1440
         TabIndex        =   44
         Top             =   360
         Width           =   615
      End
      Begin MSComCtl2.DTPicker Dtp_Tgl2 
         Height          =   315
         Left            =   3600
         TabIndex        =   48
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   12582912
         CalendarTitleForeColor=   16777215
         Format          =   20578305
         CurrentDate     =   39211
      End
      Begin IsButton_Ard.isButton Cmd_Ok 
         Height          =   495
         Left            =   3600
         TabIndex        =   49
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         Icon            =   "Frm_Member.frx":8DC8
         Style           =   8
         Caption         =   "&OK"
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
      Begin IsButton_Ard.isButton Cmd_Cancel 
         Height          =   495
         Left            =   4560
         TabIndex        =   50
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         Icon            =   "Frm_Member.frx":8DE4
         Style           =   8
         Caption         =   "&CANCEL"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S/D"
         Height          =   210
         Index           =   19
         Left            =   3120
         TabIndex        =   47
         Top             =   720
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   210
         Index           =   18
         Left            =   2160
         TabIndex        =   45
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   17
         Left            =   1320
         TabIndex        =   43
         Top             =   720
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   16
         Left            =   1320
         TabIndex        =   42
         Top             =   360
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   210
         Index           =   15
         Left            =   360
         TabIndex        =   41
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disc"
         Height          =   210
         Index           =   14
         Left            =   360
         TabIndex        =   40
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   3960
      TabIndex        =   33
      Top             =   6480
      Width           =   4695
      Begin IsButton_Ard.isButton Cmd_Add 
         Height          =   375
         Left            =   1080
         TabIndex        =   34
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Icon            =   "Frm_Member.frx":8E00
         Style           =   8
         Caption         =   "Tambah &Aktivasi"
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
      Begin IsButton_Ard.isButton Cmd_Dell 
         Height          =   375
         Left            =   2880
         TabIndex        =   35
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Icon            =   "Frm_Member.frx":8E1C
         Style           =   8
         Caption         =   "&Hapus Aktivasi"
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
   End
   Begin VB.Frame Frame_Nav 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cari Member"
      Height          =   735
      Left            =   120
      TabIndex        =   28
      Top             =   6480
      Width           =   3735
      Begin IsButton_Ard.isButton Cmd_Nav 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Icon            =   "Frm_Member.frx":8E38
         Style           =   8
         Caption         =   "<<"
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
      Begin IsButton_Ard.isButton Cmd_Nav 
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   30
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Icon            =   "Frm_Member.frx":8E54
         Style           =   8
         Caption         =   "<"
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
      Begin IsButton_Ard.isButton Cmd_Nav 
         Height          =   375
         Index           =   2
         Left            =   1440
         TabIndex        =   31
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Icon            =   "Frm_Member.frx":8E70
         Style           =   8
         Caption         =   ">"
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
      Begin IsButton_Ard.isButton Cmd_Nav 
         Height          =   375
         Index           =   3
         Left            =   2040
         TabIndex        =   32
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Icon            =   "Frm_Member.frx":8E8C
         Style           =   8
         Caption         =   ">>"
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
   End
   Begin TrueOleDBGrid60.TDBGrid Grid_Member 
      Height          =   2535
      Left            =   120
      OleObjectBlob   =   "Frm_Member.frx":8EA8
      TabIndex        =   20
      Top             =   3840
      Width           =   8535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.CheckBox Check_Aktif 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Aktif"
         Height          =   375
         Left            =   1920
         TabIndex        =   38
         Top             =   3000
         Width           =   855
      End
      Begin IsButton_Ard.isButton Cmd_Tambah 
         Height          =   615
         Left            =   7320
         TabIndex        =   21
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         Icon            =   "Frm_Member.frx":D564
         Style           =   8
         Caption         =   "&Tambah"
         iNonThemeStyle  =   0
         ShowFocus       =   -1  'True
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
      Begin MSMask.MaskEdBox Txt_TglLhr 
         Height          =   315
         Left            =   1920
         TabIndex        =   19
         Top             =   2640
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Txt_Telp2 
         Height          =   315
         Left            =   3600
         TabIndex        =   18
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Txt_Telp1 
         Height          =   315
         Left            =   1920
         TabIndex        =   17
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Txt_Alamat 
         Height          =   795
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox Txt_Ktp 
         Height          =   315
         Left            =   1920
         TabIndex        =   13
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox Txt_Nama 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox Txt_Nom 
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin IsButton_Ard.isButton Cmd_Rubah 
         Height          =   615
         Left            =   7320
         TabIndex        =   22
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         Icon            =   "Frm_Member.frx":D580
         Style           =   8
         Caption         =   "&Rubah"
         iNonThemeStyle  =   0
         ShowFocus       =   -1  'True
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
      Begin IsButton_Ard.isButton Cmd_Hapus 
         Height          =   615
         Left            =   7320
         TabIndex        =   23
         Top             =   1440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         Icon            =   "Frm_Member.frx":D59C
         Style           =   8
         Caption         =   "&Hapus"
         iNonThemeStyle  =   0
         ShowFocus       =   -1  'True
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
      Begin IsButton_Ard.isButton Cmd_Keluar 
         Height          =   615
         Left            =   7320
         TabIndex        =   24
         Top             =   2640
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         Icon            =   "Frm_Member.frx":D5B8
         Style           =   8
         Caption         =   "&Keluar"
         iNonThemeStyle  =   0
         ShowFocus       =   -1  'True
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
         Height          =   615
         Left            =   7320
         TabIndex        =   25
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         Icon            =   "Frm_Member.frx":D5D4
         Style           =   8
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
      Begin IsButton_Ard.isButton Cmd_Batal 
         Height          =   615
         Left            =   7320
         TabIndex        =   26
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         Icon            =   "Frm_Member.frx":D5F0
         Style           =   8
         Caption         =   "&Batal"
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
      Begin IsButton_Ard.isButton Cmd_Info 
         Height          =   615
         Left            =   7320
         TabIndex        =   27
         Top             =   2040
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         Icon            =   "Frm_Member.frx":D60C
         Style           =   8
         Caption         =   "&Info"
         iNonThemeStyle  =   0
         ShowFocus       =   -1  'True
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
      Begin VB.Label Lbl_Info 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lbl_Info"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   7425
         TabIndex        =   75
         Top             =   3360
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   13
         Left            =   1800
         TabIndex        =   37
         Top             =   3000
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Member"
         Height          =   210
         Index           =   12
         Left            =   360
         TabIndex        =   36
         Top             =   3000
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   11
         Left            =   1800
         TabIndex        =   16
         Top             =   2280
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Telp."
         Height          =   210
         Index           =   10
         Left            =   360
         TabIndex        =   15
         Top             =   2280
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   9
         Left            =   1800
         TabIndex        =   10
         Top             =   2640
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   8
         Left            =   1800
         TabIndex        =   9
         Top             =   1440
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   7
         Left            =   1800
         TabIndex        =   8
         Top             =   1080
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   6
         Left            =   1800
         TabIndex        =   7
         Top             =   720
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   210
         Index           =   5
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No KTP/SIM"
         Height          =   210
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Lahir"
         Height          =   210
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Top             =   2640
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   210
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   210
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Member"
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1005
      End
   End
End
Attribute VB_Name = "Frm_Member"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rubah As Boolean
Dim Arr_Member As New XArrayDB

Dim Moving As Boolean
Dim yold, xold As Long

Dim Arr_Rubah As New XArrayDB
Dim Arr_Hapus As New XArrayDB
Dim Arr_Info As New XArrayDB
Dim Rs_Nav As Recordset
Dim Tambah_Akt As Boolean


Private Sub Isi_Semua(ByVal rec As Recordset)
On Error Resume Next

    With rec
    
        If .BOF Then .MoveFirst
        If .EOF Then .MoveLast
        
        Txt_Nom.Text = IIf(Not IsNull(!No_Member), !No_Member, "")
        Txt_Nama.Text = IIf(Not IsNull(!nama), !nama, "")
        Txt_Ktp.Text = IIf(Not IsNull(!No_KtpSim), !No_KtpSim, "")
        Txt_Alamat.Text = IIf(Not IsNull(!alamat), !alamat, "")
        Txt_Telp1.Text = IIf(Not IsNull(!No_Telp1), !No_Telp1, "")
        Txt_Telp2.Text = IIf(Not IsNull(!No_Telp2), !No_Telp2, "")
        
        Dim Tgl_Lhr As String
            Tgl_Lhr = IIf(Not IsNull(!Tgl_Lhr), !Tgl_Lhr, "")
            
            If Tgl_Lhr = "11/11/1111" Then
                Txt_TglLhr.Text = "__/__/____"
            Else
                Txt_TglLhr.Text = Tgl_Lhr
            End If
        
        Dim aktif As Integer
            aktif = IIf(Not IsNull(!aktif), !aktif, 0)
            
            If aktif = 0 Then
                Check_Aktif.Value = vbUnchecked
            Else
                Check_Aktif.Value = vbChecked
            End If
        
        Isi_Grid_Member
        
        If .RecordCount = 0 Then
            Lbl_Info.Caption = "Record ke " & 0 & " Dari " & .RecordCount & " Record"
        Else
            Lbl_Info.Caption = "Record ke " & .AbsolutePosition & " Dari " & .RecordCount & " Record"
        End If
        
    End With
    
End Sub

Private Sub Isi_Grid_Member()
    
    Dim sql As String
    Dim rs As Recordset
    
    Dim a As Long
    Dim tgl1, tgl2, disc, akt, yid As String
    
    sql = "select * from Tb_Member_Detail where Kode_Member='" & Trim(Txt_Nom.Text) & "'"
    
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
            
            tgl1 = IIf(Not IsNull(!Tanggal1), !Tanggal1, "")
            tgl2 = IIf(Not IsNull(!Tanggal2), !Tanggal2, "")
            disc = IIf(Not IsNull(!disc), !disc, 0)
            akt = IIf(Not IsNull(!aktif), !aktif, 0)
            yid = !id
            
            Arr_Member(a, 0) = a
            Arr_Member(a, 1) = tgl1
            Arr_Member(a, 2) = "S/D"
            Arr_Member(a, 3) = tgl2
            Arr_Member(a, 4) = disc
            Arr_Member(a, 5) = akt
            Arr_Member(a, 6) = yid
            
        a = a + 1
        .MoveNext
        Loop
        
        Grid_Member.ReBind
        Grid_Member.Refresh
        
        Grid_Member.MoveFirst
    
    End With
    
End Sub

Sub isi_grid(ByVal nama_grid As Object, ByVal arra As Object, ByVal rec As Recordset)
    
    arra.ReDim 0, 0, 0, 0
        nama_grid.ReBind
        nama_grid.Refresh
    
    Dim a As Long
    Dim nom, nama As String
    
        a = 1
        With rec
            Do While Not .EOF
                arra.ReDim 1, a, 0, nama_grid.Columns.Count
                nama_grid.ReBind
                nama_grid.Refresh
                    
                    nom = IIf(Not IsNull(!No_Member), !No_Member, "")
                    nama = IIf(Not IsNull(!nama), !nama, "")
                    
                    arra(a, 0) = nom
                    arra(a, 1) = nama
'                    arra(a, 2) = nama_pro
'                    arra(a, 3) = nama_cust
                
            
            a = a + 1
            .MoveNext
            Loop
            
            nama_grid.ReBind
            nama_grid.Refresh
            
            nama_grid.MoveFirst
            
        End With
    
End Sub

Private Sub Check_Aktif_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
    
        If Rubah = False Then
            Cmd_Add_Click
        Else
            Cmd_Add.TabIndex = Check_Aktif.TabIndex + 1
        End If
    
    End If
End Sub

Private Sub Cmd_Add_Click()
    
    Tambah_Akt = False
    
    With TDB_Aktivasi
        
        If .Visible = False Then
        
            txt_disc.Text = 0
            txt_disc.Enabled = True
            Dtp_Tgl1.Enabled = True
            Dtp_Tgl2.Enabled = True
            
            .Visible = True
                
            txt_disc.SetFocus
        
        Else
            .Visible = False
        End If
        
    End With
    
End Sub

Private Sub cmd_batal_Click()

    Rubah = False
'    Cetak = False
    
'        cmd_cetak.Enabled = True
                 
        Cmd_Tambah.Visible = True
        
      If tambah_form = True Then
        Cmd_Tambah.Enabled = True
      End If
      
        Cmd_Simpan.Visible = False
        
      Cmd_Rubah.Visible = True
      If edit_form = True Then
        Cmd_Rubah.Enabled = True
      End If
      
      If hapus_form = True Then
        Cmd_Hapus.Enabled = True
      End If
      
        Cmd_Info.Enabled = True
        Cmd_Keluar.Enabled = True
        
        Cmd_Batal.Visible = False
        
      '  Cmd_Add.Visible = False
       ' Cmd_Del.Visible = False
        
        Cmd_Add.Enabled = False
        Cmd_Dell.Enabled = False
        
        Cmd_Hapus.Visible = True
        Cmd_Info.Visible = True
        
Dim X As Object
    
    For Each X In Me
        If TypeOf X Is TextBox Then
            If UCase(Left(X.Name, 6)) <> UCase("txt_cr") Then
                X.Enabled = False
            End If
        End If
        
'            If TypeOf X Is TDBDate Then X.Enabled = False
            If TypeOf X Is CommandButton Then
                If UCase(X.Name) = UCase("cmd_browse") Then
                    X.Enabled = False
                End If
            End If
        If TypeOf X Is TDBContainer3D Then X.Visible = False
        
    Next

Set X = Nothing

Txt_TglLhr.Enabled = False
Frame_Nav.Enabled = True

Cmd_Tambah.SetFocus

Txt_Cr_Info_KeyUp 0, 0, 0
            Cmd_Nav_Click 3

End Sub

Private Sub cmd_cancel_Click()
    TDB_Aktivasi.Visible = False
    Tambah_Akt = False
    Cmd_Simpan.SetFocus
End Sub

Private Sub Cmd_Dell_Click()
On Error GoTo err_handler

    If Arr_Member.UpperBound(1) = 1 And Arr_Member(1, 1) = Empty Then Exit Sub
    
    
    Dim Konfirm As Integer
    If Arr_Member(Grid_Member.Bookmark, 5) = vbChecked Then
        Konfirm = CInt(MsgBox("Data aktivasi ini tidah bisa dihapus karena masih aktif", vbOKOnly + vbInformation, "Informasi"))
        
        On Error GoTo 0
        Exit Sub
    End If
    
    If Rubah = False Then
        
      If Arr_Member.UpperBound(1) > 1 Then
        Grid_Member.Delete
      Else
        Arr_Member.ReDim 0, 0, 0, 0
        Arr_Member.ReDim 1, 1, 1, 1
      End If
        
        Grid_Member.ReBind
        Grid_Member.Refresh
    
    Else
        
        cn.BeginTrans
        
        Dim sql As String
        Dim rs As Recordset
            sql = "delete from Tb_Member_Detail where Id=" & Arr_Member(Grid_Member.Bookmark, 6)
            
            Set rs = New ADODB.Recordset
                rs.Open sql, cn
            
        cn.CommitTrans
        Cmd_Batal.Enabled = False
           
      If Arr_Member.UpperBound(1) > 1 Then
        Grid_Member.Delete
      Else
        Arr_Member.ReDim 0, 0, 0, 0
        Arr_Member.ReDim 1, 1, 1, 1
      End If
        
        Grid_Member.ReBind
        Grid_Member.Refresh
            
    End If
        
    On Error GoTo 0
    Exit Sub
    
err_handler:
    
    If Rubah = True Then cn.RollbackTrans
        
            Konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
                Err.Clear

End Sub

Private Sub cmd_hapus_Click()

Frame_Nav.Enabled = False
Rubah = False
'Cetak = False

Cmd_Tambah.Enabled = False

Cmd_Rubah.Visible = False
Cmd_Batal.Visible = True

Cmd_Hapus.Enabled = False

Cmd_Info.Enabled = False

Cmd_Keluar.Enabled = False
'cmd_cetak.Enabled = False

With TDB_Hapus
    
    If .Visible = False Then
        
        Txt_Cr_Hapus(0).Text = ""
        Txt_Cr_Hapus(1).Text = ""
'        Txt_Cr_Hapus(2).Text = ""
'        Txt_Cr_Hapus(3).Text = ""
        
        Txt_Cr_Hapus_KeyUp 0, 0, 0
        
        .Visible = True
        
        Txt_Cr_Hapus(0).SetFocus
        
    Else
        .Visible = False
    End If
    
End With

End Sub

Private Sub Cmd_Info_Click()

Frame_Nav.Enabled = False
Rubah = False
'Cetak = False

Cmd_Tambah.Enabled = False

Cmd_Rubah.Visible = False
Cmd_Batal.Visible = True
Cmd_Batal.Enabled = True

Cmd_Hapus.Enabled = False

Cmd_Info.Enabled = False

Cmd_Keluar.Enabled = False
'cmd_cetak.Enabled = False

With TDB_Info
    
    If .Visible = False Then
        
        Txt_Cr_Info(0).Text = ""
        Txt_Cr_Info(1).Text = ""
'        Txt_Cr_Info(2).Text = ""
'        Txt_Cr_Info(3).Text = ""
        
        Txt_Cr_Info_KeyUp 0, 0, 0
        
        .Visible = True
        
        Txt_Cr_Info(0).SetFocus
        
    Else
        .Visible = False
    End If
    
End With

End Sub

Private Sub cmd_keluar_Click()
    Unload Me
    utama.Enabled = True
End Sub

Private Sub Cmd_Nav_Click(Index As Integer)

On Error Resume Next

With Rs_Nav
    Select Case Index
    
        Case 0
        
            .MoveFirst
            
        Case 1
            
            If .BOF Then .MoveFirst
                
                .MovePrevious
            
            If .BOF Then .MoveFirst
            
       Case 2
            
            If .EOF Then .MoveLast
                
                .MoveNext
                
            If .EOF Then .MoveLast
            
       Case 3
            
            .MoveLast
            
    End Select
End With

Isi_Semua Rs_Nav

End Sub

Private Sub cmd_ok_Click()
    
    Dim a As Long
    
    If Arr_Member.UpperBound(1) = 1 And Arr_Member(1, 1) = Empty Then
        a = 1
    Else
        a = Arr_Member.UpperBound(1) + 1
    End If
    
    Arr_Member.ReDim 1, a, 0, Grid_Member.Columns.Count
        Grid_Member.ReBind
        Grid_Member.Refresh
    
    If Arr_Member.UpperBound(1) = 1 And Arr_Member(1, 1) = Empty Then
        Arr_Member(a, 0) = 1
    Else
        Arr_Member(a, 0) = Arr_Member.UpperBound(1)
    End If
    
    Arr_Member(a, 1) = Dtp_Tgl1.Value
    Arr_Member(a, 2) = "S/D"
    Arr_Member(a, 3) = Dtp_Tgl2.Value
    Arr_Member(a, 4) = Trim(txt_disc.Text)
    If Arr_Member.UpperBound(1) > 1 Then
        Arr_Member(a - 1, 5) = vbUnchecked
    Else
        Arr_Member(a, 5) = vbUnchecked
    End If
    
    Arr_Member(a, 5) = vbChecked
    
    Grid_Member.ReBind
    Grid_Member.Refresh
    
    Tambah_Akt = True
    TDB_Aktivasi.Visible = False
    Cmd_Add.Enabled = False
    Cmd_Simpan.SetFocus
        
End Sub

Private Sub Cmd_Rubah_Click()

 Frame_Nav.Enabled = False

Cmd_Tambah.Visible = False
Cmd_Simpan.Visible = True
Cmd_Simpan.Enabled = False

Cmd_Rubah.Visible = False
Cmd_Batal.Visible = True

Cmd_Hapus.Enabled = False
Cmd_Add.Visible = True
Cmd_Add.Enabled = False

Cmd_Info.Enabled = False
Cmd_Dell.Visible = True
Cmd_Dell.Enabled = False

Cmd_Keluar.Enabled = False
'cmd_cetak.Enabled = False

With TDB_Rubah
    
    If .Visible = False Then
        
        Txt_Cr_Rubah(0).Text = ""
        Txt_Cr_Rubah(1).Text = ""
'        Txt_Cr_Rubah(2).Text = ""
'        Txt_Cr_Rubah(3).Text = ""
        
        Txt_Cr_Rubah_KeyUp 0, 0, 0
        
        .Visible = True
        
        Txt_Cr_Rubah(0).SetFocus
        
    Else
        .Visible = False
    End If
    
End With

End Sub

Private Sub cmd_simpan_Click()
On Error GoTo err_handler
   
    Dim sql As String
    Dim rs As Recordset
    
    Dim sql1 As String
    Dim rs1 As Recordset
    
    Dim Konfirm As Integer
        If Txt_Nom.Text = "" Then
            
            Konfirm = CInt(MsgBox("No. Member tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
            
            Txt_Nom.SetFocus
            On Error GoTo 0
            Exit Sub
        End If
        
        Dim tanggal As String
            If Txt_TglLhr.Text = "__/__/____" Then
                tanggal = "11/11/1111"
            Else
                tanggal = Trim(Txt_TglLhr.Text)
            End If
            
            Dim aktif As Integer
                If Check_Aktif = vbChecked Then
                    aktif = 1
                Else
                    aktif = 0
                End If
            
        If Arr_Member.UpperBound(1) = 1 And Arr_Member(1, 1) = Empty Then
        
            Konfirm = CInt(MsgBox("Tidak ada data aktivasi yang akan disimpan", vbOKOnly + vbInformation, "Informasi"))
            
            Txt_Nom.SetFocus
            On Error GoTo 0
            Exit Sub
        End If
    
    cn.BeginTrans
    
    If Rubah = False Then
    
        sql1 = "select No_Member from Tb_Member where No_Member='" & Trim(Txt_Nom.Text) & "'"
            
            Set rs1 = New ADODB.Recordset
                rs1.Open sql1, cn
                
                With rs1
                    
                    If Not .EOF Then
                        Konfirm = CInt(MsgBox("No. Member yang anda masukkan tidak ditemukan", vbOKOnly + vbInformation, "Informasi"))
                        
                        Txt_Nom.SetFocus
                        On Error GoTo 0
                        Exit Sub
                    Else
                        
                        sql = "insert into Tb_Member (No_Member,Nama,No_KtpSim,Alamat,No_Telp1,No_Telp2,Tgl_Lhr,Aktif)"
                        sql = sql & " values('" & Trim(Txt_Nom.Text) & "','" & Trim(Txt_Nama.Text) & "','" & Trim(Txt_Ktp.Text) & "','" & Trim(Txt_Alamat.Text) & "','" & Trim(Txt_Telp1.Text) & "','" & Trim(Txt_Telp2.Text) & "','" & tanggal & "'," & aktif & ")"
                        
                        Set rs = New ADODB.Recordset
                            rs.Open sql, cn
                        
                     If Tambah_Akt = True Then
                        
'                        Dim MyAkt As Integer
'                            If Arr_Member(Arr_Member.UpperBound(1), 5) = vbChecked Then
'                                MyAkt = 1
'                            Else
'                                MyAkt = 0
'                            End If
                        
                        sql = "insert into Tb_Member_Detail (Kode_Member,Tanggal1,Tanggal2,Disc,Aktif)"
                        sql = sql & " values ('" & Trim(Txt_Nom.Text) & "','" & Arr_Member(Arr_Member.UpperBound(1), 1) & "','" & Arr_Member(Arr_Member.UpperBound(1), 3) & "'," & Arr_Member(Arr_Member.UpperBound(1), 4) & ",1)"
                        
                        Set rs = New ADODB.Recordset
                            rs.Open sql, cn
                        
                      End If
                      
                        cn.CommitTrans
                        
                        Konfirm = CInt(MsgBox("No. Member " & Trim(Txt_Nom.Text) & " telah disimpan", vbOKOnly + vbInformation, "Informasi"))
                        
                    End If
                    
                End With
    Else
    
        sql = "update Tb_Member set Nama='" & Trim(Txt_Nama.Text) & "',No_KtpSim='" & Trim(Txt_Ktp.Text) & "',Alamat='" & Trim(Txt_Alamat.Text) & "',No_Telp1='" & Trim(Txt_Telp1.Text) & "',No_Telp2='" & Trim(Txt_Telp2.Text) & "',Tgl_Lhr='" & Trim(tanggal) & "',Aktif=" & aktif & " where No_Member='" & Trim(Txt_Nom.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, cn
        
        If Tambah_Akt = True Then
                        
'                        Dim MyAkt As Integer
'                            If Arr_Member(Arr_Member.UpperBound(1), 5) = vbChecked Then
'                                MyAkt = 1
'                            Else
'                                MyAkt = 0
'                            End If
                        
                        sql1 = "select Max(Id) as mx,Aktif from Tb_Member_Detail where Kode_Member='" & Trim(Txt_Nom.Text) & "' group by Id,Aktif"
                        Set rs1 = New ADODB.Recordset
                            rs1.Open sql1, cn
                            
                        With rs1
                            If Not .EOF Then
                                
                                sql = "update Tb_Member_Detail set Aktif=0 where Id=" & !mx
                                
                                Set rs = New ADODB.Recordset
                                    rs.Open sql, cn
                                
                            End If
                        End With
                        
                        sql = "insert into Tb_Member_Detail (Kode_Member,Tanggal1,Tanggal2,Disc,Aktif)"
                        sql = sql & " values ('" & Trim(Txt_Nom.Text) & "','" & Arr_Member(Arr_Member.UpperBound(1), 1) & "','" & Arr_Member(Arr_Member.UpperBound(1), 3) & "'," & Arr_Member(Arr_Member.UpperBound(1), 4) & ",1)"
                        
                        Set rs = New ADODB.Recordset
                            rs.Open sql, cn
                        
                      End If
        
            cn.CommitTrans
            
            Konfirm = CInt(MsgBox("No. Member " & Trim(Txt_Nom.Text) & " telah dirubah", vbOKOnly + vbInformation, "Informasi"))
    End If
    
    cmd_batal_Click
    On Error GoTo 0
    Exit Sub
    
err_handler:
    
    cn.RollbackTrans
    
    Konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
        Err.Clear
    
End Sub

Private Sub cmd_tambah_Click()

    Rubah = False
    Frame_Nav.Enabled = False
                
    Cmd_Tambah.Visible = False
    Cmd_Simpan.Visible = True
    Cmd_Simpan.Enabled = False
    
     Cmd_Rubah.Visible = False
     Cmd_Batal.Visible = True
     
     Cmd_Hapus.Enabled = False
     Cmd_Info.Enabled = False
     Cmd_Keluar.Enabled = False
        
    Txt_Nom.Text = ""
    Txt_Nom.Enabled = True
    Txt_Nom.SetFocus
    

End Sub

Private Sub Dtp_Tgl1_GotFocus()
    Call Focus_(Dtp_Tgl1)
End Sub

Private Sub Dtp_Tgl1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Dtp_Tgl2.SetFocus
End Sub

Private Sub Dtp_Tgl2_GotFocus()
    Call Focus_(Dtp_Tgl2)
End Sub

Private Sub Dtp_Tgl2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Cmd_Ok.SetFocus
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        Cmd_Tambah.SetFocus
End Sub

Private Sub Form_Load()

Rubah = False
With Me
    .Left = Screen.Width / 2 - .Width / 2
    .Top = Screen.Height / 2 - .Height / 2
End With

With TDB_Aktivasi
    .Left = 2760
    .Top = 4800
End With

     Call cari_wewenang("Form Data Member")
        
        If tambah_form = True Then
            Cmd_Tambah.Enabled = True
            Cmd_Add.Enabled = False
        Else
            Cmd_Add.Enabled = False
            Cmd_Tambah.Enabled = False
        End If
        
        If edit_form = True Then
            Cmd_Rubah.Enabled = True
        Else
            Cmd_Rubah.Enabled = False
        End If
        
        If hapus_form = True Then
            Cmd_Dell.Enabled = True
            Cmd_Hapus.Enabled = True
        Else
            Cmd_Dell.Enabled = False
            Cmd_Hapus.Enabled = False
        End If
        
        Txt_Nom.Enabled = False
        Txt_Nama.Enabled = False
        Txt_Ktp.Enabled = False
        Txt_Alamat.Enabled = False
        Txt_Telp1.Enabled = False
        Txt_Telp2.Enabled = False
        Txt_TglLhr.Enabled = False
         
        Cmd_Add.Enabled = False
        Cmd_Dell.Enabled = False
        
        Txt_Telp1.TabIndex = Txt_Alamat.TabIndex + 1
        
        Grid_Member.Array = Arr_Member
        
        Arr_Member.ReDim 0, 0, 0, 0
        Arr_Member.ReDim 1, 1, 1, 1
            Grid_Member.ReBind
            Grid_Member.Refresh
                        
        Grid_Rubah.Array = Arr_Rubah
        Grid_Hapus.Array = Arr_Hapus
        Grid_Info.Array = Arr_Info
        
        With TDB_Rubah
            .Left = Me.Width / 2 - .Width / 2
            .Top = Me.Height / 2 - .Height / 2
        End With
                                
        With TDB_Hapus
            .Left = Me.Width / 2 - .Width / 2
            .Top = Me.Height / 2 - .Height / 2
        End With
                               
        With TDB_Info
            .Left = Me.Width / 2 - .Width / 2
            .Top = Me.Height / 2 - .Height / 2
        End With
                               
        Txt_Cr_Info_KeyUp 0, 0, 0
            Cmd_Nav_Click 3
                               
End Sub


Private Sub Form_Unload(Cancel As Integer)
    utama.Enabled = True
End Sub

Private Sub Grid_Hapus_DblClick()

On Error GoTo err_handler
    
    If Arr_Hapus.UpperBound(1) = 1 And Arr_Hapus(1, 1) = Empty Then Exit Sub
    
    If MsgBox("Yakin akan hapus " & Arr_Hapus(Grid_Hapus.Bookmark, 0), vbYesNo + vbQuestion, "Hapus") = vbNo Then
    
    Grid_Hapus.SetFocus
    On Error GoTo 0
    Exit Sub
    End If
    
    Dim Konfirm As Integer
    Dim sql As String
    Dim rs As Recordset
        
    Dim sql1 As String
    Dim rs1 As Recordset
    
    Dim jumlah As Double
    
        cn.BeginTrans
                       
        sql = "delete from Tb_Member_Detail where Kode_Member='" & Arr_Hapus(Grid_Hapus.Bookmark, 0) & "'"
            Set rs = New ADODB.Recordset
                rs.Open sql, cn
        
        sql = "delete from Tb_Member where No_Member='" & Arr_Hapus(Grid_Hapus.Bookmark, 0) & "'"
            Set rs = New ADODB.Recordset
                rs.Open sql, cn
        
        
        Konfirm = CInt(MsgBox("No Faktur " & Arr_Hapus(Grid_Hapus.Bookmark, 0) & " telah dihapus", vbOKOnly + vbInformation, "Informasi"))
        
        cn.CommitTrans
        
        cmd_batal_Click
        On Error GoTo 0
        Exit Sub

err_handler:
        
        cn.RollbackTrans
            
        If Err.Number = -2147467259 Then
            Konfirm = CInt(MsgBox("Data tidak bisa dihapus karna sedang dipakai oleh transaksi lain", vbOKOnly + vbInformation, "Informasi"))
        Else
            Konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
        End If

                Err.Clear

End Sub

Private Sub Grid_Hapus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Hapus_DblClick
    If KeyCode = vbKeyEscape Then cmd_batal_Click
End Sub

Private Sub Grid_Info_DblClick()
    
    If Arr_Info.UpperBound(1) = 1 And Arr_Info(1, 1) = Empty Then Exit Sub
    
    With Rs_Nav
        
        .MoveFirst
        .Find "No_Member='" & Arr_Info(Grid_Info.Bookmark, 0) & "'"
        
    End With
    
    Isi_Semua Rs_Nav
    
    TDB_Info.Visible = False
    Frame_Nav.Enabled = True
    Cmd_Nav(0).SetFocus
    
End Sub

Private Sub Grid_Info_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Info_DblClick
    If KeyCode = vbKeyEscape Then cmd_batal_Click
End Sub

Private Sub Grid_Rubah_DblClick()
    
    If Arr_Rubah.UpperBound(1) = 1 And Arr_Rubah(1, 1) = Empty Then Exit Sub
    
    With Rs_Nav
        
        .MoveFirst
        .Find "No_Member='" & Arr_Rubah(Grid_Rubah.Bookmark, 0) & "'"
        
    End With
    
    Isi_Semua Rs_Nav
    
    Txt_Nama.Enabled = True
    Txt_Ktp.Enabled = True
    Txt_Alamat.Enabled = True
    Txt_Telp1.Enabled = True
    Txt_Telp2.Enabled = True
    Txt_TglLhr.Enabled = True
    Check_Aktif.Enabled = True
    
    If tambah_form = True Then
        Cmd_Add.Enabled = True
    End If
    If hapus_form = True Then
        Cmd_Dell.Enabled = True
    End If
    
    Cmd_Simpan.Enabled = True
    
    TDB_Rubah.Visible = False
    Rubah = True
    Txt_Nama.SetFocus
    
End Sub

Private Sub Grid_Rubah_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Rubah_DblClick
    If KeyCode = vbKeyEscape Then cmd_batal_Click
End Sub

Private Sub TDB_Aktivasi_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDB_Aktivasi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDB_Aktivasi.Top = TDB_Aktivasi.Top - (yold - Y)
   TDB_Aktivasi.Left = TDB_Aktivasi.Left - (xold - X)
End If

End Sub

Private Sub TDB_Aktivasi_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub TDB_Hapus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDB_Hapus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDB_Hapus.Top = TDB_Hapus.Top - (yold - Y)
   TDB_Hapus.Left = TDB_Hapus.Left - (xold - X)
End If

End Sub

Private Sub TDB_Hapus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub TDB_Info_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDB_Info_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDB_Info.Top = TDB_Info.Top - (yold - Y)
   TDB_Info.Left = TDB_Info.Left - (xold - X)
End If

End Sub

Private Sub TDB_Info_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub TDB_Rubah_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub TDB_Rubah_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   TDB_Rubah.Top = TDB_Rubah.Top - (yold - Y)
   TDB_Rubah.Left = TDB_Rubah.Left - (xold - X)
End If

End Sub

Private Sub TDB_Rubah_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub txt_alamat_GotFocus()
    Call Focus_(Txt_Alamat)
End Sub

Private Sub Txt_Cr_Hapus_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim sql As String

sql = "select top 100 * from Tb_Member"

Select Case Index
    Case 0
        sql = sql & " where No_Member like '%" & Trim(Txt_Cr_Hapus(0).Text) & "%'"
    Case 1
        sql = sql & " where Nama like '%" & Trim(Txt_Cr_Hapus(1).Text) & "%'"
End Select

sql = sql & " order by No_Member asc"

Set Rs_Nav = New ADODB.Recordset
    Rs_Nav.Open sql, cn, adOpenKeyset
    
    isi_grid Grid_Hapus, Arr_Hapus, Rs_Nav

End Sub

Private Sub Txt_Cr_Hapus_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Hapus.SetFocus
    If KeyCode = vbKeyEscape Then cmd_batal_Click
End Sub


Private Sub Txt_Cr_Info_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Info.SetFocus
    If KeyCode = vbKeyEscape Then cmd_batal_Click
End Sub


Private Sub Txt_Cr_Info_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim sql As String

sql = "select top 100 * from Tb_Member"

Select Case Index
    Case 0
        sql = sql & " where No_Member like '%" & Trim(Txt_Cr_Info(0).Text) & "%'"
    Case 1
        sql = sql & " where Nama like '%" & Trim(Txt_Cr_Info(1).Text) & "%'"
End Select

sql = sql & " order by No_Member asc"

Set Rs_Nav = New ADODB.Recordset
    Rs_Nav.Open sql, cn, adOpenKeyset
    
    isi_grid Grid_Info, Arr_Info, Rs_Nav

End Sub

Private Sub Txt_Cr_Rubah_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim sql As String

sql = "select top 100 * from Tb_Member"

Select Case Index
    Case 0
        sql = sql & " where No_Member like '%" & Trim(Txt_Cr_Rubah(0).Text) & "%'"
    Case 1
        sql = sql & " where Nama like '%" & Trim(Txt_Cr_Rubah(1).Text) & "%'"
End Select

sql = sql & " order by No_Member asc"

Set Rs_Nav = New ADODB.Recordset
    Rs_Nav.Open sql, cn, adOpenKeyset
    
    isi_grid Grid_Rubah, Arr_Rubah, Rs_Nav

End Sub

Private Sub Txt_Cr_Rubah_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Grid_Rubah.SetFocus
    If KeyCode = vbKeyEscape Then cmd_batal_Click
End Sub

Private Sub txt_disc_GotFocus()
    Call Focus_(txt_disc)
End Sub

Private Sub txt_disc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Dtp_Tgl1.SetFocus
End Sub

Private Sub txt_disc_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",")) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt_disc_LostFocus()
    With txt_disc
        If .Text = "" Then .Text = 0
    End With
End Sub

Private Sub Txt_Ktp_GotFocus()
    Call Focus_(Txt_Ktp)
End Sub

Private Sub Txt_Ktp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Alamat.SetFocus
End Sub

Private Sub txt_nama_GotFocus()
    Call Focus_(Txt_Nama)
End Sub

Private Sub txt_nama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Ktp.SetFocus
End Sub

Private Sub Txt_Nom_GotFocus()
    Call Focus_(Txt_Nom)
End Sub

Private Sub Txt_Nom_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
    
    Dim Konfirm As Integer
    If Txt_Nom.Text = "" Then
        
        Konfirm = CInt(MsgBox("No. Member tidak boleh kosong", vbOKOnly + vbInformation, "Informasi"))
        
        On Error GoTo 0
        Exit Sub
    End If
    
    Dim sql As String
    Dim rs As Recordset
        sql = "select No_Member from Tb_Member where No_Member='" & Trim(Txt_Nom.Text) & "'"
        
        Set rs = New ADODB.Recordset
            rs.Open sql, cn
            
            With rs
                
                If Not .EOF Then
                    Konfirm = CInt(MsgBox("No. Bukti yang anda masukkan sudah ada", vbOKOnly + vbInformation, "Informasi"))
                    
                    Txt_Nama.Enabled = False
                    Txt_Alamat.Enabled = False
                    Txt_Ktp.Enabled = False
                    Txt_Telp1.Enabled = False
                    Txt_Telp2.Enabled = False
                    Txt_TglLhr.Enabled = False
                    
                    Cmd_Add.Enabled = False
                    
                    Cmd_Dell.Enabled = False
                    
                    Cmd_Simpan.Enabled = False
                    
                Else
                    
                    Txt_Nama.Text = ""
                    Txt_Alamat.Text = ""
                    Txt_Ktp.Text = ""
                    Txt_Telp1.Text = ""
                    Txt_Telp2.Text = ""
                    Txt_TglLhr.Text = "__/__/____"
                    Check_Aktif.Value = vbChecked
                    
                    Arr_Member.ReDim 0, 0, 0, 0
                    Arr_Member.ReDim 1, 1, 1, 1
                        Grid_Member.ReBind
                        Grid_Member.Refresh
                    
                    Txt_Nama.Enabled = True
                    Txt_Alamat.Enabled = True
                    Txt_Ktp.Enabled = True
                    Txt_Telp1.Enabled = True
                    Txt_Telp2.Enabled = True
                    Txt_TglLhr.Enabled = True
                    
                    Cmd_Add.Enabled = True
                    
                    If hapus_form = True Then
                        Cmd_Dell.Enabled = True
                    End If
                    
                    Cmd_Simpan.Enabled = True
                    
                    Txt_Nama.SetFocus
                    
                End If
                
            End With
    
    
    End If
    
End Sub

Private Sub Txt_Telp1_GotFocus()
    Call Focus_(Txt_Telp1)
End Sub

Private Sub Txt_Telp1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_Telp2.SetFocus
End Sub

Private Sub Txt_Telp2_GotFocus()
    Call Focus_(Txt_Telp2)
End Sub

Private Sub Txt_Telp2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_TglLhr.SetFocus
End Sub

Private Sub Txt_TglLhr_GotFocus()
    Call Focus_(Txt_TglLhr)
End Sub

Private Sub Txt_TglLhr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Check_Aktif.SetFocus
End Sub
