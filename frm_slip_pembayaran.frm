VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_slip_pembayaran 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic_air 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   -5160
      ScaleHeight     =   6465
      ScaleWidth      =   5505
      TabIndex        =   32
      Top             =   3360
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton pic_x_air 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   34
         Top             =   0
         Width           =   495
      End
      Begin TrueOleDBGrid60.TDBGrid grd_air 
         Height          =   5895
         Left            =   120
         OleObjectBlob   =   "frm_slip_pembayaran.frx":0000
         TabIndex        =   33
         Top             =   480
         Width           =   5295
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   5505
         TabIndex        =   35
         Top             =   0
         Width           =   5535
      End
   End
   Begin VB.PictureBox Pic_listrik 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   -5400
      ScaleHeight     =   6465
      ScaleWidth      =   5505
      TabIndex        =   28
      Top             =   3960
      Visible         =   0   'False
      Width           =   5535
      Begin TrueOleDBGrid60.TDBGrid grd_listrik 
         Height          =   5895
         Left            =   120
         OleObjectBlob   =   "frm_slip_pembayaran.frx":2810
         TabIndex        =   31
         Top             =   480
         Width           =   5295
      End
      Begin VB.CommandButton cmd_x 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   30
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
         ScaleWidth      =   5505
         TabIndex        =   29
         Top             =   0
         Width           =   5535
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   120
      ScaleHeight     =   6615
      ScaleWidth      =   8535
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin VB.CheckBox cek_lain 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Perhitungkan Disc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   40
         Top             =   720
         Width           =   4095
      End
      Begin VB.CommandButton cmd_simpan 
         Caption         =   "Simpan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CommandButton cmd_baru 
         Caption         =   "Baru"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   6000
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtp_tgl 
         Height          =   375
         Left            =   5880
         TabIndex        =   37
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   -2147483635
         CalendarTitleForeColor=   16777215
         Format          =   60948480
         CurrentDate     =   38664
      End
      Begin VB.TextBox txt_total 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5880
         TabIndex        =   8
         Top             =   4680
         Width           =   2415
      End
      Begin VB.CommandButton cmd_preview 
         Caption         =   "Preview"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6000
         Width           =   1215
      End
      Begin VB.TextBox txt_lain 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         TabIndex        =   7
         Top             =   5640
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txt_air 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         TabIndex        =   6
         Top             =   5160
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txt_listrik 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         TabIndex        =   5
         Top             =   4680
         Visible         =   0   'False
         Width           =   2415
      End
      Begin Crystal.CrystalReport report 
         Left            =   7560
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.CommandButton cmd_tampil 
         Caption         =   "Tampil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txt_kode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   1680
         Width           =   1575
      End
      Begin MSMask.MaskEdBox msk_tgl1 
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_tgl2 
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5400
         TabIndex        =   36
         Top             =   720
         Width           =   330
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4800
         TabIndex        =   27
         Top             =   4680
         Width           =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pot Lain"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   600
         TabIndex        =   26
         Top             =   5640
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pot. Air"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   600
         TabIndex        =   25
         Top             =   5160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pot. Listrik"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   600
         TabIndex        =   24
         Top             =   4680
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   480
         X2              =   8280
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label lbl_jml 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   23
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Label lbl_jumlah 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4800
         TabIndex        =   22
         Top             =   3960
         Width           =   660
      End
      Begin VB.Label lbl_ppn 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   21
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Label lbl_nilai 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   20
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ppn"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   600
         TabIndex        =   19
         Top             =   3960
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nilai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   600
         TabIndex        =   18
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label lbl_persentase 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   17
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Persentase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   600
         TabIndex        =   16
         Top             =   3000
         Width           =   1110
      End
      Begin VB.Label lbl_tot_jual 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   15
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tot Jual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   600
         TabIndex        =   14
         Top             =   2520
         Width           =   750
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   480
         X2              =   8280
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kd. Counter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   480
         TabIndex        =   13
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3840
         TabIndex        =   12
         Top             =   1320
         Width           =   345
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   480
         X2              =   8280
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laporan Slip Pembayaran"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   2925
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   480
         TabIndex        =   10
         Top             =   1200
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm_slip_pembayaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nama_pemilik, perse As String
Dim jml_jual As Double, jml_nilai As Double, jml_ppn As Double, jml_tot As Double
Dim kd_counter, nm_counter As String
Dim biaya_listrik As String
Dim biaya_air As String
Dim arr_air As New XArrayDB
Dim arr_listrik As New XArrayDB

    
Private Sub isi()

On Error GoTo er_isi

    Dim sql, sql1, sql2 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset
    Dim nilai As Double, ppn As Double, tot As Double, tot_jual As Double, persen
    Dim a As Long
    
        
        
        If msk_tgl1.Text = "__/__/____" And msk_tgl2.Text = "__/__/____" Then
            MsgBox ("Periode laporan hrs diisi semua")
            msk_tgl1.SetFocus
            Exit Sub
        End If
        
        If txt_kode.Text = "" Then
            MsgBox ("Kode counter hrs diisi")
            txt_kode.SetFocus
            Exit Sub
        End If
        
        sql = "select distinct(kode_counter)as kd_c from qr_penjualan_sebenarnya where"
        sql = sql & " tgl >= datevalue('" & Trim(msk_tgl1.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl2.Text) & "')"
        sql = sql & " and kode_counter='" & Trim(txt_kode.Text) & "'"
        sql = sql & " order by kode_counter"
        
        rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                
                a = 1
                jml_jual = 0
                jml_nilai = 0
                jml_ppn = 0
                jml_tot = 0
                
                Do While Not rs.EOF
                    
                    sql1 = "select nama_counter,nama_pemilik,presentasi_p from tbl_counter where kode='" & Trim(rs!kd_c) & "'"
                    rs1.Open sql1, cn
                        If Not rs1.EOF Then
                        
                         If cek_lain.Value = vbUnchecked Then
                            sql2 = "select sum(harga_sebenarnya) as benar from qr_penjualan_sebenarnya where"
                            sql2 = sql2 & " kode_counter='" & Trim(rs!kd_c) & "' and tgl >= datevalue('" & Trim(msk_tgl1.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl2.Text) & "')"
                         End If
                                
                         If cek_lain.Value = vbChecked Then
                           sql2 = "select sum(total_harga) as benar from qr_penjualan_sebenarnya where"
                           sql2 = sql2 & " kode_counter='" & Trim(rs!kd_c) & "' and tgl >= datevalue('" & Trim(msk_tgl1.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl2.Text) & "')"
                         End If
                                
                                rs2.Open sql2, cn
                                    If Not rs2.EOF Then
                                        
                                        
                                        
                                        
                                        kd_counter = rs!kd_c
                                        nm_counter = rs1!nama_counter
                                        nama_pemilik = rs1!nama_pemilik
                                        perse = rs1!presentasi_p
                                        
                                        tot_jual = rs2!benar
                                        persen = Mid(rs1!presentasi_p, 1, Len(rs1!presentasi_p) - 1)
                                        nilai = (CDbl(tot_jual) * CDbl(persen)) / 100
                                        ppn = (CDbl(tot_jual) - CDbl(nilai)) * (10 / 100)
                                        tot = CDbl(tot_jual) - (CDbl(nilai) + CDbl(ppn))
                                        
                                            
                                            jml_jual = CDbl(jml_jual) + CDbl(tot_jual)
                                            jml_nilai = CDbl(jml_nilai) + CDbl(nilai)
                                            jml_ppn = CDbl(jml_ppn) + CDbl(ppn)
                                            jml_tot = CDbl(jml_tot) + CDbl(tot)
                                            
                                      
                                            
                                    a = a + 1
                                    End If
                                rs2.Close
                        Else
                            MsgBox ("Nama Counter Tidak ditemukan")
                            Exit Sub
                        End If
                      rs1.Close
                 rs.MoveNext
                 Loop
            End If
         rs.Close
                    
               
                  lbl_tot_jual.Caption = Format(jml_jual, "###,###,###")
                  lbl_persentase.Caption = Trim(persen) & "%"
                  
                
                    
                  lbl_nilai.Caption = Format(jml_nilai, "###,###,###")
                
                
                    
                  lbl_ppn.Caption = Format(jml_ppn, "###,###,###")
                  
                
        
                  lbl_jml.Caption = Format(jml_tot, "###,###,###")
                  
                  txt_total.Text = Format(jml_tot, "###,###,###")
                  
                  
Exit Sub

er_isi:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub cmd_baru_Click()
    msk_tgl1.Text = "__/__/____"
    msk_tgl2.Text = "__/__/____"
    txt_kode.Text = ""
    lbl_tot_jual.Caption = ""
    lbl_persentase.Caption = ""
    lbl_jml.Caption = ""
    lbl_nilai.Caption = ""
    lbl_ppn.Caption = ""
    txt_listrik.Text = 0
    txt_air.Text = 0
    txt_lain.Text = 0
    txt_total.Text = 0
    msk_tgl1.SetFocus
End Sub

Private Sub cmd_preview_Click()

On Error GoTo preview

Dim sql As String
Dim rs As New ADODB.Recordset
    
    If msk_tgl1.Text = "__/__/____" And msk_tgl2.Text = "__/__/____" And txt_kode.Text = "" Then
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    utama.MousePointer = vbHourglass

If cek_lain.Value = vbUnchecked Then

sql = "SELECT"
    sql = sql & " kode_counter,"
    sql = sql & "tgl,"
    sql = sql & "harga_sebenarnya"
    sql = sql & " From "
    sql = sql & "qr_penjualan_sebenarnya"
    sql = sql & " where tgl >= datevalue('" & Trim(msk_tgl1.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl2.Text) & "') and kode_counter='" & Trim(txt_kode.Text) & "'"
    sql = sql & " order by tgl"
    rs.Open sql, cn
    
End If

If cek_lain.Value = vbChecked Then

sql = "SELECT"
    sql = sql & " kode_counter,"
    sql = sql & "nama_counter,"
    sql = sql & "tgl,"
    sql = sql & "harga_sebenarnya,"
    sql = sql & "total_harga"
    sql = sql & " From"
    sql = sql & " qr_penjualan_sebenarnya"
    sql = sql & " where tgl >= datevalue('" & Trim(msk_tgl1.Text) & "') and tgl <= datevalue('" & Trim(msk_tgl2.Text) & "') and kode_counter='" & Trim(txt_kode.Text) & "'"
    sql = sql & " order by tgl"
    rs.Open sql, cn
    
End If
    
If cek_lain.Value = vbUnchecked Then
    Report.ReportFileName = App.path & "\Laporan\lap_total_pertangga1.rpt"
End If

If cek_lain.Value = vbChecked Then
    Report.ReportFileName = App.path & "\Laporan\lap_total_pertangga2.rpt"
End If

    Report.Connect = cn
    Report.RetrieveSQLQuery
    Report.SQLQuery = sql
    
        Report.Formulas(0) = "tgl_awal='" & Trim(msk_tgl1.Text) & "'"
        Report.Formulas(1) = "tgl_akhir='" & Trim(msk_tgl2.Text) & "'"
        
        Report.Formulas(2) = "pemakai='" & id_user & "'"
        Report.Formulas(3) = "nama_pemilik='" & Trim(nama_pemilik) & "'"
        Report.Formulas(4) = "nama_counter='" & Trim(nm_counter) & "'"
         
        Report.Formulas(5) = "penjualan='" & Trim(lbl_tot_jual.Caption) & "'"
        
        Dim saring As String
        Dim lang
            lang = Len(lbl_persentase.Caption)
            saring = Mid(lbl_persentase.Caption, 1, CDbl(lang) - 1)
            
        Report.Formulas(6) = "persen='" & saring & "'"

        Report.Formulas(7) = "nilai='" & Trim(lbl_nilai.Caption) & "'"
        
        Report.Formulas(8) = "ppn='" & Trim(lbl_ppn.Caption) & "'"
        
        Report.Formulas(9) = "jumlah='" & Trim(lbl_jml.Caption) & "'"
        
        Report.Formulas(10) = "pot_listrik='" & Trim(txt_listrik.Text) & "'"
        
        Report.Formulas(11) = "pot_air='" & Trim(txt_air.Text) & "'"
        
        Report.Formulas(12) = "pot_lain='" & Trim(txt_lain.Text) & "'"
        
        Report.Formulas(13) = "total='" & Trim(txt_total.Text) & "'"
        
    Report.DiscardSavedData = True
    Report.WindowState = crptMaximized
    Report.Action = 1
    
    If Me.MousePointer = vbHourglass Then
        Me.MousePointer = vbDefault
        utama.MousePointer = vbDefault
    End If
    
    Exit Sub
    
preview:
    
    If Me.MousePointer = vbHourglass Then
        Me.MousePointer = vbDefault
        utama.MousePointer = vbDefault
    End If
    
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
    
End Sub

Private Sub Cmd_Simpan_Click()

On Error GoTo err_simpan

Dim sql As String
Dim rs As New ADODB.Recordset
    
    If msk_tgl1.Text = "__/__/____" Or msk_tgl2.Text = "__/__/____" Or txt_kode.Text = "" Then
        Dim cek_dulu
        cek_dulu = MsgBox("Periode dan Kode Counter hrs diisi", , App.Title)
        Exit Sub
    End If
    
    Dim tgl As Date
    Dim persen
        tgl = dtp_tgl.Value
        persen = Trim(lbl_persentase.Caption)
        persen = Len(persen)
        persen = Mid(lbl_persentase.Caption, 1, CDbl(persen) - 1)
        
    Dim seleksi As Long
        If cek_lain.Value = vbChecked Then
            seleksi = 1
        Else
            seleksi = 0
        End If
        
    sql = "insert into tr_slip_pembayaran (tgl_bayar,periode1,periode2,kode_counter,tot_jual,persentase,ppn,nilai,jumlah,pot_listrik,pot_air,pot_lain,total,ket,nama_user)"
    sql = sql & " values ('" & Trim(tgl) & "','" & Trim(msk_tgl1.Text) & "','" & Trim(msk_tgl2.Text) & "','" & Trim(txt_kode.Text) & "'," & CCur(Trim(lbl_tot_jual.Caption)) & ","
    sql = sql & "'" & persen & "'," & CCur(Trim(lbl_ppn.Caption)) & "," & CCur(Trim(lbl_nilai.Caption)) & "," & CCur(Trim(lbl_jml.Caption)) & "," & CCur(Trim(txt_listrik.Text)) & "," & CCur(Trim(txt_air.Text)) & "," & CCur(Trim(txt_lain.Text)) & "," & CCur(Trim(txt_total.Text)) & "," & seleksi & " ,'" & Trim(utama.lbl_user.Caption) & "')"
    rs.Open sql, cn
    
    MsgBox ("Data berhasil disimpan")
    
    On Error GoTo 0
    
    Exit Sub
    
err_simpan:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub Cmd_Tampil_Click()
    Call isi
    Call isi_listrik
    Call isi_air
End Sub

Private Sub cmd_x_Click()
    Pic_listrik.Visible = False
    txt_listrik.SetFocus
End Sub

Private Sub Form_Activate()
    msk_tgl1.SetFocus
End Sub

Private Sub Form_Load()
    
    grd_listrik.Array = arr_listrik
    
    grd_air.Array = arr_air
    
    Me.Left = utama.Width / 2 - Me.Width / 2
    Me.Top = utama.Height / 2 - Me.Height / 2 - 1250
    
    dtp_tgl.Value = Date
    
    txt_listrik.Text = 0
    txt_air.Text = 0
    txt_lain.Text = 0
    txt_total.Text = 0
    
End Sub

Private Sub kosong_listrik()
    arr_listrik.ReDim 0, 0, 0, 0
    grd_listrik.ReBind
    grd_listrik.Refresh
End Sub

Private Sub kosong_air()
    arr_air.ReDim 0, 0, 0, 0
    grd_air.ReBind
    grd_air.Refresh
End Sub

Private Sub isi_air()

On Error GoTo air

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim tgl, biaya As String
    Dim a As Long
        
        kosong_air
        
        sql = "select tgl,harga from qr_biling_air where kode='" & Trim(txt_kode.Text) & "' order by tgl desc"
        rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
                    
                a = 1
                Do While Not rs.EOF
                    arr_air.ReDim 1, a, 0, 2
                    grd_air.ReBind
                    grd_air.Refresh
                             
                        If Not IsNull(rs!tgl) Then
                            tgl = rs!tgl
                        Else
                            tgl = ""
                        End If
                        
                        If Not IsNull(rs!harga) Then
                            biaya = Format(rs!harga, "###,###,###")
                        Else
                            biaya = 0
                        End If
                        
                        arr_air(a, 0) = tgl
                        arr_air(a, 1) = biaya
                        
                a = a + 1
                rs.MoveNext
                Loop
                grd_air.ReBind
                grd_air.Refresh
                    
            End If
        rs.Close
        
        On Error GoTo 0
        
        Exit Sub
    
air:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub

Private Sub isi_listrik()

On Error GoTo er_listrik

    Dim sql As String
    Dim rs As New ADODB.Recordset
    
        
        kosong_listrik
        
        sql = "select tgl,harga from qr_biling_listrik where kode='" & Trim(txt_kode.Text) & "' order by tgl desc"
        rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
                    
                lanjut_listrik rs
            End If
        rs.Close
        
        On Error GoTo 0
        
        Exit Sub
        
er_listrik:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
 End Sub
 Sub lanjut_listrik(rs As Recordset)
 
 Dim tgl, biaya As String
 Dim a As Long
                        a = 1
                        Do While Not rs.EOF
                            arr_listrik.ReDim 1, a, 0, 2
                            grd_listrik.ReBind
                            grd_listrik.Refresh
                            
                                If Not IsNull(rs!tgl) Then
                                    tgl = rs!tgl
                                Else
                                    tgl = ""
                                End If
                                
                                If Not IsNull(rs!harga) Then
                                    biaya = Format(rs!harga, "###,###,###")
                                Else
                                    biaya = 0
                                End If
                                
                                arr_listrik(a, 0) = tgl
                                arr_listrik(a, 1) = biaya
                        a = a + 1
                        rs.MoveNext
                        Loop
                        grd_listrik.ReBind
                        grd_listrik.Refresh
End Sub

Private Sub grd_air_Click()
On Error Resume Next
    If arr_air.UpperBound(1) > 0 Then
        biaya_air = arr_air(grd_air.Bookmark, 1)
    End If
End Sub

Private Sub grd_air_DblClick()
On Error Resume Next
    If arr_air.UpperBound(1) > 0 Then
        txt_air.Text = Format(biaya_air, "###,###,###")
        pic_air.Visible = False
        txt_air.SetFocus
    End If
End Sub

Private Sub grd_air_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        grd_air_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        pic_air.Visible = False
        txt_air.SetFocus
    End If
    
End Sub

Private Sub grd_air_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_air_Click
End Sub

Private Sub grd_listrik_Click()
On Error Resume Next
    If arr_listrik.UpperBound(1) > 0 Then
        biaya_listrik = arr_listrik(grd_listrik.Bookmark, 1)
    End If
End Sub

Private Sub grd_listrik_DblClick()
On Error Resume Next
    If arr_listrik.UpperBound(1) > 0 Then
        txt_listrik.Text = Format(biaya_listrik, "###,###,###")
        Pic_listrik.Visible = False
        txt_listrik.SetFocus
    End If
End Sub

Private Sub grd_listrik_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        grd_listrik_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        Pic_listrik.Visible = False
        txt_listrik.SetFocus
    End If
    
End Sub

Private Sub grd_listrik_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_listrik_Click
End Sub

Private Sub msk_tgl1_GotFocus()
    msk_tgl1.SelStart = 0
    msk_tgl1.SelLength = Len(msk_tgl1)
End Sub
Private Sub msk_tgl2_GotFocus()
    msk_tgl2.SelStart = 0
    msk_tgl2.SelLength = Len(msk_tgl2)
End Sub

Private Sub pic_air_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_air.Visible = False
        txt_air.SetFocus
    End If
End Sub

Private Sub Pic_listrik_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Pic_listrik.Visible = False
        txt_listrik.SetFocus
    End If
End Sub

Private Sub pic_x_air_Click()
    pic_air.Visible = False
    txt_air.SetFocus
End Sub

Private Sub txt_air_GotFocus()
    txt_air.SelStart = 0
    txt_air.SelLength = Len(txt_air)
End Sub

Private Sub txt_air_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And txt_kode.Text <> "" Then
        pic_air.Visible = True
        grd_air.SetFocus
    End If
End Sub

Private Sub txt_air_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
    Beep
    KeyAscii = 0
End If
End Sub

Private Sub txt_air_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    'If txt_air.Text <> "" Then
        
        If txt_air.Text = "" Then
            txt_air.Text = 0
        End If
        
        If txt_air.Text <> 0 Then
            txt_air.Text = Format(txt_air.Text, "###,###,###")
            txt_air.SelStart = Len(txt_air)
        End If
        
        Dim jml_s As Double
        txt_total.Text = 0
            jml_s = CDbl(lbl_jml.Caption) - CDbl(txt_listrik.Text) - CDbl(txt_air.Text) - CDbl(txt_lain.Text)
            txt_total.Text = Format(jml_s, "###,###,###")
        
    'End If
End Sub

Private Sub txt_air_LostFocus()
    If txt_air.Text = "" Then
        txt_air.Text = 0
    End If
End Sub
Private Sub txt_kode_GotFocus()
    txt_kode.SelStart = 0
    txt_kode.SelLength = Len(txt_kode)
End Sub

Private Sub txt_lain_GotFocus()
    txt_lain.SelStart = 0
    txt_lain.SelLength = Len(txt_lain)
End Sub

Private Sub txt_lain_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
    Beep
    KeyAscii = 0
End If
End Sub

Private Sub txt_lain_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    'If txt_lain.Text <> "" Then
        
        If txt_lain.Text = "" Then
            txt_lain.Text = 0
        End If
        
        If txt_lain.Text <> 0 Then
            txt_lain.Text = Format(txt_lain.Text, "###,###,###")
            txt_lain.SelStart = Len(txt_lain)
        End If
        
        Dim jml_s As Double
        txt_total.Text = 0
            jml_s = CDbl(lbl_jml.Caption) - CDbl(txt_listrik.Text) - CDbl(txt_air.Text) - CDbl(txt_lain.Text)
            txt_total.Text = Format(jml_s, "###,###,###")
        
     'End If
End Sub

Private Sub txt_lain_LostFocus()
    If txt_lain.Text = "" Then
        txt_lain.Text = 0
    End If
End Sub

Private Sub txt_listrik_GotFocus()
    txt_listrik.SelStart = 0
    txt_listrik.SelLength = Len(txt_listrik)
End Sub

Private Sub txt_listrik_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And txt_kode.Text <> "" Then
        Pic_listrik.Visible = True
        grd_listrik.SetFocus
    End If
End Sub

Private Sub txt_listrik_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
    Beep
    KeyAscii = 0
End If
End Sub

Private Sub txt_listrik_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    'If txt_listrik.Text <> "" Then
        
        If txt_listrik.Text = "" Then
            txt_listrik.Text = 0
        End If
        
        If txt_listrik.Text <> 0 Then
            txt_listrik.Text = Format(txt_listrik.Text, "###,###,###")
            txt_listrik.SelStart = Len(txt_listrik)
        End If
        
        Dim jml_s As Double
        txt_total.Text = 0
            jml_s = CDbl(lbl_jml.Caption) - CDbl(txt_listrik.Text) - CDbl(txt_air.Text) - CDbl(txt_lain.Text)
            txt_total.Text = Format(jml_s, "###,###,###")
            
    'End If
End Sub

Private Sub txt_listrik_LostFocus()
    If txt_listrik.Text = "" Then
        txt_listrik.Text = 0
    End If
End Sub

Private Sub txt_total_GotFocus()
    txt_total.SelStart = 0
    txt_total.SelLength = Len(txt_total)
End Sub

Private Sub txt_total_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
    Beep
    KeyAscii = 0
End If
End Sub

Private Sub txt_total_KeyUp(KeyCode As Integer, Shift As Integer)
    If txt_total.Text <> "" Then
        
        txt_total.Text = Format(txt_total.Text, "###,###,###")
        txt_total.SelStart = Len(txt_total.Text)
        
    End If
End Sub

Private Sub txt_total_LostFocus()
    If txt_total.Text = "" Then
        txt_total.Text = 0
    End If
End Sub
