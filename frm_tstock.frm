VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Begin VB.Form frm_tstock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Stock"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   Begin TDBContainer3D6Ctl.TDBContainer3D pic_counter 
      Height          =   5415
      Left            =   -3120
      TabIndex        =   36
      Top             =   5040
      Visible         =   0   'False
      Width           =   5775
      _Version        =   65536
      _ExtentX        =   10186
      _ExtentY        =   9551
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "frm_tstock.frx":0000
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_tstock.frx":001C
      Childs          =   "frm_tstock.frx":00C8
      Begin VB.TextBox txt_c 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   240
         TabIndex        =   39
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txt_c 
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
         Index           =   1
         Left            =   2520
         TabIndex        =   40
         Top             =   960
         Width           =   2895
      End
      Begin VB.Frame Frame5 
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   37
         Top             =   480
         Width           =   5175
      End
      Begin TrueDBGrid60.TDBGrid grd_counter 
         Height          =   3735
         Left            =   240
         OleObjectBlob   =   "frm_tstock.frx":00E4
         TabIndex        =   43
         Top             =   1440
         Width           =   5175
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Nama Counter"
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
         Left            =   3345
         TabIndex        =   42
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Kode Counter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   600
         Width           =   2115
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
         Index           =   1
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   1065
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D pic_barang 
      Height          =   5295
      Left            =   3000
      TabIndex        =   28
      Top             =   1080
      Visible         =   0   'False
      Width           =   7095
      _Version        =   65536
      _ExtentX        =   12515
      _ExtentY        =   9340
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "frm_tstock.frx":2F01
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_tstock.frx":2F1D
      Childs          =   "frm_tstock.frx":2FC9
      Begin VB.TextBox Text1 
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
         Left            =   1680
         TabIndex        =   31
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
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
         Left            =   4560
         TabIndex        =   32
         Top             =   720
         Width           =   2175
      End
      Begin VB.Frame Frame5 
         Height          =   135
         Index           =   4
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   6615
      End
      Begin TrueDBGrid60.TDBGrid grd_barang 
         Height          =   3855
         Left            =   240
         OleObjectBlob   =   "frm_tstock.frx":2FE5
         TabIndex        =   35
         Top             =   1080
         Width           =   6615
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Kode Barang"
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
         Left            =   345
         TabIndex        =   34
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Nama Barang"
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
         Left            =   3210
         TabIndex        =   33
         Top             =   720
         Width           =   1185
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
         Index           =   16
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   2460
      End
   End
   Begin VB.PictureBox g 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   -6120
      ScaleHeight     =   4905
      ScaleWidth      =   6225
      TabIndex        =   24
      Top             =   3840
      Visible         =   0   'False
      Width           =   6255
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   6225
         TabIndex        =   25
         Top             =   0
         Width           =   6255
         Begin VB.CommandButton cmd_x_c 
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
            Left            =   5760
            TabIndex        =   26
            Top             =   0
            Width           =   495
         End
      End
   End
   Begin VB.PictureBox ppp 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   -6000
      ScaleHeight     =   4905
      ScaleWidth      =   6225
      TabIndex        =   17
      Top             =   4560
      Visible         =   0   'False
      Width           =   6255
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   6225
         TabIndex        =   18
         Top             =   0
         Width           =   6255
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
            Left            =   5760
            TabIndex        =   19
            Top             =   0
            Width           =   495
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   120
      ScaleHeight     =   6105
      ScaleWidth      =   9945
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin MSComCtl2.DTPicker dtp_tgl 
         Height          =   375
         Left            =   720
         TabIndex        =   27
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   12582912
         CalendarTitleForeColor=   16777215
         Format          =   54394881
         CurrentDate     =   39211
      End
      Begin VB.TextBox txt_kode_counter 
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
         Left            =   2400
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmd_simpan 
         Caption         =   "Simpan"
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
         TabIndex        =   4
         Top             =   5520
         Width           =   1335
      End
      Begin VB.TextBox txt_jml 
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
         Left            =   2400
         TabIndex        =   3
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txt_kode_barang 
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
         Left            =   2400
         TabIndex        =   2
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lbl_nama_counter 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   23
         Top             =   1320
         Width           =   4695
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Counter"
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
         TabIndex        =   22
         Top             =   1440
         Width           =   1425
      End
      Begin VB.Label Label11 
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
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl."
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
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   9600
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label lbl_jml 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   495
         Left            =   2520
         TabIndex        =   16
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jml Seluruh"
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
         Left            =   2520
         TabIndex        =   15
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label lbl_tersedia 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barang yang tersedia"
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
         TabIndex        =   13
         Top             =   4440
         Width           =   2085
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   9720
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jml Barang"
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
         Top             =   3720
         Width           =   1080
      End
      Begin VB.Label lbl_max 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   11
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max Barang"
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
         Left            =   3240
         TabIndex        =   10
         Top             =   3120
         Width           =   1185
      End
      Begin VB.Label lbl_min 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   9
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min Barang"
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
         TabIndex        =   8
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label lbl_nama 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   7
         Top             =   2400
         Width           =   4695
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   1350
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frm_tstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_barang As New XArrayDB
Dim arr_counter As New XArrayDB
Dim id_b As String
Dim id_counter As String
Dim Moving As Boolean
Dim yold, xold As Long

Private Sub Cmd_Simpan_Click()
    
On Error GoTo er_s
    
    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    Dim sql2 As String
    Dim rs2 As New ADODB.Recordset
        
        
        If txt_kode_counter.Text = "" Or Txt_Kode_Barang.Text = "" Or txt_jml.Text = "" Or txt_jml.Text = 0 Then
            MsgBox ("Semua data harus diisi")
            Exit Sub
        End If
               
        If CDbl(lbl_jml.Caption) > CDbl(lbl_max.Caption) Then
            MsgBox ("Stock yang dimasukkan melebihi Stock maximum")
            txt_jml.SetFocus
            Exit Sub
        End If
        
        If CDbl(lbl_jml.Caption) < CDbl(lbl_min.Caption) Then
            MsgBox ("Stock yang dimasukkan kurang dari stock minimum")
            txt_jml.SetFocus
            Exit Sub
        End If
               
 If mdl_stock = True Then
          
        If MsgBox("Yakin data yang dimasukkan sudah benar", vbYesNo + vbQuestion, "Pesan") = vbNo Then
            Exit Sub
        End If
        
        cn.BeginTrans
        
            sql1 = "insert into tr_stock (id_barang,brg_in,brg_out,tgl,ket,nama_user,kemana)"
            sql1 = sql1 & " values (" & id_b & "," & Trim(txt_jml.Text) & ",0,'" & Trim(dtp_tgl.Value) & "',0,'" & Trim(utama.lbl_user.Caption) & "','-' )"
            rs1.Open sql1, cn
                    
            sql = "select id_barang from tr_jml_stock where id_barang=" & id_b
            rs.Open sql, cn
                If rs.EOF Then
                    sql1 = "insert into tr_jml_stock (id_barang,jml_stock) values(" & id_b & "," & Trim(lbl_jml.Caption) & ")"
                    rs1.Open sql1, cn
                Else
                    sql1 = "update tr_jml_stock set jml_stock=" & Trim(lbl_jml.Caption) & " where id_barang=" & id_b
                    rs1.Open sql1, cn
                End If
            rs.Close
                
        MsgBox ("Data Berhasil disimpan")
        cn.CommitTrans
        kosong_semua
        txt_kode_counter.SetFocus
        Exit Sub
        
'   ElseIf mdl_stock = False Then
'
'        If CDbl(txt_jml.Text) > CDbl(lbl_max.Caption) Then
'            MsgBox ("Data stock yang akan diedit melebihi stock maximum")
'            Exit Sub
'        End If
'
'        If MsgBox("Yakin data yang dimasukkan sudah benar", vbYesNo + vbQuestion, "Pesan") = vbNo Then
'            Exit Sub
'        End If
'
'    cn.BeginTrans
'
'        sql = "select id from tr_jml_stock where id=" & id_sk1
'        rs.Open sql, cn
'            If Not rs.EOF Then
'
'                sql1 = "update tr_jml_stock set jml_stock=" & Trim(txt_jml.Text) & " where id=" & id_sk1
'                rs1.Open sql1, cn
'
'
'
'
'                MsgBox ("Data berhasil disimpan")
'            Else
'                MsgBox ("Data yang akan diedit tidak ditemukan")
'
'            End If
'       rs.Close
'       frm_browse_tstock.cmd_tampil.Default = True
'       Unload Me
'       Exit Sub
'
   End If
        
        
er_s:
    cn.RollbackTrans
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub

Private Sub kosong_semua()
    txt_kode_counter.Text = ""
    lbl_nama_counter.Caption = ""
    Txt_Kode_Barang.Text = ""
    Lbl_Nama.Caption = ""
    lbl_min.Caption = 0
    lbl_max.Caption = 0
    txt_jml.Text = 0
    lbl_tersedia.Caption = 0
    lbl_jml.Caption = 0
    
End Sub

Private Sub cmd_x_c_Click()
    pic_counter.Visible = False
    txt_kode_counter.SetFocus
End Sub

Private Sub cmd_x_Click()
    pic_barang.Visible = False
    Txt_Kode_Barang.SetFocus
End Sub

Private Sub Form_Activate()
    txt_kode_counter.SetFocus
End Sub

Private Sub Form_Load()
        
   
        
    grd_counter.Array = arr_counter

    grd_barang.Array = arr_barang
    
    dtp_tgl.Value = Format(Date, "dd/mm/yyyy")
    
    With pic_barang
        .Left = 3000
        .Top = 600
    End With
    
    With pic_counter
        .Left = 2760
        .Top = 720
    End With
    
    isi_counter
    
If mdl_stock = True Then

    Txt_Kode_Barang.Enabled = True
        
    lbl_max.Caption = 0
    lbl_min.Caption = 0
    lbl_tersedia.Caption = 0
    txt_jml.Text = 0
    
    kosong_semua
    
ElseIf mdl_stock = False Then
    
    Txt_Kode_Barang.Enabled = False
    
    Call tolong_diisi
    
End If
    
   Me.Left = Screen.Width \ 2 - Me.Width \ 2 + 750
   Me.Top = Screen.Height \ 2 - Me.Height \ 2 - 2100
    
End Sub

Private Sub kosong_counter()
    arr_counter.ReDim 0, 0, 0, 0
    grd_counter.ReBind
    grd_counter.Refresh
End Sub

Private Sub isi_counter()

On Error GoTo er_counter

    Dim sql As String
    Dim rs_c As New ADODB.Recordset
        
        kosong_counter
        
        sql = "select id,kode,nama_counter from tbl_counter  order by kode"
        rs_c.Open sql, cn, adOpenKeyset
            If Not rs_c.EOF Then
                
                rs_c.MoveLast
                rs_c.MoveFirst
                
                lanjut_counter rs_c
                
            End If
       rs_c.Close
       Exit Sub
       
er_counter:
       Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
       
End Sub

Private Sub lanjut_counter(rs_c As Recordset)
    Dim id_c, kode, nama As String
    Dim a As Long
        
        a = 1
            Do While Not rs_c.EOF
                arr_counter.ReDim 1, a, 0, 3
                grd_counter.ReBind
                grd_counter.Refresh
                    
                    id_c = rs_c("id")
                    
                    If Not IsNull(rs_c("kode")) Then
                        kode = rs_c("kode")
                    Else
                        kode = ""
                    End If
                    
                    If Not IsNull(rs_c("nama_counter")) Then
                        nama = rs_c("nama_counter")
                    Else
                        nama = ""
                    End If
                     
            arr_counter(a, 0) = id_c
            arr_counter(a, 1) = kode
            arr_counter(a, 2) = nama
           a = a + 1
           rs_c.MoveNext
           Loop
           grd_counter.ReBind
           grd_counter.Refresh
End Sub

Private Sub tolong_diisi()
    
On Error GoTo tolong
    
Dim sql As String
Dim rs As New ADODB.Recordset
    
    sql = "select nama_barang,stock_min,stock_max,kode from tbl_barang where id=" & id_sk
        rs.Open sql, cn
            If Not rs.EOF Then
                Txt_Kode_Barang.Text = rs("kode")
                Lbl_Nama.Caption = rs("nama_barang")
                lbl_min.Caption = rs("stock_min")
                lbl_max.Caption = rs("stock_max")
                
                pic_barang.Visible = False
                isi_tersedia (False)
            End If
        rs.Close
Exit Sub

tolong:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear

End Sub

Private Sub kosong()
    arr_barang.ReDim 0, 0, 0, 0
    grd_barang.ReBind
    grd_barang.Refresh
End Sub

Private Sub isi()

On Error GoTo er_isi

    Dim sql As String
    Dim rs_barang As New ADODB.Recordset
        
        kosong
        
        sql = "select id_barang,nama_counter,nama_barang,kode from qr_barang where ket=1 and id_counter=" & id_counter & " order by nama_counter,kode"
        rs_barang.Open sql, cn, adOpenKeyset
            If Not rs_barang.EOF Then
                
                rs_barang.MoveLast
                rs_barang.MoveFirst
                
                isi_barang rs_barang
                     
            End If
        rs_barang.Close
        Exit Sub
        
er_isi:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub isi_barang(rs_barang As Recordset)
    Dim id_barang, counter, barang, kode As String
    Dim a As Long
        
        a = 1
            Do While Not rs_barang.EOF
                arr_barang.ReDim 1, a, 0, 5
                grd_barang.ReBind
                grd_barang.Refresh
                    
                    id_barang = rs_barang("id_barang")
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
                    barang = rs_barang("nama_barang")
                  End If
                  
               arr_barang(a, 0) = id_barang
               arr_barang(a, 1) = counter
               arr_barang(a, 2) = kode
               arr_barang(a, 3) = barang
               
           a = a + 1
           rs_barang.MoveNext
           Loop
           grd_barang.ReBind
           grd_barang.Refresh
           
                  
End Sub

Private Sub grd_barang_Click()
    On Error Resume Next
        If arr_barang.UpperBound(1) > 0 Then
            id_b = arr_barang(grd_barang.Bookmark, 0)
        End If
End Sub

Private Sub grd_barang_DblClick()

On Error GoTo er_d

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        sql = "select nama_barang,stock_min,stock_max,kode from tbl_barang where id=" & id_b
        rs.Open sql, cn
            If Not rs.EOF Then
                Txt_Kode_Barang.Text = rs("kode")
                Lbl_Nama.Caption = rs("nama_barang")
                lbl_min.Caption = rs("stock_min")
                lbl_max.Caption = rs("stock_max")
                
                pic_barang.Visible = False
                isi_tersedia (True)
                Txt_Kode_Barang.SetFocus
            End If
        rs.Close
    Exit Sub
    
er_d:
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
        
        isi
        
    End If
End Sub

Private Sub grd_counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        grd_counter_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        pic_counter.Visible = False
        txt_kode_counter.SetFocus
    End If
End Sub

Private Sub grd_counter_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_counter_Click
End Sub

Private Sub pic_barang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_barang.Visible = False
        Txt_Kode_Barang.SetFocus
    End If
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

Private Sub TDBContainer3D1_Click()

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

Private Sub Text1_GotFocus(Index As Integer)
    Select Case Index
        
        Case 0
            Text1(0).SelStart = 0
            Text1(0).SelLength = Len(Text1(0))
        Case 1
            Text1(1).SelStart = 0
            Text1(1).SelLength = Len(Text1(1))
            
   End Select
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_barang.Visible = False
        Txt_Kode_Barang.SetFocus
    End If
    
    If KeyCode = 13 Then
        grd_barang_DblClick
    End If
    
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo er_up

    Dim sql As String
    Dim rs_barang As New ADODB.Recordset
        
    If arr_barang.UpperBound(1) = 0 Then
        Exit Sub
    End If
        
        sql = "select id_barang,nama_counter,nama_barang,kode from qr_barang where ket=1 and kode_counter='" & Trim(txt_kode_counter.Text) & "'"
        
    Select Case Index
        
        Case 0
            
            sql = sql & " and kode like '%" & Trim(Text1(0).Text) & "%'"
            
        Case 1
            
            sql = sql & " and nama_barang like '%" & Trim(Text1(1).Text) & "%'"
   End Select
        
        sql = sql & " order by nama_counter"
            
        rs_barang.Open sql, cn, adOpenKeyset
        
        arr_barang.ReDim 0, 0, 0, 0
        arr_barang.ReDim 1, 1, 1, 1
            grd_barang.ReBind
            grd_barang.Refresh
        
            If Not rs_barang.EOF Then
                
                rs_barang.MoveLast
                rs_barang.MoveFirst
                
                isi_barang rs_barang
            End If
       rs_barang.Close
                
Exit Sub

er_up:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
                
End Sub

Private Sub txt_c_GotFocus(Index As Integer)
    Select Case Index
        Case 0
            txt_c(0).SelStart = 0
            txt_c(1).SelLength = Len(txt_c(0))
        Case 1
            txt_c(1).SelStart = 0
            txt_c(1).SelLength = Len(txt_c(1))
   End Select
End Sub

Private Sub txt_c_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        grd_counter_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        pic_counter.Visible = -False
        txt_kode_counter.SetFocus
    End If
End Sub

Private Sub txt_c_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
On Error GoTo er_c
    
    Dim sql As String
    Dim rs_c As New ADODB.Recordset
        
        kosong_counter
        
        sql = "select id,kode,nama_counter from tbl_counter"
        
    If txt_c(0).Text <> "" Or txt_c(1).Text <> "" Then
        sql = sql & " where"
        Select Case Index
            Case 0
                sql = sql & " kode like '%" & Trim(txt_c(0).Text) & "%'"
            Case 1
                sql = sql & " nama_counter like '%" & Trim(txt_c(1).Text) & "%'"
        End Select
   End If
   
   sql = sql & " order by kode"
   rs_c.Open sql, cn, adOpenKeyset
    If Not rs_c.EOF Then
        
        rs_c.MoveLast
        rs_c.MoveFirst
        
        lanjut_counter rs_c
    End If
  rs_c.Close
  
  Exit Sub
  
er_c:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
  
End Sub

Private Sub txt_jml_Change()
If mdl_stock = True Then

lbl_jml.Caption = 0

    If txt_jml.Text <> "" Then
    
        Dim jml_semua As Double
        jml_semua = CDbl(txt_jml.Text) + CDbl(lbl_tersedia.Caption)
        lbl_jml.Caption = jml_semua
           
    End If
    
End If
End Sub

Private Sub txt_jml_GotFocus()
    txt_jml.SelStart = 0
    txt_jml.SelLength = Len(txt_jml)
End Sub

Private Sub txt_jml_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub


Private Sub txt_kode_barang_GotFocus()
    Txt_Kode_Barang.SelStart = 0
    Txt_Kode_Barang.SelLength = Len(Txt_Kode_Barang)
End Sub

Private Sub txt_kode_barang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        
        If txt_kode_counter.Text = "" Then
            Dim messg As Integer
                messg = CInt(MsgBox("Kode Counter harus diisi", vbOKOnly + vbInformation, "Konfimasi"))
                txt_kode_counter.SetFocus
                Exit Sub
        End If
        
        Txt_Kode_Barang.Text = ""
        pic_barang.Visible = True
        Text1(0).Text = ""
        Text1(1).Text = ""
        grd_barang.MoveFirst
        Text1(0).SetFocus
    End If
End Sub

Private Sub txt_kode_barang_LostFocus()

On Error GoTo er_los

    If Txt_Kode_Barang.Text <> "" Then
        Dim sql, sql1 As String
        Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
            
            sql = "select id,nama_barang,stock_min,stock_max from tbl_barang where kode='" & Trim(Txt_Kode_Barang.Text) & "'and id_counter=" & id_counter & " and ket=1"
            rs.Open sql, cn
                If Not rs.EOF Then
                        id_b = rs("id")
                        Lbl_Nama.Caption = rs("nama_barang")
                        lbl_min.Caption = rs("stock_min")
                        lbl_max.Caption = rs("stock_max")
                    
                        isi_tersedia (True)
                Else
                        MsgBox ("kode barang dan kode counter yang anda masukkan tidak menyediakan perhitungan stock")
                        Txt_Kode_Barang.SetFocus
                        Exit Sub
                End If
            rs.Close
    End If
    
    Exit Sub
    
er_los:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub

Private Sub isi_tersedia(mana As Boolean)

On Error GoTo er_t

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
If mana = True Then
    
        sql = "select jml_stock from tr_jml_stock where id_barang=" & id_b
   
End If

        rs.Open sql, cn
            If Not rs.EOF Then
                lbl_tersedia.Caption = rs("jml_stock")
                  
            Else
            
                lbl_tersedia.Caption = 0
                
            End If
        rs.Close
        Exit Sub
        
er_t:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description)
        Err.Clear
        
End Sub

Private Sub txt_kode_counter_GotFocus()
    txt_kode_counter.SelStart = 0
    txt_kode_counter.SelLength = Len(txt_kode_counter)
End Sub

Private Sub txt_kode_counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txt_kode_counter.Text = ""
        lbl_nama_counter.Caption = ""
        txt_c(0).Text = ""
        txt_c(1).Text = ""
        pic_counter.Visible = True
        txt_c(0).SetFocus
    End If
End Sub

Private Sub txt_kode_counter_LostFocus()

On Error GoTo lagi

If txt_kode_counter.Text <> "" Then
    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        sql = "select nama_counter,id_counter from qr_barang where kode_counter='" & Trim(txt_kode_counter.Text) & "'and ket=1"
        rs.Open sql, cn
            If Not rs.EOF Then
                id_counter = rs("id_counter")
                lbl_nama_counter.Caption = rs("nama_counter")
                
                isi
                
            Else
                MsgBox ("Kode yang anda masukkan tidak ditemukan / Tidak menyediakan perhitungan stock")
                txt_kode_counter.SetFocus
            End If
        rs.Close
End If

Exit Sub

lagi:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear

End Sub
