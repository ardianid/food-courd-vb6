VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_tr_biling 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic_counter 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   2400
      ScaleHeight     =   5865
      ScaleWidth      =   5625
      TabIndex        =   19
      Top             =   1200
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox txt_nm 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txt_nm 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   2520
         TabIndex        =   21
         Top             =   840
         Width           =   2895
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   0
         Width           =   375
      End
      Begin TrueDBGrid60.TDBGrid grd_counter 
         Height          =   4455
         Left            =   120
         OleObjectBlob   =   "frm_tr_biling.frx":0000
         TabIndex        =   23
         Top             =   1320
         Width           =   5415
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   5625
         TabIndex        =   22
         Top             =   0
         Width           =   5655
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
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
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
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
         Left            =   2520
         TabIndex        =   24
         Top             =   480
         Width           =   3015
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   7185
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.TextBox txt_bayar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   28
         Top             =   3480
         Width           =   4335
      End
      Begin VB.CommandButton cmd_simpan 
         Caption         =   "Simpan"
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
         Left            =   5760
         TabIndex        =   18
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox txt_bln_ini 
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
         Left            =   5160
         TabIndex        =   17
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txt_jml_akhir 
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
         Left            =   1920
         TabIndex        =   15
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txt_kode 
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
         Left            =   2040
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtp_tgl 
         Height          =   375
         Left            =   4560
         TabIndex        =   5
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   55050241
         CurrentDate     =   38631
      End
      Begin VB.Frame Frame1 
         Caption         =   "Jenis Transaksi Billing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6855
         Begin VB.OptionButton opt_air 
            Caption         =   "Air"
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
            Left            =   2880
            TabIndex        =   3
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton opt_listrik 
            Caption         =   "Listrik"
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
            Left            =   240
            TabIndex        =   2
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Bayar"
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
         Left            =   240
         TabIndex        =   27
         Top             =   3600
         Width           =   1410
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penilaian bln ini"
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
         Left            =   3120
         TabIndex        =   16
         Top             =   3120
         Width           =   1635
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penilaian Akhir"
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
         Left            =   240
         TabIndex        =   14
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label lbl_jumlah_awal 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1920
         TabIndex        =   13
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Awal"
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
         Left            =   240
         TabIndex        =   12
         Top             =   2520
         Width           =   1305
      End
      Begin VB.Label lbl_harga 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5160
         TabIndex        =   11
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7080
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga"
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
         Left            =   3120
         TabIndex        =   10
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label lbl_counter 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2040
         TabIndex        =   9
         Top             =   1560
         Width           =   4935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   270
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   270
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
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
         Left            =   4080
         TabIndex        =   4
         Top             =   1080
         Width           =   375
      End
   End
End
Attribute VB_Name = "frm_tr_biling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_counter As New XArrayDB
Dim id_cntr As String
Dim id_biling As String

Private Sub cmd_simpan_Click()
Dim sql, sql1 As String
Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    
On Error GoTo er_simpan
    
    If txt_kode.Text = "" Then
        MsgBox ("Kode counter harus diisi")
        Exit Sub
    End If
    
    If MsgBox("Yakin data yang dimasukkan sudah benar", vbYesNo + vbQuestion, "Pesan") = vbNo Then
        Exit Sub
    End If
    
    If opt_listrik.Value = True Then
       sql = "insert into tr_biling_listrik (id_biling,tgl,pakai,harga,nama_user)"
       sql = sql & " values (" & id_biling & ",'" & Trim(dtp_tgl.Value) & "'," & Trim(txt_bln_ini.Text) & "," & CCur(txt_bayar.Text) & ",'" & Trim(utama.lbl_user.Caption) & "')"
       rs.Open sql, cn
    End If
    
    If opt_air.Value = True Then
       sql = "insert into tr_biling_air (id_biling,tgl,pakai,harga,nama_user)"
       sql = sql & " values (" & id_biling & ",'" & Trim(dtp_tgl.Value) & "'," & Trim(txt_bln_ini.Text) & "," & CCur(txt_bayar.Text) & ",'" & Trim(utama.lbl_user.Caption) & "')"
       rs.Open sql, cn
    End If
    
    MsgBox ("Data berhasil disimpan")
    txt_kode.Text = ""
    lbl_counter.Caption = ""
    lbl_jumlah_awal.Caption = 0
    lbl_harga.Caption = 0
    txt_jml_akhir.Text = 0
    txt_bln_ini.Text = 0
    txt_bayar.Text = 0
    txt_kode.SetFocus
    Exit Sub
    
er_simpan:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub Form_Activate()
    txt_kode.SetFocus
End Sub

Private Sub Form_Load()
    
    grd_counter.Array = arr_counter
    
    opt_listrik.Value = True
    
    dtp_tgl.Value = Format(Date, "dd/mm/yyyy")
    
    Label4.Caption = "Harga Listrik/Kwh"
    lbl_jumlah_awal.Caption = 0
    lbl_harga.Caption = 0
    txt_jml_akhir.Text = 0
    txt_bln_ini.Text = 0
    txt_bayar.Text = 0
    
    isi_counter
    
    besar (True)
    
    txt_jml_akhir.Text = 0
    txt_bln_ini.Text = 0
    
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2 - 2300
    
End Sub

Private Sub cari_lama(param As Boolean)

On Error GoTo er_lama

    Dim sql, sql1, sql2 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset
    
            Select Case param
                Case True
                 sql2 = "select id_biling_listrik from qr_biling_listrik where id_counter=" & id_cntr
                 rs2.Open sql2, cn
                  If Not rs2.EOF Then
                    sql = "select max(id_biling_listrik) as id_listrik from qr_biling_listrik where id_counter=" & id_cntr
                    rs.Open sql, cn
                        If Not rs.EOF Then
                            sql1 = "select pakai from qr_biling_listrik where id_biling_listrik=" & rs("id_listrik")
                            rs1.Open sql1, cn
                                If Not rs1.EOF Then
                                    lbl_jumlah_awal.Caption = rs1("pakai")
                                End If
                            rs1.Close
                        Else
                            lbl_jumlah_awal.Caption = 0
                        End If
                    rs.Close
                  Else
                    lbl_jumlah_awal.Caption = 0
                  End If
                Case False
                 sql2 = "select id_biling_air from qr_biling_air where id_counter=" & id_cntr
                 rs2.Open sql2, cn
                  If Not rs2.EOF Then
                    sql = "select max(id_biling_air) as id_air from qr_biling_air where id_counter=" & id_cntr
                    rs.Open sql, cn
                        If Not rs.EOF Then
                            sql1 = "select pakai from qr_biling_air where id_biling_air=" & rs("id_air")
                            rs1.Open sql1, cn
                                If Not rs1.EOF Then
                                    lbl_jumlah_awal.Caption = rs1("pakai")
                                End If
                            rs1.Close
                        Else
                            lbl_jumlah_awal.Caption = 0
                        End If
                    rs.Close
                  Else
                    lbl_jumlah_awal.Caption = 0
                  End If
           End Select
           
           Exit Sub
           
er_lama:
           Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
           
End Sub

Private Sub isi_harga(param As Boolean)

On Error GoTo er_harga

Dim sql As String
Dim rs As New ADODB.Recordset
    
    sql = "select harga_air,harga_listrik,id from tbl_biling where id_counter=" & id_cntr
    rs.Open sql, cn
        If Not rs.EOF Then
            Select Case param
                Case True
                    lbl_harga.Caption = Format(rs("harga_listrik"), "currency")
                Case False
                    lbl_harga.Caption = Format(rs("harga_air"), "currency")
            End Select
            
            id_biling = rs("id")
        Else
        
            MsgBox ("Data biling listrik dan air counter tidak ditemukan")
            
        End If
    rs.Close
    
    Exit Sub
    
er_harga:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub

Private Sub kosong_counter()
    arr_counter.ReDim 0, 0, 0, 0
    grd_counter.ReBind
    grd_counter.Refresh
End Sub


Private Sub isi_counter()

On Error GoTo isi_counter

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
        
isi_counter:
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

Private Sub opt_air_Click()
Label4.Caption = "Harga Air/M3"
lbl_harga.Caption = 0
txt_jml_akhir.Text = 0
txt_bln_ini.Text = 0
txt_bayar.Text = 0
lbl_jumlah_awal.Caption = 0
    If txt_kode.Text <> "" Then
        txt_kode_LostFocus
    End If
End Sub

Private Sub opt_listrik_Click()
Label4.Caption = "Harga Listrik/Kwh"
lbl_harga.Caption = 0
txt_jml_akhir.Text = 0
txt_bln_ini.Text = 0
txt_bayar.Text = 0
lbl_jumlah_awal.Caption = 0
    If txt_kode.Text <> "" Then
        txt_kode_LostFocus
    End If
End Sub

Private Sub pic_counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_counter.Visible = False
        besar (True)
        txt_kode.SetFocus
    End If
End Sub

Private Sub txt_bayar_GotFocus()
    txt_bayar.SelStart = 0
    txt_bayar.SelLength = Len(txt_bayar)
End Sub

Private Sub txt_bayar_KeyUp(KeyCode As Integer, Shift As Integer)
    If txt_bayar.Text <> "" Then
        txt_bayar.Text = Format(txt_bayar.Text, "###,###,###")
        txt_bayar.SelStart = Len(txt_bayar.Text)
    Else
        txt_bayar.Text = 0
    End If
End Sub

Private Sub txt_bayar_LostFocus()
    If txt_bayar.Text = "" Then
        txt_bayar.Text = 0
    End If
End Sub

Private Sub txt_bln_ini_GotFocus()
    txt_bln_ini.SelStart = 0
    txt_bln_ini.SelLength = Len(txt_bln_ini)
End Sub

Private Sub txt_bln_ini_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
    Beep
    KeyAscii = 0
End If
End Sub

Private Sub txt_bln_ini_KeyUp(KeyCode As Integer, Shift As Integer)
    If txt_bln_ini.Text <> "" Then
        Dim jumlah As Double
            jumlah = CDbl(txt_bln_ini.Text) * CDbl(lbl_harga.Caption)
            txt_bayar.Text = Format(jumlah, "###,###,###")
    Else
        txt_bln_ini.Text = 0
    End If
End Sub

Private Sub txt_jml_akhir_GotFocus()
    txt_jml_akhir.SelStart = 0
    txt_jml_akhir.SelLength = Len(txt_jml_akhir)
End Sub

Private Sub txt_jml_akhir_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
    Beep
    KeyAscii = 0
End If
End Sub

Private Sub txt_jml_akhir_KeyUp(KeyCode As Integer, Shift As Integer)

txt_bln_ini.Text = 0
    If txt_jml_akhir.Text <> "" Then
        Dim sementara, jumlah As Double
            If lbl_jumlah_awal.Caption = 0 Then
                sementara = CDbl(txt_jml_akhir.Text) - CDbl(lbl_jumlah_awal.Caption)
            Else
                sementara = CDbl(txt_jml_akhir.Text) - CDbl(lbl_jumlah_awal.Caption)
            End If
            
            txt_bln_ini.Text = sementara
            jumlah = CDbl(txt_bln_ini.Text) * CDbl(lbl_harga.Caption)
            txt_bayar.Text = Format(jumlah, "###,###,###")
    Else
        txt_jml_akhir.Text = 0
    End If
End Sub

Private Sub txt_kode_GotFocus()
    txt_kode.SelStart = 0
    txt_kode.SelLength = Len(txt_kode)
End Sub

Private Sub txt_kode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txt_kode.Text = ""
        lbl_counter.Caption = ""
        txt_nm(0).Text = ""
        txt_nm(1).Text = ""
        besar (False)
        pic_counter.Visible = True
        txt_nm(1).SetFocus
    End If
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
        besar (True)
        txt_kode.SetFocus
    End If
    
    If KeyCode = 13 Then
        txt_kode.Text = arr_counter(grd_counter.Bookmark, 1)
        lbl_counter.Caption = arr_counter(grd_counter.Bookmark, 2)
        pic_counter.Visible = False
        besar (True)
        txt_kode.SetFocus
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

Private Sub grd_counter_Click()
    On Error Resume Next
        If arr_counter.UpperBound(1) > 0 Then
            id_cntr = arr_counter(grd_counter.Bookmark, 0)
        End If
End Sub

Private Sub grd_counter_DblClick()
If arr_counter.UpperBound(1) > 0 Then
    
    txt_kode.Text = arr_counter(grd_counter.Bookmark, 1)
    lbl_counter.Caption = arr_counter(grd_counter.Bookmark, 2)
    pic_counter.Visible = False
    besar (True)
    txt_kode.SetFocus
End If
    
End Sub

Private Sub grd_counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        grd_counter_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        pic_counter.Visible = False
        besar (True)
        txt_kode.SetFocus
    End If
End Sub

Private Sub grd_counter_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_counter_Click
End Sub

Private Sub txt_kode_LostFocus()

On Error GoTo los

If txt_kode.Text <> "" Then

    Dim sql As String
    Dim rs As New ADODB.Recordset
                
        lbl_harga.Caption = 0
                
        sql = "select id,kode,nama_counter from tbl_counter where kode='" & Trim(txt_kode.Text) & "'"
        rs.Open sql, cn
            If Not rs.EOF Then
                id_cntr = rs("id")
                lbl_counter.Caption = rs("nama_counter")
                
                If opt_listrik.Value = True Then
                    Call isi_harga(True)
                    Call cari_lama(True)
                ElseIf opt_air.Value = True Then
                    Call isi_harga(False)
                    Call cari_lama(False)
                End If
                
            Else
                MsgBox ("Kode counter yang anda masukkan tidak ditemukan")
                txt_kode.SetFocus
            End If
        rs.Close
End If

Exit Sub

los:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear

End Sub


Private Sub cmd_x_Click()
    pic_counter.Visible = False
    besar (True)
    txt_kode.SetFocus
End Sub

Private Sub besar(param As Boolean)
    Select Case param
        Case False
            Me.Height = 7695
            Me.Width = 8265
            Me.ScaleHeight = 7215
            Me.ScaleWidth = 8175
        Case True
            Me.Height = 5415
            Me.Width = 7560
            Me.ScaleHeight = 4935
            Me.ScaleWidth = 7470
    End Select
End Sub
