VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Object = "{EC76FE26-BAFD-4E89-AA40-E748DA83A570}#1.0#0"; "IsButton_Ard.ocx"
Begin VB.Form frm_masuk 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5610
   ControlBox      =   0   'False
   Icon            =   "frm_masuk.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin TDBContainer3D6Ctl.TDBContainer3D pic_karyawan 
      Height          =   5535
      Left            =   -480
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   9763
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "frm_masuk.frx":27C92
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_masuk.frx":27CAE
      Childs          =   "frm_masuk.frx":27D5A
      Begin VB.TextBox txt_cari 
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
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   4335
      End
      Begin VB.Frame Frame5 
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   4335
      End
      Begin TrueDBGrid60.TDBGrid grd_karyawan 
         Height          =   3975
         Left            =   240
         OleObjectBlob   =   "frm_masuk.frx":27D76
         TabIndex        =   6
         Top             =   1320
         Width           =   4335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Nama Karyawan"
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
         TabIndex        =   8
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN KARYAWAN"
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
         TabIndex        =   5
         Top             =   240
         Width           =   2190
      End
   End
   Begin TDBContainer3D6Ctl.TDBContainer3D TDBContainer3D1 
      Height          =   3015
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   5318
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "frm_masuk.frx":2A46C
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_masuk.frx":2A488
      Childs          =   "frm_masuk.frx":2A534
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   240
         ScaleHeight     =   2505
         ScaleWidth      =   5025
         TabIndex        =   10
         Top             =   240
         Width           =   5055
         Begin VB.Frame Frame1 
            BackColor       =   &H00FF0000&
            Height          =   135
            Left            =   240
            TabIndex        =   17
            Top             =   1560
            Width           =   4575
         End
         Begin VB.TextBox txt_nama 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   2040
            TabIndex        =   13
            Top             =   360
            Width           =   2775
         End
         Begin VB.TextBox txt_pwd 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   2040
            PasswordChar    =   "*"
            TabIndex        =   12
            Top             =   840
            Width           =   2775
         End
         Begin IsButton_Ard.isButton cmd_ok 
            Height          =   615
            Left            =   2040
            TabIndex        =   11
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   1085
            Icon            =   "frm_masuk.frx":2A550
            Style           =   10
            Caption         =   "OK"
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
         Begin IsButton_Ard.isButton cmd_batal 
            Height          =   615
            Left            =   3480
            TabIndex        =   14
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   1085
            Icon            =   "frm_masuk.frx":2A56C
            Style           =   10
            Caption         =   "Cancel"
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
            Caption         =   "Nama Karyawan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   1560
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   240
            TabIndex        =   15
            Top             =   840
            Width           =   945
         End
      End
   End
   Begin VB.PictureBox C 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   -4800
      ScaleHeight     =   5625
      ScaleWidth      =   5025
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   5055
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   5025
         TabIndex        =   1
         Top             =   0
         Width           =   5055
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
            Left            =   4680
            TabIndex        =   2
            Top             =   0
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "frm_masuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_karyawan As New XArrayDB
Dim id_kr As String, jml As Integer

Dim Moving As Boolean
Dim yold, xold As Long
Dim status As String

Private Sub normal()

    Me.BackColor = &HFFFFFF
    Me.Height = 3045
    Me.Width = 5700
    Me.ScaleHeight = 3045
    Me.ScaleWidth = 5700
    
End Sub

Private Sub besar()
    
    Me.BackColor = &HC0C0C0
    Me.Height = 5895
    Me.Width = 7455
    Me.ScaleHeight = 5895
    Me.ScaleWidth = 7455
    
End Sub

Private Sub cmd_batal_Click()
    Unload Me
    cn.Close
    Set cn = Nothing
    End
End Sub

Private Sub cmd_batal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmd_batal_Click
    End If
End Sub

Private Sub cmd_ok_Click()
    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
        
   On Error GoTo er_masuk
        
        If Txt_Nama.Text = "" Then
            MsgBox ("Nama karyawan harus diisi")
            Txt_Nama.SetFocus
            Exit Sub
        End If
        
        sql = "select id_karyawan from tbl_user where aktif=1 and id_karyawan=" & id_kr
        rs.Open sql, cn
            If Not rs.EOF Then
                
                Dim pwd_masuk As String
                    pwd_masuk = encrypt(txt_pwd.Text)
                
                sql1 = "select id,pwd,nama_karyawan,foto from qr_user where pwd='" & Trim(pwd_masuk) & "'  and id_karyawan=" & id_kr
                rs1.Open sql1, cn
                    If Not rs1.EOF Then
    
                            id_user = rs1("id")
                            utama.lbl_user.Caption = rs1("nama_karyawan")
                            utama.Show
                            Unload Me
                            Exit Sub
                            
                  Else
                    Dim lg
                    lg = MsgBox("Password anda salah", vbOKOnly + vbInformation, "Pesan")
                    jml = jml + 1
                    
                    If jml = 3 Then
                        cmd_batal_Click
                        Exit Sub
                    End If
                    
                    txt_pwd.SetFocus
                    txt_pwd_GotFocus
                    Exit Sub
                  End If
                rs1.Close
            Else
                Dim slh
                slh = MsgBox("Nama anda tidak tercantum dalam daftar pemakai program/Password anda sudah tidak aktif lagi ", vbOKOnly + vbInformation, "Pesan")
                jml = jml + 1
                    
                    If jml = 3 Then
                        cmd_batal_Click
                    End If
                    
                Txt_Nama.SetFocus
                Exit Sub
            End If
      rs.Close
      Exit Sub
      
er_masuk:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub cmd_ok_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmd_batal_Click
    End If
End Sub

Private Sub cmd_x_Click()
    pic_karyawan.Visible = False
    normal
    Txt_Nama.SetFocus
End Sub





Private Sub Form_Activate()
    Txt_Nama.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
        cmd_batal_Click
    End If
    
End Sub

Private Sub Form_Load()
    
'    Call buka_lap

status = Buka_Koneksi
 If status = "-2147467259" Then
    Dim konfirm As Integer
        konfirm = CInt(MsgBox("Database tidak ditemukan, silahkan periksa kembali", vbOKOnly + vbInformation, "Informasi"))
        
        Load Frm_Seting_Letak_Database
        Frm_Seting_Letak_Database.Show
        
        Unload Me
        Exit Sub
        
 End If
 
' If Lokasi_foto = False Then
'    konfirm = CInt(MsgBox("Lokasi foto tidak ditemukan, silahkan periksa kembali", vbOKOnly + vbInformation, "Informasi"))
'
'        Load Frm_Seting_Letak_Foto
'        Frm_Seting_Letak_Foto.Show
'
'        Unload Me
'        Exit Sub
'
' End If
 
 
    Call buka_path_foto

'    Call buka_path

'    Call Buka_Koneksi
    
    grd_karyawan.Array = arr_karyawan
    
    isi_karyawan
    
    normal
    
    jml = 0
    
    With pic_karyawan
        .Left = 2400
        .Top = 120
    End With
    
    Me.Left = Screen.Width \ 2 - Me.Width \ 2
    Me.Top = Screen.Height \ 2 - Me.Height \ 2
    
End Sub

Private Sub kosong_karyawan()
    arr_karyawan.ReDim 0, 0, 0, 0
    grd_karyawan.ReBind
    grd_karyawan.Refresh
End Sub

Private Sub isi_karyawan()

On Error GoTo er_isi

    Dim sql As String
    Dim rs_karyawan As New ADODB.Recordset
        
        kosong_karyawan
        
        sql = "select id,nama_karyawan from tbl_karyawan order by nama_karyawan"
        rs_karyawan.Open sql, cn, adOpenKeyset
            If Not rs_karyawan.EOF Then
                
                rs_karyawan.MoveLast
                rs_karyawan.MoveFirst
                
                lanjut_karyawan rs_karyawan
                
            End If
       rs_karyawan.Close
        
       Exit Sub
       
er_isi:
       Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub

Private Sub lanjut_karyawan(rs_karyawan As Recordset)
    Dim id_k, nm As String
    Dim a  As Long
        
        a = 1
            Do While Not rs_karyawan.EOF
                arr_karyawan.ReDim 1, a, 0, 2
                grd_karyawan.ReBind
                grd_karyawan.Refresh
                    
                    id_k = rs_karyawan("id")
                    nm = rs_karyawan("nama_karyawan")
                    
               arr_karyawan(a, 0) = id_k
               arr_karyawan(a, 1) = nm
           a = a + 1
           rs_karyawan.MoveNext
           Loop
           grd_karyawan.ReBind
           grd_karyawan.Refresh
                
End Sub

Private Sub pic_karyawan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_karyawan.Visible = False
        Txt_Nama.SetFocus
    End If
End Sub

Private Sub pic_karyawan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub pic_karyawan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   pic_karyawan.Top = pic_karyawan.Top - (yold - Y)
   pic_karyawan.Left = pic_karyawan.Left - (xold - X)
End If

End Sub

Private Sub pic_karyawan_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmd_batal_Click
    End If
End Sub

Private Sub txt_cari_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        normal
        pic_karyawan.Visible = False
        Txt_Nama.SetFocus
    End If
    
    If KeyCode = 13 Then
        normal

        Txt_Nama.Text = arr_karyawan(grd_karyawan.Bookmark, 1)
        pic_karyawan.Visible = False
        Txt_Nama.SetFocus
    End If
End Sub
Private Sub grd_karyawan_DblClick()
    If arr_karyawan.UpperBound(1) > 0 Then
        normal
        Txt_Nama.Text = arr_karyawan(grd_karyawan.Bookmark, 1)
        pic_karyawan.Visible = False
        Txt_Nama.SetFocus
    End If
End Sub
Private Sub grd_karyawan_Click()
    On Error Resume Next
        If arr_karyawan.UpperBound(1) > 0 Then
            id_kr = arr_karyawan(grd_karyawan.Bookmark, 0)
        End If
End Sub

Private Sub grd_karyawan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        grd_karyawan_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        normal
        pic_karyawan.Visible = False
        Txt_Nama.SetFocus
    End If
    
End Sub
Private Sub txt_cari_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo er_k

    Dim sql As String
    Dim rs_karyawan As New ADODB.Recordset
        
        sql = "select id,nama_karyawan from tbl_karyawan"
            
            If txt_cari.Text <> "" Then
                sql = sql & " where nama_karyawan like '%" & Trim(txt_cari.Text) & "%'"
            End If
                
       sql = sql & " order by nama_karyawan"
       rs_karyawan.Open sql, cn, adOpenKeyset
        If Not rs_karyawan.EOF Then
            
            rs_karyawan.MoveLast
            rs_karyawan.MoveFirst
            
            lanjut_karyawan rs_karyawan
        End If
      rs_karyawan.Close
        
      Exit Sub
      
er_k:
      Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub

Private Sub grd_karyawan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_karyawan_Click
End Sub

Private Sub txt_nama_GotFocus()
    Txt_Nama.SelStart = 0
    Txt_Nama.SelLength = Len(Txt_Nama)
End Sub

Private Sub txt_nama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        besar
        Txt_Nama.Text = ""
        txt_cari.Text = ""
        pic_karyawan.Visible = True
        txt_cari.SetFocus
    End If
    
    If KeyCode = 13 Then
        txt_pwd.SetFocus
    End If
    
    If KeyCode = vbKeyEscape Then
        cmd_batal_Click
    End If
        
End Sub
Private Sub txt_nama_LostFocus()

On Error GoTo er_l

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
    If Txt_Nama.Text <> "" Then
        
        sql = "select id_karyawan from qr_user where nama_karyawan='" & Trim(Txt_Nama.Text) & "'"
        rs.Open sql, cn
            If Not rs.EOF Then
                id_kr = rs("id_karyawan")
                
            Else
              If MsgBox("Nama karyawan tidak ditemukan dalam daftar akses program", vbOKOnly + vbInformation, "Pesan") = vbOK Then
                Txt_Nama.SetFocus
              End If
            End If
        rs.Close
    End If
    
    Exit Sub
    
er_l:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub

Private Sub txt_pwd_GotFocus()
    txt_pwd.SelStart = 0
    txt_pwd.SelLength = Len(txt_pwd)
    Cmd_Ok.Default = True
End Sub

Private Sub txt_pwd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmd_batal_Click
    End If
End Sub

Private Sub txt_pwd_LostFocus()
    Cmd_Ok.Default = False
End Sub
