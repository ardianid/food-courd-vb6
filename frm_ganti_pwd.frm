VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Begin VB.Form frm_ganti_pwd 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin TDBContainer3D6Ctl.TDBContainer3D pic_karyawan 
      Height          =   4575
      Left            =   1800
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   4695
      _Version        =   65536
      _ExtentX        =   8281
      _ExtentY        =   8070
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "frm_ganti_pwd.frx":0000
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_ganti_pwd.frx":001C
      Childs          =   "frm_ganti_pwd.frx":00C8
      Begin VB.TextBox txt_cari 
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
         Left            =   2040
         TabIndex        =   22
         Top             =   600
         Width           =   2415
      End
      Begin VB.Frame Frame5 
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   4215
      End
      Begin TrueDBGrid60.TDBGrid grd_karyawan 
         Height          =   3375
         Left            =   240
         OleObjectBlob   =   "frm_ganti_pwd.frx":00E4
         TabIndex        =   23
         Top             =   1080
         Width           =   4215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Nama Karyawan"
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
         TabIndex        =   25
         Top             =   600
         Width           =   1635
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
         TabIndex        =   24
         Top             =   120
         Width           =   1065
      End
   End
   Begin VB.PictureBox pic_karyawan1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   480
      ScaleHeight     =   4665
      ScaleWidth      =   5025
      TabIndex        =   10
      Top             =   6480
      Width           =   5055
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   5025
         TabIndex        =   14
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
            TabIndex        =   13
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.TextBox txt_cari1 
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
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   4815
      End
      Begin TrueDBGrid60.TDBGrid grd_karyawan1 
         Height          =   3375
         Left            =   120
         OleObjectBlob   =   "frm_ganti_pwd.frx":27DA
         TabIndex        =   12
         Top             =   1200
         Width           =   4815
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Nama Karyawan"
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
         TabIndex        =   15
         Top             =   480
         Width           =   4815
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   120
      ScaleHeight     =   5385
      ScaleWidth      =   6585
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   6135
         TabIndex        =   16
         Top             =   240
         Width           =   6135
         Begin VB.TextBox txt_nama 
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
            Left            =   2160
            TabIndex        =   1
            Top             =   120
            Width           =   3855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Karyawan :"
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
            TabIndex        =   17
            Top             =   120
            Width           =   1770
         End
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
         Left            =   5040
         TabIndex        =   5
         Top             =   4680
         Width           =   1335
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   120
         ScaleHeight     =   2775
         ScaleWidth      =   6375
         TabIndex        =   6
         Top             =   1320
         Width           =   6375
         Begin VB.Frame Frame1 
            Height          =   135
            Left            =   240
            TabIndex        =   18
            Top             =   720
            Width           =   5655
         End
         Begin VB.TextBox txt_pwd1 
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
            IMEMode         =   3  'DISABLE
            Left            =   3240
            PasswordChar    =   "*"
            TabIndex        =   4
            Top             =   1560
            Width           =   2655
         End
         Begin VB.TextBox txt_pwd_baru 
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
            IMEMode         =   3  'DISABLE
            Left            =   3240
            PasswordChar    =   "*"
            TabIndex        =   3
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox txt_pwd_lama 
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
            IMEMode         =   3  'DISABLE
            Left            =   2160
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Konfirmasi Password Baru :"
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
            TabIndex        =   9
            Top             =   1560
            Width           =   2745
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password Baru :"
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
            Left            =   1440
            TabIndex        =   8
            Top             =   1080
            Width           =   1635
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password Lama :"
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
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   1725
         End
      End
   End
End
Attribute VB_Name = "frm_ganti_pwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_karyawan As New XArrayDB
Dim id_kr As String
Dim Moving As Boolean
Dim yold, xold As Long

Private Sub normal()
'    Me.Height = 5625
'    Me.Width = 5145
'    Me.ScaleHeight = 5145
'    Me.ScaleWidth = 5055
End Sub

Private Sub besar()
'    Me.Height = 5625
'    Me.Width = 7545
'    Me.ScaleHeight = 5145
'    Me.ScaleWidth = 7455
End Sub


Private Sub kosong_karyawan()
    arr_karyawan.ReDim 0, 0, 0, 0
    grd_karyawan.ReBind
    grd_karyawan.Refresh
End Sub

Private Sub isi_karyawan()

On Error GoTo er_handler

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
       
er_handler:
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

Private Sub Cmd_Simpan_Click()
    Dim sql, sql1, sql2 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset
        
    On Error GoTo er_simpan
        
        If txt_nama.Text = "" Then
            MsgBox ("Nama karyawan harus diisi")
            txt_nama.SetFocus
            Exit Sub
        End If
        
        
        sql = "select id_karyawan from tbl_user where id_karyawan=" & id_kr
        rs.Open sql, cn
            If Not rs.EOF Then
                
                Dim pwd_lama As String
                    pwd_lama = encrypt(txt_pwd_lama.Text)
                
                sql1 = "select pwd from tbl_user where pwd='" & Trim(pwd_lama) & "' and id_karyawan=" & id_kr
                rs1.Open sql1, cn
                    If Not rs1.EOF Then
                        
                        Dim pwd_baru
                            pwd_baru = encrypt(txt_pwd_baru.Text)
                        sql2 = "update tbl_user set pwd='" & Trim(pwd_baru) & "' where id_karyawan=" & id_kr
                        rs2.Open sql2, cn
                        
                    Else
                        MsgBox ("Password yang anda masukkan tidak ditemukan")
                        Unload Me
                        Exit Sub
                    End If
                rs1.Close
            Else
                MsgBox ("Nama karyawan tidak ditemukan")
                Unload Me
                Exit Sub
            End If
        rs.Close
        MsgBox ("Data berhasil disimpan")
        txt_nama.Text = ""
        
        txt_pwd_lama.Text = ""
        txt_pwd_baru.Text = ""
        txt_pwd1.Text = ""
        txt_nama.SetFocus
        Exit Sub
        
er_simpan:
            Dim psn
                psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
                Err.Clear
End Sub

Private Sub cmd_x_Click()
    pic_karyawan.Visible = False
    normal
    txt_nama.SetFocus
End Sub

Private Sub Form_Activate()
    txt_nama.SetFocus
End Sub

Private Sub Form_Load()
    grd_karyawan.Array = arr_karyawan
    
    isi_karyawan
    
    normal
    
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = 250 'Screen.Height / 2 - Me.Height / 2 - 3000
    
End Sub

Private Sub grd_karyawan_Click()
    On Error Resume Next
        If arr_karyawan.UpperBound(1) > 0 Then
            id_kr = arr_karyawan(grd_karyawan.Bookmark, 0)
        End If
End Sub

Private Sub pic_karyawan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        normal
        pic_karyawan.Visible = False
        txt_nama.SetFocus
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

Private Sub txt_cari_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        normal
        pic_karyawan.Visible = False
        txt_nama.SetFocus
    End If
    
    If KeyCode = 13 Then
        normal
        txt_nama.Text = arr_karyawan(grd_karyawan.Bookmark, 1)
        pic_karyawan.Visible = False
        txt_nama.SetFocus
    End If
End Sub
Private Sub grd_karyawan_DblClick()
    If arr_karyawan.UpperBound(1) > 0 Then
        normal
        txt_nama.Text = arr_karyawan(grd_karyawan.Bookmark, 1)
        pic_karyawan.Visible = False
        txt_nama.SetFocus
    End If
End Sub

Private Sub grd_karyawan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        grd_karyawan_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        normal
        pic_karyawan.Visible = False
        txt_nama.SetFocus
    End If
    
End Sub
Private Sub txt_cari_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo er_handler

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

er_handler:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub

Private Sub grd_karyawan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_karyawan_Click
End Sub

Private Sub txt_nama_GotFocus()
    txt_nama.SelStart = 0
    txt_nama.SelLength = Len(txt_nama)
End Sub

Private Sub txt_nama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        besar
        txt_nama.Text = ""
        
        txt_cari.Text = ""
        pic_karyawan.Visible = True
        txt_cari.SetFocus
    End If
    
    If KeyCode = 13 Then
        txt_pwd_lama.SetFocus
    End If
    
End Sub
Private Sub txt_nama_LostFocus()

On Error GoTo er_lo

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
    If txt_nama.Text <> "" Then
        
        sql = "select id_karyawan from qr_user where nama_karyawan='" & Trim(txt_nama.Text) & "'"
        rs.Open sql, cn
            If Not rs.EOF Then
                id_kr = rs("id_karyawan")
                
            Else
                MsgBox ("Nama karyawan tidak ditemukan")
                txt_nama.SetFocus
            End If
        rs.Close
    End If
    
    Exit Sub
    
er_lo:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
            
End Sub

Private Sub txt_pwd_baru_GotFocus()
    txt_pwd_baru.SelStart = 0
    txt_pwd_baru.SelLength = Len(txt_pwd_baru)
End Sub

Private Sub txt_pwd_baru_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_pwd1.SetFocus
    End If
End Sub

Private Sub txt_pwd_lama_GotFocus()
    txt_pwd_lama.SelStart = 0
    txt_pwd_lama.SelLength = Len(txt_pwd_lama)
End Sub

Private Sub txt_pwd_lama_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_pwd_baru.SetFocus
    End If
End Sub

Private Sub txt_pwd1_GotFocus()
    txt_pwd1.SelStart = 0
    txt_pwd1.SelLength = Len(txt_pwd1)
End Sub

Private Sub txt_pwd1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmd_simpan.SetFocus
    End If
End Sub

Private Sub txt_pwd1_LostFocus()
    If txt_pwd1.Text <> "" And txt_pwd_baru.Text <> "" Then
        If Trim(txt_pwd_baru.Text) <> Trim(txt_pwd1.Text) Then
            MsgBox ("Konfirmasi password harus sama dengan password")
            txt_pwd1.SetFocus
        End If
    End If
End Sub
