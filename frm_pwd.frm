VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Begin VB.Form frm_pwd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   Begin TDBContainer3D6Ctl.TDBContainer3D pic_karyawan 
      Height          =   4575
      Left            =   2160
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   8070
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "frm_pwd.frx":0000
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_pwd.frx":001C
      Childs          =   "frm_pwd.frx":00C8
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
         TabIndex        =   13
         Top             =   600
         Width           =   3015
      End
      Begin VB.Frame Frame5 
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   4815
      End
      Begin TrueDBGrid60.TDBGrid grd_karyawan 
         Height          =   3375
         Left            =   240
         OleObjectBlob   =   "frm_pwd.frx":00E4
         TabIndex        =   15
         Top             =   1080
         Width           =   4815
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
         TabIndex        =   14
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
         TabIndex        =   12
         Top             =   120
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   120
      ScaleHeight     =   4905
      ScaleWidth      =   8025
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.CheckBox cek_aktif 
         Caption         =   "Aktif"
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
         Left            =   240
         TabIndex        =   8
         Top             =   3960
         Width           =   735
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
         Left            =   6480
         TabIndex        =   4
         Top             =   4200
         Width           =   1215
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
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txt_pwd 
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
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txt_karyawan 
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
         Left            =   3120
         TabIndex        =   1
         Top             =   360
         Width           =   3495
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   7680
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Konfirmasi Password"
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
         Left            =   840
         TabIndex        =   7
         Top             =   1560
         Width           =   1995
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   840
         TabIndex        =   6
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   0
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   1620
      End
   End
   Begin VB.CommandButton cmd_hak 
      Caption         =   "Atur Hak Akses"
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
      Left            =   6000
      TabIndex        =   9
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "frm_pwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_karyawan As New XArrayDB
Dim id_kr As String
Dim Moving As Boolean
Dim yold, xold As Long
Private Sub form_semula()
'    Me.Height = 3375
'    Me.Width = 8400
'    Me.ScaleHeight = 2895
'    Me.ScaleWidth = 8310
End Sub

Private Sub ubah_besar()
'    Me.Height = 5655
'    Me.Width = 8400
'    Me.ScaleHeight = 5175
'    Me.ScaleWidth = 8310
End Sub

Private Sub Cmd_Simpan_Click()
    Dim sql As String
    Dim rs As New ADODB.Recordset
        
    On Error GoTo er_simpan
        
        If txt_karyawan.Text = "" Or txt_pwd.Text = "" Or txt_pwd1.Text = "" Then
            MsgBox ("Semua data harus diisi")
            Exit Sub
        End If
        
        Dim aktif As Integer
            If cek_aktif.Value = vbChecked Then
                aktif = 1
            Else
                aktif = 0
            End If
            
      Dim sql1 As String
      Dim rs1 As New ADODB.Recordset
            
      sql1 = "select id_karyawan from tbl_user where id_karyawan=" & id_kr
      rs1.Open sql1, cn
      If rs1.EOF Then
        
        Dim pwd_karyawan As String
            pwd_karyawan = encrypt(txt_pwd.Text)
            
        sql = "insert into tbl_user (id_karyawan,pwd,aktif) values(" & id_kr & ",'" & Trim(pwd_karyawan) & "'," & aktif & ")"
        rs.Open sql, cn
      Else
        MsgBox ("Karyawan yang anda masukkan sudah terdaftar sebagai pengguna program")
        txt_karyawan.SetFocus
        Exit Sub
      End If
      rs1.Close
        
        MsgBox ("Data berhasil disimpan")
        txt_karyawan.Text = ""
        txt_pwd.Text = ""
        txt_pwd1.Text = ""
        txt_karyawan.SetFocus
        Exit Sub
            
er_simpan:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub cmd_x_Click()
    form_semula
    pic_karyawan.Visible = False
    txt_karyawan.SetFocus
End Sub

Private Sub Form_Activate()
    txt_karyawan.SetFocus
End Sub

Private Sub Form_Load()
    
    grd_karyawan.Array = arr_karyawan
    
    isi_karyawan
'    form_semula
    
    cek_aktif.Value = vbChecked
    
    Me.Left = utama.Width / 2 - Me.Width / 2
    Me.Top = 450
    
    
End Sub

Private Sub kosong_karyawan()
    arr_karyawan.ReDim 0, 0, 0, 0
    grd_karyawan.ReBind
    grd_karyawan.Refresh
End Sub

Private Sub isi_karyawan()

On Error GoTo isi_karyawan

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
       
isi_karyawan:
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

Private Sub grd_karyawan_Click()
    On Error Resume Next
        If arr_karyawan.UpperBound(1) > 0 Then
            id_kr = arr_karyawan(grd_karyawan.Bookmark, 0)
        End If
End Sub

Private Sub grd_karyawan_DblClick()
    If arr_karyawan.UpperBound(1) > 0 Then
        form_semula
        txt_karyawan.Text = arr_karyawan(grd_karyawan.Bookmark, 1)
        pic_karyawan.Visible = False
        txt_karyawan.SetFocus
    End If
End Sub

Private Sub grd_karyawan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        grd_karyawan_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        form_semula
        pic_karyawan.Visible = False
        txt_karyawan.SetFocus
    End If
    
End Sub

Private Sub grd_karyawan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_karyawan_Click
End Sub

Private Sub pic_karyawan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        form_semula
        pic_karyawan.Visible = False
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
        form_semula
        pic_karyawan.Visible = False
        txt_karyawan.SetFocus
    End If
    
    If KeyCode = 13 Then
        form_semula
        txt_karyawan.Text = arr_karyawan(grd_karyawan.Bookmark, 1)
        pic_karyawan.Visible = False
        txt_karyawan.SetFocus
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

Private Sub txt_karyawan_GotFocus()
    txt_karyawan.SelStart = 0
    txt_karyawan.SelLength = Len(txt_karyawan)
End Sub

Private Sub txt_karyawan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txt_karyawan.Text = ""
'        ubah_besar
        txt_cari.Text = ""
        pic_karyawan.Visible = True
        txt_cari.SetFocus
    End If
End Sub

Private Sub txt_karyawan_LostFocus()

On Error GoTo er_h

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
    If txt_karyawan.Text <> "" Then
        
        sql = "select id from tbl_karyawan where nama_karyawan='" & Trim(txt_karyawan.Text) & "'"
        rs.Open sql, cn
            If Not rs.EOF Then
                id_kr = rs("id")
            Else
                MsgBox ("Nama karyawan tidak ditemukan")
                txt_karyawan.SetFocus
            End If
        rs.Close
    End If
    
    Exit Sub
    
er_h:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub


Private Sub txt_pwd_GotFocus()
    txt_pwd.SelStart = 0
    txt_pwd.SelLength = Len(txt_pwd)
End Sub

Private Sub txt_pwd1_GotFocus()
   txt_pwd1.SelStart = 0
   txt_pwd1.SelLength = Len(txt_pwd1)
End Sub

Private Sub txt_pwd1_LostFocus()
    If txt_pwd1.Text <> "" And txt_pwd.Text <> "" Then
        If Trim(txt_pwd.Text) <> Trim(txt_pwd1.Text) Then
            MsgBox ("Konfirmasi password harus sama dengan password")
            txt_pwd1.SetFocus
            Exit Sub
        End If
    End If
End Sub
