VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_slip_penggajian 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport cr 
      Left            =   720
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox pic_karyawan 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   2640
      ScaleHeight     =   5745
      ScaleWidth      =   5385
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   5415
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   5385
         TabIndex        =   10
         Top             =   0
         Width           =   5415
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
            TabIndex        =   11
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.TextBox txt_cari 
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
         TabIndex        =   9
         Top             =   840
         Width           =   5175
      End
      Begin TrueDBGrid60.TDBGrid grd_karyawan 
         Height          =   4335
         Left            =   120
         OleObjectBlob   =   "frm_slip_penggajian.frx":0000
         TabIndex        =   12
         Top             =   1320
         Width           =   5175
      End
      Begin VB.Label Label5 
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
         TabIndex        =   13
         Top             =   480
         Width           =   5175
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2745
      ScaleWidth      =   7905
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton cmd_tampil 
         Caption         =   "Tampil"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6120
         TabIndex        =   7
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txt_karyawan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         TabIndex        =   6
         Top             =   1320
         Width           =   5535
      End
      Begin VB.ComboBox cbo_bulan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txt_thn 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   7680
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   7680
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Karyawan"
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
         TabIndex        =   5
         Top             =   1320
         Width           =   1725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
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
         TabIndex        =   3
         Top             =   840
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thn"
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
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
   End
End
Attribute VB_Name = "frm_slip_penggajian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_karyawan As New XArrayDB
Dim id_kr As String

Private Sub form_penuh()
    Me.Height = 3495
    Me.Width = 8280
    Me.ScaleHeight = 3015
    Me.ScaleWidth = 8190
End Sub

Private Sub besar()
   Me.Height = 6450
   Me.Width = 8280
   Me.ScaleHeight = 5970
   Me.ScaleWidth = 8190
End Sub

Private Sub kosong_karyawan()
    arr_karyawan.ReDim 0, 0, 0, 0
    grd_karyawan.ReBind
    grd_karyawan.Refresh
End Sub

Private Sub isi_karyawan()

On Error GoTo isi

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
       
isi:
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

Private Sub cmd_tampil_Click()

On Error GoTo er_printing

Dim sql As String
Dim rs As New ADODB.Recordset
    
    If txt_thn.Text = "" Then
        MsgBox ("Tahun hrs diisi")
        txt_thn.SetFocus
        Exit Sub
    End If
    
    If cbo_bulan.Text = "" Then
        MsgBox ("Bulan harus dipilih")
        cbo_bulan.SetFocus
        Exit Sub
    End If
    
    If txt_karyawan.Text = "" Then
        MsgBox ("Nama karyawan harus diisi")
        txt_karyawan.SetFocus
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    utama.MousePointer = vbHourglass
    
    sql = "select nama_karyawan,bulan,thn,gaji_pokok,tunjangan,lain_lain,potongan,gaji_diterima from qr_gaji"
    sql = sql & " where nama_karyawan='" & Trim(txt_karyawan.Text) & "' and bulan=" & bulan(cbo_bulan.Text) & " and thn=" & Trim(txt_thn.Text)
    
    rs.Open sql, cn
        
    cr.ReportFileName = path_lap & "\lap_penggajian.rpt"
    cr.Connect = cn
    cr.Formulas(2) = "pemakai='" & id_user & "'"
    cr.RetrieveSQLQuery
    cr.SQLQuery = sql
    cr.DiscardSavedData = True
    cr.Action = 1
    
    If Me.MousePointer = vbHourglass Then
        Me.MousePointer = vbDefault
        utama.MousePointer = vbDefault
    End If
    Exit Sub
    
er_printing:
        
        If Me.MousePointer = vbHourglass Then
            Me.MousePointer = vbDefault
            utama.MousePointer = vbDefault
        End If
        
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub

Private Sub cmd_x_Click()
    pic_karyawan.Visible = False
    txt_karyawan.SetFocus
End Sub

Private Sub Form_Load()

    grd_karyawan.Array = arr_karyawan
    
    isi_karyawan
    
    isi_combo
    
    txt_thn.Text = Year(Now)
    
    form_penuh
    
    
    
End Sub

Private Sub isi_combo()
    With cbo_bulan
         .AddItem "Januari"
         .AddItem "Februari"
         .AddItem "Maret"
         .AddItem "April"
         .AddItem "Mei"
         .AddItem "Juni"
         .AddItem "Juli"
         .AddItem "Agustus"
         .AddItem "September"
         .AddItem "Oktober"
         .AddItem "Nopember"
         .AddItem "Desember"
    End With
End Sub

Private Sub grd_karyawan_Click()
    On Error Resume Next
        If arr_karyawan.UpperBound(1) > 0 Then
            id_kr = arr_karyawan(grd_karyawan.Bookmark, 0)
        End If
End Sub

Private Sub grd_karyawan_DblClick()
    If arr_karyawan.UpperBound(1) > 0 Then
        form_penuh
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
        form_penuh
        pic_karyawan.Visible = False
        txt_karyawan.SetFocus
    End If
    
End Sub

Private Sub grd_karyawan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_karyawan_Click
End Sub

Private Sub pic_karyawan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        form_penuh
        pic_karyawan.Visible = False
        txt_karyawan.SetFocus
    End If
End Sub

Private Sub txt_cari_GotFocus()
    txt_cari.SelStart = 0
    txt_cari.SelLength = Len(txt_cari)
End Sub

Private Sub txt_cari_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        form_penuh
        pic_karyawan.Visible = False
        txt_karyawan.SetFocus
    End If
    
    If KeyCode = 13 Then
        form_penuh
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
        txt_cari.Text = ""
        besar
        pic_karyawan.Visible = True
        txt_cari.SetFocus
    End If
End Sub
