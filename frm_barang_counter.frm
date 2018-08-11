VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frm_barang_counter 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic_counter 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   2160
      ScaleHeight     =   5865
      ScaleWidth      =   4785
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txt_nm 
         Height          =   405
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   4575
      End
      Begin TrueDBGrid60.TDBGrid grd_counter 
         Height          =   4335
         Left            =   120
         OleObjectBlob   =   "frm_barang_counter.frx":0000
         TabIndex        =   24
         Top             =   1440
         Width           =   4575
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   4785
         TabIndex        =   22
         Top             =   0
         Width           =   4815
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Nama Counter"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   4575
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   825
      ScaleWidth      =   9225
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.CommandButton cmd_tampil 
         Caption         =   "Tampil"
         Height          =   615
         Left            =   6960
         TabIndex        =   2
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox txt_nama_counter 
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
      Begin VB.Line Line1 
         X1              =   6480
         X2              =   6480
         Y1              =   0
         Y2              =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Counter"
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.PictureBox pic_tambah 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2865
      ScaleWidth      =   9225
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   9255
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   7680
         TabIndex        =   19
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmd_ok 
         Caption         =   "Ok"
         Height          =   495
         Left            =   6240
         TabIndex        =   18
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txt_stock 
         Height          =   375
         Left            =   1920
         TabIndex        =   17
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txt_harga 
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
         Left            =   1920
         TabIndex        =   16
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txt_nama_barang 
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox txt_kode_barang 
         Height          =   375
         Left            =   1920
         TabIndex        =   13
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Awal"
         Height          =   195
         Left            =   600
         TabIndex        =   20
         Top             =   2160
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Jual"
         Height          =   195
         Left            =   600
         TabIndex        =   15
         Top             =   1680
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         Height          =   195
         Left            =   600
         TabIndex        =   12
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         Height          =   195
         Left            =   600
         TabIndex        =   11
         Top             =   720
         Width           =   930
      End
      Begin VB.Line Line2 
         X1              =   360
         X2              =   9000
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Barang"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   1485
      End
   End
   Begin VB.PictureBox pic_barang 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   120
      ScaleHeight     =   5625
      ScaleWidth      =   9225
      TabIndex        =   7
      Top             =   1080
      Width           =   9255
      Begin VB.CommandButton cmd_tambah 
         Caption         =   "Tambah"
         Height          =   495
         Left            =   4920
         TabIndex        =   4
         Top             =   5040
         Width           =   1335
      End
      Begin VB.CommandButton cmd_edit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   6360
         TabIndex        =   5
         Top             =   5040
         Width           =   1335
      End
      Begin VB.CommandButton cmd_hapus 
         Caption         =   "Hapus"
         Height          =   495
         Left            =   7800
         TabIndex        =   6
         Top             =   5040
         Width           =   1335
      End
      Begin TrueDBGrid60.TDBGrid grd_barang 
         Height          =   4815
         Left            =   120
         OleObjectBlob   =   "frm_barang_counter.frx":2CE9
         TabIndex        =   3
         Top             =   120
         Width           =   9015
      End
   End
End
Attribute VB_Name = "frm_barang_COUNTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_counter As New XArrayDB
Dim arr_barang As New XArrayDB
Dim id_cntr As String, sql_barang As String, id_brg As String
Dim arah_simpan As Boolean

Private Sub besar_form()
    Me.Height = 7350
    Me.Width = 9585
    Me.ScaleHeight = 6870
    Me.ScaleWidth = 9495
    pic_tambah.Visible = False
    pic_barang.Visible = True
End Sub

Private Sub setengah_form()
    Me.Height = 4530
    Me.Width = 9585
    Me.ScaleHeight = 4050
    Me.ScaleWidth = 9495
    pic_barang.Visible = False
    pic_tambah.Visible = True
End Sub

Private Sub kosong_counter()
    arr_counter.ReDim 0, 0, 0, 0
    grd_counter.ReBind
    grd_counter.Refresh
End Sub

Private Sub kosong_barang()
    arr_barang.ReDim 0, 0, 0, 0
    grd_barang.ReBind
    grd_barang.Refresh
End Sub

Private Sub isi_counter()
    Dim rs_counter As New ADODB.Recordset
    Dim sql As String
        
        kosong_counter
        
        sql = "select id,nama_counter from tbl_counter"
        rs_counter.Open sql, cn, adOpenKeyset
            If Not rs_counter.EOF Then
                
                rs_counter.MoveLast
                rs_counter.MoveFirst
                    
                    lanjut_counter rs_counter
            End If
        rs_counter.Close
        
End Sub

Private Sub lanjut_counter(rs_counter As Recordset)
    Dim i_c, n_c As String
    Dim a As Long
        
        a = 1
            Do While Not rs_counter.EOF
                
                arr_counter.ReDim 1, a, 0, 3
                grd_counter.ReBind
                grd_counter.Refresh
                    
                    i_c = rs_counter("id")
                    n_c = rs_counter("nama_counter")
                    
                arr_counter(a, 0) = i_c
                arr_counter(a, 1) = a
                arr_counter(a, 2) = n_c
                
            a = a + 1
            rs_counter.MoveNext
            Loop
            grd_counter.ReBind
            grd_counter.Refresh
                    
End Sub

Private Sub cmd_cancel_Click()
    besar_form
    Picture1.Enabled = True
    cmd_tampil_Click
End Sub

Private Sub cmd_edit_Click()
    If arr_barang.UpperBound(1) > 0 Then
        arah_simpan = False
        setengah_form
        Picture1.Enabled = False
        txt_kode_barang.Text = arr_barang(grd_barang.Bookmark, 2)
        txt_nama_barang.Text = arr_barang(grd_barang.Bookmark, 3)
        txt_harga.Text = arr_barang(grd_barang.Bookmark, 4)
        txt_stock.Text = arr_barang(grd_barang.Bookmark, 5)
        txt_kode_barang.SetFocus
    End If
End Sub

Private Sub cmd_hapus_Click()

On Error GoTo er_h:

Dim sql, sql1 As String
Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    
If arr_barang.UpperBound(1) > 0 Then
    If MsgBox("Yakin akan hapus barang " & arr_barang(grd_barang.Bookmark, 2), vbYesNo + vbQuestion, "Pesan") = vbYes Then
        
        sql = "select id from tbl_barang where id=" & id_brg
        rs.Open sql, cn
            If Not rs.EOF Then
                
                sql1 = "delete from tbl_barang where id=" & id_brg
                rs1.Open sql1, cn
                
            Else
                MsgBox ("Data yang akan dihapus tidak ditemukan")
            End If
            
        rs.Close
        cmd_tampil_Click
        Exit Sub
    Else
        Exit Sub
    End If
End If

Exit Sub
    
er_h:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
    
End Sub

Private Sub cmd_ok_Click()

On Error GoTo er_s

Dim sql, sql1 As String
Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    
If txt_kode_barang.Text = "" Then
    MsgBox ("Kode barang harus diisi")
    Exit Sub
End If

If txt_nama_barang.Text = "" Then
    MsgBox ("Nama barang harus diisi")
    Exit Sub
End If

If MsgBox("Yakin semua data yang dimasukkan sudah benar.....?", vbYesNo + vbQuestion, "Pesan") = vbNo Then
    Exit Sub
End If

If arah_simpan = True Then
    
    sql = "select kode from tbl_barang where kode='" & Trim(txt_kode_barang.Text) & "'"
    rs.Open sql, cn
        If Not rs.EOF Then
            If MsgBox("Kode barang " & txt_kode_barang.Text & " sudah ada,anda yakin akan menyimpannya", vbYesNo + vbQuestion, "Pesan") = vbYes Then
                simpan_aja
            Else
                Exit Sub
            End If
        Else
            simpan_aja
        End If
    rs.Close
 MsgBox ("Data berhasil_dismpan")
 arah_simpan = True
 kosong_text
 txt_kode_barang.SetFocus
 Exit Sub
 
 ElseIf arah_simpan = False Then
    
    sql = "select id from tbl_barang where id=" & id_brg
    rs.Open sql, cn
        If Not rs.EOF Then
        
            sql1 = "update tbl_barang set kode='" & Trim(txt_kode_barang.Text) & "',nama_barang='" & Trim(txt_nama_barang.Text) & "',harga_jual=" & Trim(txt_harga.Text) & ",stock_awal=" & Trim(txt_stock.Text) & " where id=" & id_brg
            rs1.Open sql1, cn
            
        Else
            
            MsgBox ("Data yang akan diedit tidak ditemukan")
            
        End If
    rs.Close
    cmd_cancel_Click
    Exit Sub
    
End If
    
er_s:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub simpan_aja()
    Dim sql As String
    Dim rs As New ADODB.Recordset
     
     sql = "insert into tbl_barang (id_counter,kode,nama_barang,harga_jual,stock_awal)"
     sql = sql & " values(" & id_cntr & ",'" & Trim(txt_kode_barang.Text) & "','" & Trim(txt_nama_barang.Text) & "'," & Trim(txt_harga.Text) & "," & Trim(txt_stock.Text) & ")"
     rs.Open sql, cn
     
End Sub

Private Sub cmd_tambah_Click()
    arah_simpan = True
    setengah_form
    kosong_text
    Picture1.Enabled = False
    txt_kode_barang.SetFocus
End Sub

Private Sub kosong_text()
    txt_nama_barang.Text = ""
    txt_kode_barang.Text = ""
    txt_harga.Text = ""
    txt_stock.Text = ""
End Sub

Private Sub cmd_tampil_Click()
    If txt_nama_counter.Text <> "" And id_cntr <> "" Then
        
        Dim rs_barang As New ADODB.Recordset
            
            kosong_barang
            
            sql_barang = "select id,kode,nama_barang,harga_jual,stock_awal from tbl_barang where id_counter=" & id_cntr
            rs_barang.Open sql_barang, cn, adOpenKeyset
                If Not rs_barang.EOF Then
                    
                    rs_barang.MoveLast
                    rs_barang.MoveFirst
                        
                    lanjut_barang rs_barang
                End If
             rs_barang.Close
    End If
           
End Sub

Private Sub lanjut_barang(rs_barang As Recordset)
    
    Dim id_b, kd_b, nm_b, hg_jl, stk As String
    Dim a As Long
        
        a = 1
            Do While Not rs_barang.EOF
                arr_barang.ReDim 1, a, 0, 6
                grd_barang.ReBind
                grd_barang.Refresh
                    
                    id_b = rs_barang("id")
                    
                    If Not IsNull(rs_barang("kode")) Then
                        kd_b = rs_barang("kode")
                    Else
                        kd_b = ""
                    End If
                    
                    If Not IsNull(rs_barang("nama_barang")) Then
                        nm_b = rs_barang("nama_barang")
                    Else
                        nm_b = ""
                    End If
                    
                    If Not IsNull(rs_barang("harga_jual")) Then
                        hg_jl = rs_barang("harga_jual")
                    Else
                        hg_jl = ""
                    End If
                    
                    If Not IsNull(rs_barang("stock_awal")) Then
                        stk = rs_barang("stock_awal")
                    Else
                        stk = 0
                    End If
                    
                arr_barang(a, 0) = id_b
                arr_barang(a, 1) = a
                arr_barang(a, 2) = kd_b
                arr_barang(a, 3) = nm_b
                arr_barang(a, 4) = hg_jl
                arr_barang(a, 5) = stk
                
            a = a + 1
            rs_barang.MoveNext
            
            Loop
            
            grd_barang.ReBind
            grd_barang.Refresh
        
    
End Sub

Private Sub cmd_x_Click()
    pic_counter.Visible = False
End Sub



Private Sub Form_Load()

    grd_counter.Array = arr_counter
    
    grd_barang.Array = arr_barang
    
    isi_counter
    
    kosong_barang
    
    besar_form
    
End Sub

Private Sub grd_barang_Click()
On Error Resume Next
    If arr_barang.UpperBound(1) > 0 Then
        id_brg = arr_barang(grd_barang.Bookmark, 0)
    End If
End Sub

Private Sub grd_barang_HeadClick(ByVal ColIndex As Integer)
    
If arr_barang.UpperBound(1) > 0 Then

    Dim rs_barang As New ADODB.Recordset
    Dim sql As String
    
    If sql_barang = "" Then
        Exit Sub
    End If
    
    sql = sql_barang
    
    Select Case ColIndex
        
        Case 2
            
            sql = sql & " order by kode"
            
        Case 3
            
            sql = sql & " order by nama_barang"
            
        Case 4
            
            sql = sql & " order by harga_jual"
            
        Case 5
            
            sql = sql & " order by stock_awal"
            
    End Select
    
    rs_barang.Open sql, cn
    If Not rs_barang.EOF Then
        
        rs_barang.MoveLast
        rs_barang.MoveFirst
        
        lanjut_barang rs_barang
    End If
    rs_barang.Close
    
End If

End Sub

Private Sub grd_barang_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_barang_Click
End Sub

Private Sub grd_counter_Click()
    On Error Resume Next
        If arr_counter.UpperBound(1) > 0 Then
            id_cntr = arr_counter(grd_counter.Bookmark, 0)
        End If
End Sub

Private Sub grd_counter_DblClick()
If arr_counter.UpperBound(1) > 0 Then
    
    txt_nama_counter.Text = arr_counter(grd_counter.Bookmark, 2)
    pic_counter.Visible = False
    
End If
    
End Sub

Private Sub grd_counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        grd_counter_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        pic_counter.Visible = False
    End If
End Sub

Private Sub grd_counter_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_counter_Click
End Sub

Private Sub pic_counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_counter.Visible = False
    End If
End Sub

Private Sub txt_harga_GotFocus()
    txt_harga.SelStart = 0
    txt_harga.SelLength = Len(txt_harga)
End Sub

Private Sub txt_harga_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txt_kode_barang_GotFocus()
    txt_kode_barang.SelStart = 0
    txt_kode_barang.SelLength = Len(txt_kode_barang)
End Sub

Private Sub txt_nama_barang_GotFocus()
    txt_nama_barang.SelStart = 0
    txt_nama_barang.SelLength = Len(txt_nama_barang)
End Sub

Private Sub txt_nama_counter_GotFocus()
    txt_nama_counter.SelStart = 0
    txt_nama_counter.SelLength = Len(txt_nama_counter)
End Sub

Private Sub txt_nama_counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txt_nama_counter.Text = ""
        pic_counter.Visible = True
        txt_nm.Text = ""
        txt_nm.SetFocus
    End If
End Sub

Private Sub txt_nama_counter_LostFocus()

If txt_nama_counter.Text <> "" Then

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        sql = "select id from tbl_counter where nama_counter='" & Trim(txt_nama_counter.Text) & "'"
        rs.Open sql, cn
            If Not rs.EOF Then
                id_cntr = rs("id")
            Else
                MsgBox ("Nama counter yang anda masukkan tidak ditemukan")
                txt_nama_counter.SetFocus
            End If
        rs.Close
        
End If

End Sub

Private Sub txt_nm_GotFocus()
    txt_nm.SelStart = 0
    txt_nm.SelLength = Len(txt_nm)
End Sub

Private Sub txt_nm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_counter.Visible = False
    End If
End Sub

Private Sub txt_nm_KeyUp(KeyCode As Integer, Shift As Integer)
        
        Dim sql As String
        Dim rs_counter As New ADODB.Recordset
            
            sql = "select id,nama_counter from tbl_counter where nama_counter like '%" & Trim(txt_nm.Text) & "%'"
            rs_counter.Open sql, cn, adOpenKeyset
                If Not rs_counter.EOF Then
                    
                    rs_counter.MoveLast
                    rs_counter.MoveFirst
                    
                    lanjut_counter rs_counter
                End If
            rs_counter.Close
End Sub
Private Sub txt_stock_GotFocus()
    txt_stock.SelStart = 0
    txt_stock.SelLength = Len(txt_stock)
End Sub

Private Sub txt_stock_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub
