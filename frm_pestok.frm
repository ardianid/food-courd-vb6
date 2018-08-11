VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Begin VB.Form frm_pestok 
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin TDBContainer3D6Ctl.TDBContainer3D pic_barang 
      Height          =   6255
      Left            =   1680
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   11033
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "frm_pestok.frx":0000
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_pestok.frx":001C
      Childs          =   "frm_pestok.frx":00C8
      Begin VB.TextBox txt 
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
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txt 
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
         Left            =   1920
         TabIndex        =   16
         Top             =   960
         Width           =   3855
      End
      Begin VB.Frame Frame5 
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   5535
      End
      Begin TrueDBGrid60.TDBGrid grd_barang 
         Height          =   4695
         Left            =   240
         OleObjectBlob   =   "frm_pestok.frx":00E4
         TabIndex        =   19
         Top             =   1320
         Width           =   5535
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   600
         Width           =   3855
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
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   2460
      End
   End
   Begin VB.PictureBox j 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   -5640
      ScaleHeight     =   5625
      ScaleWidth      =   5865
      TabIndex        =   8
      Top             =   7320
      Visible         =   0   'False
      Width           =   5895
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   0
         Width           =   5895
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   360
      ScaleHeight     =   8265
      ScaleWidth      =   14505
      TabIndex        =   0
      Top             =   120
      Width           =   14535
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
         Left            =   12960
         TabIndex        =   7
         Top             =   7680
         Width           =   1455
      End
      Begin TrueDBGrid60.TDBGrid grd_daftar 
         Height          =   6495
         Left            =   120
         OleObjectBlob   =   "frm_pestok.frx":2DCC
         TabIndex        =   3
         Top             =   1080
         Width           =   14295
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   14295
         Begin MSComCtl2.DTPicker dtp_tgl 
            Height          =   375
            Left            =   840
            TabIndex        =   11
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
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
            Format          =   20709377
            CurrentDate     =   39211
         End
         Begin VB.TextBox txt_kode 
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
            Left            =   4920
            TabIndex        =   6
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton cmd_tampil 
            Caption         =   "Tampil"
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
            Left            =   12480
            TabIndex        =   2
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label2 
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
            Left            =   3480
            TabIndex        =   5
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label Label1 
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
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "frm_pestok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB
Dim arr_barang As New XArrayDB
Dim sql As String, kode_b As String

Dim Moving As Boolean
Dim yold, xold As Long

Private Sub Cmd_Simpan_Click()

On Error GoTo err_simpan

    Dim sql1, sql2 As String
    Dim rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset
    Dim a As Long
        
    If MsgBox("Yakin semua data yang anda masukkan sudah benar", vbYesNo + vbQuestion, "Pesan") = vbNo Then
        Exit Sub
    End If
        
        
        cn.BeginTrans
        
        grd_daftar.MoveFirst
        
        For a = 1 To arr_daftar.UpperBound(1)
            If (arr_daftar(a, 6) <> 0 And arr_daftar(a, 6) <> Empty) And (arr_daftar(a, 0) <> "" And arr_daftar(a, 0) <> Empty) Then
                        
                sql2 = "select stock_min,stock_max from tbl_barang where id=" & arr_daftar(a, 0)
                rs2.Open sql2, cn
                    If Not rs2.EOF Then
                     Dim jangan
                        If CDbl(arr_daftar(a, 6)) < CDbl(rs2("stock_min")) Then
                            jangan = MsgBox("Stock kode barang " & arr_daftar(a, 3) & " Kurang dari stock minimum" & Chr(13) & "Stock Minimum Barang tersebut " & rs2("stock_min"))
                            cn.RollbackTrans
                            Exit Sub
                            grd_daftar.MoveFirst
                        End If
                        
                        If CDbl(arr_daftar(a, 6)) > CDbl(rs2("stock_max")) Then
                            jangan = MsgBox("Stock kode barang " & arr_daftar(a, 3) & " Lebih dari stock maximum" & Chr(13) & "Stock Maximum Barang tersebut " & rs2("stock_max"))
                            cn.RollbackTrans
                            Exit Sub
                            grd_daftar.MoveFirst
                        End If
                    End If
                rs2.Close
                        
                sql1 = "insert into tr_stock (id_barang,brg_in,brg_out,tgl,ket,nama_user,kemana)"
                sql1 = sql1 & " values  (" & arr_daftar(a, 0) & "," & arr_daftar(a, 7) & "," & arr_daftar(a, 8) & ",'" & Trim(dtp_tgl.Value) & "',1,'" & Trim(utama.lbl_user.Caption) & "','" & arr_daftar(a, 9) & "')"
                rs1.Open sql1, cn
                
                If arr_daftar(a, 7) <> 0 Then
                    sql2 = "select id_barang,jml_stock from tr_jml_stock where id_barang=" & arr_daftar(a, 0)
                    rs2.Open sql2, cn
                        If Not rs2.EOF Then
                            'Dim jml_sekarang As Double
                            'jml_sekarang = CDbl(rs2("jml_stock")) + CDbl(arr_daftar(a, 7))
                            sql1 = "update tr_jml_stock set jml_stock=" & arr_daftar(a, 6) & " where id_barang=" & arr_daftar(a, 0)
                            rs1.Open sql1, cn
                        End If
                    rs2.Close
                End If
                
                If arr_daftar(a, 8) <> 0 Then
                    sql2 = "select id_barang,jml_stock from tr_jml_stock where id_barang=" & arr_daftar(a, 0)
                    rs2.Open sql2, cn
                        If Not rs2.EOF Then
                          '  Dim jml_aja As Double
                           ' jml_aja = CDbl(rs2("jml_stock")) - CDbl(arr_daftar(a, 8))
                            sql1 = "update tr_jml_stock set jml_stock=" & arr_daftar(a, 6) & " where id_barang=" & arr_daftar(a, 0)
                            rs1.Open sql1, cn
                        End If
                    rs2.Close
                End If
            End If
        Next a
        
        MsgBox ("Data berhasil disimpan")
        cn.CommitTrans
        kosong_daftar
        txt_kode.Text = "Semua"
        txt_kode.SetFocus
        Exit Sub
        
err_simpan:
    cn.RollbackTrans
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub

Private Sub Cmd_Tampil_Click()
    isi
End Sub

Private Sub cmd_x_Click()
    pic_barang.Visible = False
    txt_kode.SetFocus
End Sub

Private Sub Form_Load()

grd_daftar.Array = arr_daftar

grd_barang.Array = arr_barang

With pic_barang
    .Left = 5760
    .Top = 600
End With

txt_kode.Text = ""
txt_kode.Text = "Semua"

kosong_daftar

dtp_tgl.Value = Format(Date, "dd/mm/yyyy")

isi_barang

End Sub

Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub

Private Sub kosong_barang()
    arr_barang.ReDim 0, 0, 0, 0
    grd_barang.ReBind
    grd_barang.Refresh
End Sub

Private Sub isi_barang()

On Error GoTo er_handler

    Dim sql1 As String
    Dim rs_barang As New ADODB.Recordset
        
        kosong_barang
        
        sql1 = "select nama_counter,kode,nama_barang from qr_barang where ket=1 order by nama_counter"
        rs_barang.Open sql1, cn, adOpenKeyset
            If Not rs_barang.EOF Then
                
                rs_barang.MoveLast
                rs_barang.MoveFirst
                
                lanjut_barang rs_barang
            End If
        rs_barang.Close
        
        Exit Sub
        
er_handler:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub

Private Sub lanjut_barang(rs_barang As Recordset)
    Dim nama_counter, kode_barang, nama_barang As String
    Dim a As Long
            
            a = 1
                Do While Not rs_barang.EOF
                    arr_barang.ReDim 1, a, 0, 3
                    grd_barang.ReBind
                    grd_barang.Refresh
                        
                        If Not IsNull(rs_barang("nama_counter")) Then
                            nama_counter = rs_barang("nama_counter")
                        Else
                            nama_counter = ""
                        End If
                        
                        If Not IsNull(rs_barang("kode")) Then
                            kode_barang = rs_barang("kode")
                        Else
                            kode_barang = ""
                        End If
                        
                        If Not IsNull(rs_barang("nama_barang")) Then
                            nama_barang = rs_barang("nama_barang")
                        Else
                            nama_barang = ""
                        End If
                        
                     arr_barang(a, 0) = nama_counter
                     arr_barang(a, 1) = kode_barang
                     arr_barang(a, 2) = nama_barang
                     
                     a = a + 1
                     rs_barang.MoveNext
                     Loop
                     grd_barang.ReBind
                     grd_barang.Refresh
End Sub
Private Sub isi()

On Error GoTo er_isi

Dim rs_daftar As New ADODB.Recordset
    
    kosong_daftar
    
    If txt_kode.Text = "" Then
        txt_kode.Text = "Semua"
    End If
    
    If txt_kode.Text = "Semua" Then
                
        sql = "select * from qr_penyesuaian"
        
    ElseIf txt_kode.Text <> "Semua" And txt_kode.Text <> "" Then
        
        sql = "select * from qr_penyesuaian where kode_barang='" & Trim(txt_kode.Text) & "'"
        
    End If
    
    sql = sql & " order by kode_counter"
    rs_daftar.Open sql, cn, adOpenKeyset
        If Not rs_daftar.EOF Then
            
            rs_daftar.MoveLast
            rs_daftar.MoveFirst
            
            isi_daftar rs_daftar
            
        End If
   rs_daftar.Close
            
   Exit Sub
   
er_isi:
   Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub isi_daftar(rs_daftar As Recordset)
    
    Dim id_barang, kode_counter, nama_counter, kode_barang, nama_barang, stock As String
    Dim a As Long
        
        a = 1
            Do While Not rs_daftar.EOF
                
                arr_daftar.ReDim 1, a, 0, 10
                grd_daftar.ReBind
                grd_daftar.Refresh
                    
                    id_barang = rs_daftar("id_barang")
                    
                    If Not IsNull(rs_daftar("kode_counter")) Then
                        kode_counter = rs_daftar("kode_counter")
                    Else
                        kode_counter = ""
                    End If
                    
                    If Not IsNull(rs_daftar("nama_counter")) Then
                        nama_counter = rs_daftar("nama_counter")
                    Else
                        nama_counter = ""
                    End If
                    
                    If Not IsNull(rs_daftar("kode_barang")) Then
                        kode_barang = rs_daftar("kode_barang")
                    Else
                        kode_barang = ""
                    End If
                    
                    If Not IsNull(rs_daftar("nama_barang")) Then
                        nama_barang = rs_daftar("nama_barang")
                    Else
                        nama_barang = ""
                    End If
                    
                    If Not IsNull(rs_daftar("jml_stock")) Then
                        stock = rs_daftar("jml_stock")
                    Else
                        stock = ""
                    End If
                    
               arr_daftar(a, 0) = id_barang
               arr_daftar(a, 1) = kode_counter
               arr_daftar(a, 2) = nama_counter
               arr_daftar(a, 3) = kode_barang
               arr_daftar(a, 4) = nama_barang
               arr_daftar(a, 5) = stock
               arr_daftar(a, 6) = 0
               arr_daftar(a, 7) = 0
               arr_daftar(a, 8) = 0
               arr_daftar(a, 9) = "-"
               
            a = a + 1
            rs_daftar.MoveNext
            Loop
            
            If arr_daftar.UpperBound(1) = 1 Then
                arr_daftar.ReDim 1, arr_daftar.UpperBound(1) + 1, 0, 10
                grd_daftar.ReBind
                grd_daftar.Refresh
                    
               arr_daftar(a, 0) = ""
               arr_daftar(a, 1) = ""
               arr_daftar(a, 2) = ""
               arr_daftar(a, 3) = ""
               arr_daftar(a, 4) = ""
               arr_daftar(a, 5) = ""
               arr_daftar(a, 6) = 0
               arr_daftar(a, 7) = 0
               arr_daftar(a, 8) = 0
               arr_daftar(a, 9) = "-"
            End If
                
            grd_daftar.ReBind
            grd_daftar.Refresh
    
End Sub

Private Sub grd_barang_Click()
    On Error Resume Next
        If arr_barang.UpperBound(1) > 0 Then
            kode_b = arr_barang(grd_barang.Bookmark, 1)
        End If
End Sub

Private Sub grd_barang_DblClick()

If arr_barang.UpperBound(1) > 0 Then
    txt_kode.Text = kode_b
    pic_barang.Visible = False
    txt_kode.SetFocus
End If
    
End Sub

Private Sub grd_barang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        grd_barang_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        pic_barang.Visible = False
    End If
End Sub

Private Sub grd_barang_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_barang_Click
End Sub

Private Sub grd_daftar_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo er_c

    If ColIndex = 6 Then
        
        arr_daftar(grd_daftar.Bookmark, ColIndex) = grd_daftar.Columns(ColIndex).Text
        
        
        If CDbl(arr_daftar(grd_daftar.Bookmark, 5)) > CDbl(arr_daftar(grd_daftar.Bookmark, ColIndex)) Then
            
            arr_daftar(grd_daftar.Bookmark, 7) = 0
        
            arr_daftar(grd_daftar.Bookmark, 8) = CDbl(arr_daftar(grd_daftar.Bookmark, 5)) - CDbl(arr_daftar(grd_daftar.Bookmark, ColIndex))
            
        Else
            
            arr_daftar(grd_daftar.Bookmark, 8) = 0
            
            arr_daftar(grd_daftar.Bookmark, 7) = CDbl(arr_daftar(grd_daftar.Bookmark, ColIndex)) - CDbl(arr_daftar(grd_daftar.Bookmark, 5))
        End If
        Exit Sub
        
    End If
        
    If ColIndex = 9 Then
        arr_daftar(grd_daftar.Bookmark, ColIndex) = grd_daftar.Columns(ColIndex).Text
    End If
    
    
    
    Exit Sub
        
er_c:
    
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub grd_daftar_HeadClick(ByVal ColIndex As Integer)

On Error GoTo er_h

    Dim sql2 As String
    Dim rs_daftar As New ADODB.Recordset
    
    
    If sql = "" Then
        Exit Sub
    End If
    
    
    If arr_daftar.UpperBound(1) = 0 Then
        Exit Sub
    End If
        
    sql2 = ""
    sql2 = sql2 & sql
        
        Select Case ColIndex
            
            Case 2
                   
                sql2 = sql2 & ",nama_counter"
                
            Case 3
                
                sql2 = sql2 & ",kode_barang"
                
            Case 4
                
                sql2 = sql2 & ",nama_barang"
                
            Case 5
                
                sql2 = sql2 & ",jml_stock"
                
        End Select
        
        rs_daftar.Open sql2, cn, adOpenKeyset
            If Not rs_daftar.EOF Then
                
                rs_daftar.MoveLast
                rs_daftar.MoveFirst
                
                isi_daftar rs_daftar
            End If
        rs_daftar.Close
        
    Exit Sub
    
er_h:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub


Private Sub pic_barang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        pic_barang.Visible = False
        txt_kode.Visible = False
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

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_barang.Visible = False
        txt_kode.SetFocus
    End If
    
    If KeyCode = 13 Then
        grd_barang_DblClick
    End If
    
End Sub

Private Sub txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo er_u

    Dim sql1 As String
    Dim rs_barang As New ADODB.Recordset
        
' If arr_barang.UpperBound(1) > 0 Then
 
        kosong_barang
                
        sql1 = "select nama_counter,kode,nama_barang from qr_barang where ket=1"
        
    Select Case Index
        Case 0
            sql1 = sql1 & " and kode like '%" & Trim(txt(0).Text) & "%'"
        Case 1
            sql1 = sql1 & " and nama_barang like '%" & Trim(txt(1).Text) & "%'"
    End Select
        
        sql1 = sql1 & " order by nama_counter"
        rs_barang.Open sql1, cn, adOpenKeyset
            If Not rs_barang.EOF Then
                
                rs_barang.MoveLast
                rs_barang.MoveFirst
                
                lanjut_barang rs_barang
            End If
        rs_barang.Close
        
'End If
      
Exit Sub

er_u:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
                       
End Sub

Private Sub txt_kode_GotFocus()
    txt_kode.SelStart = 0
    txt_kode.SelLength = Len(txt_kode)
End Sub

Private Sub txt_kode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        
        txt_kode.Text = ""
        txt(0).Text = ""
        txt(1).Text = ""
        pic_barang.Visible = True
        txt(0).SetFocus
        
    End If
        
End Sub
