VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{1ABFD380-C196-11D2-B0EA-00A024695830}#1.0#0"; "ticon3d6.ocx"
Begin VB.Form frm_penyesuaian_stock 
   Caption         =   "Form Penyesuaian Inventori"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin TDBContainer3D6Ctl.TDBContainer3D pc1 
      Height          =   6255
      Left            =   -2280
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   11033
      ApplyEffect     =   0
      AutoSize        =   0
      Enabled         =   -1  'True
      Redraw          =   -1  'True
      MouseIcon       =   "frm_penyesuaian_stock.frx":0000
      MousePointer    =   0
      CtrlEffectType  =   8
      CtrlEffectValue =   "Raised"
      ChildsEffectType=   8
      ChildsEffectValue=   "Inset"
      Effects         =   "frm_penyesuaian_stock.frx":001C
      Childs          =   "frm_penyesuaian_stock.frx":00C8
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
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   2055
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
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   21
         Top             =   960
         Width           =   3495
      End
      Begin VB.Frame Frame5 
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   5655
      End
      Begin TrueOleDBGrid60.TDBGrid grd_info_invent 
         Height          =   4575
         Left            =   240
         OleObjectBlob   =   "frm_penyesuaian_stock.frx":00E4
         TabIndex        =   23
         Top             =   1440
         Width           =   5655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENCARIAN DATA INVENTORI"
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
         Index           =   5
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   2685
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Inventori"
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
         Left            =   2400
         TabIndex        =   19
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
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
         Top             =   720
         Width           =   2025
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   240
      ScaleHeight     =   8265
      ScaleWidth      =   14625
      TabIndex        =   3
      Top             =   120
      Width           =   14655
      Begin MSComCtl2.DTPicker dt_tgl 
         Height          =   375
         Left            =   2760
         TabIndex        =   15
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
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
         Format          =   53149697
         CurrentDate     =   39211
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   11400
         TabIndex        =   9
         Top             =   7320
         Width           =   3015
         Begin VB.CommandButton cmd_keluar 
            Caption         =   "&Keluar"
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
            Left            =   1560
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmd_simpan 
            Caption         =   "&Simpan"
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
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.PictureBox hh 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   5775
         Left            =   -5520
         ScaleHeight     =   5745
         ScaleWidth      =   5745
         TabIndex        =   6
         Top             =   1920
         Visible         =   0   'False
         Width           =   5775
         Begin VB.CommandButton cmd_x 
            Caption         =   "X"
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
            Left            =   5400
            TabIndex        =   8
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Inventori"
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
            Index           =   4
            Left            =   2040
            TabIndex        =   14
            Top             =   480
            Width           =   3570
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Kode"
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
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   1770
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            Caption         =   "Data Info Inventori"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   375
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   5775
         End
      End
      Begin VB.CommandButton cmd_tampil 
         Caption         =   "&Tampil"
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
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin TrueOleDBGrid60.TDBGrid grd_invent 
         Height          =   6255
         Left            =   240
         OleObjectBlob   =   "frm_penyesuaian_stock.frx":2928
         TabIndex        =   2
         Top             =   1080
         Width           =   14175
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
         Left            =   6840
         TabIndex        =   0
         Top             =   600
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   14400
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penyesuaian Inventori"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   120
         Width           =   2565
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Inventori"
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
         Index           =   1
         Left            =   5160
         TabIndex        =   5
         Top             =   600
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Sekarang"
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
         Left            =   720
         TabIndex        =   4
         Top             =   600
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frm_penyesuaian_stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arr_invent As New XArrayDB
Dim sql_info_invent As String
Dim rs_info_invent As New ADODB.Recordset

Dim Moving As Boolean
Dim yold, xold As Long


Private Sub cmd_keluar_Click()
    Unload Me
End Sub

Private Sub cmd_simpan_Click()
    Dim sql1, sql2 As String
    Dim rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset
    Dim a As Long
    
    On Error GoTo eror_simpan
    
    If MsgBox("Apakah yakin data sudah valid??", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    cn.BeginTrans
        For a = 1 To arr_invent.UpperBound(1)
            If arr_invent(a, 3) <> 0 And arr_invent(a, 3) <> Empty Then
                If arr_invent(a, 4) <> 0 Then
                    sql1 = "insert into tbl_tr_inventori (id_invent,invent_in,tgl_tr,ket,nama_user)"
                    sql1 = sql1 & " values  (" & arr_invent(a, 6) & "," & Val(arr_invent(a, 4)) & "," & _
                        "'" & Format(Date, "dd/mm/yyyy") & "','-','" & Trim(utama.lbl_user.Caption) & "')"
                    rs1.Open sql1, cn
                End If
                
                If arr_invent(a, 5) <> 0 Then
                    sql1 = "insert into tbl_tr_inventori (id_invent,invent_out,tgl_tr,ket,nama_user)"
                    sql1 = sql1 & " values  (" & arr_invent(a, 6) & "," & Val(arr_invent(a, 5)) & "," & _
                        "'" & Format(Date, "dd/mm/yyyy") & "','-','" & Trim(utama.lbl_user.Caption) & "')"
                    rs1.Open sql1, cn
                End If

                If arr_invent(a, 3) <> 0 Then
                    sql2 = "update tbl_stock_invent set stock_invent=" & arr_invent(a, 3) & "  where id_invent= " & arr_invent(a, 6)
                    rs2.Open sql2, cn
                End If
            End If
        Next a
        
        MsgBox ("Data berhasil disimpan")
        cn.CommitTrans
        kosong_invent
    Exit Sub
    
        
eror_simpan:
    cn.RollbackTrans
    Dim konfirm
        konfirm = MsgBox(Err.Number & Chr(13) & Err.Description, vbExclamation + vbOKOnly, "Error Program")
        Err.Clear

            
                    
                
End Sub

Private Sub Cmd_Tampil_Click()
    arr_invent.ReDim 0, 0, 0, 0
    grd_invent.ReBind
    grd_invent.Refresh
    
    grd_invent.Array = arr_invent
    isi_invent
End Sub

Private Sub cmd_tmp_invent_Click()
    pc1.Visible = True
End Sub

Private Sub cmd_x_Click()
    pc1.Visible = False
    txt_kode.SetFocus
End Sub

Private Sub Form_Load()
    '================================================================================'
    'mengatur tampilan layar form
    '================================================================================'
    'Me.Height = 9060
    'Me.Width = 10920
    'Me.Left = (utama.Width - frm_penyesuaian_stock.Width) / 2
    'Me.Top = (utama.Height - utama.Height) / 4
       
    '========================================'
    'isi grid inventdengan data invent
    '======================================='
    
    With pc1
        .Left = 7560
        .Top = 720
    End With
    
'        sql_info_invent = "select  * from tbl_inventori order by kode_invent"
'        Set rs_info_invent = cn.Execute(sql_info_invent)
'        Set grd_info_invent.DataSource = rs_info_invent
'        grd_info_invent.ReBind
'        grd_info_invent.Refresh

    
    dt_tgl.Value = Format(Date, "dd/mm/yyyy")
    
    grd_invent.Array = arr_invent
    isi_invent 'pemaggilan procedur isi invent
    
    
    
    
End Sub

Private Sub isi_invent()

On Error GoTo er_i

    Dim sql1 As String
    Dim rs_invent As New ADODB.Recordset
        
    '==========================================='
    'isi dari procedur isi invent menampilkan data berdasarkan jika text kosong maka
    'ditampilkan semua if ada maka ditampilkan berdasarkan data kode invent
    '================================================================================'
    
        kosong_invent
        If txt_kode.Text = "" Then
            sql1 = "select kode_invent,nama_invent,stock_invent,id_invent from qr_sesuai  order by kode_invent"
           Else
            sql1 = "select kode_invent,nama_invent,stock_invent,id_invent from qr_sesuai  where kode_invent='" & txt_kode.Text & "' order by kode_invent"
        End If
        rs_invent.Open sql1, cn, adOpenKeyset
            If Not rs_invent.EOF Then
                
                rs_invent.MoveLast
                rs_invent.MoveFirst
                
                lanjut_invent rs_invent
            End If
        rs_invent.Close
        Exit Sub
        
er_i:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub lanjut_invent(rs_invent As Recordset)
    Dim kode_invent, nama_invent, stock_invent, id_invent As String
    Dim a As Long
            
    '=============================================================================='
    'melkukan pengamblan data inventori dengan fungsi array'
    '=============================================================================='
            a = 1
                Do While Not rs_invent.EOF
                    arr_invent.ReDim 1, a, 0, 6
                    grd_invent.ReBind
                    grd_invent.Refresh
                        
                        If Not IsNull(rs_invent("kode_invent")) Then
                            kode_invent = rs_invent("kode_invent")
                        Else
                            kode_invent = ""
                        End If
                        
                        If Not IsNull(rs_invent("nama_invent")) Then
                            nama_invent = rs_invent("nama_invent")
                        Else
                            nama_invent = ""
                        End If
                        
                        If Not IsNull(rs_invent("stock_invent")) Then
                            stock_invent = rs_invent("stock_invent")
                        Else
                            stock_invent = ""
                        End If
                        
                        If Not IsNull(rs_invent("id_invent")) Then
                            id_invent = rs_invent("id_invent")
                        Else
                            id_invent = ""
                        End If
                     arr_invent(a, 0) = kode_invent
                     arr_invent(a, 1) = nama_invent
                     arr_invent(a, 2) = stock_invent
                     arr_invent(a, 3) = 0
                     arr_invent(a, 4) = 0
                     arr_invent(a, 5) = 0
                     arr_invent(a, 6) = id_invent
                     
                     a = a + 1
                     rs_invent.MoveNext
                     Loop
                     grd_invent.ReBind
                     grd_invent.Refresh
End Sub

Public Sub kosong_invent()
    arr_invent.ReDim 0, 0, 0, 0
    grd_invent.ReBind
    grd_invent.Refresh
    
End Sub


Private Sub grd_info_invent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If rs_info_invent.RecordCount <> 0 Then
            txt_kode.Text = rs_info_invent.Fields("kode_invent")
            pc1.Visible = False
            txt_kode.SetFocus
           Else
            Exit Sub
        End If
    End If
       
    If KeyCode = vbKeyEscape Then
        pc1.Visible = False
        txt_kode.SetFocus
    End If
    
End Sub

Private Sub grd_invent_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    
    
    '==============================================================================='
    'melakukan perubahan data manipulasi penyesuaian stock, apabila stock akhir
    'bertambah maka terjadi pemasukan stock if stock akhir berkurang maka terjadi pengeluaran inventori
    '======================================================================================'
    
    On Error GoTo er_c
    If ColIndex = 3 Then
        
        arr_invent(grd_invent.Bookmark, ColIndex) = grd_invent.Columns(ColIndex).Text
                
        
        If CDbl(arr_invent(grd_invent.Bookmark, ColIndex)) > CDbl(arr_invent(grd_invent.Bookmark, 2)) Then
             grd_invent.Columns(4).Text = CDbl(arr_invent(grd_invent.Bookmark, ColIndex)) - CDbl(arr_invent(grd_invent.Bookmark, 2))
             arr_invent(grd_invent.Bookmark, 4) = grd_invent.Columns(4).Text
             grd_invent.Columns(5).Text = 0
             arr_invent(grd_invent.Bookmark, 5) = grd_invent.Columns(5).Text
            Else
             grd_invent.Columns(5).Text = CDbl(arr_invent(grd_invent.Bookmark, 2)) - CDbl(arr_invent(grd_invent.Bookmark, ColIndex))
             arr_invent(grd_invent.Bookmark, 5) = grd_invent.Columns(5).Text
             grd_invent.Columns(4).Text = 0
             arr_invent(grd_invent.Bookmark, 4) = grd_invent.Columns(4).Text
        End If
        Exit Sub

        
    End If
        
er_c:
    
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub pc1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
        pc1.Visible = False
        txt_kode.SetFocus
    End If
End Sub

Private Sub pc1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = True
If Moving = True Then
   yold = Y
   xold = X
End If
End Sub

Private Sub pc1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Moving = True Then
   pc1.Top = pc1.Top - (yold - Y)
   pc1.Left = pc1.Left - (xold - X)
End If

End Sub

Private Sub pc1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
        pc1.Visible = False
        txt_kode.SetFocus
    End If
    
    If KeyCode = 13 Then
        grd_info_invent.SetFocus
    End If
    
    
End Sub

Private Sub txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    '================================================================================='
    'melakukan pencarian berdasarkan data text perketikan huruf
    '================================================================================='
    
    On Error GoTo er_k
    
       sql_info_invent = "select * from tbl_inventori"
        
  If txt(0).Text <> "" Or txt(1).Text <> "" Then
  
    sql_info_invent = sql_info_invent & " where"
  
        Select Case Index
            Case 0
                sql_info_invent = sql_info_invent & " kode_invent like '%" & Trim(txt(0).Text) & "%'"
            Case 1
                sql_info_invent = sql_info_invent & " nama_invent like '%" & Trim(txt(1).Text) & "%'"
        End Select
  End If
            sql_info_invent = sql_info_invent & " order by kode_invent"
            Set rs_info_invent = cn.Execute(sql_info_invent)
            Set grd_info_invent.DataSource = rs_info_invent
            grd_info_invent.ReBind
            grd_info_invent.Refresh
            
   Exit Sub
   
er_k:
   Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub

Private Sub txt_kode_GotFocus()
    txt_kode.SelStart = 0
    txt_kode.SelLength = Len(txt_kode)
End Sub

Private Sub txt_kode_KeyDown(KeyCode As Integer, Shift As Integer)
    
On Error GoTo er_dow
    
    If KeyCode = vbKeyF3 Then
        sql_info_invent = "select  * from tbl_inventori order by kode_invent"
        Set rs_info_invent = cn.Execute(sql_info_invent)
        Set grd_info_invent.DataSource = rs_info_invent
        grd_info_invent.ReBind
        grd_info_invent.Refresh
        txt(0).Text = ""
        txt(1).Text = ""
        pc1.Visible = True
        txt(0).SetFocus
    End If
    
    If KeyCode = 13 Then
        arr_invent.ReDim 0, 0, 0, 0
        grd_invent.ReBind
        grd_invent.Refresh
        
        grd_invent.Array = arr_invent
        isi_invent
    End If
    
    Exit Sub
    
er_dow:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

