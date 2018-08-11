VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frm_biling 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   LinkTopic       =   "frm_bilig"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic_counter 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   1440
      ScaleHeight     =   5865
      ScaleWidth      =   5625
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   5655
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
         TabIndex        =   14
         Top             =   0
         Width           =   375
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
         TabIndex        =   12
         Top             =   840
         Width           =   2895
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
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   2415
      End
      Begin TrueDBGrid60.TDBGrid grd_counter 
         Height          =   4455
         Left            =   120
         OleObjectBlob   =   "frm_biling.frx":0000
         TabIndex        =   13
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
         TabIndex        =   15
         Top             =   0
         Width           =   5655
      End
      Begin VB.Label Label7 
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
         TabIndex        =   17
         Top             =   480
         Width           =   3015
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
         TabIndex        =   16
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3105
      ScaleWidth      =   6105
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton cmd_simpan 
         Caption         =   "Simpan"
         Height          =   495
         Left            =   4680
         TabIndex        =   4
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txt_harga_air 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txt_harga_listrik 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1920
         Width           =   1935
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
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Air /M3"
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
         Left            =   3240
         TabIndex        =   9
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   5880
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Listrik /Kwh"
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
         Top             =   1560
         Width           =   1905
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
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label2 
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
         TabIndex        =   6
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label1 
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
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frm_biling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_counter As New XArrayDB
Dim id_cntr As String

Private Sub cmd_simpan_Click()
    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
        
On Error GoTo ers

        If txt_kode.Text = "" Then
            MsgBox ("Kode counter hrs diisi")
            Exit Sub
        End If
        
        sql = "select id_counter from tbl_biling where id_counter=" & id_cntr
        rs.Open sql, cn
            If rs.EOF Then
                
                sql1 = "insert into tbl_biling (id_counter,harga_listrik,harga_air)"
                sql1 = sql1 & " values (" & id_cntr & "," & CCur(txt_harga_listrik.Text) & "," & CCur(txt_harga_air.Text) & ")"
                rs1.Open sql1, cn
                
                MsgBox ("Data berhasil disimpan")
                txt_kode.Text = ""
                lbl_counter.Caption = ""
                txt_harga_air.Text = ""
                txt_harga_listrik.Text = ""
                txt_kode.SetFocus
                
            Else
                MsgBox ("Data biling air dan listrik counter sudah ada")
                txt_kode.SetFocus
            End If
        rs.Close
        Exit Sub
        
ers:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub

Private Sub cmd_x_Click()
    pic_counter.Visible = False
    benar (True)
    txt_kode.SetFocus
End Sub

Private Sub Form_Activate()
    txt_kode.SetFocus
End Sub

Private Sub Form_Load()
    grd_counter.Array = arr_counter
    
    isi_counter
    
    txt_harga_listrik.Text = 0
    txt_harga_air.Text = 0
    
    benar (True)
    
    Me.Left = utama.Width / 2 - Me.Width / 2
    Me.Top = utama.Height / 2 - Me.Height / 2 - 2750
    
End Sub

Private Sub kosong_counter()
    arr_counter.ReDim 0, 0, 0, 0
    grd_counter.ReBind
    grd_counter.Refresh
End Sub


Private Sub isi_counter()

On Error GoTo er_isi_counter

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
        
er_isi_counter:
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

Private Sub pic_counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pic_counter.Visible = False
        benar (True)
        txt_kode.SetFocus
    End If
End Sub

Private Sub txt_harga_air_GotFocus()
    txt_harga_air.SelStart = 0
    txt_harga_air.SelLength = Len(txt_harga_air)
End Sub

Private Sub txt_harga_air_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
    Beep
    KeyAscii = 0
End If
End Sub

Private Sub txt_harga_air_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If txt_harga_air.Text <> "" Then
       txt_harga_air.Text = Format(txt_harga_air.Text, "###,###,###")
       txt_harga_air.SelStart = Len(txt_harga_air.Text)
    End If
    
End Sub

Private Sub txt_harga_air_LostFocus()
If txt_harga_air.Text = "" Then
    txt_harga_air.Text = 0
End If
End Sub

Private Sub txt_harga_listrik_GotFocus()
    txt_harga_listrik.SelStart = 0
    txt_harga_listrik.SelLength = Len(txt_harga_listrik)
End Sub

Private Sub txt_harga_listrik_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
    Beep
    KeyAscii = 0
End If
End Sub

Private Sub txt_harga_listrik_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If txt_harga_listrik.Text <> "" Then
        txt_harga_listrik.Text = Format(txt_harga_listrik.Text, "###,###,###")
        txt_harga_listrik.SelStart = Len(txt_harga_listrik.Text)
    End If
    
End Sub

Private Sub txt_harga_listrik_LostFocus()
    If txt_harga_listrik.Text = "" Then
        txt_harga_listrik.Text = 0
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
        benar (False)
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
        benar (True)
        txt_kode.SetFocus
    End If
    
    If KeyCode = 13 Then
        txt_kode.Text = arr_counter(grd_counter.Bookmark, 1)
        lbl_counter.Caption = arr_counter(grd_counter.Bookmark, 2)
        pic_counter.Visible = False
        benar (True)
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
    benar (True)
    txt_kode.SetFocus
End If
    
End Sub

Private Sub grd_counter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        grd_counter_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        pic_counter.Visible = False
        txt_kode.SetFocus
    End If
End Sub

Private Sub grd_counter_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_counter_Click
End Sub

Private Sub benar(param As Boolean)
    Select Case param
        Case True
            Me.Height = 3855
            Me.Width = 6480
            Me.ScaleHeight = 3375
            Me.ScaleWidth = 6390
        Case False
            Me.Height = 6840
            Me.Width = 8250
            Me.ScaleHeight = 6360
            Me.ScaleWidth = 8160
    End Select
End Sub

Private Sub txt_kode_LostFocus()

On Error GoTo er_kode

If txt_kode.Text <> "" Then

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        sql = "select id,kode,nama_counter from tbl_counter where kode='" & Trim(txt_kode.Text) & "'"
        rs.Open sql, cn
            If Not rs.EOF Then
                id_cntr = rs("id")
                lbl_counter.Caption = rs("nama_counter")
            Else
                MsgBox ("Kode counter yang anda masukkan tidak ditemukan")
                txt_kode.SetFocus
            End If
        rs.Close
        
End If

Exit Sub

er_kode:
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear

End Sub
