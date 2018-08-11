VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_cfaktur 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   13710
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   120
      ScaleHeight     =   6705
      ScaleWidth      =   13425
      TabIndex        =   0
      Top             =   120
      Width           =   13455
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   10440
         ScaleHeight     =   825
         ScaleWidth      =   2745
         TabIndex        =   12
         Top             =   120
         Width           =   2775
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cetak Faktur = F2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   720
            TabIndex        =   15
            Top             =   480
            Width           =   1860
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   405
         End
         Begin VB.Label lbl_tgl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_tgl"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   720
            TabIndex        =   13
            Top             =   120
            Width           =   645
         End
      End
      Begin VB.CommandButton cmd_cetak 
         Caption         =   "Cetak Faktur"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11880
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
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
         Height          =   615
         Left            =   10440
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin TrueDBGrid60.TDBGrid grd_daftar 
         Height          =   4815
         Left            =   120
         OleObjectBlob   =   "frm_cfaktur.frx":0000
         TabIndex        =   7
         Top             =   1800
         Width           =   13095
      End
      Begin MSMask.MaskEdBox msk_tgl 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_faktur 
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
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
      Begin MSMask.MaskEdBox msk_jam 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_jam2 
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
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
         Left            =   2640
         TabIndex        =   11
         Top             =   1200
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam"
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
         TabIndex        =   10
         Top             =   1200
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Faktur"
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
         TabIndex        =   9
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label2 
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
         TabIndex        =   8
         Top             =   720
         Width           =   375
      End
   End
End
Attribute VB_Name = "frm_cfaktur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB


Private Sub cmd_cetak_Click()

On Error GoTo er_print

Cmd_Tampil_Click

            noff = Trim(txt_faktur.Text)
            
            byyr = 0
            kemm = 0
            
            htu = False
            
            Load Frm_Lap_BuktiByar
                Frm_Lap_BuktiByar.Show
            
            On Error GoTo 0
            Exit Sub
            
'Dim a As Long
'Dim grs
'    Printer.Font = "Arial"
'
'
'        Printer.CurrentX = 0
'        Printer.CurrentY = 0
'
'            Printer.Print
'            Printer.Print Tab((55 / 2 - Len("Bukti Pembayaran") / 2) - 25); "Bukti Pembayaran"
'            Printer.Print Tab((55 / 2 - Len("Simpur Food Center") / 2) - 25); "Simpur Food Center"
'            Printer.Print
'
'            grs = String$(71, "-")
'
'            Printer.Print grs
'
'            Printer.Print "Tgl. " & lbl_tgl.Caption; Tab(21); "Jam. " & frm_jual.lbl_jam.Caption
'            Printer.Print grs
'            Printer.Print "No Faktur " & txt_faktur.Text
'            Printer.Print
'            Printer.Print "Qty"; Tab(5); "Nama Barang"; Tab(23); "Harga"; Tab(35); "Disc"; Tab(41); "Charge"
'            Printer.Print grs
'            Printer.Print grs
'
'       For a = 1 To arr_daftar.UpperBound(1)
'
'            Printer.Print arr_daftar(a, 4); Tab(5); arr_daftar(a, 3); Tab(23); Format(arr_daftar(a, 5), "currency"); Tab(35); arr_daftar(a, 6); Tab(41); arr_daftar(a, 7)
'
'      Next a
'
'            Printer.Print grs
'            Printer.Print "Total Discount"; Tab(21); grd_daftar.Columns(6).FooterText
'            Printer.Print "Total Charge"; Tab(21); grd_daftar.Columns(7).FooterText
'            Printer.Print grs
'            Printer.Print "Total"; Tab(25); grd_daftar.Columns(8).FooterText
'            Printer.Print grs
'            Printer.Print "jml Bayar "; Tab(25); 0 'Format(Space(Len(txt_jml_bayar.Text) - Len(txt_jml_bayar.Text)) + txt_jml_bayar.Text, "Currency")
'            Printer.Print "Kembali"; Tab(25); 0 'Space(Len(lbl_kembali.Caption) - Len(lbl_kembali.Caption)) + lbl_kembali.Caption
'            Printer.Print
'            Printer.Print " *********** Terima Kasih ***********"
'
'
'      Printer.EndDoc
'      Exit Sub


er_print:
           Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Information")
            Err.Clear
            

End Sub

Private Sub cmd_cetak_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Cmd_Tampil_Click()

On Error GoTo er_tampil

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        sql = "select tgl,jam,no_faktur,qty,nama_barang,harga_satuan,disc,cash,total_harga from qr_semua_penjualan where ket=0"
            
            If txt_faktur.Text <> "" Or msk_tgl.Text <> "__/__/____" Or msk_jam.Text <> "__:__:__" Or msk_jam2.Text <> "__:__:__" Then
                
                    If txt_faktur.Text <> "" Then
                        sql = sql & " and no_faktur='" & Trim(txt_faktur.Text) & "'"
                    End If
                    
                    If msk_tgl.Text <> "__/__/____" And txt_faktur.Text = "" Then
                        sql = sql & " and tgl = datevalue('" & Trim(msk_tgl.Text) & "')"
                    End If
                    
                    If msk_tgl.Text <> "__/__/____" And txt_faktur.Text <> "" Then
                        sql = sql & " and tgl = datevalue('" & Trim(msk_tgl.Text) & "')"
                    End If
                    
                    If msk_jam.Text <> "__:__:__" And txt_faktur.Text = "" And msk_tgl.Text = "__/__/____" And msk_jam2.Text = "__:__:__" Then
                        sql = sql & " and jam =TimeValue  ('" & Trim(msk_jam.Text) & "')"
                    End If
                    
                    If msk_jam.Text <> "__:__:__" And msk_jam2.Text = "__:__:__" And (txt_faktur.Text <> "" Or msk_tgl.Text <> "__/__/____") Then
                        sql = sql & " and jam = timevalue('" & Trim(msk_jam.Text) & "')"
                    End If
                    
                    If msk_jam2.Text <> "__:__:__" And txt_faktur.Text = "" And msk_tgl.Text = "__/__/____" And msk_jam.Text = "__:__:__" Then
                        sql = sql & " and jam =TimeValue  ('" & Trim(msk_jam2.Text) & "')"
                    End If
                    
                    If msk_jam2.Text <> "__:__:__" And msk_jam.Text = "__:__:__" And (txt_faktur.Text <> "" Or msk_tgl.Text <> "__/__/____") Then
                        sql = sql & " and jam = timevalue('" & Trim(msk_jam2.Text) & "')"
                    End If
                    
                    If msk_jam.Text <> "__:__:__" And msk_jam2.Text <> "__:__:__" Then
                        
                         If txt_faktur.Text = "" And msk_tgl.Text = "__/__/____" Then
                            sql = sql & " jam >= timevalue('" & Trim(msk_jam.Text) & "') and jam <= timevalue('" & Trim(msk_jam2.Text) & "')"
                         End If
                         
                         If txt_faktur.Text <> "" Or msk_tgl.Text <> "__/__/____" Then
                            sql = sql & " and jam >= timevalue('" & Trim(msk_jam.Text) & "') and jam <= timevalue('" & Trim(msk_jam2.Text) & "')"
                         End If
                    End If
                    
            End If
        sql = sql & " order by tgl,no_faktur"
        rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
                
                lanjut rs
            End If
        rs.Close
        
        Exit Sub
        
er_tampil:
            Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub lanjut(rs As Recordset)
    Dim tgl, jam, faktur, qty, nama_barang, harga, disc, cash, total As String
    Dim a As Long
    Dim jml_qty, jml_harga, jml_disc, jml_cash, jml_total As Double
        
        a = 1
        jml_qty = 0
        jml_harga = 0
        jml_disc = 0
        jml_cash = 0
        jml_total = 0
            Do While Not rs.EOF
                
                arr_daftar.ReDim 1, a, 0, 10
                grd_daftar.ReBind
                grd_daftar.Refresh
                    
                    If Not IsNull(rs("tgl")) Then
                        tgl = rs("tgl")
                    Else
                        tgl = ""
                    End If
                    
                    If Not IsNull(rs("jam")) Then
                        jam = rs("jam")
                    Else
                        jam = ""
                    End If
                    
                    If Not IsNull(rs("no_faktur")) Then
                        faktur = rs("no_faktur")
                    Else
                        faktur = ""
                    End If
                    
                    If Not IsNull(rs("qty")) Then
                        qty = rs("qty")
                    Else
                        qty = ""
                    End If
                    
                    If Not IsNull(rs("nama_barang")) Then
                        nama_barang = rs("nama_barang")
                    Else
                        nama_barang = ""
                    End If
                    
                    If Not IsNull(rs("harga_satuan")) Then
                        harga = rs("harga_satuan")
                    Else
                        harga = 0
                    End If
                    
                    If Not IsNull(rs("disc")) Then
                        disc = rs("disc")
                    Else
                        disc = 0
                    End If
                    
                    If Not IsNull(rs("cash")) Then
                        cash = rs("cash")
                    Else
                        cash = 0
                    End If
                    
                    If Not IsNull(rs("total_harga")) Then
                        total = rs("total_harga")
                    Else
                        total = 0
                    End If
                    
                    jml_qty = jml_qty + CDbl(qty)
                    jml_harga = jml_harga + CDbl(harga)
                    
                        If Right(disc, 1) = "%" Then
                            jml_disc = jml_disc + CDbl(Mid(disc, 1, Len(disc) - 1))
                        Else
                            jml_disc = jml_disc + CDbl(disc)
                        End If
                        
                        If Right(cash, 1) = "%" Then
                            jml_cash = jml_cash + CDbl(Mid(cash, 1, Len(cash) - 1))
                        Else
                            jml_cash = jml_cash + CDbl(cash)
                        End If
                        
                    jml_total = jml_total + CDbl(total)
                    
                arr_daftar(a, 0) = tgl
                arr_daftar(a, 1) = faktur
                arr_daftar(a, 2) = jam
                arr_daftar(a, 3) = nama_barang
                arr_daftar(a, 4) = qty
                arr_daftar(a, 5) = harga
                arr_daftar(a, 6) = disc
                arr_daftar(a, 7) = cash
                arr_daftar(a, 8) = total
             a = a + 1
             rs.MoveNext
             Loop
             
             grd_daftar.Columns(3).FooterText = "Total"
             grd_daftar.Columns(4).FooterText = jml_qty
             grd_daftar.Columns(5).FooterText = Format(jml_harga, "currency")
             grd_daftar.Columns(6).FooterText = jml_disc & "%"
             grd_daftar.Columns(7).FooterText = jml_cash & "%"
             grd_daftar.Columns(8).FooterText = Format(jml_total, "currency")
             
             grd_daftar.ReBind
             grd_daftar.Refresh
             
             grd_daftar.MoveFirst
             grd_daftar_Click
             
End Sub

Private Sub cmd_tampil_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    txt_faktur.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    If KeyCode = vbKeyF2 Then
        Cmd_Tampil_Click
        Cmd_Tampil_Click
        cmd_cetak_Click
    End If
End Sub

Private Sub Form_Load()

    grd_daftar.Array = arr_daftar
    
    lbl_tgl.Caption = Date
    
    kosong
    
    Me.Left = Screen.Width \ 2 - Me.Width \ 2
    Me.Top = Screen.Height \ 2 - Me.Height \ 2
    
End Sub

Private Sub kosong()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub

Private Sub grd_daftar_Click()
    On Error Resume Next
        If arr_daftar.UpperBound(1) > 0 Then
            txt_faktur.Text = arr_daftar(grd_daftar.Bookmark, 1)
        End If
End Sub

Private Sub grd_daftar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    If KeyCode = vbKeyF2 Then
        Cmd_Tampil_Click
        Cmd_Tampil_Click
        cmd_cetak_Click
    End If
    
End Sub

Private Sub grd_daftar_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_daftar_Click
End Sub

Private Sub msk_jam_GotFocus()
    msk_jam.SelStart = 0
    msk_jam.SelLength = Len(msk_jam)
End Sub

Private Sub msk_jam_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If msk_jam <> "__:__:__" Then
            Cmd_Tampil_Click
        End If
            msk_jam2.SetFocus
    End If
    
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    If KeyCode = vbKeyF2 Then
        Cmd_Tampil_Click
        Cmd_Tampil_Click
        cmd_cetak_Click
    End If
End Sub

Private Sub msk_jam2_GotFocus()
    msk_jam2.SelStart = 0
    msk_jam2.SelLength = Len(msk_jam2)
End Sub

Private Sub msk_jam2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Cmd_Tampil_Click
    End If
    
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    If KeyCode = vbKeyF2 Then
        Cmd_Tampil_Click
        Cmd_Tampil_Click
        cmd_cetak_Click
    End If
End Sub

Private Sub msk_tgl_GotFocus()
    msk_tgl.SelStart = 0
    msk_tgl.SelLength = Len(msk_tgl)
End Sub

Private Sub msk_tgl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If msk_tgl <> "__/__/____" Then
            Cmd_Tampil_Click
        End If
        msk_jam.SetFocus
    End If
    
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    If KeyCode = vbKeyF2 Then
        Cmd_Tampil_Click
        Cmd_Tampil_Click
        cmd_cetak_Click
    End If
    
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    If KeyCode = vbKeyF2 Then
        Cmd_Tampil_Click
        Cmd_Tampil_Click
        cmd_cetak_Click
    End If
    
End Sub


Private Sub txt_faktur_GotFocus()
    txt_faktur.SelStart = 0
    txt_faktur.SelLength = Len(txt_faktur)
End Sub

Private Sub txt_faktur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
     If txt_faktur.Text <> "" Then
        Cmd_Tampil_Click
     End If
        msk_tgl.SetFocus
    End If
    
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    If KeyCode = vbKeyF2 Then
        Cmd_Tampil_Click
        Cmd_Tampil_Click
        cmd_cetak_Click
    End If
    
End Sub
