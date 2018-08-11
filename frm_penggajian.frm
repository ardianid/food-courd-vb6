VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_penggajian 
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   120
      ScaleHeight     =   8265
      ScaleWidth      =   14985
      TabIndex        =   0
      Top             =   240
      Width           =   15015
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2385
         ScaleWidth      =   2745
         TabIndex        =   6
         Top             =   5640
         Width           =   2775
         Begin VB.TextBox txt_nama 
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
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   2535
         End
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
            Height          =   495
            Left            =   120
            TabIndex        =   8
            Top             =   1800
            Width           =   1215
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
            Left            =   1440
            TabIndex        =   7
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label2 
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
            Left            =   480
            TabIndex        =   9
            Top             =   120
            Width           =   1725
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7815
         Left            =   3000
         ScaleHeight     =   7785
         ScaleWidth      =   11865
         TabIndex        =   5
         Top             =   240
         Width           =   11895
         Begin TrueDBGrid60.TDBGrid grd_daftar 
            Height          =   7575
            Left            =   120
            OleObjectBlob   =   "frm_penggajian.frx":0000
            TabIndex        =   11
            Top             =   120
            Width           =   11655
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   4815
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2775
         Begin VB.ListBox lst_bulan 
            BackColor       =   &H00400000&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   4200
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   2535
         End
      End
      Begin MSComCtl2.DTPicker dtp_tgl 
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   8421504
         CalendarForeColor=   16777215
         Format          =   19726337
         CurrentDate     =   38639
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "frm_penggajian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB

Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub


Private Sub cmd_simpan_Click()
    
On Error GoTo er_simpan
    
    If arr_daftar.UpperBound(1) = 0 Then
        Exit Sub
    End If
    
    If MsgBox("Yakin semua data yang dimasukkan sudah benar", vbYesNo + vbQuestion, "Pesan") = vbNo Then
        Exit Sub
    End If
    
    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    Dim a As Long
    Dim pakai
    Dim thn, bln
    
    cn.BeginTrans
    
    thn = Year(Now)
    
    bln = bulan(lst_bulan.Text)
    
    For a = 1 To arr_daftar.UpperBound(1)
        pakai = arr_daftar(a, 8)
            
            If pakai <> 0 Then
                
                sql = "select id from tbl_gaji where bulan=" & bln & " and id_karyawan=" & arr_daftar(a, 0) & " and thn=" & thn
                rs.Open sql, cn
                    If rs.EOF Then
                        
                        sql1 = "insert into tbl_gaji (id_karyawan,bulan,thn,tgl,gaji_pokok,tunjangan,lain_lain,potongan,gaji_diterima,nama_user)"
                        sql1 = sql1 & " values(" & arr_daftar(a, 0) & "," & bln & "," & thn & ",'" & Trim(dtp_tgl.Value) & "'," & CCur(arr_daftar(a, 3)) & "," & CCur(arr_daftar(a, 4)) & "," & CCur(arr_daftar(a, 5)) & "," & CCur(arr_daftar(a, 6)) & "," & CCur(arr_daftar(a, 7)) & ",'" & Trim(utama.lbl_user.Caption) & "')"
                        rs1.Open sql1, cn
                        
                    Else
                        Dim j
                        j = MsgBox("Gaji karyawan : " & arr_daftar(a, 2) & " Pada bulan : " & lst_bulan.Text & Chr(13) & " sudah ada,data batal disimpan")
                        cn.RollbackTrans
                        Exit Sub
                        
                    End If
                rs.Close
                
            End If
    Next a
    
    MsgBox ("Data berhasil disimpan")
    cn.CommitTrans
    kosong_daftar
    Exit Sub
    
er_simpan:
            cn.RollbackTrans
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub cmd_tampil_Click()
    
    If txt_nama.Text = "" Then
        txt_nama.Text = "Semua"
    End If
    
    isi_daftar
End Sub

Private Sub Form_Load()

grd_daftar.Array = arr_daftar

kosong_daftar

dtp_tgl.Value = Format(Date, "dd/mm/yyyy")

isi_lst
lst_bulan.ListIndex = 0

txt_nama.Text = "Semua"

End Sub

Private Sub isi_lst()
    With lst_bulan
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

Private Sub isi_daftar()

On Error GoTo er_daftar

    Dim sql As String
    Dim rsa As New ADODB.Recordset
        
        kosong_daftar
        
        sql = "select id,nama_karyawan,gaji_pokok from tbl_karyawan"
        
        If txt_nama.Text <> "Semua" Then
            sql = sql & " where nama_karyawan='" & Trim(txt_nama.Text) & "'"
        End If
      
        sql = sql & " order by nama_karyawan"
        
        rsa.Open sql, cn, adOpenKeyset
            If Not rsa.EOF Then
                
                rsa.MoveLast
                rsa.MoveFirst
                    
                    lanjut_isi rsa
            End If
        rsa.Close
            
        Exit Sub
        
er_daftar:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
            
End Sub

Private Sub lanjut_isi(rsa As Recordset)
    Dim id_k, nm, pk As String
    Dim a As Long
        
        a = 1
            Do While Not rsa.EOF
                arr_daftar.ReDim 1, a, 0, 9
                grd_daftar.ReBind
                grd_daftar.Refresh
                    
                    id_k = rsa("id")
                    If Not IsNull(rsa("nama_karyawan")) Then
                        nm = rsa("nama_karyawan")
                    Else
                        nm = ""
                    End If
                    
                    If Not IsNull(rsa("gaji_pokok")) Then
                        pk = rsa("gaji_pokok")
                    Else
                        pk = 0
                    End If
                    
            arr_daftar(a, 0) = id_k
            arr_daftar(a, 1) = a
            arr_daftar(a, 2) = nm
            If pk <> 0 Then
                arr_daftar(a, 3) = Format(pk, "###,###,###")
            Else
                arr_daftar(a, 3) = pk
            End If
            arr_daftar(a, 4) = 0
            arr_daftar(a, 5) = 0
            arr_daftar(a, 6) = 0
            
            If pk <> 0 Then
                arr_daftar(a, 7) = Format(pk, "###,###,###")
            Else
                arr_daftar(a, 7) = pk
            End If
            
            arr_daftar(a, 8) = vbChecked
            
            a = a + 1
            rsa.MoveNext
            Loop
            grd_daftar.ReBind
            grd_daftar.Refresh
            grd_daftar.MoveFirst
            
End Sub


Private Sub grd_daftar_AfterColUpdate(ByVal ColIndex As Integer)

On Error GoTo er_data

If ColIndex = 4 Then
    
        arr_daftar(grd_daftar.Bookmark, ColIndex) = grd_daftar.Columns(ColIndex).Text
                
    End If
    
    If ColIndex = 5 Then
    
        arr_daftar(grd_daftar.Bookmark, ColIndex) = grd_daftar.Columns(ColIndex).Text
        
    End If
   
   
   If ColIndex = 6 Then
        
        arr_daftar(grd_daftar.Bookmark, ColIndex) = grd_daftar.Columns(ColIndex).Text
        
   End If
   
   If ColIndex = 8 Then
        arr_daftar(grd_daftar.Bookmark, ColIndex) = grd_daftar.Columns(ColIndex).Text
   End If
   
   Exit Sub
   
er_data:
    
    Dim bukan_tipe
        bukan_tipe = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
End Sub

Private Sub grd_daftar_Change()
    
On Error GoTo er_c
    
    If grd_daftar.Columns(4).Text = "" Then
        grd_daftar.Columns(4).Text = 0
        arr_daftar(grd_daftar.Bookmark, 4) = grd_daftar.Columns(4).Text
        
        Dim ss01
        ss01 = (CDbl(arr_daftar(grd_daftar.Bookmark, 3)) + CDbl(arr_daftar(grd_daftar.Bookmark, 4)) + CDbl(arr_daftar(grd_daftar.Bookmark, 5))) - CDbl(arr_daftar(grd_daftar.Bookmark, 6))
            grd_daftar.Columns(7).Text = Format(ss01, "###,###,###")
            arr_daftar(grd_daftar.Bookmark, 7) = grd_daftar.Columns(7).Text
        
        Exit Sub
    ElseIf grd_daftar.Columns(4).Text <> "" And grd_daftar.Columns(4).Text <> 0 Then
        grd_daftar.Columns(4).Text = Format(grd_daftar.Columns(4).Text, "##,##,###")
        
        arr_daftar(grd_daftar.Bookmark, 4) = grd_daftar.Columns(4).Text
        
        Dim ss1
        ss1 = (CDbl(arr_daftar(grd_daftar.Bookmark, 3)) + CDbl(arr_daftar(grd_daftar.Bookmark, 4)) + CDbl(arr_daftar(grd_daftar.Bookmark, 5))) - CDbl(arr_daftar(grd_daftar.Bookmark, 6))
            grd_daftar.Columns(7).Text = Format(ss1, "###,###,###")
            arr_daftar(grd_daftar.Bookmark, 7) = grd_daftar.Columns(7).Text
    End If
    
    If grd_daftar.Columns(5).Text = "" Then
        grd_daftar.Columns(5).Text = 0
        arr_daftar(grd_daftar.Bookmark, 5) = grd_daftar.Columns(5).Text
        
        Dim ss02
        ss02 = (CDbl(arr_daftar(grd_daftar.Bookmark, 3)) + CDbl(arr_daftar(grd_daftar.Bookmark, 4)) + CDbl(arr_daftar(grd_daftar.Bookmark, 5))) - CDbl(arr_daftar(grd_daftar.Bookmark, 6))
            grd_daftar.Columns(7).Text = Format(ss02, "###,###,###")
            arr_daftar(grd_daftar.Bookmark, 7) = grd_daftar.Columns(7).Text
        
        Exit Sub
    ElseIf grd_daftar.Columns(5).Text <> "" And grd_daftar.Columns(5).Text <> 0 Then
        
        grd_daftar.Columns(5).Text = Format(grd_daftar.Columns(5).Text, "##,##,###")
        
        arr_daftar(grd_daftar.Bookmark, 5) = grd_daftar.Columns(5).Text
        
        Dim ss2
        ss2 = (CDbl(arr_daftar(grd_daftar.Bookmark, 3)) + CDbl(arr_daftar(grd_daftar.Bookmark, 4)) + CDbl(arr_daftar(grd_daftar.Bookmark, 5))) - CDbl(arr_daftar(grd_daftar.Bookmark, 6))
            grd_daftar.Columns(7).Text = Format(ss2, "###,###,###")
            arr_daftar(grd_daftar.Bookmark, 7) = grd_daftar.Columns(7).Text
        
    End If
    
    If grd_daftar.Columns(6).Text = "" Then
        grd_daftar.Columns(6).Text = 0
        arr_daftar(grd_daftar.Bookmark, 6) = grd_daftar.Columns(6).Text
        
        Dim ss03
        ss03 = (CDbl(arr_daftar(grd_daftar.Bookmark, 3)) + CDbl(arr_daftar(grd_daftar.Bookmark, 4)) + CDbl(arr_daftar(grd_daftar.Bookmark, 5))) - CDbl(arr_daftar(grd_daftar.Bookmark, 6))
            grd_daftar.Columns(7).Text = Format(ss03, "###,###,###")
            arr_daftar(grd_daftar.Bookmark, 7) = grd_daftar.Columns(7).Text
        
        Exit Sub
    ElseIf grd_daftar.Columns(6).Text <> "" And grd_daftar.Columns(6).Text <> 0 Then
    
        grd_daftar.Columns(6).Text = Format(grd_daftar.Columns(6).Text, "##,##,###")
        
        arr_daftar(grd_daftar.Bookmark, 6) = grd_daftar.Columns(6).Text
        
        Dim ss3
        ss3 = (CDbl(arr_daftar(grd_daftar.Bookmark, 3)) + CDbl(arr_daftar(grd_daftar.Bookmark, 4)) + CDbl(arr_daftar(grd_daftar.Bookmark, 5))) - CDbl(arr_daftar(grd_daftar.Bookmark, 6))
            grd_daftar.Columns(7).Text = Format(ss3, "###,###,###")
            arr_daftar(grd_daftar.Bookmark, 7) = grd_daftar.Columns(7).Text
        
    End If
    
  Exit Sub
  
er_c:
    Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
    
End Sub
Private Sub txt_nama_GotFocus()
    txt_nama.SelStart = 0
    txt_nama.SelLength = Len(txt_nama)
End Sub
