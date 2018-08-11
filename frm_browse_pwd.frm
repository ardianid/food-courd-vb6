VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frm_browse_pwd 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4545
      ScaleWidth      =   9465
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1305
         ScaleWidth      =   1545
         TabIndex        =   7
         Top             =   3000
         Width           =   1575
         Begin VB.CommandButton cmd_hapus 
            Caption         =   "Hapus"
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
            TabIndex        =   9
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmd_edit 
            Caption         =   "Edit"
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
            TabIndex        =   8
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1545
         ScaleWidth      =   1545
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin TrueDBGrid60.TDBGrid grd_daftar 
         Height          =   3015
         Left            =   1800
         OleObjectBlob   =   "frm_browse_pwd.frx":0000
         TabIndex        =   5
         Top             =   1320
         Width           =   7575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Pencarian"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   9255
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
            Left            =   7560
            TabIndex        =   4
            Top             =   360
            Width           =   1455
         End
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
            Left            =   2280
            TabIndex        =   3
            Top             =   480
            Width           =   2655
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
            Left            =   480
            TabIndex        =   2
            Top             =   480
            Width           =   1620
         End
      End
   End
End
Attribute VB_Name = "frm_browse_pwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_daftar As New XArrayDB
Dim id_user As String



Private Sub cmd_edit_Click()

    If cmd_edit.Caption = "Edit" Then
        cmd_edit.Caption = "Read Only"
        grd_daftar.Columns(3).Locked = False
        Exit Sub
    End If
    
    If cmd_edit.Caption = "Read Only" Then
        cmd_edit.Caption = "Edit"
        grd_daftar.Columns(3).Locked = True
        Cmd_Tampil_Click
        Exit Sub
    End If
   
End Sub

Private Sub cmd_hapus_Click()
    Dim sql As String
    Dim rss As New ADODB.Recordset
        
    On Error GoTo er_hapus
        
    If MsgBox("Yakin akan menghapus password karyawan " & arr_daftar(grd_daftar.Bookmark, 2), vbYesNo + vbQuestion, "Pesan") = vbNo Then
        Exit Sub
    End If
        
        sql = "delete from tbl_user where id=" & id_user
        rss.Open sql, cn
            
        Cmd_Tampil_Click
        Exit Sub
            
er_hapus:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbQuestion, "Pesan")
            Err.Clear
        
End Sub

Private Sub Cmd_Tampil_Click()

On Error GoTo er_tampil

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        
        sql = "select * from qr_user"
            
            If txt_nama.Text <> "" Then
                sql = sql & " where nama_karyawan like '%" & Trim(txt_nama.Text) & "%'"
            End If
                
       sql = sql & " order by nama_karyawan"
       rs.Open sql, cn, adOpenKeyset
        If Not rs.EOF Then
             rs.MoveLast
             rs.MoveFirst
             
             lanjut_isi rs
         End If
       rs.Close
       
       Exit Sub
       
er_tampil:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
       
End Sub

Private Sub Form_Load()
 
    grd_daftar.Array = arr_daftar
    
    isi_daftar
    
    Me.Height = 5295
    Me.Width = 9825
    Me.ScaleHeight = 4815
    Me.ScaleWidth = 9735
    
    Me.Left = utama.Width / 2 - Me.Width / 2
    Me.Top = utama.Height / 2 - Me.Height / 2 - 2750
    
End Sub

Private Sub kosong_daftar()
    arr_daftar.ReDim 0, 0, 0, 0
    grd_daftar.ReBind
    grd_daftar.Refresh
End Sub

Private Sub isi_daftar()

On Error GoTo er_isi

    Dim sql As String
    Dim rs As New ADODB.Recordset
        
        kosong_daftar
        
        sql = "select * from qr_user order by nama_karyawan"
        rs.Open sql, cn, adOpenKeyset
            If Not rs.EOF Then
                
                rs.MoveLast
                rs.MoveFirst
                
                lanjut_isi rs
                
            End If
        rs.Close
        
        Exit Sub
        
er_isi:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
        
End Sub

Private Sub lanjut_isi(rs As Recordset)
     Dim id_u, nm, aks, akt As String
     Dim a As Long
        
        a = 1
            Do While Not rs.EOF
                arr_daftar.ReDim 1, a, 0, 6
                grd_daftar.ReBind
                grd_daftar.Refresh
                    
                    id_u = rs("id")
                    If Not IsNull(rs("nama_karyawan")) Then
                        nm = rs("nama_karyawan")
                    Else
                        nm = ""
                    End If
                    
                    If Not IsNull(rs("aktif")) Then
                        akt = rs("aktif")
                    Else
                        akt = ""
                    End If
                    
                arr_daftar(a, 0) = a
                arr_daftar(a, 1) = id_u
                arr_daftar(a, 2) = nm
                
                If akt = 1 Then
                    arr_daftar(a, 3) = vbChecked
                Else
                    arr_daftar(a, 3) = vbUnchecked
                End If
                
            a = a + 1
            rs.MoveNext
            Loop
            grd_daftar.ReBind
            grd_daftar.Refresh
                    
End Sub

Private Sub grd_daftar_AfterColUpdate(ByVal ColIndex As Integer)

On Error GoTo er_a

    If ColIndex = 3 Then
        arr_daftar(grd_daftar.Bookmark, ColIndex) = grd_daftar.Columns(ColIndex).Text
            
            Dim sql As String
            Dim rs As New ADODB.Recordset
            Dim cek
                cek = arr_daftar(grd_daftar.Bookmark, ColIndex)

                If cek = 0 Then
                    cek = 0
                Else
                    cek = 1
                End If

            sql = "update tbl_user set aktif=" & cek & " where id=" & arr_daftar(grd_daftar.Bookmark, 1)
            rs.Open sql, cn
    End If
    
    Exit Sub
    
er_a:
        Dim psn
            psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
            Err.Clear
End Sub

Private Sub grd_daftar_Click()
    On Error Resume Next
        If arr_daftar.UpperBound(1) > 0 Then
            id_user = arr_daftar(grd_daftar.Bookmark, 1)
        End If
End Sub

Private Sub grd_daftar_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    grd_daftar_Click
End Sub

Private Sub txt_nama_GotFocus()
    txt_nama.SelStart = 0
    txt_nama.SelLength = Len(txt_nama)
    Cmd_Tampil.Default = True
End Sub

Private Sub txt_nama_LostFocus()
    Cmd_Tampil.Default = False
End Sub
