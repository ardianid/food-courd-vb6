VERSION 5.00
Begin VB.Form frm_alat 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2265
      ScaleWidth      =   7425
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton cmd_simpan 
         Caption         =   "Simpan"
         Height          =   495
         Left            =   5880
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txt_stock 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txt_alat 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   720
         Width           =   5295
      End
      Begin VB.TextBox txt_kode_alat 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Awal "
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Inventori"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Inventori"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frm_alat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_simpan_Click()

On Error GoTo er_simpan

Dim sql, sql1 As String
Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
     
    If txt_kode_alat.Text = "" And txt_alat.Text = "" And txt_stock.Text = "" Then
        MsgBox ("Semua Data harus diisi")
        Exit Sub
    End If
     
    sql = "select kode from tbl_alat where kode='" & Trim(txt_kode_alat.Text) & "'"
    rs.Open sql, cn
        If Not rs.EOF Then
            MsgBox ("Kode yang anda masukkan sudah ada")
        Else
            
            sql1 = "insert into tbl_alat (kode,nama_alat,stock_awal)"
            sql1 = sql1 & " values('" & Trim(txt_kode_alat.Text) & "','" & Trim(txt_alat.Text) & "'," & Trim(txt_stock.Text) & ")"
            rs1.Open sql1, cn
                
            MsgBox ("Data berhasil disimpan")
            kosong
            txt_kode_alat.SetFocus
       End If
    rs.Close
    Exit Sub
                        
er_simpan:
    
    Dim psn
        psn = MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbExclamation, "Error")
        Err.Clear
        
End Sub

Private Sub kosong()
    txt_kode_alat.Text = ""
    txt_alat.Text = ""
    txt_stock.Text = ""
End Sub
Private Sub txt_alat_GotFocus()
    txt_alat.SelStart = 0
    txt_alat.SelLength = Len(txt_alat)
End Sub

Private Sub txt_kode_alat_GotFocus()
    txt_kode_alat.SelStart = 0
    txt_kode_alat.SelLength = Len(txt_kode_alat)
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
