Attribute VB_Name = "mdl_semua"
Option Explicit
Public cn As New ADODB.Connection
Public path
Public frm As Form
Public path_lap
Public mdl_karyawan As Boolean, id_kar As String
Public mdl_counter As Boolean, id_counter As String
Public mdl_stock As Boolean, id_sk, id_sk1 As String
Public path_foto As String, id_user As String
Public tambah_form As Boolean, edit_form As Boolean, hapus_form As Boolean, lap_form As Boolean

Public sqlku, noff As String
Public byyr, kemm As Double
Public htu As Boolean

Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public tgl1, tgl2 As String

Public Const RGN_DIFF = 4
Public Const RGN_OR = 2

Public Function Lokasi_foto() As Boolean

On Error GoTo err_handler


    Lokasi_Database = GetSetting("path_my_food_foto", "my_food_foto", "my_food_foto", 0)
    
    Lokasi_foto = True
    
    On Error GoTo 0
    Exit Function
    
err_handler:
    
    Lokasi_foto = False
    
End Function

Public Function Set_Lokasi_foto(ByVal Letak As String) As Boolean
    
On Error GoTo err_handler

    SaveSetting "path_my_food_foto", "my_food_foto", "my_food_foto", Letak
       
    Set_Lokasi_foto = True
    
    On Error GoTo 0
    Exit Function

err_handler:
    
    Set_Lokasi_foto = False
    
    Dim p As Integer
        p = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
        Err.Clear
       
End Function

Public Function Lokasi_Database() As Variant
    
    Lokasi_Database = GetSetting("path_my_food", "my_food", "my_food", 0)
    
End Function

Public Function Set_Lokasi_Database(ByVal Letak As String) As Boolean
    
On Error GoTo err_handler

    SaveSetting "path_my_food", "my_food", "my_food", Letak
       
    Set_Lokasi_Database = True
    
    On Error GoTo 0
    Exit Function

err_handler:
    
    Set_Lokasi_Database = False
    
    Dim p As Integer
        p = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
        Err.Clear
       
End Function

Public Sub Focus_(ByVal obj As Object)
On Error Resume Next
    
    With obj
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Public Function Buka_Koneksi() As String
    
    On Error Resume Next
    
    If cn.State = adStateOpen Then cn.Close
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Lokasi_Database & ";Persist Security info=False"
    
    Buka_Koneksi = Err.Number
    Exit Function
    
    
'err_handler:
'
'    Dim Konfirm As Integer
'        Konfirm = CInt(MsgBox(Err.Number & Chr(13) & Err.Description, vbOKOnly + vbInformation, "Information"))
'        Err.Clear
'
'        End
        
End Function

Public Sub buka_path()

On Error GoTo buka

Dim i As Long, s As String
path = ""
    Open App.path & "\path.txt" For Input As #1
    
        path = Input(LOF(1), #1)
        
    Close #1
    
    Dim j As Long
    j = Len(path)
   If j <> 0 And j <> Empty Then
    path = Mid(path, 1, j - 4)
   End If
    path = Left(path, Len(path))
   Exit Sub
    
buka:
    Dim er
        er = MsgBox("Path tidak ditemukan", vbOKOnly + vbInformation, "Pesan")
        Err.Clear
        End
End Sub

Public Sub buka_path_foto()

On Error GoTo buka

Dim i As Long, s As String
path_foto = ""
    Open App.path & "\path_foto.txt" For Input As #1
    
        path_foto = Input(LOF(1), #1)
        
    Close #1
    
    path_foto = Left(path_foto, Len(path_foto))
   Exit Sub
    
buka:
    Dim er
        er = MsgBox("Path Foto tidak ditemukan", vbOKOnly + vbInformation, "Pesan")
        Err.Clear
        End
End Sub

Public Sub buka_lap()
    On Error GoTo buka

Dim i As Long, s As String
path_lap = ""
    Open App.path & "\path_lap.txt" For Input As #1
    
        path_lap = Input(LOF(1), #1)
        
    Close #1
    
    path_lap = Left(path_lap, Len(path_lap))
   Exit Sub
    
buka:
    Dim er
        er = MsgBox("Path Laporan tidak ditemukan", vbOKOnly + vbInformation, "Pesan")
        Err.Clear
        End
End Sub

Public Sub cari_wewenang(param As String)
    Dim sql, sql1 As String
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
        
        
        
        sql = "select id_wewenang from qr_wewenang where nama_form='" & param & "' and id_user=" & id_user
        rs.Open sql, cn
            If Not rs.EOF Then
                
                sql1 = "select tambah,edit,hapus,lap from qr_hak where id_wewenang=" & rs!id_wewenang
                rs1.Open sql1, cn
                    
                    If Not rs1.EOF Then
                        
                        If rs1!tambah = 1 Then
                            tambah_form = True
                        Else
                            tambah_form = False
                        End If
                        
                        If rs1!edit = 1 Then
                            edit_form = True
                        Else
                            edit_form = False
                        End If
                        
                        If rs1!hapus = 1 Then
                            hapus_form = True
                        Else
                            hapus_form = False
                        End If
                        
                        If rs1!lap = 1 Then
                            lap_form = True
                        Else
                            lap_form = False
                        End If
                    End If
                rs1.Close
            End If
        rs.Close
        
End Sub

Public Function rubah_apsen(param)
    Select Case param
        Case "Hadir"
            rubah_apsen = 1
        Case "Alpha"
            rubah_apsen = 2
        Case "Izin"
            rubah_apsen = 3
        Case "Sakit"
            rubah_apsen = 4
    End Select
End Function

Public Function absen_sebenarnya(param)
    Select Case param
        Case 1
            absen_sebenarnya = "Hadir"
        Case 2
            absen_sebenarnya = "Alpha"
        Case 3
            absen_sebenarnya = "Izin"
        Case 4
            absen_sebenarnya = "Sakit"
    End Select
End Function

Public Function bulan(param)
    Select Case param
        Case "Januari"
            bulan = 1
        Case "Februari"
            bulan = 2
        Case "Maret"
            bulan = 3
        Case "April"
            bulan = 4
        Case "Mei"
            bulan = 5
        Case "Juni"
            bulan = 6
        Case "Juli"
            bulan = 7
        Case "Agustus"
            bulan = 8
        Case "September"
            bulan = 9
        Case "Oktober"
            bulan = 10
        Case "Nopember"
            bulan = 11
        Case "Desember"
            bulan = 12
    End Select
End Function

Public Function balik_bulan(param)
    Select Case param
        Case 1
            balik_bulan = "Januari"
        Case 2
            balik_bulan = "Februari"
        Case 3
            balik_bulan = "Maret"
        Case 4
            balik_bulan = "April"
        Case 5
            balik_bulan = "Mei"
        Case 6
            balik_bulan = "Juni"
        Case 7
            balik_bulan = "Juli"
        Case 8
            balik_bulan = "Agustus"
        Case 9
            balik_bulan = "September"
        Case 10
            balik_bulan = "Oktober"
        Case 11
            balik_bulan = "Nopember"
        Case 12
            balik_bulan = "Desember"
    End Select
End Function

Public Function encrypt(param)
    Dim panjang As Long
    Dim potong
    Dim a As Long
    Dim huruf, encrypt_sementara As String
        
        potong = Trim(param)
        panjang = Len(potong)
            
            encrypt_sementara = ""
                For a = 1 To panjang
                    huruf = Mid(potong, a, 1)
                        If encrypt_sementara <> "" Then
                            encrypt_sementara = encrypt_sementara & "." & Asc(huruf)
                        Else
                            encrypt_sementara = encrypt_sementara & Asc(huruf)
                        End If
                Next a
             encrypt = ""
             encrypt = encrypt_sementara
        
End Function
    
Public Function decrypt(param)
    Dim panjang As Long
    Dim potong
    Dim a As Long
    Dim huruf As String
    Dim decrypt_sementara As String
    Dim kata
    Dim ubah
        
        potong = Trim(param)
        panjang = Len(potong)
            
            decrypt_sementara = ""
            kata = ""
                For a = 1 To panjang
                    huruf = Mid(potong, a, 1)
                    If huruf <> "." Then
                        kata = kata & huruf
                    End If
                    If huruf = "." Or a = panjang Then
                        ubah = ""
                        ubah = Chr(kata)
                        decrypt_sementara = decrypt_sementara & ubah
                        kata = ""
                    End If
                Next a
            decrypt = ""
            decrypt = decrypt_sementara
End Function
