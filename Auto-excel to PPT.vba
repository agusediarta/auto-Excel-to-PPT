' PASTIKAN KODE INI BERADA DI BARIS PALING ATAS
#If VBA7 Then
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
        (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
        ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#Else
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
        (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
        ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

Sub ExcelKePPT_TembusDrive()
    Dim pptApp As Object, pptPres As Object, pptSlideAsli As Object
    Dim slideRange As Object, pptSlide As Object
    Dim ws As Worksheet
    Dim BarisTerakhir As Long, i As Long
    Dim strNIM As String, strNama As String, strTTL As String, strJudul As String, linkFoto As String
    Dim directLink As String, fileID As String, tempPath As String
    Dim newPic As Object, errBox As Object
    Dim hasilDownload As Boolean

    ' Set Sheet yang sedang aktif sebagai sumber data
    Set ws = ActiveSheet
    
    ' Mencari baris paling bawah yang ada isinya (patokannya kolom B / kolom ke-2)
    BarisTerakhir = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    If BarisTerakhir < 2 Then Exit Sub

    ' Menyambungkan Excel ke PowerPoint yang sedang terbuka
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    Set pptPres = pptApp.ActivePresentation
    On Error GoTo 0

    If pptPres Is Nothing Then
        MsgBox "Buka PowerPoint-nya dulu ya!", vbExclamation
        Exit Sub
    End If

    ' Menggunakan Slide ke-1 sebagai Template
    Set pptSlideAsli = pptPres.Slides(1)

    ' =========================================================================
    ' 1. AREA PENGATURAN KOLOM EXCEL (SILAKAN SESUAIKAN ANGKA KOLOMNYA DI SINI)
    ' =========================================================================
    ' i = Mulai dari baris ke-2 (karena baris 1 biasanya judul tabel / header)
    For i = 2 To BarisTerakhir
        
        ' Angka di dalam kurung menunjukkan urutan kolom. 
        ' Contoh: 2 = Kolom B (NIM/No Induk), 3 = Kolom C (Nama), dst.
        strNIM = Trim(ws.Cells(i, 2).Value)
        
        If strNIM <> "" Then
            strNama = Trim(ws.Cells(i, 3).Value)
            strTTL = Trim(ws.Cells(i, 4).Value)
            strJudul = Trim(ws.Cells(i, 5).Value)
            linkFoto = Trim(ws.Cells(i, 6).Value) ' Kolom link Google Drive
            
            ' Duplikat Slide Template lalu pindahkan ke urutan paling belakang
            Set slideRange = pptSlideAsli.Duplicate
            Set pptSlide = slideRange.Item(1)
            pptSlide.MoveTo pptPres.Slides.Count
            
            ' =========================================================================
            ' 2. AREA PENGATURAN TEKS PPT (GANTI TEKS DALAM KURUNG SIKU JIKA PERLU)
            ' =========================================================================
            ' GantiTeksLugas [Slide Tujuan], [Teks Template di PPT], [Data dari Excel]
            ' Contoh: Jika di PPT tertulis "[NO_INDUK]", ubah "[NIM]" di bawah menjadi "[NO_INDUK]"
            GantiTeksLugas pptSlide, "[NAMA]", strNama
            GantiTeksLugas pptSlide, "[NIM]", strNIM
            GantiTeksLugas pptSlide, "[TTL]", strTTL
            GantiTeksLugas pptSlide, "[JUDUL]", strJudul
            GantiTeksLugas pptSlide, "[Judul]", strJudul ' Antisipasi typo huruf kecil
            
            ' =========================================================================
            ' 3. AREA PENGATURAN FOTO OTOMATIS
            ' =========================================================================
            If InStr(linkFoto, "/d/") > 0 Then
                ' Mengambil ID unik dari link Google Drive
                fileID = Split(Split(linkFoto, "/d/")(1), "/")(0)
                
                ' TRIK RAHASIA: Pakai Thumbnail API agar tidak diblokir saat download massal
                directLink = "https://drive.google.com/thumbnail?id=" & fileID & "&sz=w800"
                tempPath = Environ("TEMP") & "\" & fileID & ".jpg"
                
                ' Memanggil fungsi mesin download
                hasilDownload = DownloadGambarModern(directLink, tempPath)
                
                If hasilDownload = True Then
                    On Error Resume Next
                    
                    ' ---------------------------------------------------------------
                    ' MENGATUR KOORDINAT & UKURAN FOTO DI PPT (Satuan dalam Points)
                    ' Left = Jarak dari kiri, Top = Jarak dari atas
                    ' Width = Lebar Foto, Height = Tinggi Foto
                    ' ---------------------------------------------------------------
                    Set newPic = pptSlide.Shapes.AddPicture(FileName:=tempPath, _
                        LinkToFile:=0, SaveWithDocument:=-1, _
                        Left:=115, Top:=175, Width:=94, Height:=126)
                        
                    ' Matikan rasio asli gambar agar pas dengan ukuran bingkai yang kita atur
                    newPic.LockAspectRatio = msoFalse
                    On Error GoTo 0
                    
                    ' Hapus file foto dari komputer setelah tertempel di PPT
                    If Dir(tempPath) <> "" Then Kill tempPath
                Else
                    ' Jika link bermasalah atau di-private, munculkan teks peringatan merah
                    Set errBox = pptSlide.Shapes.AddTextbox(1, 115, 175, 94, 126)
                    errBox.TextFrame.TextRange.Text = "FOTO GAGAL!" & vbCrLf & "Link Rusak/Private"
                    errBox.TextFrame.TextRange.Font.Color.RGB = RGB(255, 0, 0)
                    errBox.TextFrame.TextRange.Font.Size = 12
                    errBox.TextFrame.TextRange.Font.Bold = msoTrue
                End If
            End If
        End If
    Next i

    MsgBox "PROSES SELESAI! Semua slide berhasil digenerate.", vbInformation
End Sub

' =======================================================
' FUNGSI GANTI TEKS (Dapat menembus Group Shape di PPT)
' =======================================================
Sub GantiTeksLugas(sld As Object, cari As String, ganti As String)
    Dim shp As Object, i As Integer
    For Each shp In sld.Shapes
        If shp.Type = 6 Then ' Jika bentuknya gabungan (Group)
            For i = 1 To shp.GroupItems.Count
                ProsesReplace shp.GroupItems(i), cari, ganti
            Next i
        Else
            ProsesReplace shp, cari, ganti
        End If
    Next shp
End Sub

Sub ProsesReplace(shp As Object, cari As String, ganti As String)
    On Error Resume Next
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            shp.TextFrame.TextRange.Replace cari, ganti
        End If
    End If
    On Error GoTo 0
End Sub

' =======================================================
' MESIN DOWNLOAD MODERN WINHTTP (Anti Blokir Google)
' =======================================================
Function DownloadGambarModern(ByVal url As String, ByVal savePath As String) As Boolean
    Dim http As Object
    Dim stream As Object

    On Error GoTo ErrHandler

    ' Menggunakan WinHTTP agar lebih stabil
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False

    ' Menyamar sebagai Browser biasa agar Google tidak mencurigainya sebagai Bot/Script
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    http.send

    ' Status 200 artinya sukses merespon dan mendownload
    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        stream.Type = 1 ' adTypeBinary
        stream.Write http.responseBody
        stream.SaveToFile savePath, 2 ' Angka 2 berarti menimpa (overwrite) file jika sudah ada
        stream.Close
        DownloadGambarModern = True
    Else
        DownloadGambarModern = False
    End If

    Exit Function

ErrHandler:
    DownloadGambarModern = False
End Function