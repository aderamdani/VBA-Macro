Attribute VB_Name = "Module2"
Sub MailMergeWithAttachments()
    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlSheet As Object
    Dim i As Integer
    Dim emailBody As String
    Dim emailSubject As String
    Dim emailTo As String
    Dim attachment1 As String
    Dim attachment2 As String
    Dim OutApp As Object
    Dim OutMail As Object
    Dim recipientName As String
  
    ' Buka Excel dan file yang berisi data
    Set xlApp = CreateObject("Excel.Application")
    Set xlWorkbook = xlApp.Workbooks.Open("D:\OneDrive - B One Corporation\Sertifikasi\ABSS\ABSS Certification\Sertifikat\Data Exam Score Report & Certificate  ABSS Certified User.xlsx") ' Ganti dengan path file Excel Anda
    Set xlSheet = xlWorkbook.Sheets("31122024 IAI Tazkia") ' Menggunakan sheet tertentu
  
    ' Inisialisasi Outlook
    Set OutApp = CreateObject("Outlook.Application")
  
    ' Loop melalui setiap baris di Excel
    For i = 2 To xlSheet.UsedRange.Rows.Count ' Mulai dari baris kedua
        recipientName = xlSheet.Cells(i, 2).Value ' Kolom Nama (Kolom B)
        emailTo = xlSheet.Cells(i, 4).Value ' Kolom Email (Kolom D)
        attachment1 = xlSheet.Cells(i, 22).Value ' Kolom Lampiran1 (Kolom V)
        attachment2 = xlSheet.Cells(i, 23).Value ' Kolom Lampiran2 (Kolom W)
  
        ' Ambil isi email dari dokumen Word yang sedang terbuka
        emailBody = ActiveDocument.Content.Text
        emailBody = "Dear " & recipientName & "," & vbCrLf & vbCrLf & emailBody ' Tambahkan sapaan
  
        emailSubject = "Result Exam ABSS Certified User - Accounting v.28.10" ' Ganti dengan subjek email yang diinginkan
  
        ' Buat email baru
        Set OutMail = OutApp.CreateItem(0)
        With OutMail
            .To = emailTo
            .Subject = emailSubject
            .Body = emailBody
            ' Tambahkan lampiran jika tidak kosong
            If attachment1 <> "" Then .Attachments.Add attachment1
            .Attachments.Add attachment2
            .Send ' Kirim email
        End With
  
        Set OutMail = Nothing
    Next i
  
    ' Tutup Excel
    xlWorkbook.Close False
    xlApp.Quit
  
    ' Bersihkan objek
    Set xlSheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    Set OutApp = Nothing
  
    MsgBox "Email telah dikirim!"
End Sub

