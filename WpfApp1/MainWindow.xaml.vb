Imports System.Collections.ObjectModel
Imports System.IO

Public Class JadwalImunisasi
    Public Property Id As String
    Public Property NamaAnak As String
    Public Property TanggalLahir As String
    Public Property JenisImunisasi As String
    Public Property JadwalImunisasi As String
    Public Property Status As String

    Public Sub New(ByVal id As String, ByVal nama As String, ByVal tglLahir As String,
                   ByVal jenis As String, ByVal jadwal As String, ByVal status As String)
        With Me
            .Id = id
            .NamaAnak = nama
            .TanggalLahir = tglLahir
            .JenisImunisasi = jenis
            .JadwalImunisasi = jadwal
            .Status = status
        End With
    End Sub
End Class

Class MainWindow
    Private daftarJadwal As New ObservableCollection(Of JadwalImunisasi)

    Private isEditMode As Boolean = False
    Private editingId As String = ""

    Public Sub MainWindow_Loaded() Handles Me.Loaded
        InitializeComponent()
        lvJadwal.ItemsSource = daftarJadwal

        dtpTanggalLahir.SelectedDate = Nothing
        dtpTanggalLahir.DisplayDateEnd = Date.Today
        dtpJadwalImunisasi.SelectedDate = Nothing

        UpdateJumlahData()
    End Sub

    Private Sub UpdateJumlahData()
        txtJumlahData.Text = $"Total: {daftarJadwal.Count} jadwal"
    End Sub

    Private Sub lvJadwal_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles lvJadwal.MouseDoubleClick
        If lvJadwal.SelectedItem IsNot Nothing Then
            Dim selectedJadwal = CType(lvJadwal.SelectedItem, JadwalImunisasi)
            MulaiEditMode(selectedJadwal)
        End If
    End Sub

    Private Sub MulaiEditMode(ByVal jadwal As JadwalImunisasi)
        isEditMode = True
        editingId = jadwal.Id

        txtIdJadwal.Text = jadwal.Id
        txtNamaAnak.Text = jadwal.NamaAnak

        Dim tglLahir As DateTime
        If DateTime.TryParse(jadwal.TanggalLahir, tglLahir) Then
            dtpTanggalLahir.SelectedDate = tglLahir
        End If

        For Each item As ComboBoxItem In cmbJenisImunisasi.Items
            If item.Content.ToString() = jadwal.JenisImunisasi Then
                cmbJenisImunisasi.SelectedItem = item
                Exit For
            End If
        Next

        Dim tglJadwal As DateTime
        If DateTime.TryParse(jadwal.JadwalImunisasi, tglJadwal) Then
            dtpJadwalImunisasi.SelectedDate = tglJadwal
        End If

        If jadwal.Status = "Terjadwal" Then
            rbTerjadwal.IsChecked = True
        Else
            rbSelesai.IsChecked = True
        End If

        btnTambah.Visibility = Visibility.Collapsed
        btnUpdate.Visibility = Visibility.Visible
        btnBatal.Visibility = Visibility.Visible
        txtModeIndicator.Text = "MODE EDIT - Sedang mengedit data"
        txtModeIndicator.Visibility = Visibility.Visible

        txtNamaAnak.Focus()
        txtNamaAnak.SelectAll()
    End Sub

    Private Sub KeluarEditMode()
        isEditMode = False
        editingId = ""

        btnTambah.Visibility = Visibility.Visible
        btnUpdate.Visibility = Visibility.Collapsed
        btnBatal.Visibility = Visibility.Collapsed
        txtModeIndicator.Visibility = Visibility.Collapsed

        BersihkanForm()
    End Sub

    Private Sub BtnEdit_Click(sender As Object, e As RoutedEventArgs)
        Dim button = CType(sender, Button)
        Dim id = button.Tag.ToString()

        Dim jadwal = daftarJadwal.FirstOrDefault(Function(j) j.Id = id)
        If jadwal IsNot Nothing Then
            MulaiEditMode(jadwal)
        End If
    End Sub

    Private Sub BtnHapus_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim button = CType(sender, Button)
            Dim id = button.Tag.ToString()

            Dim jadwal = daftarJadwal.FirstOrDefault(Function(j) j.Id = id)

            If jadwal IsNot Nothing Then
                Dim result = MessageBox.Show(
                    $"Apakah Anda yakin ingin menghapus jadwal?" & vbCrLf &
                    $"Nama: {jadwal.NamaAnak}" & vbCrLf &
                    $"Jenis: {jadwal.JenisImunisasi}",
                    "Konfirmasi Hapus",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question)
                Select Case result
                    Case MessageBoxResult.Yes
                        daftarJadwal.Remove(jadwal)

                        UpdateJumlahData()

                        If isEditMode AndAlso editingId = id Then
                            KeluarEditMode()
                        End If

                        MessageBox.Show("Jadwal berhasil dihapus!", "Sukses",
                                      MessageBoxButton.OK, MessageBoxImage.Information)
                    Case MessageBoxResult.No
                        Exit Sub
                End Select
                If result = MessageBoxResult.Yes Then

                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error menghapus jadwal: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub BtnBatal_Click(sender As Object, e As RoutedEventArgs) Handles btnBatal.Click
        KeluarEditMode()
    End Sub

    Private Sub BtnUpdate_Click(sender As Object, e As RoutedEventArgs) Handles btnUpdate.Click
        Try
            If String.IsNullOrWhiteSpace(txtNamaAnak.Text) Then
                Throw New Exception("Nama anak harus diisi!")
            End If

            If Not dtpTanggalLahir.SelectedDate.HasValue Then
                Throw New Exception("Tanggal lahir harus dipilih!")
            End If

            If cmbJenisImunisasi.SelectedItem Is Nothing Then
                Throw New Exception("Jenis imunisasi harus dipilih!")
            End If

            If Not dtpJadwalImunisasi.SelectedDate.HasValue Then
                Throw New Exception("Jadwal imunisasi harus dipilih!")
            End If

            Dim jenisImunisasi As String = CType(cmbJenisImunisasi.SelectedItem, ComboBoxItem).Content.ToString()
            Dim namaAnak As String = txtNamaAnak.Text.Trim()

            For Each jadwal In daftarJadwal
                ' Skip jika ini adalah data yang sedang diedit
                If jadwal.Id <> editingId Then
                    If jadwal.NamaAnak.Equals(namaAnak) And jadwal.JenisImunisasi.Equals(jenisImunisasi) Then
                        Throw New Exception("Jenis imunisasi telah dijadwalkan")
                    End If
                End If
            Next

            Dim jadwalEdit = daftarJadwal.FirstOrDefault(Function(j) j.Id = editingId)

            If jadwalEdit IsNot Nothing Then
                jadwalEdit.NamaAnak = namaAnak
                jadwalEdit.TanggalLahir = dtpTanggalLahir.SelectedDate.Value.ToString("dd/MM/yyyy")
                jadwalEdit.JenisImunisasi = jenisImunisasi
                jadwalEdit.JadwalImunisasi = dtpJadwalImunisasi.SelectedDate.Value.ToString("dd/MM/yyyy")

                If rbTerjadwal.IsChecked = True Then
                    jadwalEdit.Status = "Terjadwal"
                Else
                    jadwalEdit.Status = "Selesai"
                End If

                lvJadwal.Items.Refresh()

                MessageBox.Show("Jadwal berhasil diupdate!", "Sukses",
                              MessageBoxButton.OK, MessageBoxImage.Information)

                KeluarEditMode()
            Else
                Throw New Exception("Data tidak ditemukan!")
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Kesalahan",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub BtnTambah_Click(sender As Object, e As RoutedEventArgs) Handles btnTambah.Click
        Try
            If String.IsNullOrWhiteSpace(txtNamaAnak.Text) Then
                Throw New Exception("Nama anak harus diisi!")
            End If

            If Not dtpTanggalLahir.SelectedDate.HasValue Then
                Throw New Exception("Tanggal lahir harus dipilih!")
            End If

            If cmbJenisImunisasi.SelectedItem Is Nothing Then
                Throw New Exception("Jenis imunisasi harus dipilih!")
            End If

            If Not dtpJadwalImunisasi.SelectedDate.HasValue Then
                Throw New Exception("Jadwal imunisasi harus dipilih!")
            End If

            Dim id As String = Guid.NewGuid().ToString()

            Dim namaAnak As String = txtNamaAnak.Text.Trim()

            Dim tglLahir As String = dtpTanggalLahir.SelectedDate.Value.ToString("dd/MM/yyyy")

            Dim jenisImunisasi As String = CType(cmbJenisImunisasi.SelectedItem, ComboBoxItem).Content.ToString()
            Dim jadwal As String = dtpJadwalImunisasi.SelectedDate.Value.ToString("dd/MM/yyyy")

            Dim status As String
            If rbTerjadwal.IsChecked = True Then
                status = "Terjadwal"
            Else
                status = "Selesai"
            End If

            Dim jadwalBaru As New JadwalImunisasi(id, namaAnak, tglLahir, jenisImunisasi, jadwal, status)

            daftarJadwal.Add(jadwalBaru)

            UpdateJumlahData()

            MessageBox.Show("Jadwal imunisasi berhasil ditambahkan!", "Sukses",
                          MessageBoxButton.OK, MessageBoxImage.Information)

            BersihkanForm()

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Kesalahan",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub BersihkanForm()
        txtNamaAnak.Clear()

        dtpTanggalLahir.SelectedDate = Nothing

        cmbJenisImunisasi.SelectedItem = Nothing
        cmbJenisImunisasi.SelectedIndex = -1

        dtpJadwalImunisasi.SelectedDate = Nothing

        rbTerjadwal.IsChecked = True
        txtIdJadwal.Clear()
    End Sub

    Private Sub BtnBersihkan_Click(sender As Object, e As RoutedEventArgs) Handles btnBersihkan.Click
        If isEditMode Then
            KeluarEditMode()
        Else
            BersihkanForm()
        End If
    End Sub

    Private Sub BtnHapusSemua_Click(sender As Object, e As RoutedEventArgs) Handles btnHapusSemua.Click
        Try
            If daftarJadwal.Count = 0 Then
                MessageBox.Show("Tidak ada data untuk dihapus!", "Informasi",
                          MessageBoxButton.OK, MessageBoxImage.Information)
                Return
            End If

            Dim result = MessageBox.Show(
            $"Apakah Anda yakin ingin menghapus SEMUA data ({daftarJadwal.Count} jadwal)?",
            "Konfirmasi Hapus Semua",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning)

            If result = MessageBoxResult.Yes Then
                daftarJadwal.Clear()
                UpdateJumlahData()

                If isEditMode Then
                    KeluarEditMode()
                End If

                MessageBox.Show("Semua data berhasil dihapus!", "Sukses",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("Error menghapus data: " & ex.Message, "Error",
                      MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub SaveFile(dataArray As Array)
        Dim dirPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Personal) & "\imunisasi"

        Dim fileName As String = dirPath & "\JadwalImunisasi_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt"
        FileOpen(1, fileName, OpenMode.Output)
        PrintLine(1, "===== E-HEALTH: JADWAL IMUNISASI ANAK =====")
        PrintLine(1, "Tanggal Export: " & DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"))
        PrintLine(1, "Jumlah Record: " & dataArray.GetUpperBound(0).ToString())
        PrintLine(1, New String("="c, 50))
        PrintLine(1)
        For i As Integer = 0 To dataArray.GetUpperBound(0)
            PrintLine(1, "Record #" & (i + 1).ToString())
            PrintLine(1, "Nama Anak        : " & dataArray(i).NamaAnak)
            PrintLine(1, "Tanggal Lahir    : " & dataArray(i).TanggalLahir)
            PrintLine(1, "Jenis Imunisasi  : " & dataArray(i).JenisImunisasi)
            PrintLine(1, "Jadwal           : " & dataArray(i).JadwalImunisasi)
            PrintLine(1, "Status           : " & dataArray(i).Status)
        Next
        PrintLine(1)
        PrintLine(1, "===== AKHIR DOKUMEN =====")
        FileClose(1)

        MessageBox.Show("Data berhasil disimpan ke: " & vbCrLf & fileName,
                          "Sukses", MessageBoxButton.OK, MessageBoxImage.Information)
    End Sub

    Private Sub BtnSimpanFile_Click(sender As Object, e As RoutedEventArgs) Handles btnSimpanFile.Click
        Try
            If daftarJadwal.Count = 0 Then
                Throw New Exception("Tidak ada data untuk disimpan!")
            End If

            Dim pathDir As String = Environment.GetFolderPath(Environment.SpecialFolder.Personal) & "\imunisasi"
            Directory.CreateDirectory(pathDir)
            Dim dataArray As Array = daftarJadwal.ToArray()

            If (Directory.GetFiles(pathDir).Length > 0) Then
                Dim latestFile As String = Directory.GetFiles(pathDir).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).FirstOrDefault().ToString()
                Dim fileDate As Date = FileDateTime(latestFile)
                Dim result = MessageBox.Show(
                    $"File terakhir disimpan pada: {fileDate}" & vbCrLf &
                    $"Apakah anda ingin menyimpan lagi?",
                    "Konfirmasi simpan",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question)

                Select Case result
                    Case MessageBoxResult.Yes
                        SaveFile(dataArray)
                    Case MessageBoxResult.No
                        Exit Sub
                End Select
            Else
                SaveFile(dataArray)
            End If
        Catch ex As Exception
            MessageBox.Show("Gagal menyimpan file: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub
End Class