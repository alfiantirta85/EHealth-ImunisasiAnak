Imports System.Collections.ObjectModel
Imports System.IO
Imports Microsoft.Win32

Public Class JadwalImunisasi
    Public Property Id As String
    Public Property NamaAnak As String
    Public Property TanggalLahir As Date
    Public Property JenisImunisasi As String
    Public Property JadwalImunisasi As Date
    Public Property Status As String

    Public Sub New(ByVal id As String, ByVal nama As String, ByVal tglLahir As String,
                   ByVal jenis As String, ByVal jadwal As String, ByVal status As String)
        Me.Id = id
        Me.NamaAnak = nama
        Me.TanggalLahir = tglLahir
        Me.JenisImunisasi = jenis
        Me.JadwalImunisasi = jadwal
        Me.Status = status
    End Sub
End Class

Class MainWindow
    Private daftarJadwal As New ObservableCollection(Of JadwalImunisasi)

    Private isEditMode As Boolean = False
    Private editingId As String = ""

    Public Sub New()
        InitializeComponent()

        lvJadwal.ItemsSource = daftarJadwal

        dpTanggalLahir.SelectedDate = DateTime.Now
        dpJadwalImunisasi.SelectedDate = DateTime.Now.AddDays(7)

        AddHandler lvJadwal.MouseDoubleClick, AddressOf lvJadwal_MouseDoubleClick

        UpdateJumlahData()
    End Sub

    Private Sub UpdateJumlahData()
        txtJumlahData.Text = $"Total: {daftarJadwal.Count} jadwal"
    End Sub

    Private Sub lvJadwal_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
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
            dpTanggalLahir.SelectedDate = tglLahir
        End If

        For Each item As ComboBoxItem In cmbJenisImunisasi.Items
            If item.Content.ToString() = jadwal.JenisImunisasi Then
                cmbJenisImunisasi.SelectedItem = item
                Exit For
            End If
        Next

        Dim tglJadwal As DateTime
        If DateTime.TryParse(jadwal.JadwalImunisasi, tglJadwal) Then
            dpJadwalImunisasi.SelectedDate = tglJadwal
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

    Private Sub btnEdit_Click(sender As Object, e As RoutedEventArgs)
        Dim button = CType(sender, Button)
        Dim id = button.Tag.ToString()

        Dim jadwal = daftarJadwal.FirstOrDefault(Function(j) j.Id = id)
        If jadwal IsNot Nothing Then
            MulaiEditMode(jadwal)
        End If
    End Sub

    Private Sub btnHapus_Click(sender As Object, e As RoutedEventArgs)
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

                If result = MessageBoxResult.Yes Then
                    daftarJadwal.Remove(jadwal)

                    UpdateJumlahData()

                    If isEditMode AndAlso editingId = id Then
                        KeluarEditMode()
                    End If

                    MessageBox.Show("Jadwal berhasil dihapus!", "Sukses",
                                  MessageBoxButton.OK, MessageBoxImage.Information)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error menghapus jadwal: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub btnBatal_Click(sender As Object, e As RoutedEventArgs) Handles btnBatal.Click
        KeluarEditMode()
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As RoutedEventArgs) Handles btnUpdate.Click
        Try
            If String.IsNullOrWhiteSpace(txtNamaAnak.Text) Then
                Throw New Exception("Nama anak harus diisi!")
            End If

            If Not dpTanggalLahir.SelectedDate.HasValue Then
                Throw New Exception("Tanggal lahir harus dipilih!")
            End If

            If cmbJenisImunisasi.SelectedItem Is Nothing Then
                Throw New Exception("Jenis imunisasi harus dipilih!")
            End If

            If Not dpJadwalImunisasi.SelectedDate.HasValue Then
                Throw New Exception("Jadwal imunisasi harus dipilih!")
            End If

            Dim jadwal = daftarJadwal.FirstOrDefault(Function(j) j.Id = editingId)

            If jadwal IsNot Nothing Then
                jadwal.NamaAnak = txtNamaAnak.Text.Trim()

                jadwal.TanggalLahir = dpTanggalLahir.SelectedDate.Value.ToString("dd/MM/yyyy")

                jadwal.JenisImunisasi = CType(cmbJenisImunisasi.SelectedItem, ComboBoxItem).Content.ToString()
                jadwal.JadwalImunisasi = dpJadwalImunisasi.SelectedDate.Value.ToString("dd/MM/yyyy")

                If rbTerjadwal.IsChecked = True Then
                    jadwal.Status = "Terjadwal"
                Else
                    jadwal.Status = "Selesai"
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

    Private Sub btnTambah_Click(sender As Object, e As RoutedEventArgs) Handles btnTambah.Click
        Try
            If String.IsNullOrWhiteSpace(txtNamaAnak.Text) Then
                Throw New Exception("Nama anak harus diisi!")
            End If

            If Not dpTanggalLahir.SelectedDate.HasValue Then
                Throw New Exception("Tanggal lahir harus dipilih!")
            End If

            If cmbJenisImunisasi.SelectedItem Is Nothing Then
                Throw New Exception("Jenis imunisasi harus dipilih!")
            End If

            If Not dpJadwalImunisasi.SelectedDate.HasValue Then
                Throw New Exception("Jadwal imunisasi harus dipilih!")
            End If

            Dim id As String = Guid.NewGuid().ToString()

            Dim namaAnak As String = txtNamaAnak.Text.Trim()

            Dim tglLahir As String = dpTanggalLahir.SelectedDate.Value.ToString("dd/MM/yyyy")

            Dim jenisImunisasi As String = CType(cmbJenisImunisasi.SelectedItem, ComboBoxItem).Content.ToString()
            Dim jadwal As String = dpJadwalImunisasi.SelectedDate.Value.ToString("dd/MM/yyyy")

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

        dpTanggalLahir.SelectedDate = DateTime.Now

        cmbJenisImunisasi.SelectedItem = Nothing

        dpJadwalImunisasi.SelectedDate = DateTime.Now.AddDays(7)

        rbTerjadwal.IsChecked = True
        txtIdJadwal.Clear()
    End Sub

    Private Sub btnBersihkan_Click(sender As Object, e As RoutedEventArgs) Handles btnBersihkan.Click
        If isEditMode Then
            KeluarEditMode()
        Else
            BersihkanForm()
        End If
    End Sub

    Private Sub btnSimpanFile_Click(sender As Object, e As RoutedEventArgs) Handles btnSimpanFile.Click
        Try
            If daftarJadwal.Count = 0 Then
                Throw New Exception("Tidak ada data untuk disimpan!")
            End If

            Dim saveDialog As New SaveFileDialog()
            saveDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            saveDialog.DefaultExt = "txt"

            saveDialog.FileName = "JadwalImunisasi_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt"

            If saveDialog.ShowDialog() = True Then
                Using writer As New StreamWriter(saveDialog.FileName)
                    writer.WriteLine("===== E-HEALTH: JADWAL IMUNISASI ANAK =====")

                    writer.WriteLine("Tanggal Export: " & DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"))

                    writer.WriteLine("Jumlah Record: " & daftarJadwal.Count.ToString())
                    writer.WriteLine(New String("="c, 50))
                    writer.WriteLine()

                    Dim counter As Integer = 1
                    For Each jadwal As JadwalImunisasi In daftarJadwal
                        writer.WriteLine("Record #" & counter.ToString())
                        writer.WriteLine("Nama Anak        : " & jadwal.NamaAnak)
                        writer.WriteLine("Tanggal Lahir    : " & jadwal.TanggalLahir)
                        writer.WriteLine("Jenis Imunisasi  : " & jadwal.JenisImunisasi)
                        writer.WriteLine("Jadwal           : " & jadwal.JadwalImunisasi)
                        writer.WriteLine("Status           : " & jadwal.Status)
                        writer.WriteLine(New String("-"c, 50))

                        counter += 1
                    Next

                    writer.WriteLine()
                    writer.WriteLine("===== AKHIR DOKUMEN =====")
                End Using

                MessageBox.Show("Data berhasil disimpan ke: " & vbCrLf & saveDialog.FileName,
                              "Sukses", MessageBoxButton.OK, MessageBoxImage.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("Gagal menyimpan file: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub
End Class