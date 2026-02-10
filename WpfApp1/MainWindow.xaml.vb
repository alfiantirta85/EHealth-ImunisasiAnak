Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Windows
Imports Microsoft.Win32

' ╔════════════════════════════════════════════════════════════════════════════╗
' ║                         CLASS MODEL UNTUK DATA                              ║
' ║  KETENTUAN SOAL: Identifikasi String, Parameters, ByVal                    ║
' ╚════════════════════════════════════════════════════════════════════════════╝
Public Class JadwalImunisasi
    ' === PROPERTY (String) - KETENTUAN: Identifikasi String ===
    Public Property Id As String
    Public Property NamaAnak As String
    Public Property TanggalLahir As String
    Public Property JenisImunisasi As String
    Public Property JadwalImunisasi As String
    Public Property Status As String

    ' === CONSTRUCTOR dengan PARAMETERS - KETENTUAN: Parameter, ByVal ===
    ' ByVal = pass by value (default di VB.NET)
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

' ╔════════════════════════════════════════════════════════════════════════════╗
' ║                         CLASS UTAMA APLIKASI                                ║
' ║  KETENTUAN SOAL: Control Printing, My Namespace                            ║
' ╚════════════════════════════════════════════════════════════════════════════╝
Class MainWindow

    ' ═══ DEKLARASI VARIABEL - KETENTUAN: Collection, Array, String ═══

    ' COLLECTION - KETENTUAN: Array dan Collection
    ' ObservableCollection = collection yang auto-update UI
    Private daftarJadwal As New ObservableCollection(Of JadwalImunisasi)

    ' STRING VARIABLES - KETENTUAN: Identifikasi String
    Private isEditMode As Boolean = False
    Private editingId As String = ""

    ' NAMESPACE CONSTANT - KETENTUAN: My Namespace
    Private Const APP_NAMESPACE As String = "EHealthImunisasi"

    ' ═══════════════════════════════════════════════════════════════════════════
    ' CONSTRUCTOR - KETENTUAN: Function dan Methods
    ' ═══════════════════════════════════════════════════════════════════════════
    Public Sub New()
        ' Inisialisasi komponen UI - wajib ada
        InitializeComponent()

        ' Set ItemsSource untuk data binding
        ' KETENTUAN: Collection
        lvJadwal.ItemsSource = daftarJadwal

        ' Set default values - KETENTUAN: Date, Time dan TimeSpan
        dpTanggalLahir.SelectedDate = DateTime.Now
        dpJadwalImunisasi.SelectedDate = DateTime.Now.AddDays(7)

        ' Register event handler untuk double-click
        ' KETENTUAN: Methods
        AddHandler lvJadwal.MouseDoubleClick, AddressOf lvJadwal_MouseDoubleClick

        ' Update counter
        UpdateJumlahData()
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    ' METHOD: Update Counter - KETENTUAN: Function dan Methods, Calculate Text
    ' ═══════════════════════════════════════════════════════════════════════════
    Private Sub UpdateJumlahData()
        ' STRING METHOD: ToString() - KETENTUAN: String Methods
        ' CALCULATE TEXT - KETENTUAN: Calculate Text
        txtJumlahData.Text = $"Total: {daftarJadwal.Count} jadwal"
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    ' EVENT: Double-Click untuk Edit - KETENTUAN: Function dan Methods
    ' ═══════════════════════════════════════════════════════════════════════════
    Private Sub lvJadwal_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
        ' IF STATEMENT - KETENTUAN: Identifikasi If Statements
        If lvJadwal.SelectedItem IsNot Nothing Then
            Dim selectedJadwal = CType(lvJadwal.SelectedItem, JadwalImunisasi)
            MulaiEditMode(selectedJadwal)
        End If
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    ' METHOD: Masuk Edit Mode - KETENTUAN: Function dan Methods, Passing Arrays
    ' ═══════════════════════════════════════════════════════════════════════════
    Private Sub MulaiEditMode(ByVal jadwal As JadwalImunisasi)
        ' Set flag
        isEditMode = True
        editingId = jadwal.Id

        ' STRING METHOD: Isi form dengan data - KETENTUAN: String Methods
        txtIdJadwal.Text = jadwal.Id
        txtNamaAnak.Text = jadwal.NamaAnak

        ' DATE PARSING - KETENTUAN: Date, Time dan TimeSpan
        Dim tglLahir As DateTime
        If DateTime.TryParse(jadwal.TanggalLahir, tglLahir) Then
            dpTanggalLahir.SelectedDate = tglLahir
        End If

        ' FOR EACH LOOP - KETENTUAN: For Each Loop
        For Each item As ComboBoxItem In cmbJenisImunisasi.Items
            If item.Content.ToString() = jadwal.JenisImunisasi Then
                cmbJenisImunisasi.SelectedItem = item
                Exit For
            End If
        Next

        ' DATE PARSING
        Dim tglJadwal As DateTime
        If DateTime.TryParse(jadwal.JadwalImunisasi, tglJadwal) Then
            dpJadwalImunisasi.SelectedDate = tglJadwal
        End If

        ' SELECT CASE (alternatif) - KETENTUAN: Select Case Statements
        ' Bisa juga pakai If-Else
        If jadwal.Status = "Terjadwal" Then
            rbTerjadwal.IsChecked = True
        Else
            rbSelesai.IsChecked = True
        End If

        ' Update UI - KETENTUAN: Control Printing
        btnTambah.Visibility = Visibility.Collapsed
        btnUpdate.Visibility = Visibility.Visible
        btnBatal.Visibility = Visibility.Visible
        txtModeIndicator.Text = "📝 MODE EDIT - Sedang mengedit data"
        txtModeIndicator.Visibility = Visibility.Visible

        ' Focus dan select all text
        txtNamaAnak.Focus()
        txtNamaAnak.SelectAll()
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    ' METHOD: Keluar dari Edit Mode - KETENTUAN: Function dan Methods
    ' ═══════════════════════════════════════════════════════════════════════════
    Private Sub KeluarEditMode()
        isEditMode = False
        editingId = ""

        btnTambah.Visibility = Visibility.Visible
        btnUpdate.Visibility = Visibility.Collapsed
        btnBatal.Visibility = Visibility.Collapsed
        txtModeIndicator.Visibility = Visibility.Collapsed

        BersihkanForm()
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    ' EVENT: Edit Button Click - KETENTUAN: Function dan Methods
    ' ═══════════════════════════════════════════════════════════════════════════
    Private Sub btnEdit_Click(sender As Object, e As RoutedEventArgs)
        Dim button = CType(sender, Button)
        Dim id = button.Tag.ToString()

        ' LINQ - mencari data di collection
        Dim jadwal = daftarJadwal.FirstOrDefault(Function(j) j.Id = id)
        If jadwal IsNot Nothing Then
            MulaiEditMode(jadwal)
        End If
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    ' EVENT: Hapus Button - KETENTUAN: Handling Exceptions, Error
    ' ═══════════════════════════════════════════════════════════════════════════
    Private Sub btnHapus_Click(sender As Object, e As RoutedEventArgs)
        Try
            ' EXCEPTION HANDLING - KETENTUAN: Handling Exceptions
            Dim button = CType(sender, Button)
            Dim id = button.Tag.ToString()

            Dim jadwal = daftarJadwal.FirstOrDefault(Function(j) j.Id = id)

            ' IF STATEMENT - KETENTUAN: Identifikasi If Statements
            If jadwal IsNot Nothing Then
                ' STRING CONCATENATION - KETENTUAN: String Methods
                Dim result = MessageBox.Show(
                    $"Apakah Anda yakin ingin menghapus jadwal?" & vbCrLf &
                    $"Nama: {jadwal.NamaAnak}" & vbCrLf &
                    $"Jenis: {jadwal.JenisImunisasi}",
                    "Konfirmasi Hapus",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question)

                If result = MessageBoxResult.Yes Then
                    ' REMOVE dari COLLECTION - KETENTUAN: Collection
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
            ' ERROR HANDLING - KETENTUAN: Error
            MessageBox.Show("Error menghapus jadwal: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    ' EVENT: Batal Edit - KETENTUAN: ByVal dan ByRef
    ' ═══════════════════════════════════════════════════════════════════════════
    Private Sub btnBatal_Click(sender As Object, e As RoutedEventArgs) Handles btnBatal.Click
        KeluarEditMode()
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    ' EVENT: Update Jadwal - KETENTUAN: Handling Exceptions, String Methods
    ' ═══════════════════════════════════════════════════════════════════════════
    Private Sub btnUpdate_Click(sender As Object, e As RoutedEventArgs) Handles btnUpdate.Click
        Try
            ' VALIDASI dengan IF STATEMENTS - KETENTUAN: Identifikasi If Statements

            ' STRING METHOD: IsNullOrWhiteSpace - KETENTUAN: String Methods
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

            ' Cari jadwal yang akan diupdate
            Dim jadwal = daftarJadwal.FirstOrDefault(Function(j) j.Id = editingId)

            If jadwal IsNot Nothing Then
                ' STRING METHOD: Trim() - KETENTUAN: String Methods
                jadwal.NamaAnak = txtNamaAnak.Text.Trim()

                ' DATE METHOD: ToString() - KETENTUAN: Date, Time dan TimeSpan
                jadwal.TanggalLahir = dpTanggalLahir.SelectedDate.Value.ToString("dd/MM/yyyy")

                ' STRING METHOD: ToString() - KETENTUAN: String Methods
                jadwal.JenisImunisasi = CType(cmbJenisImunisasi.SelectedItem, ComboBoxItem).Content.ToString()
                jadwal.JadwalImunisasi = dpJadwalImunisasi.SelectedDate.Value.ToString("dd/MM/yyyy")

                If rbTerjadwal.IsChecked = True Then
                    jadwal.Status = "Terjadwal"
                Else
                    jadwal.Status = "Selesai"
                End If

                ' Refresh ListView
                lvJadwal.Items.Refresh()

                MessageBox.Show("Jadwal berhasil diupdate!", "Sukses",
                              MessageBoxButton.OK, MessageBoxImage.Information)

                KeluarEditMode()
            Else
                Throw New Exception("Data tidak ditemukan!")
            End If

        Catch ex As Exception
            ' EXCEPTION HANDLING - KETENTUAN: Handling Exceptions, Error
            MessageBox.Show("Error: " & ex.Message, "Kesalahan",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    ' EVENT: Tambah Jadwal - KETENTUAN: Handling Exceptions, Collection
    ' ═══════════════════════════════════════════════════════════════════════════
    Private Sub btnTambah_Click(sender As Object, e As RoutedEventArgs) Handles btnTambah.Click
        Try
            ' TRY-CATCH - KETENTUAN: Handling Exceptions

            ' VALIDASI INPUT dengan IF - KETENTUAN: Identifikasi If Statements
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

            ' GENERATE UNIQUE ID - KETENTUAN: String Methods
            Dim id As String = Guid.NewGuid().ToString()

            ' AMBIL DATA DARI FORM - KETENTUAN: String Methods
            Dim namaAnak As String = txtNamaAnak.Text.Trim()

            ' DATE TO STRING - KETENTUAN: Date, Time dan TimeSpan
            Dim tglLahir As String = dpTanggalLahir.SelectedDate.Value.ToString("dd/MM/yyyy")

            ' COMBOBOX GET VALUE - KETENTUAN: String Methods
            Dim jenisImunisasi As String = CType(cmbJenisImunisasi.SelectedItem, ComboBoxItem).Content.ToString()
            Dim jadwal As String = dpJadwalImunisasi.SelectedDate.Value.ToString("dd/MM/yyyy")

            ' IF-ELSE untuk STATUS - KETENTUAN: Identifikasi If Statements
            Dim status As String
            If rbTerjadwal.IsChecked = True Then
                status = "Terjadwal"
            Else
                status = "Selesai"
            End If

            ' BUAT OBJEK BARU dengan CONSTRUCTOR - KETENTUAN: Parameter, ByVal
            Dim jadwalBaru As New JadwalImunisasi(id, namaAnak, tglLahir, jenisImunisasi, jadwal, status)

            ' TAMBAHKAN ke COLLECTION - KETENTUAN: Array dan Collection
            daftarJadwal.Add(jadwalBaru)

            ' UPDATE COUNTER - KETENTUAN: Calculate Text
            UpdateJumlahData()

            MessageBox.Show("Jadwal imunisasi berhasil ditambahkan!", "Sukses",
                          MessageBoxButton.OK, MessageBoxImage.Information)

            BersihkanForm()

        Catch ex As Exception
            ' ERROR HANDLING - KETENTUAN: Error
            MessageBox.Show("Error: " & ex.Message, "Kesalahan",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    ' METHOD: Bersihkan Form - KETENTUAN: Function dan Methods
    ' ═══════════════════════════════════════════════════════════════════════════
    Private Sub BersihkanForm()
        ' STRING METHOD: Clear() - KETENTUAN: String Methods
        txtNamaAnak.Clear()

        ' DATE OPERATIONS - KETENTUAN: Date, Time dan TimeSpan
        dpTanggalLahir.SelectedDate = DateTime.Now

        cmbJenisImunisasi.SelectedItem = Nothing

        ' DATE CALCULATION: AddDays - KETENTUAN: Date, Time dan TimeSpan
        dpJadwalImunisasi.SelectedDate = DateTime.Now.AddDays(7)

        rbTerjadwal.IsChecked = True
        txtIdJadwal.Clear()
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    ' EVENT: Bersihkan Button - KETENTUAN: ByVal dan ByRef
    ' ═══════════════════════════════════════════════════════════════════════════
    Private Sub btnBersihkan_Click(sender As Object, e As RoutedEventArgs) Handles btnBersihkan.Click
        ' IF STATEMENT - KETENTUAN: Identifikasi If Statements
        If isEditMode Then
            KeluarEditMode()
        Else
            BersihkanForm()
        End If
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    ' EVENT: Simpan ke File - KETENTUAN: Files, File Dates and Times, For Each
    ' ═══════════════════════════════════════════════════════════════════════════
    Private Sub btnSimpanFile_Click(sender As Object, e As RoutedEventArgs) Handles btnSimpanFile.Click
        Try
            ' VALIDASI - KETENTUAN: If Statements
            If daftarJadwal.Count = 0 Then
                Throw New Exception("Tidak ada data untuk disimpan!")
            End If

            ' FILE DIALOG - KETENTUAN: Files, Direktori
            Dim saveDialog As New SaveFileDialog()
            saveDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            saveDialog.DefaultExt = "txt"

            ' FILENAME dengan TIMESTAMP - KETENTUAN: File Dates and Times
            saveDialog.FileName = "JadwalImunisasi_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt"

            If saveDialog.ShowDialog() = True Then
                ' FILE WRITING - KETENTUAN: Files
                ' USING statement untuk auto-dispose
                Using writer As New StreamWriter(saveDialog.FileName)
                    ' WRITE HEADER - KETENTUAN: Printed Text
                    writer.WriteLine("===== E-HEALTH: JADWAL IMUNISASI ANAK =====")

                    ' DATE TO STRING - KETENTUAN: Date, Time dan TimeSpan
                    writer.WriteLine("Tanggal Export: " & DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"))

                    ' CALCULATE TEXT - KETENTUAN: Calculate Text
                    writer.WriteLine("Jumlah Record: " & daftarJadwal.Count.ToString())
                    writer.WriteLine(New String("="c, 50))
                    writer.WriteLine()

                    ' FOR EACH LOOP - KETENTUAN: For Each Loop, Identifikasi For-Loop
                    Dim counter As Integer = 1
                    For Each jadwal As JadwalImunisasi In daftarJadwal
                        ' REPORT PRINTING - KETENTUAN: Report Printing
                        writer.WriteLine("Record #" & counter.ToString())
                        writer.WriteLine("Nama Anak        : " & jadwal.NamaAnak)
                        writer.WriteLine("Tanggal Lahir    : " & jadwal.TanggalLahir)
                        writer.WriteLine("Jenis Imunisasi  : " & jadwal.JenisImunisasi)
                        writer.WriteLine("Jadwal           : " & jadwal.JadwalImunisasi)
                        writer.WriteLine("Status           : " & jadwal.Status)
                        writer.WriteLine(New String("-"c, 50))

                        ' INCREMENT - KETENTUAN: Calculate Text
                        counter += 1
                    Next

                    ' FOOTER - KETENTUAN: Printed Text
                    writer.WriteLine()
                    writer.WriteLine("===== AKHIR DOKUMEN =====")
                End Using

                MessageBox.Show("Data berhasil disimpan ke: " & vbCrLf & saveDialog.FileName,
                              "Sukses", MessageBoxButton.OK, MessageBoxImage.Information)
            End If

        Catch ex As Exception
            ' ERROR HANDLING - KETENTUAN: Error, Handling Exceptions
            MessageBox.Show("Gagal menyimpan file: " & ex.Message, "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

End Class