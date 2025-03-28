import sys
import pandas as pd
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel,
    QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QComboBox,
    QMessageBox, QDateEdit, QHeaderView, QFileDialog, QGroupBox, 
    QStatusBar, QSplashScreen, QSizePolicy
)
from PyQt5.QtCore import QDate, Qt, QTimer
from PyQt5.QtGui import QIcon, QPixmap, QColor

DATA_PATH = 'data/kost_data.xlsx'

nomor_kamar = [
    "1A", "1B", "1C", "1D", "1E", "1F", "1G", "1H", "1I", "1J",
    "1K", "1L", "1M", "1N", "1O", "1P", "1Q", "1R",
    "2A", "2B", "2C", "2D", "2E", "2F", "2G", "2H", "2I", "2J",
    "2K", "2L", "2M", "2N", "2O", "2P", "2Q", "2R",
    "3A", "3B", "3C", "3D"
]

harga_kamar = [
    "Rp. 450.000", "Rp. 500.000", "Rp. 550.000", "Rp. 600.000",
    "Rp. 650.000", "Rp. 700.000", "Rp. 750.000", "Rp. 800.000",
    "Rp. 850.000", "Rp. 900.000", "Rp. 950.000", "Rp. 1.000.000"
]

def load_data():
    if os.path.exists(DATA_PATH):
        try:
            df = pd.read_excel(DATA_PATH, dtype={'No Kamar': str, 'Nomor WhatsApp': str})
            df = df.fillna('')
            df['No Kamar'] = df['No Kamar'].str.strip().str.upper()
            df['Nama Penghuni'] = df['Nama Penghuni'].str.strip()
            
            df['Sorting'] = df['No Kamar'].apply(lambda x: nomor_kamar.index(x) if x in nomor_kamar else len(nomor_kamar))
            df = df.sort_values('Sorting')
            df = df.drop('Sorting', axis=1)
            
            return df
        except Exception as e:
            print(f"Error loading data: {e}")
            return pd.DataFrame(columns=[
                'No Kamar', 'Nama Penghuni', 'Nomor WhatsApp', 'Tanggal Masuk',
                'Status Kamar', 'Status Pembayaran', 'Harga Kamar'
            ])
    else:
        return pd.DataFrame(columns=[
            'No Kamar', 'Nama Penghuni', 'Nomor WhatsApp', 'Tanggal Masuk',
            'Status Kamar', 'Status Pembayaran', 'Harga Kamar'
        ])

def save_data(df):
    try:
        os.makedirs('data', exist_ok=True)
        df['No Kamar'] = df['No Kamar'].str.strip().str.upper()
        
        df['Sorting'] = df['No Kamar'].apply(lambda x: nomor_kamar.index(x) if x in nomor_kamar else len(nomor_kamar))
        df = df.sort_values('Sorting')
        df = df.drop('Sorting', axis=1)
        
        df.to_excel(DATA_PATH, index=False)
        return True
    except Exception as e:
        print(f"Error saving data: {e}")
        return False

def format_whatsapp_number(nomor):
    if pd.isna(nomor) or nomor == '':
        return ''
    nomor = ''.join(filter(str.isdigit, str(nomor)))
    if nomor.startswith('0'):
        nomor = '62' + nomor[1:]
    elif not nomor.startswith('62'):
        nomor = '62' + nomor
    return nomor

class KostApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Aplikasi Manajemen Kost")
        self.setWindowIcon(QIcon('icon.png'))
        self.resize(1200, 700)  # Ukuran landscape
        
        # Tampilkan splash screen
        self.show_splash()
        
        # Setup UI
        self.init_ui()
        self.current_data = load_data()
        
        # Status bar
        self.statusBar().showMessage("Aplikasi siap digunakan", 3000)
        
        # Terapkan style
        self.set_app_style()
        
    def show_splash(self):
        if os.path.exists("splash.png"):
            splash = QSplashScreen(QPixmap("splash.png"))
            splash.show()
            QTimer.singleShot(2000, splash.close)
        
    def set_app_style(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f7fa;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QGroupBox {
                border: 1px solid #d1d5db;
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 15px;
                font-weight: bold;
                color: #374151;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
            QLabel {
                font-weight: 500;
                color: #374151;
                font-size: 13px;
            }
            QPushButton {
                background-color: #4f46e5;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: 500;
                font-size: 13px;
                min-width: 100px;
                transition: background-color 0.3s;
            }
            QPushButton:hover {
                background-color: #4338ca;
            }
            QPushButton:pressed {
                background-color: #3730a3;
            }
            QPushButton#danger {
                background-color: #ef4444;
            }
            QPushButton#danger:hover {
                background-color: #dc2626;
            }
            QPushButton#danger:pressed {
                background-color: #b91c1c;
            }
            QLineEdit, QComboBox, QDateEdit {
                padding: 8px;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                background-color: white;
                font-size: 13px;
                min-height: 20px;
            }
            QLineEdit:focus, QComboBox:focus, QDateEdit:focus {
                border: 1px solid #4f46e5;
                outline: none;
            }
            QTableWidget {
                background-color: white;
                alternate-background-color: #f9fafb;
                gridline-color: #e5e7eb;
                border: 1px solid #d1d5db;
                font-size: 13px;
                selection-background-color: #e0e7ff;
                selection-color: #1e40af;
            }
            QHeaderView::section {
                background-color: #4f46e5;
                color: white;
                padding: 8px;
                border: none;
                font-weight: 500;
                font-size: 13px;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 30px;
                border-left-width: 1px;
                border-left-color: #d1d5db;
                border-left-style: solid;
                border-top-right-radius: 6px;
                border-bottom-right-radius: 6px;
            }
            QComboBox QAbstractItemView {
                border: 1px solid #d1d5db;
                selection-background-color: #e0e7ff;
                selection-color: #1e40af;
            }
            QDateEdit::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 30px;
                border-left-width: 1px;
                border-left-color: #d1d5db;
                border-left-style: solid;
                border-top-right-radius: 6px;
                border-bottom-right-radius: 6px;
            }
            QStatusBar {
                background-color: #f3f4f6;
                color: #4b5563;
                border-top: 1px solid #d1d5db;
                font-size: 12px;
            }
        """)

    def init_ui(self):
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        # Layout utama menggunakan QHBoxLayout untuk landscape
        main_layout = QHBoxLayout(self.central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(15, 15, 15, 15)
        
        # Container kiri (form input)
        left_container = QWidget()
        left_layout = QVBoxLayout(left_container)
        left_layout.setSpacing(15)
        
        # Container kanan (tabel data)
        right_container = QWidget()
        right_layout = QVBoxLayout(right_container)
        right_layout.setSpacing(15)
        
        # Tambahkan kedua container ke layout utama
        main_layout.addWidget(left_container, 40)  # 40% width
        main_layout.addWidget(right_container, 60)  # 60% width
        
        # Buat form input di container kiri
        self.create_input_form(left_layout)
        self.create_search_section(left_layout)
        
        # Buat tabel di container kanan
        self.create_data_table(right_layout)
        self.create_action_buttons(right_layout)

        # Tampilkan data awal
        self.tampilkan_data()

    def create_input_form(self, parent_layout):
        form_group = QGroupBox("Form Data Kamar")
        form_layout = QGridLayout()
        form_layout.setSpacing(10)
        form_layout.setContentsMargins(15, 15, 15, 15)
        
        # Kolom 1
        form_layout.addWidget(QLabel("Nomor Kamar:"), 0, 0)
        self.input_no_kamar = QComboBox()
        self.input_no_kamar.addItems(nomor_kamar)
        form_layout.addWidget(self.input_no_kamar, 0, 1)

        form_layout.addWidget(QLabel("Nama Penghuni:"), 1, 0)
        self.input_nama_penghuni = QLineEdit()
        self.input_nama_penghuni.setPlaceholderText("Masukkan nama lengkap")
        form_layout.addWidget(self.input_nama_penghuni, 1, 1)

        form_layout.addWidget(QLabel("Nomor WhatsApp:"), 2, 0)
        self.input_nomor_whatsapp = QLineEdit()
        self.input_nomor_whatsapp.setPlaceholderText("628123456789")
        form_layout.addWidget(self.input_nomor_whatsapp, 2, 1)

        form_layout.addWidget(QLabel("Tanggal Masuk:"), 3, 0)
        self.input_tanggal_masuk = QDateEdit()
        self.input_tanggal_masuk.setDate(QDate.currentDate())
        self.input_tanggal_masuk.setCalendarPopup(True)
        self.input_tanggal_masuk.setDisplayFormat("dd/MM/yyyy")
        form_layout.addWidget(self.input_tanggal_masuk, 3, 1)

        # Kolom 2
        form_layout.addWidget(QLabel("Status Kamar:"), 0, 2)
        self.combo_status_kamar = QComboBox()
        self.combo_status_kamar.addItems(["Sendiri", "Berdua", "Kamar Kosong"])
        self.combo_status_kamar.currentTextChanged.connect(self.handle_status_kamar_change)
        form_layout.addWidget(self.combo_status_kamar, 0, 3)

        form_layout.addWidget(QLabel("Status Pembayaran:"), 1, 2)
        self.combo_status_pembayaran = QComboBox()
        self.combo_status_pembayaran.addItems(["Lunas", "Menunggak"])
        form_layout.addWidget(self.combo_status_pembayaran, 1, 3)

        form_layout.addWidget(QLabel("Harga Kamar:"), 2, 2)
        self.combo_harga_kamar = QComboBox()
        self.combo_harga_kamar.addItems(harga_kamar)
        form_layout.addWidget(self.combo_harga_kamar, 2, 3)

        # Tombol tambah memanjang 2 kolom
        self.button_tambah = QPushButton("Tambah/Update Data")
        self.button_tambah.clicked.connect(self.tambah_data)
        form_layout.addWidget(self.button_tambah, 4, 0, 1, 4)  # row, col, rowspan, colspan

        form_group.setLayout(form_layout)
        parent_layout.addWidget(form_group)

    def handle_status_kamar_change(self, status):
        disabled = (status == "Kamar Kosong")
        self.input_nama_penghuni.setDisabled(disabled)
        self.input_nomor_whatsapp.setDisabled(disabled)
        self.input_tanggal_masuk.setDisabled(disabled)
        self.combo_status_pembayaran.setDisabled(disabled)
        self.combo_harga_kamar.setDisabled(disabled)
        
        if disabled:
            self.input_nama_penghuni.clear()
            self.input_nomor_whatsapp.clear()
            self.input_tanggal_masuk.setDate(QDate.currentDate())
            self.combo_status_pembayaran.setCurrentIndex(0)
            self.combo_harga_kamar.setCurrentIndex(0)

    def create_search_section(self, parent_layout):
        search_group = QGroupBox("Pencarian Data")
        search_layout = QHBoxLayout()
        search_layout.setSpacing(10)
        search_layout.setContentsMargins(15, 15, 15, 15)
        
        self.combo_jenis_pencarian = QComboBox()
        self.combo_jenis_pencarian.addItems(["Nomor Kamar", "Nama Penghuni"])
        search_layout.addWidget(self.combo_jenis_pencarian)

        self.input_cari = QLineEdit()
        self.input_cari.setPlaceholderText("Masukkan nomor kamar atau nama...")
        self.input_cari.returnPressed.connect(self.cari_data)
        search_layout.addWidget(self.input_cari)

        self.button_cari = QPushButton("Cari")
        self.button_cari.clicked.connect(self.cari_data)
        search_layout.addWidget(self.button_cari)

        search_group.setLayout(search_layout)
        parent_layout.addWidget(search_group)

    def create_data_table(self, parent_layout):
        table_group = QGroupBox("Data Kamar Kost")
        table_layout = QVBoxLayout(table_group)
        
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "No Kamar", "Nama Penghuni", "Nomor WhatsApp", "Tanggal Masuk", 
            "Status Kamar", "Status Pembayaran", "Harga Kamar"
        ])
        
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.cellClicked.connect(self.tampilkan_data_terpilih)
        
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        
        table_layout.addWidget(self.table)
        parent_layout.addWidget(table_group)

    def create_action_buttons(self, parent_layout):
        button_group = QWidget()
        button_layout = QHBoxLayout(button_group)
        button_layout.setSpacing(10)
        button_layout.setContentsMargins(0, 0, 0, 0)
        
        self.button_edit = QPushButton("Edit Data")
        self.button_edit.clicked.connect(self.edit_data)
        button_layout.addWidget(self.button_edit)

        self.button_hapus = QPushButton("Hapus Data")
        self.button_hapus.setObjectName("danger")
        self.button_hapus.clicked.connect(self.hapus_data)
        button_layout.addWidget(self.button_hapus)

        self.button_import = QPushButton("Import Data")
        self.button_import.clicked.connect(self.import_data)
        button_layout.addWidget(self.button_import)

        self.button_export = QPushButton("Export Data")
        self.button_export.clicked.connect(self.export_data)
        button_layout.addWidget(self.button_export)

        self.button_save = QPushButton("Simpan Data")
        self.button_save.clicked.connect(self.save_data_manual)
        button_layout.addWidget(self.button_save)

        parent_layout.addWidget(button_group)

    def tampilkan_data(self):
        self.current_data = load_data()
        self.tampilkan_hasil_pencarian(self.current_data)
        self.statusBar().showMessage("Data ditampilkan", 3000)

    def tampilkan_hasil_pencarian(self, df):
        try:
            self.table.setRowCount(len(df))
            for i, row in df.iterrows():
                for j, value in enumerate(row):
                    display_value = str(value) if pd.notna(value) else ""
                    item = QTableWidgetItem(display_value)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    
                    # Warna untuk status pembayaran
                    if j == 5:  # Kolom Status Pembayaran
                        if "Menunggak" in str(value):
                            item.setBackground(QColor(254, 226, 226))  # Merah muda
                            item.setForeground(QColor(220, 38, 38))    # Merah
                        elif "Lunas" in str(value):
                            item.setBackground(QColor(220, 252, 231))  # Hijau muda
                            item.setForeground(QColor(22, 163, 74))    # Hijau
                    
                    self.table.setItem(i, j, item)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Gagal menampilkan data: {str(e)}")
            self.statusBar().showMessage("Gagal menampilkan data", 3000)

    def tampilkan_data_terpilih(self, row, column):
        try:
            no_kamar = self.table.item(row, 0).text()
            df = self.current_data[self.current_data['No Kamar'] == no_kamar]
            
            if not df.empty:
                data = df.iloc[0]
                
                index = self.input_no_kamar.findText(data['No Kamar'])
                if index >= 0:
                    self.input_no_kamar.setCurrentIndex(index)
                
                self.input_nama_penghuni.setText(str(data['Nama Penghuni']))
                self.input_nomor_whatsapp.setText(str(data['Nomor WhatsApp']))
                
                if data['Tanggal Masuk'] and str(data['Tanggal Masuk']) != 'nan':
                    try:
                        date = QDate.fromString(data['Tanggal Masuk'], "yyyy-MM-dd")
                        if not date.isValid():
                            date = QDate.fromString(data['Tanggal Masuk'], "dd/MM/yyyy")
                        if date.isValid():
                            self.input_tanggal_masuk.setDate(date)
                    except:
                        pass
                
                self.combo_status_kamar.setCurrentText(data['Status Kamar'])
                self.combo_status_pembayaran.setCurrentText(data['Status Pembayaran'])
                self.combo_harga_kamar.setCurrentText(data['Harga Kamar'])
                
        except Exception as e:
            print(f"Error menampilkan data terpilih: {e}")

    def cari_data(self):
        keyword = self.input_cari.text().strip()
        if not keyword:
            QMessageBox.warning(self, "Peringatan", "Masukkan kata kunci pencarian!")
            return

        df = load_data()
        if df.empty:
            QMessageBox.information(self, "Info", "Tidak ada data yang tersedia.")
            return

        jenis_pencarian = self.combo_jenis_pencarian.currentText()

        try:
            if jenis_pencarian == "Nomor Kamar":
                result = df[df['No Kamar'].str.strip().str.upper() == keyword.strip().upper()]
            else:
                result = df[df['Nama Penghuni'].str.strip().str.upper().str.contains(
                    keyword.strip().upper(), na=False)]
            
            if result.empty:
                QMessageBox.information(self, "Pencarian", 
                    f"Data dengan {jenis_pencarian} '{keyword}' tidak ditemukan.")
                self.statusBar().showMessage("Pencarian tidak ditemukan", 3000)
            else:
                self.current_data = result
                self.tampilkan_hasil_pencarian(result)
                self.statusBar().showMessage(f"Menampilkan hasil pencarian: {keyword}", 3000)
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat mencari: {str(e)}")
            self.statusBar().showMessage("Gagal melakukan pencarian", 3000)

    def tambah_data(self):
        no_kamar = self.input_no_kamar.currentText().strip().upper()
        status_kamar = self.combo_status_kamar.currentText()

        if status_kamar != "Kamar Kosong":
            nama = self.input_nama_penghuni.text().strip()
            whatsapp = self.input_nomor_whatsapp.text().strip()
            
            if not nama:
                QMessageBox.warning(self, "Peringatan", "Nama penghuni harus diisi!")
                self.statusBar().showMessage("Gagal: Nama penghuni harus diisi", 3000)
                return
            if not whatsapp:
                QMessageBox.warning(self, "Peringatan", "Nomor WhatsApp harus diisi!")
                self.statusBar().showMessage("Gagal: Nomor WhatsApp harus diisi", 3000)
                return

        new_data = {
            'No Kamar': no_kamar,
            'Nama Penghuni': self.input_nama_penghuni.text().strip() if status_kamar != "Kamar Kosong" else '',
            'Nomor WhatsApp': format_whatsapp_number(self.input_nomor_whatsapp.text().strip()) if status_kamar != "Kamar Kosong" else '',
            'Tanggal Masuk': self.input_tanggal_masuk.date().toString("dd/MM/yyyy") if status_kamar != "Kamar Kosong" else '',
            'Status Kamar': status_kamar,
            'Status Pembayaran': self.combo_status_pembayaran.currentText() if status_kamar != "Kamar Kosong" else '',
            'Harga Kamar': self.combo_harga_kamar.currentText() if status_kamar != "Kamar Kosong" else ''
        }

        df = load_data()
        
        existing_index = df.index[df['No Kamar'].str.strip().str.upper() == no_kamar].tolist()
        
        if existing_index:
            df.loc[existing_index[0]] = new_data
            action = "diupdate"
        else:
            df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
            action = "ditambahkan"
        
        if save_data(df):
            QMessageBox.information(self, "Sukses", f"Data berhasil {action}!")
            self.statusBar().showMessage(f"Data kamar {no_kamar} berhasil {action}", 3000)
            self.tampilkan_data()
            self.clear_form()
        else:
            QMessageBox.critical(self, "Error", "Gagal menyimpan data!")
            self.statusBar().showMessage("Gagal menyimpan data", 3000)

    def edit_data(self):
        selected_row = self.table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, "Peringatan", "Pilih data yang akan diedit di tabel!")
            self.statusBar().showMessage("Pilih data terlebih dahulu untuk diedit", 3000)
            return
        
        self.tambah_data()

    def hapus_data(self):
        selected_row = self.table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, "Peringatan", "Pilih data yang akan dihapus di tabel!")
            self.statusBar().showMessage("Pilih data terlebih dahulu untuk dihapus", 3000)
            return

        no_kamar = self.table.item(selected_row, 0).text()
        
        reply = QMessageBox.question(
            self, 'Konfirmasi', 
            f'Apakah Anda yakin ingin menghapus data kamar {no_kamar}?',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            df = load_data()
            df = df[df['No Kamar'].str.strip().str.upper() != no_kamar.strip().upper()]
            
            if save_data(df):
                QMessageBox.information(self, "Sukses", "Data berhasil dihapus!")
                self.statusBar().showMessage(f"Data kamar {no_kamar} berhasil dihapus", 3000)
                self.tampilkan_data()
                self.clear_form()
            else:
                QMessageBox.critical(self, "Error", "Gagal menghapus data!")
                self.statusBar().showMessage("Gagal menghapus data", 3000)

    def import_data(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Pilih File Data", "", 
            "Excel Files (*.xlsx *.xls);;CSV Files (*.csv)"
        )

        if not file_path:
            return

        try:
            if file_path.endswith(('.xlsx', '.xls')):
                new_data = pd.read_excel(file_path, dtype={'No Kamar': str, 'Nomor WhatsApp': str})
            elif file_path.endswith('.csv'):
                new_data = pd.read_csv(file_path, dtype={'No Kamar': str, 'Nomor WhatsApp': str})
            else:
                QMessageBox.warning(self, "Peringatan", "Format file tidak didukung!")
                return

            required_columns = [
                'No Kamar', 'Nama Penghuni', 'Nomor WhatsApp', 'Tanggal Masuk',
                'Status Kamar', 'Status Pembayaran', 'Harga Kamar'
            ]
            
            if not all(col in new_data.columns for col in required_columns):
                QMessageBox.warning(self, "Peringatan", 
                    "File tidak memiliki semua kolom yang diperlukan!")
                return

            current_data = load_data()
            combined_data = pd.concat([current_data, new_data], ignore_index=True)
            combined_data = combined_data.drop_duplicates('No Kamar', keep='last')
            
            if save_data(combined_data):
                QMessageBox.information(self, "Sukses", "Data berhasil diimpor!")
                self.statusBar().showMessage("Data berhasil diimpor", 3000)
                self.tampilkan_data()
            else:
                QMessageBox.critical(self, "Error", "Gagal menyimpan data yang diimpor!")
                self.statusBar().showMessage("Gagal mengimpor data", 3000)
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat mengimpor: {str(e)}")
            self.statusBar().showMessage("Gagal mengimpor data", 3000)

    def export_data(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Simpan Data", "", 
            "Excel Files (*.xlsx);;All Files (*)"
        )

        if not file_path:
            return

        if not file_path.endswith('.xlsx'):
            file_path += '.xlsx'

        try:
            df = load_data()
            df.to_excel(file_path, index=False)
            QMessageBox.information(self, "Sukses", f"Data berhasil diekspor ke:\n{file_path}")
            self.statusBar().showMessage(f"Data berhasil diekspor ke {file_path}", 3000)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Gagal mengekspor data: {str(e)}")
            self.statusBar().showMessage("Gagal mengekspor data", 3000)

    def save_data_manual(self):
        if save_data(load_data()):
            QMessageBox.information(self, "Sukses", "Data berhasil disimpan!")
            self.statusBar().showMessage("Data berhasil disimpan", 3000)
        else:
            QMessageBox.critical(self, "Error", "Gagal menyimpan data!")
            self.statusBar().showMessage("Gagal menyimpan data", 3000)

    def clear_form(self):
        self.input_nama_penghuni.clear()
        self.input_nomor_whatsapp.clear()
        self.input_tanggal_masuk.setDate(QDate.currentDate())
        self.combo_status_kamar.setCurrentIndex(0)
        self.combo_status_pembayaran.setCurrentIndex(0)
        self.combo_harga_kamar.setCurrentIndex(0)

    def closeEvent(self, event):
        save_data(load_data())
        event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = KostApp()
    window.show()
    sys.exit(app.exec_())