import sys
import winreg as reg
import re
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, 
    QVBoxLayout, QProgressBar, QMessageBox, QTextEdit, QFrame
)
from PyQt5.QtGui import QFont, QIcon, QColor
from PyQt5.QtCore import Qt, QSize
import resources_rc

class KLMApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('LIBRARY MANAGER PRO')
        self.setFixedSize(550, 550) # Ajustado para acomodar o rodapé
        
        # Ícone
        icon_path = ":/IMG/favicom.png"
        self.setWindowIcon(QIcon(icon_path))

        # --- ESTILO GLOBAL (DARK MODERN) ---
        self.setStyleSheet("""
            QWidget {
                background-color: #0f0f0f;
                color: #e0e0e0;
                font-family: 'Segoe UI', sans-serif;
            }
            QLabel {
                font-size: 12px;
                font-weight: bold;
                color: #888888;
                margin-bottom: 2px;
                text-transform: uppercase;
            }
            QLineEdit {
                background-color: #1a1a1a;
                border: 1px solid #333333;
                border-radius: 8px;
                padding: 10px;
                color: #ffffff;
                font-family: 'Consolas';
            }
            QLineEdit:focus {
                border: 1px solid #0078d4;
                background-color: #222222;
            }
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 20px;
                padding: 12px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1a8ad4;
            }
            QPushButton:pressed {
                transform: scale(0.98);
            }
            QProgressBar {
                border: 1px solid #333333;
                border-radius: 10px;
                text-align: center;
                background-color: #1a1a1a;
                height: 12px;
                color: transparent;
            }
            QProgressBar::chunk {
                background-color: #0078d4;
                border-radius: 10px;
            }
            QTextEdit {
                background-color: #050505;
                border: 1px solid #222222;
                border-radius: 8px;
                color: #00ff41; 
                font-family: 'Consolas';
                font-size: 11px;
            }
            #FooterLabel {
                color: #555555;
                font-size: 11px;
                font-weight: bold;
            }
        """)

        layout = QVBoxLayout()
        layout.setContentsMargins(35, 35, 35, 15)
        layout.setSpacing(12)

        # Nome da Biblioteca
        self.name_label = QLabel('Nome da Biblioteca')
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Ex: Kontakt Factory Library")

        # Caminho da Biblioteca
        self.path_label = QLabel('Caminho da Pasta')
        self.path_input = QLineEdit()
        self.path_input.setPlaceholderText("C:/Libraries/Samples/LibraryName")

        # Botão Iniciar
        self.start_button = QPushButton('IMPORTAR PARA O REGISTRO')
        self.start_button.setCursor(Qt.PointingHandCursor)
        
        icon_btn = QIcon(":/IMG/iniciar.png")
        if not icon_btn.isNull():
            self.start_button.setIcon(icon_btn)
            self.start_button.setIconSize(QSize(20, 20))
        
        self.start_button.clicked.connect(self.start_process)

        # Separador visual
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("background-color: #222222;")

        # Barra de Progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)

        # Log de Saída
        self.imported_libraries = QTextEdit()
        self.imported_libraries.setReadOnly(True)
        self.imported_libraries.setPlaceholderText("Console pronto...")

        # Rodapé
        self.footer_label = QLabel("Desenvolvido por: Daniel Boechat")
        self.footer_label.setObjectName("FooterLabel")
        self.footer_label.setAlignment(Qt.AlignCenter)

        # Montagem do Layout
        layout.addWidget(self.name_label)
        layout.addWidget(self.name_input)
        layout.addWidget(self.path_label)
        layout.addWidget(self.path_input)
        layout.addWidget(self.start_button)
        layout.addSpacing(15)
        layout.addWidget(line)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.imported_libraries)
        layout.addStretch() # Empurra o rodapé para o fim
        layout.addWidget(self.footer_label)

        self.setLayout(layout)

    # --- LÓGICA MANTIDA ---
    def start_process(self):
        library_name = self.name_input.text()
        library_path = self.path_input.text()
        library_name = self.remove_accents(library_name)

        if not library_name or not library_path:
            QMessageBox.warning(self, 'Atenção', 'Por favor, preencha todos os campos.')
            return

        try:
            self.add_library_to_registry(library_name, library_path)
            self.create_library_files(library_name, library_path)
            
            for i in range(101):
                self.progress_bar.setValue(i)
                QApplication.processEvents()

            self.imported_libraries.append(f'>> [OK] "{library_name}" registrada com sucesso.')
        except Exception as e:
            QMessageBox.critical(self, 'Erro', f'Erro: {str(e)}')

    def remove_accents(self, input_str):
        return re.sub(r'[^\x00-\x7F]+', '', input_str)

    def add_library_to_registry(self, library_name, library_path):
        native_instruments_registry_path = r"SOFTWARE\Native Instruments"
        kontakt_content_registry_path = r"SOFTWARE\Native Instruments\Content"
        user_registry_path = r"SOFTWARE\Native Instruments"
        try:
            key = reg.CreateKey(reg.HKEY_LOCAL_MACHINE, native_instruments_registry_path)
            library_key = reg.CreateKey(key, library_name)
            reg.SetValueEx(library_key, "ContentDir", 0, reg.REG_SZ, library_path)
            reg.SetValueEx(library_key, "Visibility", 0, reg.REG_DWORD, 3)
            reg.CloseKey(library_key)
            reg.CloseKey(key)

            key = reg.CreateKey(reg.HKEY_LOCAL_MACHINE, kontakt_content_registry_path)
            next_id = self.get_next_kontakt_library_id(key)
            k2lib_key_name = f"k2lib{next_id:04d}"
            reg.SetValueEx(key, k2lib_key_name, 0, reg.REG_SZ, library_name)
            reg.CloseKey(key)

            user_key = reg.CreateKey(reg.HKEY_CURRENT_USER, user_registry_path)
            user_library_key = reg.CreateKey(user_key, library_name)
            user_list_index = self.get_next_user_list_index(user_key)
            reg.SetValueEx(user_library_key, "UserListIndex", 0, reg.REG_DWORD, user_list_index)
            reg.SetValueEx(user_library_key, "UserRemoved", 0, reg.REG_DWORD, 0)
            reg.CloseKey(user_library_key)
            reg.CloseKey(user_key)
        except Exception as e:
            raise Exception(f"Registro: {str(e)}")

    def get_next_kontakt_library_id(self, key):
        try:
            value_count = reg.QueryInfoKey(key)[1]
            existing_ids = set()
            for i in range(value_count):
                value_name, _, _ = reg.EnumValue(key, i)
                if value_name.startswith("k2lib"):
                    existing_ids.add(value_name)
            next_id = 10  
            while f"k2lib{next_id:04d}" in existing_ids:
                next_id += 1  
            return next_id
        except Exception as e:
            raise Exception(f"ID: {str(e)}")

    def get_next_user_list_index(self, user_key):
        try:
            subkey_count = reg.QueryInfoKey(user_key)[0]
            used_indexes = set()
            for i in range(subkey_count):
                subkey_name = reg.EnumKey(user_key, i)
                subkey = reg.OpenKey(user_key, subkey_name)
                try:
                    index_value, _ = reg.QueryValueEx(subkey, "UserListIndex")
                    used_indexes.add(index_value)
                except FileNotFoundError:
                    continue
                reg.CloseKey(subkey)
            next_index = 1
            while next_index in used_indexes:
                next_index += 1
            return next_index
        except Exception as e:
            raise Exception(f"Index: {str(e)}")

    def create_library_files(self, library_name, library_path):
        try:
            nkc_path = os.path.join(library_path, f"{library_name}_info.nkc")
            nkx_path = os.path.join(library_path, f"{library_name}_info.nkx")
            with open(nkc_path, 'wb') as f:
                f.write(b'\0' * 1024)
            open(nkx_path, 'wb').close()
        except Exception as e:
            raise Exception(f"Arquivos: {str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("Fusion") 
    klm = KLMApp()
    klm.show()
    sys.exit(app.exec_())