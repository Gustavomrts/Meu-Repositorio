import sys
import calendar
import sqlite3
import os
import hashlib
import locale  
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QDateEdit, QTimeEdit, QStackedWidget,
    QTableWidget, QTableWidgetItem, QMessageBox, QDialog, QDialogButtonBox,
    QHeaderView, QAbstractItemView, QFrame, QSizePolicy, QSpacerItem,
    QFormLayout, QListWidget, QInputDialog, QFileDialog, QListWidgetItem, QCheckBox
)
from PyQt5.QtWidgets import QTabWidget
from PyQt5.QtCore import Qt, QDate, QTime, QTimer, QDateTime, QSettings
from PyQt5.QtGui import QFont, QIcon, QColor, QBrush, QDoubleValidator, QPainter, QPixmap
from PyQt5.QtChart import QChart, QChartView, QBarSeries, QBarSet, QPieSeries, QBarCategoryAxis, QValueAxis
from PyQt5.QtWidgets import QToolTip
from PyQt5.QtGui import QCursor
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtCore import QUrl

# ==============================================

class CustomChartView(QChartView):
    def __init__(self, chart, parent=None):
        super().__init__(chart, parent)
        self.chart = chart
        self.setRenderHint(QPainter.Antialiasing)

    def drawForeground(self, painter, rect):
        super().drawForeground(painter, rect)

        if not self.chart:
            return

        series_list = self.chart.series()
        if not series_list:
            return

        for series in series_list:
            if not isinstance(series, QBarSeries):
                continue

            bars = series.barSets()
            axis_x = self.chart.axisX(series)
            axis_y = self.chart.axisY(series)
    
            for i, bar_set in enumerate(bars):
                for j, value in enumerate(bar_set):
                    bar_rect = series.barGeometry(i, j)
                    if not bar_rect.isValid():
                        continue
    
                    text = str(int(value))
                    font = painter.font()
                    font.setPointSize(10)
                    painter.setFont(font)
                    painter.setPen(Qt.white)
    
                    # Ajusta a posição do texto para cima da barra
                    x = bar_rect.center().x() - painter.fontMetrics().width(text) / 2
                    y = bar_rect.top() - 10 if bar_rect.top() > 10 else bar_rect.top() + 15
        
                    painter.drawText(x, y, text)





# Configuração do locale para formatação monetária
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')  # Para sistemas com suporte a pt_BR
except locale.Error:
    try:
        locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil')  # Para Windows
    except:
        locale.setlocale(locale.LC_ALL, '')  # Usa o padrão do sistema como fallback

class LoginDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Login - Gustavo Martins")
        self.setFixedSize(400, 450)
        self.setStyleSheet("""
            QDialog {
                background-color: #2c3e50;
            }
            QLabel {
                color: white;
                font-size: 16px;
            }
            QLineEdit {
                background: #34495e;
                color: white;
                border: 1px solid #2c3e50;
                padding: 8px;
                border-radius: 4px;
                font-size: 14px;
            }
            QPushButton {
                background: #3498db;
                color: white;
                padding: 10px;
                border-radius: 4px;
                font-size: 16px;
                min-width: 100px;
            }
            QPushButton:hover {
                background: #2980b9;
            }
            QCheckBox {
                color: white;
                spacing: 5px;
            }
            QTabWidget::pane {
                border: 0;
            }
            QTabBar::tab {
                background: #34495e;
                color: white;
                padding: 8px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background: #3498db;
            }
        """)
        
        self.settings = QSettings("GustavoMartins", "CRMApp")
        
        self.tab_widget = QTabWidget()
        
        # Aba de Login
        self.login_tab = QWidget()
        self.setup_login_tab()
        self.tab_widget.addTab(self.login_tab, "Login")
        
        # Aba de Cadastro
        self.register_tab = QWidget()
        self.setup_register_tab()
        self.tab_widget.addTab(self.register_tab, "Criar Conta")
        
        # Aba de Login Google
        self.google_tab = QWidget()
        self.setup_google_tab()
        self.tab_widget.addTab(self.google_tab, "Login Google")
        
        layout = QVBoxLayout(self)
        layout.addWidget(self.tab_widget)
    
    def setup_login_tab(self):
        layout = QVBoxLayout(self.login_tab)
    
        # Logo
        logo_label = QLabel()
        logo_path = os.path.join(os.getcwd(), "logo.png")
        logo_pixmap = QPixmap(logo_path)
        if not logo_pixmap.isNull():
            logo_label.setPixmap(logo_pixmap.scaled(150, 150, Qt.KeepAspectRatio))
        else:
            logo_label.setText("Logo")
        logo_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(logo_label)
        
        # Botão do Google
        btn_google = QPushButton("Entrar com Google")
        btn_google.setStyleSheet("""
            QPushButton {
                background: #4285F4;
                color: white;
                padding: 12px;
                border-radius: 4px;
            font-size: 16px;
                min-width: 200px;
            }
            QPushButton:hover {
                background: #3367D6;
            }
        """)
        btn_google.setCursor(Qt.PointingHandCursor)
        btn_google.clicked.connect(self.login_with_google)
        layout.addWidget(btn_google, alignment=Qt.AlignCenter)
        
        # Separador
        separator = QLabel("ou")
        separator.setAlignment(Qt.AlignCenter)
        layout.addWidget(separator)
        
        # Formulário de login tradicional
        form_layout = QFormLayout()
    
        
        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("Digite seu e-mail")
        
        # Carrega e-mail salvo
        saved_email = self.settings.value("saved_email", "")
        if saved_email:
            self.username_input.setText(saved_email)
        
        form_layout.addRow("E-mail:", self.username_input)
        
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Digite sua senha")
        self.password_input.setEchoMode(QLineEdit.Password)
        form_layout.addRow("Senha:", self.password_input)
        
        # Checkbox para manter conectado
        self.remember_check = QCheckBox("Manter conectado")
        self.remember_check.setChecked(self.settings.value("remember_me", False, type=bool))
        form_layout.addRow(self.remember_check)
        
        layout.addLayout(form_layout)
        
        # Botões
        btn_box = QDialogButtonBox()
        btn_box.addButton("Login", QDialogButtonBox.AcceptRole)
        btn_box.addButton("Cancelar", QDialogButtonBox.RejectRole)
        btn_box.accepted.connect(self.verify_login)
        btn_box.rejected.connect(self.reject)
        
        layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))
        layout.addWidget(btn_box, alignment=Qt.AlignCenter)
    
    def setup_google_tab(self):
        """Configura a aba de login com Google"""
        layout = QVBoxLayout(self.google_tab)
        
        # Logo
        logo_label = QLabel()
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(__file__)
            
        logo_path = os.path.join(base_path, "logo.png")
        logo_pixmap = QPixmap(logo_path)

        if not logo_pixmap.isNull():
            logo_label.setPixmap(logo_pixmap.scaled(150, 150, Qt.KeepAspectRatio))
        else:
            logo_label.setText("Logo")
        logo_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(logo_label)
        
        # Botão de login com Google
        btn_google = QPushButton("Entrar com Google")
        btn_google.setStyleSheet("""
            QPushButton {
                background: #4285F4;
                color: white;
                padding: 12px;
                border-radius: 4px;
                font-size: 16px;
                min-width: 200px;
            }
            QPushButton:hover {
                background: #3367D6;
            }
        """)
        btn_google.setCursor(Qt.PointingHandCursor)
        btn_google.clicked.connect(self.login_with_google)
        
        layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))
        layout.addWidget(btn_google, alignment=Qt.AlignCenter)
        layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))
    
    def login_with_google(self):
        """Implementa o login com Google"""
        try:
            from google_auth_oauthlib.flow import InstalledAppFlow
            from google.oauth2.credentials import Credentials
            
            # Configurações OAuth
            SCOPES = ['https://www.googleapis.com/auth/userinfo.profile',
                     'https://www.googleapis.com/auth/userinfo.email',
                     'openid']
            
            # Verifica se já existe um token válido
            if os.path.exists('google_token.json'):
                creds = Credentials.from_authorized_user_file('google_token.json', SCOPES)
                if creds and creds.valid:
                    self.accept_google_login(creds)
                    return
            
            # Inicia o fluxo de autenticação
            flow = InstalledAppFlow.from_client_secrets_file(
                'google_credentials.json', SCOPES)
            
            creds = flow.run_local_server(port=0)
            
            # Salva as credenciais para a próxima vez
            with open('google_token.json', 'w') as token:
                token.write(creds.to_json())
            
            self.accept_google_login(creds)
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha no login com Google:\n{str(e)}")
    
    def accept_google_login(self, creds):
        """Processa o login bem-sucedido com Google"""
        try:
            from googleapiclient.discovery import build
            
            # Obtém informações do usuário
            service = build('oauth2', 'v2', credentials=creds)
            user_info = service.userinfo().get().execute()
            
            email = user_info.get('email')
            name = user_info.get('name', 'Usuário Google')
            
            # Verifica se o usuário já existe no sistema
            db_path = f"crm_database_{email}.db"
            if not os.path.exists(db_path):
                # Cria o banco de dados para o novo usuário
                self.criar_banco_google(email, name)
            
            # Salva as credenciais
            self.settings.setValue("saved_email", email)
            self.settings.setValue("remember_me", True)
            
            self.accept()
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao obter informações do usuário:\n{str(e)}")
    
    def criar_banco_google(self, email, nome):
        """Cria um novo banco de dados para usuário do Google"""
        try:
            conn = sqlite3.connect(f"crm_database_{email}.db")
            cursor = conn.cursor()
            
            # Tabela de clientes
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS clientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                empresa TEXT,
                telefone TEXT,
                email TEXT,
                status TEXT NOT NULL,
                valor_proposta REAL,
                data_cadastro TEXT,
                data_proxima_acao TEXT,
                ultima_atualizacao TEXT,
                origem TEXT,
                responsavel TEXT,
                probabilidade INTEGER,
                etapa_negocio TEXT,
                observacoes TEXT,
                usuario_id TEXT
            )
            """)
            
            # Tabela de histórico de interações
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS interacoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cliente_id INTEGER,
                tipo TEXT,
                descricao TEXT,
                data TEXT,
                duracao INTEGER,
                resultado TEXT,
                usuario_id TEXT,
                FOREIGN KEY(cliente_id) REFERENCES clientes(id)
            )
            """)
            
            # Tabela de agendamentos
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS agendamentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo TEXT,
                descricao TEXT,
                data_inicio TEXT,
                data_fim TEXT,
                cliente_id INTEGER,
                tipo TEXT,
                status TEXT,
                usuario_id TEXT,
                FOREIGN KEY(cliente_id) REFERENCES clientes(id)
            )
            """)
            
            # Tabela de usuários
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                sobrenome TEXT NOT NULL,
                email TEXT NOT NULL UNIQUE,
                senha TEXT NOT NULL,
                data_cadastro TEXT,
                is_admin INTEGER DEFAULT 0,
                google_id TEXT
            )
            """)
            
            # Insere o novo usuário (sem senha, pois é login com Google)
            cursor.execute("""
                INSERT INTO usuarios (nome, email, data_cadastro, google_id)
                VALUES (?, ?, datetime('now'), ?)
            """, (nome, email, "google_"+hashlib.sha256(email.encode()).hexdigest()))
            
            conn.commit()
            conn.close()
            
        except Exception as e:
            raise Exception(f"Erro ao criar banco de dados: {str(e)}")
    
    def verify_login(self):
        """Verifica as credenciais de login"""
        email = self.username_input.text().strip()
        senha = self.password_input.text()
        
        if not email or not senha:
            QMessageBox.warning(self, "Aviso", "Por favor, preencha todos os campos!")
            return
            
        db_path = f"crm_database_{email}.db"
        if not os.path.exists(db_path):
            QMessageBox.critical(self, "Erro", "E-mail não cadastrado!")
            return
            
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # Verifica as credenciais (senha em hash)
            cursor.execute("SELECT senha, is_admin FROM usuarios WHERE email = ?", (email,))
            result = cursor.fetchone()
            
            if not result:
                QMessageBox.critical(self, "Erro", "E-mail não encontrado!")
                return
                
            # Verifica a senha (agora usando hash)
            hashed_password = hashlib.sha256(senha.encode()).hexdigest()
            if result[0] != hashed_password:
                QMessageBox.critical(self, "Erro", "Senha incorreta!")
                return
                
            # Salva as preferências
            self.settings.setValue("saved_email", email if self.remember_check.isChecked() else "")
            self.settings.setValue("remember_me", self.remember_check.isChecked())
            self.settings.setValue("is_admin", result[1])
            
            conn.close()
            self.accept()
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao verificar login:\n{str(e)}")
    
    def setup_register_tab(self):
        layout = QVBoxLayout(self.register_tab)
        
        # Formulário de cadastro
        form_layout = QFormLayout()
        
        self.reg_nome_input = QLineEdit()
        self.reg_nome_input.setPlaceholderText("Digite seu nome")
        form_layout.addRow("Nome:", self.reg_nome_input)
        
        self.reg_sobrenome_input = QLineEdit()
        self.reg_sobrenome_input.setPlaceholderText("Digite seu sobrenome")
        form_layout.addRow("Sobrenome:", self.reg_sobrenome_input)
        
        self.reg_email_input = QLineEdit()
        self.reg_email_input.setPlaceholderText("Digite seu e-mail")
        form_layout.addRow("E-mail:", self.reg_email_input)
        
        self.reg_senha_input = QLineEdit()
        self.reg_senha_input.setPlaceholderText("Digite sua senha")
        self.reg_senha_input.setEchoMode(QLineEdit.Password)
        form_layout.addRow("Senha:", self.reg_senha_input)
        
        self.reg_confirma_senha_input = QLineEdit()
        self.reg_confirma_senha_input.setPlaceholderText("Confirme sua senha")
        self.reg_confirma_senha_input.setEchoMode(QLineEdit.Password)
        form_layout.addRow("Confirmar Senha:", self.reg_confirma_senha_input)
        
        layout.addLayout(form_layout)
        
        # Botões
        btn_box = QDialogButtonBox()
        btn_box.addButton("Criar Conta", QDialogButtonBox.AcceptRole)
        btn_box.addButton("Cancelar", QDialogButtonBox.RejectRole)
        btn_box.accepted.connect(self.criar_conta)
        btn_box.rejected.connect(self.reject)
        
        layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))
        layout.addWidget(btn_box, alignment=Qt.AlignCenter)
    
    def criar_conta(self):
        """Valida e cria uma nova conta de usuário"""
        nome = self.reg_nome_input.text().strip()
        sobrenome = self.reg_sobrenome_input.text().strip()
        email = self.reg_email_input.text().strip()
        senha = self.reg_senha_input.text()
        confirma_senha = self.reg_confirma_senha_input.text()
        
        if not nome or not sobrenome or not email or not senha:
            QMessageBox.warning(self, "Aviso", "Todos os campos são obrigatórios!")
            return
            
        if senha != confirma_senha:
            QMessageBox.warning(self, "Aviso", "As senhas não coincidem!")
            return
            
        # Validação de e-mail simples
        if "@" not in email or "." not in email:
            QMessageBox.warning(self, "Aviso", "Por favor, insira um e-mail válido!")
            return
            
        # Verifica se o usuário já existe
        db_path = f"crm_database_{email}.db"
        if os.path.exists(db_path):
            QMessageBox.warning(self, "Aviso", "Já existe uma conta com este e-mail!")
            return
            
        # Cria o banco de dados para o novo usuário
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # Tabela de clientes
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS clientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                empresa TEXT,
                telefone TEXT,
                email TEXT,
                status TEXT NOT NULL,
                valor_proposta REAL,
                data_cadastro TEXT,
                data_proxima_acao TEXT,
                ultima_atualizacao TEXT,
                origem TEXT,
                responsavel TEXT,
                probabilidade INTEGER,
                etapa_negocio TEXT,
                observacoes TEXT,
                usuario_id TEXT
            )
            """)
            
            # Tabela de histórico de interações
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS interacoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cliente_id INTEGER,
                tipo TEXT,
                descricao TEXT,
                data TEXT,
                duracao INTEGER,
                resultado TEXT,
                usuario_id TEXT,
                FOREIGN KEY(cliente_id) REFERENCES clientes(id)
            )
            """)
            
            # Tabela de agendamentos
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS agendamentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo TEXT,
                descricao TEXT,
                data_inicio TEXT,
                data_fim TEXT,
                cliente_id INTEGER,
                tipo TEXT,
                status TEXT,
                usuario_id TEXT,
                FOREIGN KEY(cliente_id) REFERENCES clientes(id)
            )
            """)
            
            # Tabela de usuários
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                sobrenome TEXT NOT NULL,
                email TEXT NOT NULL UNIQUE,
                senha TEXT NOT NULL,
                data_cadastro TEXT,
                is_admin INTEGER DEFAULT 0,
                google_id TEXT
            )
            """)
            
            # Insere o novo usuário com senha em hash
            hashed_password = hashlib.sha256(senha.encode()).hexdigest()
            cursor.execute("""
                INSERT INTO usuarios (nome, sobrenome, email, senha, data_cadastro)
                VALUES (?, ?, ?, ?, datetime('now'))
            """, (nome, sobrenome, email, hashed_password))
            
            conn.commit()
            conn.close()
            
            QMessageBox.information(self, "Sucesso", "Conta criada com sucesso!\nVocê pode fazer login agora.")
            self.tab_widget.setCurrentIndex(0)  # Volta para a aba de login
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao criar conta:\n{str(e)}")
            if 'conn' in locals():
                conn.close()
    
    def get_credentials(self):
        return self.username_input.text(), self.password_input.text()

    def show_bitrix(self):
        """Placeholder para futura integração com Bitrix24"""
        QMessageBox.information(
            self, 
            "Integração Bitrix24",
            "Esta funcionalidade está em desenvolvimento e será implementada em breve.\n\n"
            "A integração com Bitrix24 permitirá:\n"
            "- Sincronização de contatos\n"
            "Importação de oportunidades\n"
            "Gestão de pipelines",
           QMessageBox.Ok
        )

class CRMApp(QMainWindow):
    def __init__(self, email):
        super().__init__()
        self.setWindowTitle("CODash")
        self.setGeometry(100, 100, 1400, 800)
        self.email_usuario = email  # Armazena o email do usuário logado
        self.chart3 = None
        self.agendamentos_notificados = set()

        # Configuração do ícone
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(__file__)
        icon_path = os.path.join(base_path, "icon.png")
        self.setWindowIcon(QIcon(icon_path))

        # Configuração do banco de dados
        self.init_db()

        # Configuração da interface
        self.init_ui()
        self.setup_dark_mode()
        

        # Timer para verificar agendamentos
        self.notification_timer = QTimer(self)
        self.notification_timer.timeout.connect(self.verificar_agendamentos_proximos)
        self.notification_timer.start(60000)  # Verifica a cada minuto
        
        # Timer para sincronizar com Google Agenda
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.sincronizar_google_agenda_automatico)
        self.sync_timer.start(1800000)  # Sincroniza a cada 30 minutos
        
        # Configurações de notificação
        self.notification_minutes = 15  # Padrão: 15 minutos antes
        
        # Carrega dados iniciais
        self.atualizar_metricas()
        self.carregar_agendamentos()

    def init_db(self):
        db_path = f"crm_database_{self.email_usuario}.db"
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.cursor = self.conn.cursor()

        self.cursor.execute("PRAGMA table_info(usuarios)")
        colunas = [col[1] for col in self.cursor.fetchall()]
        if 'empresa' not in colunas:
            self.cursor.execute("ALTER TABLE usuarios ADD COLUMN empresa TEXT DEFAULT 'Precode'")
    
        # Verificar empresa do usuário logado
        self.cursor.execute("SELECT is_admin, empresa FROM usuarios WHERE email = ?", (self.email_usuario,))
        result = self.cursor.fetchone()
        self.is_admin = result[0] if result else 0
        self.empresa_usuario = result[1] if result and result[1] else 'Precode'
    
        # Cria a tabela 'clientes' (sem a coluna tipo inicialmente)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS clientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                empresa TEXT,
                telefone TEXT,
                email TEXT,
                status TEXT NOT NULL,
                valor_proposta REAL,
                data_cadastro TEXT,
                data_proxima_acao TEXT,
                ultima_atualizacao TEXT,
                origem TEXT,
                responsavel TEXT,
                probabilidade INTEGER,
                etapa_negocio TEXT,
                observacoes TEXT,
                usuario_id TEXT
            )
        """)
    
        # Cria outras tabelas normalmente
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS interacoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cliente_id INTEGER,
                tipo TEXT,
                descricao TEXT,
                data TEXT,
                duracao INTEGER,
                resultado TEXT,
                usuario_id TEXT,
                FOREIGN KEY(cliente_id) REFERENCES clientes(id)
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS agendamentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo TEXT,
                descricao TEXT,
                data_inicio TEXT,
                data_fim TEXT,
                cliente_id INTEGER,
                tipo TEXT,
                status TEXT,
                usuario_id TEXT,
                FOREIGN KEY(cliente_id) REFERENCES clientes(id)
            )
        """)
    
        # Adiciona a coluna 'tipo' se ainda não existir
        self.cursor.execute("PRAGMA table_info(clientes)")
        colunas = [col[1] for col in self.cursor.fetchall()]
        if 'tipo' not in colunas:
            self.cursor.execute("ALTER TABLE clientes ADD COLUMN tipo TEXT DEFAULT 'Precode'")
    
        # Verifica se é admin
        self.cursor.execute("SELECT is_admin FROM usuarios WHERE email = ?", (self.email_usuario,))
        result = self.cursor.fetchone()
        self.is_admin = result[0] if result else 0


    def init_ui(self):
        """Configura a interface do usuário"""
        # 1. Primeiro cria o widget principal e layout
        self.main_widget = QWidget()
        self.setCentralWidget(self.main_widget)
        
        # 2. Cria o layout principal
        self.main_layout = QHBoxLayout(self.main_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        
        # 3. Cria o stacked widget (content_stack) ANTES do menu
        self.content_stack = QStackedWidget()
        
        # 4. Cria as páginas (que usarão o content_stack)
        self.create_dashboard()
        self.create_clientes_page()
        self.create_relatorios_page()
        self.create_agenda_page()
        self.create_bitrix_page()
        
        # 5. Cria o menu lateral (que precisa do content_stack)
        self.menu_lateral()
        
        # 6. Adiciona os widgets ao layout na ordem correta
        self.main_layout.addWidget(self.menu_frame)  # Menu lateral
        self.main_layout.addWidget(self.content_stack)
        
        
        # 7. Define a página inicial
        self.content_stack.setCurrentIndex(1)  # Dashboard
    
    def setup_dark_mode(self):
        """Configura o tema dark mode com melhor contraste"""
        dark_palette = self.palette()
        
        # Base
        dark_palette.setColor(dark_palette.Window, QColor(25, 35, 45))
        dark_palette.setColor(dark_palette.WindowText, Qt.white)
        dark_palette.setColor(dark_palette.Base, QColor(35, 35, 35))
        dark_palette.setColor(dark_palette.AlternateBase, QColor(53, 53, 53))
        dark_palette.setColor(dark_palette.ToolTipBase, QColor(25, 25, 25))
        dark_palette.setColor(dark_palette.ToolTipText, Qt.white)
        
        # Texto
        dark_palette.setColor(dark_palette.Text, Qt.white)
        dark_palette.setColor(dark_palette.Button, QColor(53, 53, 53))
        dark_palette.setColor(dark_palette.ButtonText, Qt.white)
        dark_palette.setColor(dark_palette.BrightText, Qt.white)
        
        # Desativados
        dark_palette.setColor(dark_palette.Disabled, dark_palette.ButtonText, QColor(127, 127, 127))
        dark_palette.setColor(dark_palette.Disabled, dark_palette.Text, QColor(127, 127, 127))
        
        # Destaques
        dark_palette.setColor(dark_palette.Highlight, QColor(142, 45, 197).lighter())
        dark_palette.setColor(dark_palette.HighlightedText, Qt.black)
        
        self.setPalette(dark_palette)
        
        # Estilo adicional para melhor legibilidade
        self.setStyleSheet("""
            QMenuBar {
                background-color: #2c3e50;
                color: white;
            }
            QMenuBar::item {
                background-color: transparent;
                padding: 5px 10px;
            }
            QMenuBar::item:selected {
                background-color: #34495e;
            }
            QMenu {
                background-color: #2c3e50;
                color: white;
                border: 1px solid #34495e;
            }
            QMenu::item:selected {
                background-color: #34495e;
            }
            QTabBar::tab {
                background: #2c3e50;
                color: white;
                padding: 8px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background: #3498db;
            }
            QStatusBar {
                background-color: #2c3e50;
                color: white;
            }
            QLabel {
                color: white;
            }
            QLineEdit, QComboBox, QDateEdit, QTimeEdit {
                background: #34495e;
                color: white;
                border: 1px solid #2c3e50;
                padding: 5px;
            }
            QComboBox QAbstractItemView {
                background: #34495e;
                color: white;
                selection-background-color: #3498db;
            }
            QMessageBox {
                background: #2c3e50;
            }
            QMessageBox QLabel {
                color: white;
            }
            QInputDialog {
                background: #2c3e50;
            }
            QInputDialog QLabel {
                color: white;
            }
            QTableWidget {
                background: #2c3e50;
                color: white;
                border: 1px solid #34495e;
                gridline-color: #34495e;
            }
            QHeaderView::section {
                background: #34495e;
                color: white;
                padding: 8px;
            }
            QCheckBox {
                margin-left: 5px;
            }
            QToolTip {
                color: #ffffff;
                background-color: #2c3e50;
                border: 1px solid #34495e;
            }
        """)

       # Adicione este método na classe CRMApp (por volta da linha 800):
    def carregar_proximos_agendamentos(self):
        """Carrega os próximos agendamentos na tabela da dashboard"""
        query = """
            SELECT a.data_inicio, a.titulo, c.nome, a.tipo, a.descricao 
            FROM agendamentos a
            LEFT JOIN clientes c ON a.cliente_id = c.id
            WHERE date(a.data_inicio) >= date('now')
            AND a.status = 'Confirmado'
        """
        
        if not self.is_admin:
            query += " AND a.usuario_id = ?"
            params = [self.email_usuario]
        else:
            params = []
            
        query += " ORDER BY a.data_inicio LIMIT 10"
        
        self.cursor.execute(query, params)
        agendamentos = self.cursor.fetchall()
        
        self.agendamentos_table.setRowCount(len(agendamentos))
        for row, (data, titulo, cliente, tipo, descricao) in enumerate(agendamentos):
            try:
                data_obj = datetime.strptime(data, "%Y-%m-%dT%H:%M:%S")
                data_str = data_obj.strftime("%d/%m/%Y")
                hora_str = data_obj.strftime("%H:%M")
            except:
                data_str = data[:10]
                hora_str = "-"
            
            self.agendamentos_table.setItem(row, 0, QTableWidgetItem(data_str))
            self.agendamentos_table.setItem(row, 1, QTableWidgetItem(hora_str))
            self.agendamentos_table.setItem(row, 2, QTableWidgetItem(titulo))
            
            # Exibe cliente e observações
            cliente_text = f"{cliente if cliente else '-'}"
            if descricao:
                cliente_text += f"\nObs: {descricao}"
            cliente_item = QTableWidgetItem(cliente_text)
            cliente_item.setToolTip(descricao if descricao else "")
            self.agendamentos_table.setItem(row, 3, cliente_item)
            
            self.agendamentos_table.setItem(row, 4, QTableWidgetItem(tipo))    
    
    def menu_lateral(self):
        """Cria o menu lateral estilizado"""
        self.menu_frame = QFrame()
        self.menu_frame.setFixedWidth(220)  
        self.menu_frame.setStyleSheet("""
            QFrame {
                background-color: #2c3e50;
                border: none;
            }
            QPushButton {
                color: white;
                text-align: left;
                padding: 12px 20px;
                border: none;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #34495e;
            }
            QPushButton:pressed {
                background-color: #2980b9;
            }
        """)
        
        layout = QVBoxLayout(self.menu_frame)
        layout.setContentsMargins(0, 20, 0, 20)
        layout.setSpacing(10)

        # Logo/Cabeçalho
        logo = QLabel()
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(__file__)
            
        logo_path = os.path.join(base_path, "logo.png")
        logo_pixmap = QPixmap(logo_path)

        if not logo_pixmap.isNull():
            logo.setPixmap(logo_pixmap.scaled(150, 150, Qt.KeepAspectRatio))
        else:
            logo.setText("Logo")
        logo.setAlignment(Qt.AlignCenter)
        logo.mousePressEvent = lambda event: self.show_dashboard()
        logo.setCursor(Qt.PointingHandCursor)
        layout.addWidget(logo)

        # Botões do menu
        buttons = [
            ("🏠 Dashboard", lambda: self.show_dashboard()),
            ("👥 Clientes", lambda: self.show_clientes()),
            ("📆 Agenda", lambda: self.show_agenda()),
            ("📊 Relatórios", lambda: self.show_relatorios()),
            ("🔗 Bitrix24", lambda: self.show_bitrix()),
            ("⚙️ Configurações", lambda: self.show_configuracoes()),
        ]

        for text, callback in buttons:
            btn = QPushButton(text)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(callback)
            layout.addWidget(btn)

        # Espaçador
        layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Botão de sair
        btn_sair = QPushButton("🚪 Sair")
        btn_sair.setCursor(Qt.PointingHandCursor)
        btn_sair.clicked.connect(self.close)
        layout.addWidget(btn_sair)

        
        self.main_layout.addWidget(self.menu_frame)

    def create_dashboard(self):
        """Cria a página de dashboard com métricas e gráficos"""
        self.dashboard_page = QWidget()
        layout = QVBoxLayout(self.dashboard_page)
        layout.setContentsMargins(30, 20, 30, 20)
        layout.setSpacing(20)

        # Título
        title = QLabel("Dashboard de Vendas")
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: white;")
        layout.addWidget(title)
    
        # Filtro de data - Layout mais compacto
        filter_frame = QFrame()
        filter_frame.setStyleSheet("background: #34495e; border-radius: 5px; padding: 10px;")
        filter_layout = QHBoxLayout(filter_frame)
        filter_layout.setContentsMargins(5, 5, 5, 5)
        
        filter_layout.addWidget(QLabel("Período:"))
        
        self.month_filter = QComboBox()
        meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        self.month_filter.addItems(meses)
        self.month_filter.setCurrentIndex(datetime.now().month - 1)
        self.month_filter.setFixedWidth(120)
        filter_layout.addWidget(self.month_filter)
        
        self.year_filter = QComboBox()
        anos = [str(datetime.now().year - 1), str(datetime.now().year), str(datetime.now().year + 1)]
        self.year_filter.addItems(anos)
        self.year_filter.setCurrentText(str(datetime.now().year))
        self.year_filter.setFixedWidth(80)
        filter_layout.addWidget(self.year_filter)
        
        btn_filter = QPushButton("Filtrar")
        btn_filter.setFixedWidth(80)
        btn_filter.clicked.connect(self.atualizar_metricas)
        btn_filter.clicked.connect(lambda: self.atualizar_grafico_fechamentos(
            self.month_filter.currentIndex() + 1,
            int(self.year_filter.currentText())
        ))


        filter_layout.addWidget(btn_filter)
        
        filter_layout.addStretch()
        
        layout.addWidget(filter_frame)
    
        

        # Botão para adicionar cliente diretamente do dashboard
        btn_add_cliente = QPushButton("➕ Adicionar Novo Cliente")
        btn_add_cliente.setStyleSheet("""
            QPushButton {
                background: #2ecc71;
                color: white;
                padding: 12px;
                border-radius: 5px;
                font-size: 16px;
            }
            QPushButton:hover {
                background: #27ae60;
            }
        """)
        btn_add_cliente.setCursor(Qt.PointingHandCursor)
        btn_add_cliente.clicked.connect(self.abrir_cadastro_cliente)
        layout.addWidget(btn_add_cliente)

        # Linha de cards de métricas
        metrics_row = QHBoxLayout()
        metrics_row.setSpacing(15)
        
        # Cards de métricas (serão atualizados dinamicamente)
        self.metric_cards = {
            "total_clientes": self.create_metric_card("👥 0", "Total Clientes", "#3498db"),
            "clientes": self.create_metric_card("👤 0", "Clientes", "#9b59b6"),
            "em_atendimento": self.create_metric_card("📞 0", "Em Atendimento", "#f39c12"),
            "propostas": self.create_metric_card("📄 0", "Propostas", "#e74c3c"),
            "finalizados": self.create_metric_card("✅ 0", "Finalizados", "#2ecc71"),
            "valor_fechado": self.create_metric_card("💰 R$0,00", "Valor Fechado", "#27ae60")
        }

        for card in self.metric_cards.values():
            metrics_row.addWidget(card)
            # Conecta o clique para abrir o popup de status
            card.mousePressEvent = lambda event, card_ref=card: self.show_status_popup(card_ref)

        layout.addLayout(metrics_row)

        # Linha de gráficos
        charts_row = QHBoxLayout()
        charts_row.setSpacing(15)
        
        # Gráfico 1: Status dos Clientes
        self.chart1 = QChart()
        self.chart1.setTitle("Status dos Clientes")
        self.chart1.setAnimationOptions(QChart.SeriesAnimations)
        self.chart1.setTheme(QChart.ChartThemeDark)
        
        chart_view1 = QChartView(self.chart1)
        chart_view1.setRenderHint(QPainter.Antialiasing)
        chart_view1.setStyleSheet("background: transparent;")
        charts_row.addWidget(chart_view1)
        
        # Gráfico 2: Fechamentos do mês selecionado
        self.chart2 = QChart()
        self.chart_view2 = CustomChartView(self.chart2)
        self.chart_view2.setRenderHint(QPainter.Antialiasing)
        self.chart_view2.setStyleSheet("background: transparent;")
        self.chart_view2.mousePressEvent = self.show_full_chart
        charts_row.addWidget(self.chart_view2)

        
        layout.addLayout(charts_row)

        # Tabela de próximos agendamentos
        agendamentos_label = QLabel("📅 Próximos Agendamentos:")
        agendamentos_label.setStyleSheet("font-size: 16px; color: white;")
        layout.addWidget(agendamentos_label)
        
        self.agendamentos_table = QTableWidget()
        self.agendamentos_table.setColumnCount(5)
        self.agendamentos_table.setHorizontalHeaderLabels(["Data", "Hora", "Título", "Cliente", "Tipo"])
        self.agendamentos_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.agendamentos_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.agendamentos_table.setStyleSheet("""
            QTableWidget {
                background: #2c3e50;
                color: white;
                border: 1px solid #34495e;
                gridline-color: #34495e;
            }
            QHeaderView::section {
                background: #34495e;
                color: white;
                padding: 8px;
            }
        """)
        layout.addWidget(self.agendamentos_table)

        self.content_stack.addWidget(self.dashboard_page)

        self.atualizar_grafico_fechamentos(
            self.month_filter.currentIndex() + 1,
            int(self.year_filter.currentText())
        )



    def show_full_chart(self, event):
        ano = int(self.year_filter.currentText())

        # Consulta SQL consolidada para clientes
        self.cursor.execute(f"""
            SELECT 
                strftime('%m', data_cadastro) as mes,
                COUNT(*) as total,
                SUM(CASE WHEN status = 'Finalizado' THEN 1 ELSE 0 END) as finalizados,
                SUM(CASE WHEN status = 'Em atendimento' THEN 1 ELSE 0 END) as em_atendimento,
                SUM(CASE WHEN status = 'Proposta' THEN 1 ELSE 0 END) as propostas
            FROM clientes
            WHERE strftime('%Y', data_cadastro) = ?
              AND usuario_id = ?
            GROUP BY mes
            ORDER BY mes
        """, (str(ano), self.email_usuario))
    
        dados_clientes = {row[0]: row[1:] for row in self.cursor.fetchall()}
    
        # Consulta SQL para agendamentos
        self.cursor.execute(f"""
            SELECT 
                strftime('%m', data_inicio) as mes,
                COUNT(*) as agendamentos
            FROM agendamentos
            WHERE strftime('%Y', data_inicio) = ?
              AND usuario_id = ?
            GROUP BY mes
            ORDER BY mes
        """, (str(ano), self.email_usuario))
    
        dados_agendamentos = {row[0]: row[1] for row in self.cursor.fetchall()}
    
        # Preparar os dados para o gráfico
        meses = [calendar.month_abbr[i] for i in range(1, 13)]
    
        bar_total = QBarSet("Total Clientes")
        bar_clientes = QBarSet("Clientes")
        bar_finalizados = QBarSet("Finalizados")
        bar_em_atendimento = QBarSet("Em Atendimento")
        bar_propostas = QBarSet("Propostas")
        bar_agendamentos = QBarSet("Agendamentos Google")
    
        for i in range(1, 13):
            m = f"{i:02d}"
            total, fin, atend, prop = dados_clientes.get(m, (0, 0, 0, 0))
            agend = dados_agendamentos.get(m, 0)

            bar_total.append(total)
            bar_clientes.append(total - fin - atend - prop)  
            bar_finalizados.append(fin)
            bar_em_atendimento.append(atend)
            bar_propostas.append(prop)
            bar_agendamentos.append(agend)

    
        # Cores
        bar_total.setColor(QColor("#3498db"))
        bar_finalizados.setColor(QColor("#2ecc71"))
        bar_em_atendimento.setColor(QColor("#f39c12"))
        bar_propostas.setColor(QColor("#e74c3c"))
        bar_agendamentos.setColor(QColor("#1abc9c"))
    
        # Série de barras agrupadas
        series = QBarSeries()
        series.append(bar_total)
        series.append(bar_finalizados)
        series.append(bar_em_atendimento)
        series.append(bar_propostas)
        series.append(bar_agendamentos)
    
        chart = QChart()
        chart.addSeries(series)
        chart.setTitle(f"Fechamentos e Status - {ano}")
        chart.setAnimationOptions(QChart.SeriesAnimations)
        chart.setTheme(QChart.ChartThemeDark)
    
        axis_x = QBarCategoryAxis()
        axis_x.append(meses)
        chart.setAxisX(axis_x, series)
    
        axis_y = QValueAxis()
        axis_y.setTitleText("Quantidade")
        axis_y.setLabelFormat("%d")
        axis_y.setRange(0, max(bar_total) * 1.2 if bar_total.count() > 0 else 10)
        chart.setAxisY(axis_y, series)
    
        chart_view = QChartView(chart)
        chart_view.setRenderHint(QPainter.Antialiasing)
    
        # Criar o popup
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Fechamentos Detalhados - {ano}")
        dialog.setMinimumSize(1000, 600)
        layout = QVBoxLayout(dialog)
        layout.addWidget(chart_view)
    
        btn_fechar = QPushButton("Fechar")
        btn_fechar.clicked.connect(dialog.close)
        layout.addWidget(btn_fechar)
        dialog.exec_()
    
    def mostrar_detalhes_mes(self, mes, ano, metrics):
        """Mostra popup com detalhes do mês quando clica no gráfico"""
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Detalhes - {calendar.month_name[mes]}/{ano}")
        dialog.setMinimumSize(800, 600)  # Tamanho maior para melhor visualização
        dialog.setStyleSheet("""
            QDialog {
                background: #2c3e50;
                color: white;
            }
            QTableWidget {
                background: #34495e;
                color: white;
                gridline-color: #7f8c8d;
            }
            QHeaderView::section {
                background: #2c3e50;
                color: white;
                padding: 5px;
            }
            QTabWidget::pane {
                border: 0;
            }
        """)
        
        # Fecha quando clica fora
        dialog.setWindowFlags(dialog.windowFlags() | Qt.WindowCloseButtonHint)
        dialog.setModal(False)
        
        layout = QVBoxLayout(dialog)
        
        # Cria abas para organização
        tab_widget = QTabWidget()
        
        # Aba 1: Métricas resumidas
        metrics_tab = QWidget()
        metrics_layout = QVBoxLayout(metrics_tab)
        
        # Tabela de métricas
        metrics_table = QTableWidget()
        metrics_table.setColumnCount(2)
        metrics_table.setHorizontalHeaderLabels(["Métrica", "Valor"])
        
        # Ordem específica para exibição
        order = ["Total", "Cliente", "Em Atend.", "Proposta", "Finalizado", "Agendamentos"]
        metrics_table.setRowCount(len(order))
        
        # Preenche a tabela
        for row, name in enumerate(order):
            value = metrics.get(name, 0)
            metrics_table.setItem(row, 0, QTableWidgetItem(name))
            
            # Formata valores monetários quando aplicável
            if name in ["Total", "Proposta", "Finalizado"]:
                value_str = locale.currency(value, grouping=True) if value else "0"
            else:
                value_str = str(value)
                
            metrics_table.setItem(row, 1, QTableWidgetItem(value_str))
        
        metrics_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        metrics_table.verticalHeader().setVisible(False)
        metrics_layout.addWidget(metrics_table)
        
        # Gráfico de pizza para visualização alternativa
        pie_chart = QChart()
        pie_chart.setTitle("Distribuição por Status")
        pie_series = QPieSeries()
        
        # Cores consistentes com o gráfico principal
        colors = {
            "Total": "#3498db",
            "Cliente": "#9b59b6",
            "Em Atend.": "#f39c12",
            "Proposta": "#e74c3c",
            "Finalizado": "#2ecc71",
            "Agendamentos": "#1abc9c"
        }
        
        # Adiciona apenas as métricas relevantes (exclui Total e Agendamentos)
        relevant_metrics = ["Cliente", "Em Atend.", "Proposta", "Finalizado"]
        for name in relevant_metrics:
            value = metrics.get(name, 0)
            if value > 0:
                slice_ = pie_series.append(f"{name}: {value}", value)
                slice_.setColor(QColor(colors[name]))
                slice_.setLabelVisible(True)
                slice_.setLabel(f"{name}: {value} ({value/sum(metrics[m] for m in relevant_metrics)*100:.1f}%)")
            
        pie_chart.addSeries(pie_series)
        pie_chart.setAnimationOptions(QChart.SeriesAnimations)
        pie_chart.legend().setVisible(True)
        pie_chart.legend().setAlignment(Qt.AlignBottom)
        
        chart_view = QChartView(pie_chart)
        chart_view.setRenderHint(QPainter.Antialiasing)
        metrics_layout.addWidget(chart_view)
        
        tab_widget.addTab(metrics_tab, "Métricas")
        
        # Aba 2: Lista de agendamentos
        agendamentos_tab = QWidget()
        agendamentos_layout = QVBoxLayout(agendamentos_tab)
        
        # Obtém os agendamentos do mês
        mes_ano = f"{ano}-{mes:02d}"
        query = """
            SELECT a.id, a.data_inicio, a.titulo, c.nome, a.tipo, a.status, a.descricao 
            FROM agendamentos a
            LEFT JOIN clientes c ON a.cliente_id = c.id
            WHERE strftime('%Y-%m', a.data_inicio) = ?
            ORDER BY a.data_inicio
        """
        self.cursor.execute(query, (mes_ano,))
        agendamentos = self.cursor.fetchall()
        
        # Tabela de agendamentos
        agendamentos_table = QTableWidget()
        agendamentos_table.setColumnCount(6)
        agendamentos_table.setHorizontalHeaderLabels(["Data", "Hora", "Título", "Cliente", "Tipo", "Status"])
        
        agendamentos_table.setRowCount(len(agendamentos))
        for row, (id_, data, titulo, cliente, tipo, status, descricao) in enumerate(agendamentos):
            try:
                data_obj = datetime.strptime(data, "%Y-%m-%dT%H:%M:%S")
                data_str = data_obj.strftime("%d/%m/%Y")
                hora_str = data_obj.strftime("%H:%M")
            except:
                data_str = data[:10]
                hora_str = "-"
            
            # Preenche a tabela
            agendamentos_table.setItem(row, 0, QTableWidgetItem(data_str))
            agendamentos_table.setItem(row, 1, QTableWidgetItem(hora_str))
            
            titulo_item = QTableWidgetItem(titulo)
            if descricao:
                titulo_item.setToolTip(descricao)
            agendamentos_table.setItem(row, 2, titulo_item)
            
            agendamentos_table.setItem(row, 3, QTableWidgetItem(cliente if cliente else "-"))
            agendamentos_table.setItem(row, 4, QTableWidgetItem(tipo))
            
            status_item = QTableWidgetItem(status)
            if status == "Confirmado":
                status_item.setBackground(QColor(46, 204, 113))  # Verde
            elif status == "Cancelado":
                status_item.setBackground(QColor(231, 76, 60))   # Vermelho
            elif status == "Pendente":
                status_item.setBackground(QColor(241, 196, 15))  # Amarelo
            agendamentos_table.setItem(row, 5, status_item)
        
        agendamentos_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        agendamentos_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        agendamentos_layout.addWidget(agendamentos_table)
        
        # Botão para exportar agendamentos
        btn_exportar = QPushButton("Exportar para Excel")
        btn_exportar.setStyleSheet("""
            QPushButton {
                background: #27ae60;
                color: white;
                padding: 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background: #219653;
            }
        """)
        btn_exportar.clicked.connect(lambda: self.exportar_agendamentos_excel(agendamentos, mes, ano))
        agendamentos_layout.addWidget(btn_exportar)
        
        tab_widget.addTab(agendamentos_tab, f"Agendamentos ({len(agendamentos)})")
        
        layout.addWidget(tab_widget)
        
        # Botão de fechar
        btn_close = QPushButton("Fechar")
        btn_close.clicked.connect(dialog.close)
        layout.addWidget(btn_close)
        
        dialog.exec_()



    def exportar_agendamentos_excel(self, agendamentos, mes, ano):
        """Exporta a lista de agendamentos para Excel"""
        try:
            from pandas import DataFrame
            
            # Prepara os dados
            dados = []
            for id_, data, titulo, cliente, tipo, status, descricao in agendamentos:
                try:
                    data_obj = datetime.strptime(data, "%Y-%m-%dT%H:%M:%S")
                    data_str = data_obj.strftime("%d/%m/%Y")
                    hora_str = data_obj.strftime("%H:%M")
                except:
                    data_str = data[:10]
                    hora_str = "-"
                
                dados.append([
                    data_str, hora_str, titulo, 
                    cliente if cliente else "-", 
                    tipo, status, descricao if descricao else ""
                ])
            
            # Cria DataFrame
            df = DataFrame(dados, columns=[
                "Data", "Hora", "Título", "Cliente", "Tipo", "Status", "Descrição"
            ])
            
            # Salva arquivo
            file_name, _ = QFileDialog.getSaveFileName(
                self, "Exportar Agendamentos", 
                f"Agendamentos_{calendar.month_name[mes]}_{ano}.xlsx", 
                "Excel Files (*.xlsx)"
            )
            
            if file_name:
                df.to_excel(file_name, index=False)
                QMessageBox.information(
                    self, "Sucesso", 
                    f"Agendamentos exportados com sucesso para:\n{file_name}"
                )
                
        except Exception as e:
            QMessageBox.critical(
                self, "Erro", 
                f"Falha ao exportar agendamentos:\n{str(e)}"
            )



    def show_status_popup(self, card):
        """Abre popup para filtrar clientes por status ao clicar em um card"""
        try:
            # Determina o status baseado no título do card
            titulo = card.title_label.text().strip().lower()
            
            if "total" in titulo:
                status = "Todos"
            elif "cliente" in titulo and not "total" in titulo:
                status = "Cliente"
            elif "atendimento" in titulo:
                status = "Em atendimento"
            elif "proposta" in titulo:
                status = "Proposta"
            elif "finalizado" in titulo:
                status = "Finalizado"
            else:
                return  # Não faz nada se não for um card reconhecido

            # Se for o card de "Total Clientes", apenas mostra todos
            if status == "Todos":
                self.status_filter.setCurrentText("Todos")
                self.carregar_clientes()
                return

            # Consulta os clientes com esse status
            query = "SELECT id, nome, empresa FROM clientes WHERE status = ?"
            params = (status,)
            
            if not self.is_admin:
                query += " AND usuario_id = ?"
                params = (status, self.email_usuario)
            
            self.cursor.execute(query, params)
            clientes = self.cursor.fetchall()

            if not clientes:
                QMessageBox.information(self, "Sem clientes", f"Não há clientes com status '{status}'.")
                return

            # Cria o popup de lista
            dialog = QDialog(self)
            dialog.setWindowTitle(f"{status}s - Selecione um cliente")
            dialog.setMinimumWidth(500)
            dialog.setStyleSheet("""
                QDialog {
                    background: #2c3e50;
                    color: white;
                }
                QLineEdit {
                    background: #34495e;
                    color: white;
                    padding: 5px;
                }
                QPushButton {
                    background: #3498db;
                    color: white;
                    padding: 8px;
                    border-radius: 4px;
                }
            """)

            layout = QVBoxLayout(dialog)
            
            # Barra de pesquisa
            search_input = QLineEdit()
            search_input.setPlaceholderText("Pesquisar clientes...")
            layout.addWidget(search_input)
            
            # Lista de clientes
            lista = QListWidget()
            lista.setStyleSheet("""
                QListWidget {
                    background: #34495e;
                    color: white;
                    border: 1px solid #2c3e50;
                }
            """)
            
            self.clientes_data = []
            for id_, nome, empresa in clientes:
                item = QListWidgetItem(f"{nome} - {empresa}" if empresa else nome)
                item.setData(Qt.UserRole, id_)
                lista.addItem(item)
                self.clientes_data.append((id_, nome, empresa))
            
            # Combo para alterar status
            status_combo = QComboBox()
            status_combo.addItems(["Cliente", "Em atendimento", "Proposta", "Finalizado"])
            status_combo.setCurrentText(status)
            
            # Botão para alterar status
            btn_alterar = QPushButton("Alterar Status")
            btn_alterar.clicked.connect(lambda: self.alterar_status_em_massa(
                lista, 
                status_combo.currentText(),
                dialog
            ))
            
            # Layout dos controles de status
            status_layout = QHBoxLayout()
            status_layout.addWidget(QLabel("Novo Status:"))
            status_layout.addWidget(status_combo)
            status_layout.addWidget(btn_alterar)
            
            layout.addWidget(QLabel(f"Clientes com status: {status}"))
            layout.addWidget(search_input)
            layout.addWidget(lista)
            layout.addLayout(status_layout)

            # Filtro de pesquisa
            def filtrar_clientes():
                termo = search_input.text().lower()
                lista.clear()
                for id_, nome, empresa in self.clientes_data:
                    if termo in nome.lower() or (empresa and termo in empresa.lower()):
                        item = QListWidgetItem(f"{nome} - {empresa}" if empresa else nome)
                        item.setData(Qt.UserRole, id_)
                        lista.addItem(item)
            
            search_input.textChanged.connect(filtrar_clientes)

            # Ao dar duplo clique no nome → altera status do cliente
            def ao_selecionar(item):
                cliente_id = item.data(Qt.UserRole)
                dialog.accept()
                self.alterar_status_cliente(cliente_id)

            lista.itemDoubleClicked.connect(ao_selecionar)
            
            # Botão de fechar
            btn_fechar = QPushButton("Fechar")
            btn_fechar.clicked.connect(dialog.close)
            layout.addWidget(btn_fechar)
            
            dialog.exec_()

        except Exception as e:
            import traceback
            with open("erro_popup_card.log", "w", encoding="utf-8") as f:
                f.write(traceback.format_exc())
            QMessageBox.critical(self, "Erro", f"Erro ao abrir popup:\n{str(e)}")

    def alterar_status_em_massa(self, lista, novo_status, dialog):
        """Altera o status de vários clientes de uma vez"""
        selected_items = lista.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Aviso", "Selecione pelo menos um cliente!")
            return
            
        ids = [item.data(Qt.UserRole) for item in selected_items]
        
        try:
            placeholders = ','.join(['?'] * len(ids))
            query = f"""
                UPDATE clientes SET
                    status = ?,
                    ultima_atualizacao = datetime('now')
                WHERE id IN ({placeholders})
            """
            
            if not self.is_admin:
                query += " AND usuario_id = ?"
                params = [novo_status] + ids + [self.email_usuario]
            else:
                params = [novo_status] + ids
            
            self.cursor.execute(query, params)
            
            self.conn.commit()
            self.carregar_clientes()
            self.atualizar_metricas()
            dialog.accept()
            QMessageBox.information(self, "Sucesso", f"Status de {len(ids)} clientes atualizado para '{novo_status}'!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao atualizar status em massa:\n{str(e)}")

    def create_metric_card(self, value, title, color):
        """Cria um card de métrica estilizado"""
        return MetricCard(value, title, color, self)

    def create_clientes_page(self):
        """Cria a página de gerenciamento de clientes"""
        self.clientes_page = QWidget()
        layout = QVBoxLayout(self.clientes_page)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # Cabeçalho
        header = QHBoxLayout()
    
        title = QLabel("Gerenciamento de Clientes")
        title.setStyleSheet("font-size: 22px; font-weight: bold; color: white;")
        header.addWidget(title)
        
        # Botão para adicionar cliente
        btn_add = QPushButton("Novo Cliente")
        btn_add.setStyleSheet("""
            QPushButton {
                background: #3498db;
                color: white;
                padding: 8px 15px;
                border-radius: 4px;
                font-size: 14px;
            }
            QPushButton:hover {
                background: #2980b9;
            }
        """)
        btn_add.setCursor(Qt.PointingHandCursor)
        btn_add.clicked.connect(self.abrir_cadastro_cliente)
        header.addWidget(btn_add)
        
        # Botão para importar do Excel
        btn_import = QPushButton("Importar do Excel")
        btn_import.setStyleSheet("""
            QPushButton {
                background: #2ecc71;
                color: white;
                padding: 8px 15px;
                border-radius: 4px;
                font-size: 14px;
            }
            QPushButton:hover {
                background: #27ae60;
            }
        """)
        btn_import.setCursor(Qt.PointingHandCursor)
        btn_import.clicked.connect(self.importar_clientes_excel)
        header.addWidget(btn_import)
        
        # Botão para exportar para Bitrix24
        btn_export = QPushButton("Exportar para Bitrix24")
        btn_export.setStyleSheet("""
            QPushButton {
                background: #9b59b6;
                color: white;
                padding: 8px 15px;
                border-radius: 4px;
                font-size: 14px;
            }
            QPushButton:hover {
                background: #8e44ad;
            }
        """)
        btn_export.setCursor(Qt.PointingHandCursor)
        btn_export.clicked.connect(self.exportar_bitrix24)
        header.addWidget(btn_export)
        
        layout.addLayout(header)
    
        # Filtros
        filter_row = QHBoxLayout()
        filter_row.setSpacing(10)

        # Barra de busca
        self.busca_input = QLineEdit()
        self.busca_input.setPlaceholderText("Pesquisar clientes...")
        filter_row.addWidget(self.busca_input)
        self.busca_input.textChanged.connect(self.filtrar_clientes_ao_vivo)

        self.busca_input.setStyleSheet("""
            QLineEdit {
                background: #2c3e50;
                color: white;
                padding: 8px;
                border: 1px solid #34495e;
                border-radius: 4px;
            }
        """)
        filter_row.addWidget(self.busca_input)
        
        # Filtro de status
        self.status_filter = QComboBox()
        self.status_filter.addItems(["Todos", "Cliente", "Em atendimento", "Proposta", "Finalizado"])
        self.status_filter.setStyleSheet("""
            QComboBox {
                background: #2c3e50;
                color: white;
                padding: 8px;
                border: 1px solid #34495e;
                border-radius: 4px;
            }
            QComboBox QAbstractItemView {
                background: #34495e;
                color: white;
                selection-background-color: #3498db;
            }
        """)
        filter_row.addWidget(self.status_filter)
        
        # Filtro Precode - ADICIONEI AQUI ANTES DOS BOTÕES
        self.precode_filter = QComboBox()
        self.precode_filter.addItems(["Todos", "Precode", "Allpost"])
        self.precode_filter.setStyleSheet("""
            QComboBox {
                background: #2c3e50;
                color: white;
                padding: 8px;
                border: 1px solid #34495e;
                border-radius: 4px;
            }
        """)
        filter_row.addWidget(QLabel("Tipo:"))  # Adicionei um label para identificar
        filter_row.addWidget(self.precode_filter)
        
        # Botão Filtrar
        btn_filter = QPushButton("Filtrar")
        btn_filter.setStyleSheet("""
            QPushButton {
                background: #3498db;
                color: white;
                padding: 8px 15px;
                border-radius: 4px;
                font-size: 14px;
            }
            QPushButton:hover {
                background: #2980b9;
            }
        """)
        btn_filter.setCursor(Qt.PointingHandCursor)
        btn_filter.clicked.connect(self.carregar_clientes)
        filter_row.addWidget(btn_filter)
        
        # Botão Selecionar Todos
        self.btn_select_all = QPushButton("Selecionar Todos")
        self.btn_select_all.setStyleSheet("""
            QPushButton {
                background: #e67e22;
                color: white;
                padding: 8px 15px;
                border-radius: 4px;
                font-size: 14px;
            }
            QPushButton:hover {
                background: #d35400;
            }
        """)
        self.btn_select_all.setCursor(Qt.PointingHandCursor)
        self.btn_select_all.clicked.connect(self.toggle_select_all_clients)
        filter_row.addWidget(self.btn_select_all)
        
        # Botão Ações em Massa
        self.btn_acoes_massa = QPushButton("Ações em Massa")
        self.btn_acoes_massa.setStyleSheet("""
            QPushButton {
                background: #e67e22;
                color: white;
                padding: 8px 15px;
                border-radius: 4px;
                font-size: 14px;
            }
            QPushButton:hover {
                background: #d35400;
            }
        """)
        self.btn_acoes_massa.setCursor(Qt.PointingHandCursor)
        self.btn_acoes_massa.clicked.connect(self.mostrar_menu_acoes_massa)
        filter_row.addWidget(self.btn_acoes_massa)
        
        layout.addLayout(filter_row)
        
        # Tabela de clientes - CORRIGI A FORMATAÇÃO AQUI
        self.clientes_table = QTableWidget()
        self.clientes_table.setColumnCount(11)
        self.clientes_table.setHorizontalHeaderLabels(["", "Precode", "Nome", "Empresa", "Telefone", "Email", "Status", "Valor", "Cadastro", ""])
        self.clientes_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.clientes_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.clientes_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.clientes_table.setColumnHidden(1, True)  # Mostra a coluna Precode
        self.clientes_table.setStyleSheet("""
            QTableWidget {
                background: #2c3e50;
                color: white;
                border: 1px solid #34495e;
                gridline-color: #34495e;
            }
            QHeaderView::section {
                background: #34495e;
                color: white;
                padding: 8px;
            }
            QCheckBox {
                margin-left: 5px;
            }
        """)
        
        layout.addWidget(self.clientes_table)
        self.carregar_clientes()
        
        self.content_stack.addWidget(self.clientes_page)

        
    def filtrar_clientes_ao_vivo(self):
        """Filtra a tabela de clientes conforme o texto digitado"""
        texto = self.busca_input.text().lower()
    
        for row in range(self.clientes_table.rowCount()):
            mostrar = False
            for col in range(self.clientes_table.columnCount()):
                # Pula a coluna 0 (checkbox) e coluna 1 (ID)
                if col in (0, 1):
                    continue
                
                item = self.clientes_table.item(row, col)
                if item and texto in item.text().lower():
                    mostrar = True
                    break
                
            self.clientes_table.setRowHidden(row, not mostrar)


    def aplicar_filtros_clientes(self):
        """Aplica todos os filtros (texto, status, tipo)"""
        self.filtrar_clientes_ao_vivo()
    
        # E então conectar os outros filtros:
        self.status_filter.currentTextChanged.connect(self.aplicar_filtros_clientes)
        self.precode_filter.currentTextChanged.connect(self.aplicar_filtros_clientes)

        
    def toggle_select_all_clients(self):
        """Alterna entre selecionar e desselecionar todos os clientes"""
        if not hasattr(self, 'all_clients_selected'):
            self.all_clients_selected = False
    
        self.all_clients_selected = not self.all_clients_selected
    
        for row in range(self.clientes_table.rowCount()):
            if self.clientes_table.isRowHidden(row):
                continue
            checkbox = self.clientes_table.cellWidget(row, 0)
            if checkbox:
                checkbox.setChecked(self.all_clients_selected)
    
        # Atualiza o texto do botão
        self.btn_select_all.setText("Desselecionar Todos" if self.all_clients_selected else "Selecionar Todos")
    

    def mostrar_menu_acoes_massa(self):
        """Mostra menu de ações em massa para clientes selecionados"""
        selected_ids = []
        for row in range(self.clientes_table.rowCount()):
            checkbox = self.clientes_table.cellWidget(row, 0)
            if checkbox and checkbox.isChecked():
                id_item = self.clientes_table.item(row, 1)
                if id_item:
                    selected_ids.append(int(id_item.text()))
        
        if not selected_ids:
            QMessageBox.warning(self, "Aviso", "Selecione pelo menos um cliente!")
            return
            
        menu = QDialog(self)
        menu.setWindowTitle("Ações em Massa")
        menu.setFixedSize(300, 200)
        menu.setStyleSheet("""
            QDialog {
                background: #2c3e50;
                color: white;
            }
            QPushButton {
                text-align: left;
                padding: 8px;
                background: transparent;
                border: none;
                color: white;
            }
            QPushButton:hover {
                background: #34495e;
            }
        """)
        
        layout = QVBoxLayout(menu)
        
        # Combo para alterar status
        status_combo = QComboBox()
        status_combo.addItems(["Cliente", "Em atendimento", "Proposta", "Finalizado"])
        
        # Botão para alterar status
        btn_alterar_status = QPushButton("Alterar Status")
        btn_alterar_status.clicked.connect(lambda: self.alterar_status_em_massa_tabela(
            selected_ids,
            status_combo.currentText(),
            menu
        ))
        
        # Layout dos controles de status
        status_layout = QHBoxLayout()
        status_layout.addWidget(QLabel("Novo Status:"))
        status_layout.addWidget(status_combo)
        layout.addLayout(status_layout)
        layout.addWidget(btn_alterar_status)
        
        # Botão para excluir clientes
        btn_excluir = QPushButton("🗑️ Excluir Clientes Selecionados")
        btn_excluir.setStyleSheet("color: #e74c3c;")
        btn_excluir.clicked.connect(lambda: self.excluir_clientes_em_massa(selected_ids, menu))
        layout.addWidget(btn_excluir)
        
        # Botão de fechar
        btn_fechar = QPushButton("Fechar")
        btn_fechar.clicked.connect(menu.close)
        layout.addWidget(btn_fechar)
        
        menu.exec_()

    def alterar_status_em_massa_tabela(self, ids, novo_status, dialog):
        """Altera o status de vários clientes selecionados na tabela"""
        try:
            placeholders = ','.join(['?'] * len(ids))
            query = f"""
                UPDATE clientes SET
                    status = ?,
                    ultima_atualizacao = datetime('now')
                WHERE id IN ({placeholders})
            """
            
            if not self.is_admin:
                query += " AND usuario_id = ?"
                params = [novo_status] + ids + [self.email_usuario]
            else:
                params = [novo_status] + ids
            
            self.cursor.execute(query, params)
            
            self.conn.commit()
            self.carregar_clientes()
            self.atualizar_metricas()
            dialog.accept()
            QMessageBox.information(self, "Sucesso", f"Status de {len(ids)} clientes atualizado para '{novo_status}'!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao atualizar status em massa:\n{str(e)}")

    def excluir_clientes_em_massa(self, ids, dialog):
        """Exclui vários clientes selecionados na tabela"""
        confirm = QMessageBox.question(
            self, "Confirmar Exclusão",
            f"Tem certeza que deseja excluir {len(ids)} clientes?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if confirm == QMessageBox.Yes:
            try:
                placeholders = ','.join(['?'] * len(ids))
                
                # Primeiro exclui agendamentos relacionados
                self.cursor.execute(f"DELETE FROM agendamentos WHERE cliente_id IN ({placeholders})", ids)
                
                # Depois exclui os clientes
                query = f"DELETE FROM clientes WHERE id IN ({placeholders})"
                
                if not self.is_admin:
                    query += " AND usuario_id = ?"
                    params = ids + [self.email_usuario]
                else:
                    params = ids
                
                self.cursor.execute(query, params)
                
                self.conn.commit()
                self.carregar_clientes()
                dialog.accept()
                QMessageBox.information(self, "Sucesso", f"{len(ids)} clientes excluídos com sucesso!")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao excluir clientes:\n{str(e)}")

    def importar_clientes_excel(self):
        """Importa clientes de um arquivo Excel"""
        try:
            from pandas import read_excel
            
            file_name, _ = QFileDialog.getOpenFileName(
                self, "Importar Clientes", "", "Excel Files (*.xlsx *.xls)"
            )
            
            if file_name:
                df = read_excel(file_name)
                
                # Verifica colunas mínimas necessárias
                required_columns = ['nome', 'email']
                for col in required_columns:
                    if col not in df.columns:
                        raise ValueError(f"Coluna '{col}' não encontrada no arquivo Excel")
                
                # Processa cada linha
                settings = QSettings("GustavoMartins", "CRMApp")
                tipo_padrao = settings.value("tipo_padrao", "Precode")


                success_count = 0
                for _, row in df.iterrows():
                    try:
                        nome = str(row['nome']).strip()
                        if not nome:
                            continue
                            
                        email = str(row.get('email', '')).strip()
                        empresa = str(row.get('empresa', '')).strip()
                        telefone = str(row.get('telefone', '')).strip()
                        status = str(row.get('status', 'Cliente')).strip()
                        valor = float(row.get('valor_proposta', 0))
                        origem = str(row.get('origem', 'Site')).strip()
                        responsavel = str(row.get('responsavel', '')).strip()
                        observacoes = str(row.get('observacoes', '')).strip()
                        
                        # Verifica se cliente já existe pelo telefone
                        if telefone:
                            query = "SELECT COUNT(*) FROM clientes WHERE telefone = ?"
                            params = (telefone,)
                            
                            if not self.is_admin:
                                query += " AND usuario_id = ?"
                                params = (telefone, self.email_usuario)
                            
                            self.cursor.execute(query, params)
                            if self.cursor.fetchone()[0] > 0:
                                continue
                                
                        # Verifica se cliente já existe pelo nome
                        query = "SELECT COUNT(*) FROM clientes WHERE nome = ?"
                        params = (nome,)
                        
                        if not self.is_admin:
                            query += " AND usuario_id = ?"
                            params = (nome, self.email_usuario)
                        
                        self.cursor.execute(query, params)
                        if self.cursor.fetchone()[0] == 0:
                            self.cursor.execute("""
                                INSERT INTO clientes (
                                    nome, empresa, telefone, email, status, valor_proposta,
                                    data_cadastro, origem, responsavel, observacoes,
                                    data_proxima_acao, ultima_atualizacao, usuario_id, tipo
                                ) VALUES (?, ?, ?, ?, ?, ?, datetime('now'), ?, ?, ?, date('now', '+7 days'), datetime('now'), ?, ?)
                            """, (
                                nome, empresa, telefone, email, status, valor,
                                origem, responsavel, observacoes, self.email_usuario, tipo_padrao
                            ))

                            success_count += 1
                            
                    except Exception as e:
                        print(f"Erro ao importar cliente {nome}: {str(e)}")
                        continue
                
                self.conn.commit()
                self.carregar_clientes()
                QMessageBox.information(
                    self, "Importação Concluída",
                    f"{success_count} clientes importados com sucesso!\n"
                    f"Total de linhas no arquivo: {len(df)}"
                )
                
        except Exception as e:
            QMessageBox.critical(
                self, "Erro na Importação",
                f"Erro ao importar clientes:\n{str(e)}\n\n"
                "Certifique-se que o arquivo Excel contém pelo menos as colunas 'nome' e 'email'."
            )

    def exportar_bitrix24(self):
        """Exporta clientes selecionados para o Bitrix24"""
        try:
            selected_ids = []
            for row in range(self.clientes_table.rowCount()):
                checkbox = self.clientes_table.cellWidget(row, 0)
                if checkbox and checkbox.isChecked():
                    id_item = self.clientes_table.item(row, 1)
                    if id_item:
                        selected_ids.append(int(id_item.text()))
            
            if not selected_ids:
                QMessageBox.warning(self, "Aviso", "Selecione pelo menos um cliente para exportar!")
                return
            
            # Corrigindo a query SQL
            placeholders = ','.join(['?'] * len(selected_ids))
            query = f"""
                SELECT nome, empresa, telefone, email, status, valor_proposta, 
                       data_cadastro, origem, responsavel, observacoes
                FROM clientes
                WHERE id IN ({placeholders})
            """
            
            if not self.is_admin:
                query += " AND usuario_id = ?"
                params = selected_ids + [self.email_usuario]
            else:
                params = selected_ids
            
            self.cursor.execute(query, params)
            dados = self.cursor.fetchall()
        
        
            
            # Criar DataFrame
            from pandas import DataFrame
            df = DataFrame(dados, columns=[
                "Nome", "Empresa", "Telefone", "Email", "Status",
                "Valor Proposta", "Data Cadastro", "Origem", "Responsável", "Observações"
            ])
            
            # Salvar arquivo CSV formatado para Bitrix24
            file_name, _ = QFileDialog.getSaveFileName(
                self, "Exportar para Bitrix24", "", "CSV Files (*.csv)"
            )
            
            if file_name:
                df.to_csv(file_name, index=False, sep=';', encoding='utf-8-sig')
                QMessageBox.information(
                    self, "Sucesso", 
                    f"{len(dados)} clientes exportados para:\n{file_name}"
                )
                
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao exportar para Bitrix24:\n{str(e)}")

    def create_bitrix_page(self):
        """Cria a página de integração com Bitrix24"""
        self.bitrix_page = QWidget()
        layout = QVBoxLayout(self.bitrix_page)
        layout.setContentsMargins(20, 20, 20, 20)
        
        title = QLabel("🔗 Integração com Bitrix24")
        title.setStyleSheet("font-size: 22px; font-weight: bold; color: white;")
        layout.addWidget(title)
        
        # Mensagem de placeholder
        msg = QLabel("Esta funcionalidade estará disponível em breve.\n\n"
                    "Aqui você poderá:\n"
                    "- Sincronizar clientes com o Bitrix24\n"
                    "- Visualizar oportunidades\n"
                    "- Gerenciar tarefas e pipelines")
        msg.setStyleSheet("font-size: 16px; color: white;")
        layout.addWidget(msg)
        
        # Botão de configuração
        btn_config = QPushButton("Configurar Conexão")
        btn_config.setStyleSheet("""
            QPushButton {
                background: #3498db;
                color: white;
                padding: 12px;
                border-radius: 5px;
                font-size: 16px;
            }
            QPushButton:hover {
                background: #2980b9;
            }
        """)
        btn_config.setCursor(Qt.PointingHandCursor)
        btn_config.clicked.connect(self.configurar_bitrix)
        layout.addWidget(btn_config)
        
        layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))
        
        self.content_stack.addWidget(self.bitrix_page)
    
    def configurar_bitrix(self):
        """Mostra a janela de configuração do Bitrix24"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Configuração do Bitrix24")
        dialog.setFixedSize(500, 300)
        dialog.setStyleSheet("""
            QDialog {
                background: #2c3e50;
                color: white;
            }
            QLineEdit {
                background: #34495e;
                color: white;
                padding: 8px;
                border: 1px solid #2c3e50;
                border-radius: 4px;
            }
        """)
        
        layout = QVBoxLayout(dialog)
        
        form = QFormLayout()
        
        # Campos de configuração
        self.bitrix_url = QLineEdit()
        self.bitrix_url.setPlaceholderText("https://seu-dominio.bitrix24.com.br")
        form.addRow("URL do Bitrix24:", self.bitrix_url)
        
        self.bitrix_user = QLineEdit()
        self.bitrix_user.setPlaceholderText("ID do Usuário")
        form.addRow("Usuário:", self.bitrix_user)
        
        self.bitrix_key = QLineEdit()
        self.bitrix_key.setPlaceholderText("Chave de Acesso")
        form.addRow("Chave de Acesso:", self.bitrix_key)
        
        # Carrega configurações salvas
        settings = QSettings("GustavoMartins", "CRMApp")
        self.bitrix_url.setText(settings.value("bitrix_url", ""))
        self.bitrix_user.setText(settings.value("bitrix_user", ""))
        self.bitrix_key.setText(settings.value("bitrix_key", ""))
        
        layout.addLayout(form)
        
        # Botões
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(lambda: self.salvar_config_bitrix(dialog))
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        dialog.exec_()
    
    def salvar_config_bitrix(self, dialog):
        """Salva as configurações do Bitrix24"""
        url = self.bitrix_url.text().strip()
        user = self.bitrix_user.text().strip()
        key = self.bitrix_key.text().strip()
        
        if not url or not user or not key:
            QMessageBox.warning(self, "Aviso", "Preencha todos os campos!")
            return
        
        settings = QSettings("GustavoMartins", "CRMApp")
        settings.setValue("bitrix_url", url)
        settings.setValue("bitrix_user", user)
        settings.setValue("bitrix_key", key)
        
        dialog.accept()
        QMessageBox.information(self, "Sucesso", "Configurações salvas com sucesso!")

    def create_agenda_page(self):
        """Cria a página de agenda integrada"""
        self.agenda_page = QWidget()
        layout = QVBoxLayout(self.agenda_page)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        title = QLabel("📅 Agenda de Atendimentos")
        title.setStyleSheet("font-size: 22px; font-weight: bold; color: white;")
        layout.addWidget(title)

        # Filtros
        filter_row = QHBoxLayout()
        
        self.agenda_filter = QComboBox()
        self.agenda_filter.addItems(["Todos", "Hoje", "Esta semana", "Próximos 30 dias", "Mês atual", "Todos os agendamentos"])
        self.agenda_filter.setCurrentText("Mês atual")
        self.agenda_filter.setStyleSheet("""
            QComboBox {
                background: #2c3e50;
                color: white;
                padding: 8px;
                border: 1px solid #34495e;
                border-radius: 4px;
            }
            QComboBox QAbstractItemView {
                background: #34495e;
                color: white;
                selection-background-color: #3498db;
            }
        """)
        self.agenda_filter.currentTextChanged.connect(self.carregar_agendamentos)
        filter_row.addWidget(self.agenda_filter)
        
        btn_sync = QPushButton("Sincronizar com Google Agenda")
        btn_sync.setStyleSheet("""
            QPushButton {
                background: #3498db;
                color: white;
                padding: 8px 15px;
                border-radius: 4px;
                font-size: 14px;
            }
            QPushButton:hover {
                background: #2980b9;
            }
        """)
        btn_sync.setCursor(Qt.PointingHandCursor)
        btn_sync.clicked.connect(self.sincronizar_google_agenda)
        filter_row.addWidget(btn_sync)
        
        layout.addLayout(filter_row)

        # Tabela de agendamentos
        self.agenda_table = QTableWidget()
        self.agenda_table.setColumnCount(7)
        self.agenda_table.setHorizontalHeaderLabels(["Data", "Hora", "Título", "Cliente", "Tipo", "Status", ""])
        self.agenda_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.agenda_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.agenda_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.agenda_table.setStyleSheet("""
            QTableWidget {
                background: #2c3e50;
                color: white;
                border: 1px solid #34495e;
                gridline-color: #34495e;
            }
            QHeaderView::section {
                background: #34495e;
                color: white;
                padding: 8px;
            }
        """)


        self.agenda_search = QLineEdit()
        self.agenda_search.setPlaceholderText("Pesquisar na agenda...")
        filter_row.addWidget(self.agenda_search)
        self.agenda_search.textChanged.connect(self.filtrar_agenda_ao_vivo)
        
        layout.addWidget(self.agenda_table)
        self.carregar_agendamentos()

        self.content_stack.addWidget(self.agenda_page)

    def sincronizar_google_agenda_automatico(self):
        """Sincroniza automaticamente com o Google Agenda sem mostrar mensagens"""
        try:
            from google.oauth2.credentials import Credentials
            from googleapiclient.discovery import build
            from google.auth.transport.requests import Request
            import os.path

            SCOPES = ['https://www.googleapis.com/auth/calendar']
            creds = None

            if os.path.exists('token.json'):
                creds = Credentials.from_authorized_user_file('token.json', SCOPES)

            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    return  # Não tem credenciais válidas, não faz nada

            service = build('calendar', 'v3', credentials=creds)
            
            # Obtém eventos de todos os tempos
            events_result = service.events().list(
                calendarId='primary',
                maxResults=2500,
                singleEvents=True,
                orderBy='startTime'
            ).execute()
            
            events = events_result.get('items', [])
            
            if not events:
                return

            # Limpa a tabela de agendamentos antes de importar
            query = "DELETE FROM agendamentos"
            if not self.is_admin:
                query += " WHERE usuario_id = ?"
                self.cursor.execute(query, (self.email_usuario,))
            else:
                self.cursor.execute(query)
            
            # Importa para o banco de dados
            imported_count = 0
            for event in events:
                try:
                    start = event['start'].get('dateTime', event['start'].get('date'))
                    end = event['end'].get('dateTime', event['end'].get('date'))
                    
                    # Formata a data/hora corretamente
                    if 'dateTime' in event['start']:
                        start_dt = datetime.strptime(start, '%Y-%m-%dT%H:%M:%S%z')
                        end_dt = datetime.strptime(end, '%Y-%m-%dT%H:%M:%S%z')
                        start_str = start_dt.strftime('%Y-%m-%dT%H:%M:%S')
                        end_str = end_dt.strftime('%Y-%m-%dT%H:%M:%S')
                    else:
                        start_str = start + "T00:00:00"
                        end_str = end + "T00:00:00"
                    
                    # Verifica se é um evento de reunião com clientes
                    cliente_id = None
                    observacoes = ""
                    if 'attendees' in event:
                        for attendee in event['attendees']:
                            if 'email' in attendee:
                                # Tenta encontrar o cliente pelo email
                                self.cursor.execute(
                                    "SELECT id FROM clientes WHERE email = ?",
                                    (attendee['email'],))
                                result = self.cursor.fetchone()
                                if result:
                                    cliente_id = result[0]
                                    break
                                
                    
                    # Pega a descrição como observações
                    observacoes = event.get('description', '')
                    
                    # Insere no banco de dados
                    self.cursor.execute("""
                        INSERT INTO agendamentos 
                        (titulo, descricao, data_inicio, data_fim, tipo, status, cliente_id, usuario_id)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        event.get('summary', 'Sem título'),
                        observacoes,
                        start_str,
                        end_str,
                        event.get('eventType', 'Reunião'),
                        'Confirmado',
                        cliente_id,
                        self.email_usuario
                    ))
                    imported_count += 1
                    
                except Exception as e:
                    print(f"Erro ao importar evento {event.get('summary')}: {str(e)}")
                    continue
            
            self.conn.commit()
            self.carregar_agendamentos()
            
        except Exception as e:
            print(f"Erro na sincronização automática: {str(e)}")

    def sincronizar_google_agenda(self):
        """Sincroniza com a API REAL do Google Agenda - Versão melhorada"""
        try:
            from google.oauth2.credentials import Credentials
            from googleapiclient.discovery import build
            from google.auth.transport.requests import Request
            import os.path
            import sys

            # Configurações
            SCOPES = ['https://www.googleapis.com/auth/calendar']
            creds = None
    
            # Determina o diretório base corretamente para executável ou script
            if getattr(sys, 'frozen', False):
                base_dir = os.path.dirname(sys.executable)
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))
        
            token_path = os.path.join(base_dir, 'token.json')
            credentials_path = os.path.join(base_dir, 'credentials.json')
        
            # O arquivo token.json armazena os tokens de acesso/refresh
            if os.path.exists(token_path):
                creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        
            # Se não houver credenciais válidas, faz login
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    from google_auth_oauthlib.flow import InstalledAppFlow
                    if not os.path.exists(credentials_path):
                        QMessageBox.critical(self, "Erro", 
                            f"Arquivo 'credentials.json' não encontrado em:\n{credentials_path}\n"
                            "Por favor, coloque o arquivo no mesmo diretório do aplicativo.")
                        return
                    
                    flow = InstalledAppFlow.from_client_secrets_file(
                        credentials_path, SCOPES)
                    creds = flow.run_local_server(port=0)
                
                # Salva as credenciais para a próxima vez
                with open(token_path, 'w') as token:
                    token.write(creds.to_json())
    
            service = build('calendar', 'v3', credentials=creds)
            
            # Obtém eventos de todos os tempos
            events_result = service.events().list(
                calendarId='primary',
                maxResults=2500,
                singleEvents=True,
                orderBy='startTime'
            ).execute()
            
            events = events_result.get('items', [])
            
            if not events:
                QMessageBox.information(self, "Informação", "Nenhum evento encontrado na agenda.")
                return
            
            # Limpa a tabela de agendamentos antes de importar
            query = "DELETE FROM agendamentos"
            if not self.is_admin:
                query += " WHERE usuario_id = ?"
                self.cursor.execute(query, (self.email_usuario,))
            else:
                self.cursor.execute(query)
            
            # Importa para o banco de dados
            imported_count = 0
            for event in events:
                try:
                    start = event['start'].get('dateTime', event['start'].get('date'))
                    end = event['end'].get('dateTime', event['end'].get('date'))
                    
                    # Formata a data/hora corretamente
                    if 'dateTime' in event['start']:
                        start_dt = datetime.strptime(start, '%Y-%m-%dT%H:%M:%S%z')
                        end_dt = datetime.strptime(end, '%Y-%m-%dT%H:%M:%S%z')
                        start_str = start_dt.strftime('%Y-%m-%dT%H:%M:%S')
                        end_str = end_dt.strftime('%Y-%m-%dT%H:%M:%S')
                    else:
                        start_str = start + "T00:00:00"
                        end_str = end + "T00:00:00"
                    
                    # Verifica se é um evento de reunião com clientes
                    cliente_id = None
                    observacoes = ""
                    if 'attendees' in event:
                        for attendee in event['attendees']:
                            if 'email' in attendee:
                                # Tenta encontrar o cliente pelo email
                                self.cursor.execute(
                                    "SELECT id FROM clientes WHERE email = ?",
                                    (attendee['email'],)
                                )
                                result = self.cursor.fetchone()
                                if result:
                                    cliente_id = result[0]
                                    break
                    
                    # Pega a descrição como observações
                    observacoes = event.get('description', '')
                    
                    # Insere no banco de dados
                    self.cursor.execute("""
                        INSERT INTO agendamentos 
                        (titulo, descricao, data_inicio, data_fim, tipo, status, cliente_id, usuario_id)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        event.get('summary', 'Sem título'),
                        observacoes,
                        start_str,
                        end_str,
                        event.get('eventType', 'Reunião'),
                        'Confirmado',
                        cliente_id,
                        self.email_usuario
                    ))
                    imported_count += 1
                    
                except Exception as e:
                    print(f"Erro ao importar evento {event.get('summary')}: {str(e)}")
                    continue
            
            self.conn.commit()
            self.carregar_agendamentos()
            
            # Mostra resumo da importação
            QMessageBox.information(
                self, "Sincronização Concluída",
                f"Total de eventos encontrados: {len(events)}\n"
                f"Eventos importados com sucesso: {imported_count}\n\n"
                "Os eventos agora aparecem na sua agenda local."
            )
        
        except Exception as e:
            QMessageBox.critical(
                self, "Erro na Sincronização",
                f"Falha ao sincronizar com o Google Agenda:\n{str(e)}\n\n"
                "Certifique-se que você tem permissão para acessar o Google Agenda."
            )

    def create_relatorios_page(self):
        """Cria a página de relatórios com gráficos e dados consolidados"""
        self.relatorios_page = QWidget()
        layout = QVBoxLayout(self.relatorios_page)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
    
        title = QLabel("📊 Relatórios Consolidados")
        title.setStyleSheet("font-size: 22px; font-weight: bold; color: white;")
        layout.addWidget(title)
    
        # Filtros
        filter_row = QHBoxLayout()
        
        self.relatorio_mes_filter = QComboBox()
        meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        self.relatorio_mes_filter.addItems(meses)
        self.relatorio_mes_filter.setCurrentIndex(datetime.now().month - 1)
        filter_row.addWidget(self.relatorio_mes_filter)
        
        self.relatorio_ano_filter = QComboBox()
        anos = [str(datetime.now().year - 1), str(datetime.now().year), str(datetime.now().year + 1)]
        self.relatorio_ano_filter.addItems(anos)
        self.relatorio_ano_filter.setCurrentText(str(datetime.now().year))
        filter_row.addWidget(self.relatorio_ano_filter)
        
        btn_filter = QPushButton("Filtrar")
        btn_filter.clicked.connect(self.atualizar_relatorios)
        filter_row.addWidget(btn_filter)
        
        btn_export = QPushButton("Exportar para Excel")
        btn_export.clicked.connect(self.exportar_relatorio)
        filter_row.addWidget(btn_export)
        
        layout.addLayout(filter_row)
    
        # Gráfico de status
        self.relatorio_chart1 = QChart()
        self.relatorio_chart1.setTitle("Status dos Clientes")
        self.relatorio_chart1.setAnimationOptions(QChart.SeriesAnimations)
        
        chart_view1 = QChartView(self.relatorio_chart1)
        chart_view1.setRenderHint(QPainter.Antialiasing)
        layout.addWidget(chart_view1)
    
        # Gráfico de conversão
        self.relatorio_chart2 = QChart()
        self.relatorio_chart2.setTitle("Conversão por Mês")
        self.relatorio_chart2.setAnimationOptions(QChart.SeriesAnimations)
        
        chart_view2 = QChartView(self.relatorio_chart2)
        chart_view2.setRenderHint(QPainter.Antialiasing)
        layout.addWidget(chart_view2)
    
        # Tabela de resumo
        self.relatorio_table = QTableWidget()
        self.relatorio_table.setColumnCount(3)
        self.relatorio_table.setHorizontalHeaderLabels(["Métrica", "Valor", "Variação"])
        self.relatorio_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.relatorio_table)
    
        # Atualiza os dados iniciais
        self.atualizar_relatorios()
    
        self.content_stack.addWidget(self.relatorios_page)

    def atualizar_relatorios(self):
        """Atualiza todos os gráficos e dados na página de relatórios"""
        mes = self.relatorio_mes_filter.currentIndex() + 1
        ano = int(self.relatorio_ano_filter.currentText())
        
        try:
            # Atualiza os dados primeiro
            self.atualizar_metricas()
        
            # Força a atualização dos gráficos
            self.relatorio_chart1.update()
            self.relatorio_chart2.update()
            self.relatorio_table.viewport().update()
        
        except Exception as e:
            print(f"Erro ao atualizar relatórios: {str(e)}")


        # Gráfico 1 - Status dos Clientes
        self.relatorio_chart1.removeAllSeries()
        series1 = QPieSeries()
        
        status_counts = {
            "Cliente": self.get_count("clientes", "status = 'Cliente'"),
            "Em atendimento": self.get_count("clientes", "status = 'Em atendimento'"),
            "Proposta": self.get_count("clientes", "status = 'Proposta'"),
            "Finalizado": self.get_count("clientes", "status = 'Finalizado'")
        }
        
        for status, count in status_counts.items():
            if count > 0:
                slice_ = series1.append(f"{status}: {count}", count)
                if status == "Cliente":
                    slice_.setColor(QColor(155, 89, 182))
                elif status == "Em atendimento":
                    slice_.setColor(QColor(52, 152, 219))
                elif status == "Proposta":
                    slice_.setColor(QColor(241, 196, 15))
                elif status == "Finalizado":
                    slice_.setColor(QColor(46, 204, 113))
        
        self.relatorio_chart1.addSeries(series1)
        
        # Gráfico 2 - Conversão por Mês
        self.relatorio_chart2.removeAllSeries()
        series2 = QBarSeries()
        
        bar_set_total = QBarSet("Total Clientes")
        bar_set_final = QBarSet("Finalizados")
        
        meses = []
        totais = []
        finalizados = []
        
        for m in range(1, 13):
            mes_ano = f"{ano}-{m:02d}"
            total = self.get_count("clientes", f"strftime('%Y-%m', data_cadastro) = '{mes_ano}'")
            fin = self.get_count("clientes", f"status = 'Finalizado' AND strftime('%Y-%m', data_cadastro) = '{mes_ano}'")
            
            meses.append(calendar.month_abbr[m])
            totais.append(total)
            finalizados.append(fin)
        
        bar_set_total.append(totais)
        bar_set_final.append(finalizados)
        series2.append(bar_set_total)
        series2.append(bar_set_final)
        self.relatorio_chart2.addSeries(series2)
        
        # Configura eixos do gráfico 2
        axis_x = QBarCategoryAxis()
        axis_x.append(meses)
        self.relatorio_chart2.createDefaultAxes()
        self.relatorio_chart2.setAxisX(axis_x, series2)
        
        # Tabela de resumo
        self.relatorio_table.setRowCount(5)
        
        metricas = [
            ("Total Clientes", self.get_count("clientes")),
            ("Clientes Ativos", self.get_count("clientes", "status = 'Cliente'")),
            ("Em Atendimento", self.get_count("clientes", "status = 'Em atendimento'")),
            ("Propostas", self.get_count("clientes", "status = 'Proposta'")),
            ("Finalizados", self.get_count("clientes", "status = 'Finalizado'"))
        ]
        
        for row, (nome, valor) in enumerate(metricas):
            self.relatorio_table.setItem(row, 0, QTableWidgetItem(nome))
            self.relatorio_table.setItem(row, 1, QTableWidgetItem(str(valor)))
            
            # Cálculo de variação (simplificado)
            variacao = "+0%"
            if row > 0:
                total = metricas[0][1]
                if total > 0:
                    percent = (valor / total) * 100
                    variacao = f"{percent:.1f}%"
            
            self.relatorio_table.setItem(row, 2, QTableWidgetItem(variacao))

        
    def atualizar_grafico_conversao(self):
        """Método à prova de falhas para atualização do gráfico"""
        try:
            # Verificação EXTRA de segurança
            if not hasattr(self, 'chart3'):
                print("CRÍTICO: chart3 não existe - criando novo")
                self.chart3 = QChart()
                
            if self.chart3 is None:
                print("CRÍTICO: chart3 é None - recriando")
                self.chart3 = QChart()
                
            # Limpeza segura
            if isinstance(self.chart3, QChart):
                self.chart3.removeAllSeries()
            else:
                raise TypeError(f"chart3 não é QChart (tipo: {type(self.chart3)})")
            
            # [SEU CÓDIGO DE CONSULTA AO BANCO AQUI]
            # Exemplo:
            query = """
                SELECT strftime('%Y-%m', data_cadastro) as mes,
                       COUNT(*) as total,
                       SUM(CASE WHEN status = 'Finalizado' THEN 1 ELSE 0 END) as finalizados
                FROM clientes
                WHERE usuario_id = ?
                GROUP BY mes
                ORDER BY mes DESC
                LIMIT 6
            """
            self.cursor.execute(query, (self.email_usuario,))
            dados = self.cursor.fetchall()
            
            if not dados:
                print("AVISO: Nenhum dado encontrado para o gráfico")
                empty_series = QPieSeries()
                empty_series.append("Sem dados", 1)
                self.chart3.addSeries(empty_series)
                return
                
            # Processamento dos dados
            meses = [mes.replace("-", "/") for mes, _, _ in dados[::-1]]
            totais = [total for _, total, _ in dados[::-1]]
            finalizados = [fin for _, _, fin in dados[::-1]]
            
            # Criação das séries
            series = QBarSeries()
            
            bar_set_total = QBarSet("Total Clientes")
            bar_set_total.append(totais)
            bar_set_total.setColor(QColor(52, 152, 219))
            
            bar_set_final = QBarSet("Finalizados")
            bar_set_final.append(finalizados)
            bar_set_final.setColor(QColor(46, 204, 113))
            
            series.append(bar_set_total)
            series.append(bar_set_final)
            self.chart3.addSeries(series)
            
            # Configuração dos eixos
            axis_x = QBarCategoryAxis()
            axis_x.append(meses)
            self.chart3.setAxisX(axis_x, series)
            
            axis_y = QValueAxis()
            axis_y.setRange(0, max(max(totais), max(finalizados)) * 1.15)
            self.chart3.setAxisY(axis_y, series)
            
            print("SUCESSO: Gráfico atualizado")
            
        except Exception as e:
            print(f"FALHA NA ATUALIZAÇÃO: {str(e)}")
            # Recuperação elegante
            self.chart3 = QChart()  # Recriação completa
            error_series = QPieSeries()
            error_series.append("Erro ao carregar", 1)
            self.chart3.addSeries(error_series)
        
        

    def carregar_clientes(self):
        filtro = self.status_filter.currentText()
        busca = self.busca_input.text().lower()
        empresa_filtro = self.empresa_filter.currentText() if hasattr(self, 'empresa_filter') else None
    
        query = "SELECT id, nome, empresa, telefone, email, status, valor_proposta, data_cadastro, tipo FROM clientes"
        conditions = []
        params = []
    
        if filtro != "Todos":
            conditions.append("status = ?")
            params.append(filtro)
        
        if busca:
            conditions.append("(LOWER(nome) LIKE ? OR LOWER(empresa) LIKE ? OR LOWER(email) LIKE ?)")
            params.extend([f"%{busca}%", f"%{busca}%", f"%{busca}%"])
        
        if empresa_filtro and empresa_filtro != "Todas":
            conditions.append("tipo = ?")
            params.append(empresa_filtro)
        
        # Se não for admin, filtrar apenas pelos clientes do usuário e da sua empresa
        if not self.is_admin:
            conditions.append("usuario_id = ?")
            params.append(self.email_usuario)
            conditions.append("tipo = ?")
            params.append(self.empresa_usuario)
        
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
        
        query += " ORDER BY data_cadastro DESC"
        
        self.cursor.execute(query, params)
        clientes = self.cursor.fetchall()
        
    
        # Ajuste o número de colunas para 10 (incluindo o tipo)
        self.clientes_table.setRowCount(len(clientes))
        for row, (id_, nome, empresa, telefone, email, status, valor, cadastro, tipo) in enumerate(clientes):
            # Checkbox para seleção
            checkbox = QCheckBox()
            checkbox.setStyleSheet("margin-left:5px;")
            self.clientes_table.setCellWidget(row, 0, checkbox)
            
            # Demais colunas
            self.clientes_table.setItem(row, 1, QTableWidgetItem(str(id_)))
            self.clientes_table.setItem(row, 2, QTableWidgetItem(nome))
            self.clientes_table.setItem(row, 3, QTableWidgetItem(empresa if empresa else "-"))
            self.clientes_table.setItem(row, 4, QTableWidgetItem(telefone if telefone else "-"))
            self.clientes_table.setItem(row, 5, QTableWidgetItem(email if email else "-"))
            
            status_item = QTableWidgetItem(status)
            if status == "Cliente":
                status_item.setBackground(QColor(155, 89, 182))  # Roxo
            elif status == "Em atendimento":
                status_item.setBackground(QColor(52, 152, 219))  # Azul
            elif status == "Proposta":
                status_item.setBackground(QColor(241, 196, 15))  # Amarelo
            elif status == "Finalizado":
                status_item.setBackground(QColor(46, 204, 113))  # Verde
            self.clientes_table.setItem(row, 6, status_item)
            
            # Formata o valor
            valor_formatado = f"R${valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            valor_item = QTableWidgetItem(valor_formatado)
            valor_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.clientes_table.setItem(row, 7, valor_item)
            
            self.clientes_table.setItem(row, 8, QTableWidgetItem(cadastro[:10]))
            
            # Adiciona o tipo (Precode/Allpost)
            self.clientes_table.setItem(row, 9, QTableWidgetItem(tipo))
            
            # Botão de ações (agora na coluna 10)
            btn_acoes = QPushButton("...")
            btn_acoes.setStyleSheet("""
                QPushButton {
                    background: #95a5a6;
                    color: white;
                    border-radius: 3px;
                    padding: 3px 8px;
                }
                QPushButton:hover {
                    background: #7f8c8d;
                }
            """)
            btn_acoes.setCursor(Qt.PointingHandCursor)
            btn_acoes.setProperty("id", id_)
            btn_acoes.clicked.connect(self.mostrar_menu_cliente)
            self.clientes_table.setCellWidget(row, 10, btn_acoes)
        
        # Ajusta os cabeçalhos
        self.clientes_table.setHorizontalHeaderLabels(["", "ID", "Nome", "Empresa", "Telefone", "Email", "Status", "Valor", "Cadastro", "Tipo", ""])
        
        # Ajusta a largura da coluna de seleção
        self.clientes_table.setColumnWidth(0, 30)
        
        # Atualiza os cards de métricas
        self.atualizar_metricas()
    
    def carregar_agendamentos(self):
        filtro = self.agenda_filter.currentText()
    
        query = """
            SELECT a.id, a.data_inicio, a.titulo, c.nome, a.tipo, a.status, a.descricao 
            FROM agendamentos a
            LEFT JOIN clientes c ON a.cliente_id = c.id
        """
        
        conditions = []
        params = []
        
        if filtro == "Hoje":
            conditions.append("date(a.data_inicio) = date('now')")
        elif filtro == "Esta semana":
            conditions.append("date(a.data_inicio) BETWEEN date('now') AND date('now', '+6 days')")
        elif filtro == "Próximos 30 dias":
            conditions.append("date(a.data_inicio) BETWEEN date('now') AND date('now', '+30 days')")
        elif filtro == "Mês atual":
            # Modificado para mostrar todo o mês, incluindo dias passados
            conditions.append("strftime('%Y-%m', a.data_inicio) = strftime('%Y-%m', 'now')")
        
        if not self.is_admin:
            conditions.append("a.usuario_id = ?")
            params.append(self.email_usuario)
        
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
        
        query += " ORDER BY a.data_inicio"
        
        self.cursor.execute(query, params)
        agendamentos = self.cursor.fetchall()
        # Restante do código...
        
        self.agenda_table.setRowCount(len(agendamentos))
        for row, (agendamento_id, data, titulo, cliente, tipo, status, descricao) in enumerate(agendamentos):
            try:
                data_obj = datetime.strptime(data, "%Y-%m-%dT%H:%M:%S")
                data_str = data_obj.strftime("%d/%m/%Y")
                hora_str = data_obj.strftime("%H:%M")
            except:
                data_str = data[:10]
                hora_str = "-"
            
            self.agenda_table.setItem(row, 0, QTableWidgetItem(data_str))
            self.agenda_table.setItem(row, 1, QTableWidgetItem(hora_str))
            titulo_com_status = f"{titulo} ({status})" if status else titulo
            titulo_item = QTableWidgetItem(titulo_com_status)
            titulo_item.setToolTip(titulo_com_status)
            titulo_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            titulo_item.setTextAlignment(Qt.AlignLeft)
            titulo_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            self.agenda_table.setItem(row, 2, titulo_item)
            self.agenda_table.setRowHeight(row, 50)  # Aumenta altura da linha
            self.agenda_table.setColumnWidth(2, 600)  # índice 2 = Título

            
            # Exibe cliente e observações
            cliente_text = f"{cliente if cliente else '-'}"
            if descricao:
                cliente_text += f"\nObs: {descricao}"
            cliente_item = QTableWidgetItem(cliente_text)
            cliente_item.setToolTip(descricao if descricao else "")
            self.agenda_table.setItem(row, 3, cliente_item)
            
            self.agenda_table.setItem(row, 4, QTableWidgetItem(tipo))
        

            
            status_item = QTableWidgetItem(status)
            if status == "Confirmado":
                status_item.setBackground(QColor(46, 204, 113))
            elif status == "Cancelado":
                status_item.setBackground(QColor(231, 76, 60))
            elif status == "Pendente":
                status_item.setBackground(QColor(241, 196, 15))
            self.agenda_table.setItem(row, 5, status_item)
            
            self.agenda_table.setColumnHidden(4, True)

           

            
            # Botão de ações
            btn_acoes = QPushButton("...")
            btn_acoes.setStyleSheet("""
                QPushButton {
                    background: #95a5a6;
                    color: white;
                    border-radius: 3px;
                    padding: 3px 8px;
                }
                QPushButton:hover {
                    background: #7f8c8d;
                }
            """)
            btn_acoes.setCursor(Qt.PointingHandCursor)
            btn_acoes.setProperty("id", agendamento_id)
            btn_acoes.clicked.connect(self.mostrar_menu_agendamento)
            self.agenda_table.setCellWidget(row, 6, btn_acoes)
            
        
        # Atualiza os cards de métricas
        self.atualizar_metricas()


    def filtrar_agenda_ao_vivo(self):
        texto = self.agenda_search.text().lower()
        for row in range(self.agenda_table.rowCount()):
            mostrar = False
            for col in range(self.agenda_table.columnCount()):
                item = self.agenda_table.item(row, col)
                if item and texto in item.text().lower():
                    mostrar = True
                    break
            self.agenda_table.setRowHidden(row, not mostrar)

    def atualizar_metricas(self):
        """Atualiza todas as métricas e gráficos do dashboard"""
        try:
            mes = self.month_filter.currentIndex() + 1
            ano = int(self.year_filter.currentText())
            mes_ano = f"{ano}-{mes:02d}"
    
        # Condição base para usuários não-admin
            user_condition = "" if self.is_admin else "usuario_id = ?"
            user_params = [] if self.is_admin else [self.email_usuario]
    
            # Métricas básicas
            total_clientes = self.get_count("clientes", user_condition, user_params)
            clientes = self.get_count("clientes", f"status = 'Cliente' {self._and(user_condition)}", user_params)
            em_atendimento = self.get_count("clientes", f"status = 'Em atendimento' {self._and(user_condition)}", user_params)
            propostas = self.get_count("clientes", f"status = 'Proposta' {self._and(user_condition)}", user_params)
    
            # Fechamentos do mês (apenas futuros)
            finalizados_condition = f"status = 'Finalizado' AND strftime('%Y-%m', data_cadastro) = ? {self._and(user_condition)}"

            finalizados_params = [mes_ano] + ([] if self.is_admin else user_params)
    
            finalizados = self.get_count("clientes", finalizados_condition, finalizados_params)
            valor_fechado = self.get_sum("clientes", "valor_proposta", finalizados_condition, finalizados_params)
    
            # Atualiza os cards
            self.metric_cards["total_clientes"].update_value(f"👥 {total_clientes}")
            self.metric_cards["clientes"].update_value(f"👤 {clientes}")
            self.metric_cards["em_atendimento"].update_value(f"📞 {em_atendimento}")
            self.metric_cards["propostas"].update_value(f"📄 {propostas}")
            self.metric_cards["finalizados"].update_value(f"✅ {finalizados}")
            
            # Formata o valor monetário
            valor_formatado = locale.currency(valor_fechado, grouping=True)
            self.metric_cards["valor_fechado"].update_value(f"💰 {valor_formatado}")
    
        except Exception as e:
            print(f"Erro ao atualizar métricas: {str(e)}")
    
        # Atualiza gráficos
        self.atualizar_grafico_status()
        self.atualizar_grafico_fechamentos(mes, ano)
        self.atualizar_grafico_conversao()
        self.atualizar_grafico_valor()
        
        # Atualiza tabela de próximos agendamentos
        self.carregar_proximos_agendamentos()
        
        self.atualizar_grafico_fechamentos(
            self.month_filter.currentIndex() + 1,
            int(self.year_filter.currentText())
        )


    def _and(self, condition):
        """Helper para adicionar AND corretamente"""
        return f"AND {condition}" if condition else ""
        
        # Atualiza gráficos
        if hasattr(self, 'chart1'):
            self.atualizar_grafico_status()
        
        if hasattr(self, 'chart2'):
            self.atualizar_grafico_fechamentos(mes, ano)
        
        if hasattr(self, 'chart3'):
            self.atualizar_grafico_conversao()
        
        if hasattr(self, 'chart4'):
            self.atualizar_grafico_valor()
        
        # Atualiza tabela de próximos agendamentos
        query = """
            SELECT a.data_inicio, a.titulo, c.nome, a.tipo, a.descricao 
            FROM agendamentos a
            LEFT JOIN clientes c ON a.cliente_id = c.id
            WHERE date(a.data_inicio) >= date('now')
        """
        
        if not self.is_admin:
            query += " AND a.usuario_id = ?"
            params = [self.email_usuario]
        else:
            params = []
            
        query += " ORDER BY a.data_inicio LIMIT 10"
        
        self.cursor.execute(query, params)
        agendamentos = self.cursor.fetchall()
        
        self.agendamentos_table.setRowCount(len(agendamentos))
        for row, (data, titulo, cliente, tipo, descricao) in enumerate(agendamentos):
            try:
                data_obj = datetime.strptime(data, "%Y-%m-%dT%H:%M:%S")
                data_str = data_obj.strftime("%d/%m/%Y")
                hora_str = data_obj.strftime("%H:%M")
            except:
                data_str = data[:10]
                hora_str = "-"
            
            self.agendamentos_table.setItem(row, 0, QTableWidgetItem(data_str))
            self.agendamentos_table.setItem(row, 1, QTableWidgetItem(hora_str))
            self.agendamentos_table.setItem(row, 2, QTableWidgetItem(titulo))
            
            # Exibe cliente e observações
            cliente_text = f"{cliente if cliente else '-'}"
            if descricao:
                cliente_text += f"\nObs: {descricao}"
            cliente_item = QTableWidgetItem(cliente_text)
            cliente_item.setToolTip(descricao if descricao else "")
            self.agendamentos_table.setItem(row, 3, cliente_item)
            
            self.agendamentos_table.setItem(row, 4, QTableWidgetItem(tipo))

    def atualizar_grafico_status(self):
        """Atualiza o gráfico de status dos clientes"""
        query = """
            SELECT status, COUNT(*) 
            FROM clientes
        """
        
        if not self.is_admin:
            query += " WHERE usuario_id = ?"
            params = (self.email_usuario,)
        else:
            params = ()
            
        query += " GROUP BY status"
        
        self.cursor.execute(query, params)
        dados = self.cursor.fetchall()
        
        status_counts = {
            "Cliente": 0,
            "Em atendimento": 0,
            "Proposta": 0,
            "Finalizado": 0
        }
        
        for status, count in dados:
            if status in status_counts:
                status_counts[status] = count
        
        series = QPieSeries()
        for status, count in status_counts.items():
            if count > 0:
                slice_ = series.append(f"{status}: {count}", count)
                if status == "Cliente":
                    slice_.setColor(QColor(155, 89, 182))
                elif status == "Em atendimento":
                    slice_.setColor(QColor(52, 152, 219))
                elif status == "Proposta":
                    slice_.setColor(QColor(241, 196, 15))
                elif status == "Finalizado":
                    slice_.setColor(QColor(46, 204, 113))
                
                # Configura tooltip
                slice_.setLabelVisible(True)
                slice_.setLabel(f"{status}: {count}")
        
        # Remove a série antiga e adiciona a nova
        self.chart1.removeAllSeries()
        if series.count() > 0:
            self.chart1.addSeries(series)
        else:
            series.append("Sem dados", 1)
            self.chart1.addSeries(series)

    def atualizar_grafico_fechamentos(self, mes, ano):
        """Atualiza o gráfico da dashboard com dados do mês atual filtrado"""
        mes_str = f"{mes:02d}"

        # Consulta no banco
        self.cursor.execute("""
            SELECT 
                COUNT(*) as total,
                SUM(CASE WHEN status = 'Finalizado' THEN 1 ELSE 0 END),
                SUM(CASE WHEN status = 'Em atendimento' THEN 1 ELSE 0 END),
                SUM(CASE WHEN status = 'Proposta' THEN 1 ELSE 0 END)
            FROM clientes
            WHERE strftime('%Y', data_cadastro) = ? 
              AND strftime('%m', data_cadastro) = ?
              AND usuario_id = ?
        """, (str(ano), mes_str, self.email_usuario))
        
        total, finalizados, em_atend, propostas = self.cursor.fetchone()
    
        total = total or 0
        finalizados = finalizados or 0
        em_atend = em_atend or 0
        propostas = propostas or 0
    
        bar_total = QBarSet("Total")
        bar_clientes = QBarSet("Clientes")
        bar_finalizados = QBarSet("Finalizados")
        bar_em_atendimento = QBarSet("Em Atendimento")
        bar_propostas = QBarSet("Propostas")
        bar_agendamentos = QBarSet("Agendamentos Google")

    
        bar_total << total
        bar_finalizados << finalizados
        bar_em_atendimento << em_atend
        bar_propostas << propostas
    
        # PyQt5 não suporta exibir os valores diretamente nas barras por padrão
        # Se quiser desenhar com QPainter no futuro, posso te ajudar com isso

    
        bar_total.setColor(QColor("#3498db"))
        bar_finalizados.setColor(QColor("#2ecc71"))
        bar_em_atendimento.setColor(QColor("#f39c12"))
        bar_propostas.setColor(QColor("#e74c3c"))
    
        series = QBarSeries()
        series.append(bar_total)
        series.append(bar_clientes)
        series.append(bar_finalizados)
        series.append(bar_em_atendimento)
        series.append(bar_propostas)
        series.append(bar_agendamentos)

        chart = QChart()
        chart.addSeries(series)
        chart.setTitle(f"Fechamentos - {calendar.month_name[mes]} {ano}")
        chart.setAnimationOptions(QChart.SeriesAnimations)
        chart.setTheme(QChart.ChartThemeDark)
    
        axis_x = QBarCategoryAxis()
        axis_x.append(["Resumo"])
        chart.setAxisX(axis_x, series)
    
        axis_y = QValueAxis()
        axis_y.setRange(0, max(total, finalizados, em_atend, propostas, 5) + 2)
        chart.setAxisY(axis_y, series)
    
        
        self.chart_view2.setChart(chart)
        self.chart2 = chart





    def atualizar_grafico_conversao(self):
        try:
            # Verificação crítica (ADICIONE ESTE IF)
            if not hasattr(self, 'chart3'):
                print("AVISO: Gráfico chart3 não foi criado ainda!")
                return
            
            # Restante do seu código original...
            if self.chart3 is None:
                self.chart3 = QChart()
            else:
                self.chart3.removeAllSeries()

        
            # 2. Busca dados (SUA QUERY ORIGINAL - não mudei)
            query = """
                SELECT strftime('%Y-%m', data_cadastro) as mes,
                       COUNT(*) as total,
                       SUM(CASE WHEN status = 'Finalizado' THEN 1 ELSE 0 END) as finalizados
                FROM clientes
                WHERE usuario_id = ?
                GROUP BY mes 
                ORDER BY mes DESC 
                LIMIT 6
            """
            self.cursor.execute(query, (self.email_usuario,))
            dados = self.cursor.fetchall()
            
            if not hasattr(self, 'chart3'):  # 👈 VERIFICAÇÃO DE SEGURANÇA
                print("Erro: Gráfico não inicializado!")
                return


            if not dados:
                raise ValueError("Sem dados para exibir")
            
            # 3. Prepara séries (MANTIDO como antes)
            series = QBarSeries()
            bar_set_total = QBarSet("Total Clientes")
            bar_set_finalizados = QBarSet("Finalizados")
            
            meses = []
            for mes, total, finalizados in dados[::-1]: 
                meses.append(mes.replace("-", "/"))
                bar_set_total.append(total)
                bar_set_finalizados.append(finalizados)
            
            series.append(bar_set_total)
            series.append(bar_set_finalizados)
            
            # 4. Configuração MÍNIMA que funcionava
            self.chart3.addSeries(series)
            
            # ==== O SEGREDO ESTÁ AQUI ====
            axis_x = QBarCategoryAxis()
            axis_x.append(meses)
            self.chart3.setAxisX(axis_x, series)
            
            axis_y = QValueAxis()
            axis_y.setRange(0, max(bar_set_total) * 1.1)
            self.chart3.setAxisY(axis_y, series)
            # ==============================
            
            # 5. Estilo (igual ao seu original)
            bar_set_total.setColor(QColor(52, 152, 219))  # Azul
            bar_set_finalizados.setColor(QColor(46, 204, 113))  # Verde
            
        # No final do método, antes do except:
            print("Dados usados no gráfico:", meses, list(bar_set_total), list(bar_set_finalizados))

        except Exception as e:
            print(f"Erro no gráfico: {e}")
            # Fallback elegante
            self.chart3.setTitle("Erro ao carregar dados")

    def atualizar_grafico_valor(self):
        """Atualiza o gráfico de valor por status"""
        query = """
            SELECT status, SUM(valor_proposta) 
            FROM clientes 
            WHERE valor_proposta > 0
        """
        
        if not self.is_admin:
            query += " AND usuario_id = ?"
            params = (self.email_usuario,)
        else:
            params = ()
            
        query += " GROUP BY status"
        
        self.cursor.execute(query, params)
        dados = self.cursor.fetchall()
        
        series = QPieSeries()
        for status, valor in dados:
            if valor:
                # Formata o valor com ponto para milhar e vírgula para decimal
                valor_formatado = f"R${valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                slice_ = series.append(f"{status} ({valor_formatado})", valor)
                if status == "Cliente":
                    slice_.setColor(QColor(155, 89, 182))
                elif status == "Em atendimento":
                    slice_.setColor(QColor(52, 152, 219))
                elif status == "Proposta":
                    slice_.setColor(QColor(241, 196, 15))
                elif status == "Finalizado":
                    slice_.setColor(QColor(46, 204, 113))
                
                # Configura tooltip
                slice_.setLabelVisible(True)
                slice_.setLabel(f"{status}: {valor_formatado}")
        
        # Remove a série antiga e adiciona a nova
        self.chart2.removeAllSeries()
        if series.count() > 0:
            self.chart2.addSeries(series)

            from calendar import monthrange
            from datetime import datetime

            hoje = datetime.today()
            num_dias = monthrange(hoje.year, hoje.month)[1]


            axis_x = QBarCategoryAxis()
            axis_x.append([f"Dia {i+1}" for i in range(num_dias)])
            self.chart2.createDefaultAxes()
            self.chart2.setAxisX(axis_x, series)

            # Remover hover com segurança
            try:
                series.hovered.disconnect()
            except:
                pass

        else:
            series.append("Sem dados", 1)
            self.chart2.addSeries(series)

    def mostrar_menu_cliente(self):
        """Mostra menu de ações para um cliente"""
        try:
            btn = self.sender()
            cliente_id = btn.property("id")

            if cliente_id is None:
                raise Exception("ID do cliente não encontrado no botão.")

            # Obtém informações adicionais do cliente
            self.cursor.execute("SELECT nome, empresa, observacoes, responsavel FROM clientes WHERE id = ?", (cliente_id,))
            resultado = self.cursor.fetchone()

            if not resultado:
                raise Exception(f"Cliente com ID {cliente_id} não encontrado no banco de dados.")

            nome, empresa, observacoes, responsavel = resultado

            menu = QDialog(self)
            menu.setWindowTitle(f"Ações do Cliente - {nome}")
            menu.setFixedSize(350, 350)
            menu.setStyleSheet("""
                QDialog {
                    background: #2c3e50;
                    color: white;
                }
                QPushButton {
                    text-align: left;
                    padding: 8px;
                    background: transparent;
                    border: none;
                    color: white;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background: #34495e;
                }
                QLabel {
                    color: white;
                    padding: 5px;
                    font-size: 13px;
                }
            """)
    
            layout = QVBoxLayout(menu)
    
            # Informações do cliente
            info_label = QLabel(f"<b>Cliente:</b> {nome}")
            layout.addWidget(info_label)
            
            if empresa:
                empresa_label = QLabel(f"<b>Empresa:</b> {empresa}")
                layout.addWidget(empresa_label)
    
            if responsavel:
                resp_label = QLabel(f"<b>Responsável:</b> {responsavel}")
                layout.addWidget(resp_label)
    
            if observacoes:
                obs_label = QLabel(f"<b>Observações:</b> {observacoes}")
                obs_label.setWordWrap(True)
                layout.addWidget(obs_label)
    
            # Separador
            separator = QFrame()
            separator.setFrameShape(QFrame.HLine)
            separator.setFrameShadow(QFrame.Sunken)
            layout.addWidget(separator)
    
            # Botões de ação
            btn_editar = QPushButton("✏️ Editar Cliente")
            btn_editar.clicked.connect(lambda: (menu.close(), self.abrir_cadastro_cliente(cliente_id)))
            layout.addWidget(btn_editar)
    
            btn_status = QPushButton("🔄 Alterar Status")
            btn_status.clicked.connect(lambda: (menu.close(), self.alterar_status_cliente(cliente_id)))
            layout.addWidget(btn_status)
    
            btn_agendar = QPushButton("📅 Agendar Reunião")
            btn_agendar.clicked.connect(lambda: (menu.close(), self.agendar_reuniao(cliente_id)))
            layout.addWidget(btn_agendar)
    
            btn_interacoes = QPushButton("📝 Registrar Interação")
            btn_interacoes.clicked.connect(lambda: (menu.close(), self.registrar_interacao(cliente_id)))
            layout.addWidget(btn_interacoes)
    
            btn_excluir = QPushButton("🗑️ Excluir Cliente")
            btn_excluir.setStyleSheet("color: #e74c3c;")
            btn_excluir.clicked.connect(lambda: (menu.close(), self.excluir_cliente(cliente_id)))
            layout.addWidget(btn_excluir)
    
            btn_fechar = QPushButton("Fechar")
            btn_fechar.clicked.connect(menu.close)
            layout.addWidget(btn_fechar)
    
            menu.exec_()
    
        except Exception as e:
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Erro", f"Ocorreu um erro ao abrir o menu:\n{str(e)}")

    def registrar_interacao(self, cliente_id):
        """Registra uma nova interação com o cliente"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Registrar Interação")
        dialog.setMinimumWidth(400)
        dialog.setStyleSheet("""
            QDialog {
                background: #2c3e50;
                color: white;
            }
            QLineEdit, QComboBox, QTextEdit {
                background: #34495e;
                color: white;
                border: 1px solid #2c3e50;
                padding: 5px;
            }
        """)
        
        form = QFormLayout(dialog)
        
        tipo_combo = QComboBox()
        tipo_combo.addItems(["Ligação", "E-mail", "Reunião", "Proposta", "Outro"])
        
        descricao_input = QTextEdit()
        descricao_input.setMaximumHeight(100)
        
        duracao_spin = QSpinBox()
        duracao_spin.setRange(1, 240)
        duracao_spin.setSuffix(" minutos")
        duracao_spin.setValue(15)
        
        resultado_input = QLineEdit()
        
        form.addRow("Tipo *", tipo_combo)
        form.addRow("Descrição *", descricao_input)
        form.addRow("Duração", duracao_spin)
        form.addRow("Resultado", resultado_input)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(lambda: self.salvar_interacao(
            cliente_id,
            tipo_combo.currentText(),
            descricao_input.toPlainText().strip(),
            duracao_spin.value(),
            resultado_input.text().strip(),
            dialog
        ))
        buttons.rejected.connect(dialog.reject)
        form.addRow(buttons)
        
        dialog.exec_()

    def salvar_interacao(self, cliente_id, tipo, descricao, duracao, resultado, dialog):
        """Salva uma nova interação no banco de dados"""
        if not descricao:
            QMessageBox.warning(self, "Aviso", "A descrição é obrigatória!")
            return
            
        try:
            self.cursor.execute("""
                INSERT INTO interacoes (
                    cliente_id, tipo, descricao, data, duracao, resultado, usuario_id
                ) VALUES (?, ?, ?, datetime('now'), ?, ?, ?)
            """, (
                cliente_id,
                tipo,
                descricao,
                duracao,
                resultado,
                self.email_usuario
            ))
            
            # Atualiza a data da última interação no cliente
            self.cursor.execute("""
                UPDATE clientes SET
                    ultima_atualizacao = datetime('now')
                WHERE id = ?
            """, (cliente_id,))
            
            self.conn.commit()
            dialog.accept()
            QMessageBox.information(self, "Sucesso", "Interação registrada com sucesso!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao registrar interação:\n{str(e)}")

    def mostrar_menu_agendamento(self):
        """Mostra menu de ações para um agendamento"""
        btn = self.sender()
        agendamento_id = btn.property("id")
        
        # Obtém informações adicionais do agendamento
        self.cursor.execute("""
            SELECT a.titulo, a.data_inicio, a.status, c.nome 
            FROM agendamentos a
            LEFT JOIN clientes c ON a.cliente_id = c.id
            WHERE a.id = ?
        """, (agendamento_id,))
        resultado = self.cursor.fetchone()
        
        if not resultado:
            QMessageBox.warning(self, "Aviso", "Agendamento não encontrado!")
            return
            
        titulo, data, status, cliente = resultado
        
        menu = QDialog(self)
        menu.setWindowTitle(f"Ações do Agendamento - {titulo}")
        menu.setFixedSize(300, 200)
        menu.setStyleSheet("""
            QDialog {
                background: #2c3e50;
                color: white;
            }
            QPushButton {
                text-align: left;
                padding: 8px;
                background: transparent;
                border: none;
                color: white;
            }
            QPushButton:hover {
                background: #34495e;
            }
            QLabel {
                color: white;
                padding: 5px;
            }
        """)
        
        layout = QVBoxLayout(menu)
        
        # Informações do agendamento
        info_label = QLabel(f"<b>Data:</b> {data[:10]}")
        layout.addWidget(info_label)
        
        if cliente:
            cliente_label = QLabel(f"<b>Cliente:</b> {cliente}")
            layout.addWidget(cliente_label)
        
        status_label = QLabel(f"<b>Status:</b> {status}")
        layout.addWidget(status_label)
        
        # Separador
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        layout.addWidget(separator)
        
        # Botões de ação
        btn_confirmar = QPushButton("✅ Confirmar Agendamento")
        btn_confirmar.clicked.connect(lambda: self.alterar_status_agendamento(agendamento_id, "Confirmado", menu))
        btn_confirmar.setVisible(status != "Confirmado")
        
        btn_cancelar = QPushButton("❌ Cancelar Agendamento")
        btn_cancelar.clicked.connect(lambda: self.alterar_status_agendamento(agendamento_id, "Cancelado", menu))
        btn_cancelar.setVisible(status != "Cancelado")
        
        btn_excluir = QPushButton("🗑️ Excluir Agendamento")
        btn_excluir.setStyleSheet("color: #e74c3c;")
        btn_excluir.clicked.connect(lambda: self.excluir_agendamento(agendamento_id, menu))
        
        layout.addWidget(btn_confirmar)
        layout.addWidget(btn_cancelar)
        layout.addWidget(btn_excluir)
        
        btn_fechar = QPushButton("Fechar")
        btn_fechar.clicked.connect(menu.close)
        layout.addWidget(btn_fechar)
        
        menu.exec_()

    def alterar_status_agendamento(self, agendamento_id, novo_status, dialog):
        """Altera o status de um agendamento"""
        try:
            self.cursor.execute("""
                UPDATE agendamentos SET
                    status = ?
                WHERE id = ?
            """, (novo_status, agendamento_id))
            
            self.conn.commit()
            self.carregar_agendamentos()
            dialog.accept()
            QMessageBox.information(self, "Sucesso", f"Agendamento marcado como '{novo_status}'!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao alterar status:\n{str(e)}")

    def excluir_agendamento(self, agendamento_id, dialog=None):
        """Exclui um agendamento do banco de dados"""
        confirm = QMessageBox.question(
            self, "Confirmar Exclusão",
            "Tem certeza que deseja excluir este agendamento?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if confirm == QMessageBox.Yes:
            try:
                query = "DELETE FROM agendamentos WHERE id = ?"
                params = (agendamento_id,)
                
                if not self.is_admin:
                    query += " AND usuario_id = ?"
                    params = (agendamento_id, self.email_usuario)
                
                self.cursor.execute(query, params)
                self.conn.commit()
                self.carregar_agendamentos()
                
                if dialog:
                    dialog.accept()
                    
                QMessageBox.information(self, "Sucesso", "Agendamento excluído com sucesso!")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao excluir agendamento:\n{str(e)}")

    def abrir_cadastro_cliente(self, cliente_id=None):
        """Abre o formulário de cadastro/edição de cliente"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Novo Cliente" if not cliente_id else "Editar Cliente")
        dialog.setMinimumWidth(500)
        dialog.setStyleSheet("""
            QDialog {
                background: #2c3e50;
                color: white;
            }
            QLabel {
                color: white;
            }
            QLineEdit, QComboBox, QDateEdit {
                background: #34495e;
                color: white;
                border: 1px solid #2c3e50;
                padding: 5px;
            }
        """)
    
        form = QFormLayout(dialog)
        
        # Se estiver editando, carrega os dados do cliente
        if cliente_id:
            self.cursor.execute("SELECT * FROM clientes WHERE id = ?", (cliente_id,))
            cliente = self.cursor.fetchone()
        
        # Campos do formulário
        nome_input = QLineEdit()
        empresa_input = QLineEdit()
        telefone_input = QLineEdit()
        email_input = QLineEdit()
        
        status_combo = QComboBox()
        status_combo.addItems(["Cliente", "Em atendimento", "Proposta", "Finalizado"])
        
        valor_input = QLineEdit()
        valor_input.setValidator(QDoubleValidator(0, 9999999, 2, self))
        
        origem_combo = QComboBox()
        origem_combo.addItems(["Site", "Indicação", "Redes Sociais", "Outros"])
        
        responsavel_input = QLineEdit()
        observacoes_input = QLineEdit()
        
        # Novo campo para tipo (só aparece na edição)
        tipo_combo = None
        if cliente_id:
            tipo_combo = QComboBox()
            tipo_combo.addItems(["Precode", "Allpost"])
        
        # Preenche os campos se estiver editando
        if cliente_id and cliente:
            nome_input.setText(cliente[1])
            empresa_input.setText(cliente[2] if cliente[2] else "")
            telefone_input.setText(cliente[3] if cliente[3] else "")
            email_input.setText(cliente[4] if cliente[4] else "")
            status_combo.setCurrentText(cliente[5])
            valor_input.setText(str(cliente[6]) if cliente[6] else "0")
            origem_combo.setCurrentText(cliente[10] if cliente[10] else "Site")
            responsavel_input.setText(cliente[11] if cliente[11] else "")
            observacoes_input.setText(cliente[14] if cliente[14] else "")
            if tipo_combo:
                tipo_combo.setCurrentText(cliente[16] if len(cliente) > 16 and cliente[16] else "Precode")
        
        form.addRow("Nome *", nome_input)
        form.addRow("Empresa", empresa_input)
        form.addRow("Telefone", telefone_input)
        form.addRow("Email", email_input)
        form.addRow("Status *", status_combo)
        form.addRow("Valor Proposta", valor_input)
        form.addRow("Origem", origem_combo)
        form.addRow("Responsável", responsavel_input)
        form.addRow("Observações", observacoes_input)
        
        # Adiciona o tipo se estiver editando
        if tipo_combo:
            form.addRow("Tipo", tipo_combo)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        if cliente_id:
            buttons.accepted.connect(lambda: self.atualizar_cliente(
                cliente_id,
                nome_input.text().strip(),
                empresa_input.text().strip(),
                telefone_input.text().strip(),
                email_input.text().strip(),
                status_combo.currentText(),
                valor_input.text(),
                origem_combo.currentText(),
                responsavel_input.text().strip(),
                observacoes_input.text().strip(),
                tipo_combo.currentText() if tipo_combo else "Precode",  # Novo parâmetro
                dialog
            ))
        else:
            buttons.accepted.connect(lambda: self.salvar_cliente(
                nome_input.text().strip(),
                empresa_input.text().strip(),
                telefone_input.text().strip(),
                email_input.text().strip(),
                status_combo.currentText(),
                valor_input.text(),
                origem_combo.currentText(),
                responsavel_input.text().strip(),
                observacoes_input.text().strip(),
                dialog
            ))
        buttons.rejected.connect(dialog.reject)
        form.addRow(buttons)
        
        dialog.exec_()

    def salvar_cliente(self, nome, empresa, telefone, email, status, valor, origem, responsavel, observacoes, dialog):
        if not nome:
            QMessageBox.warning(self, "Aviso", "O campo Nome é obrigatório!")
            return

        try:
            valor_float = float(valor.replace(",", ".")) if valor else 0
        
            # Obtém o tipo padrão das configurações
            settings = QSettings("GustavoMartins", "CRMApp")
            tipo_padrao = settings.value("tipo_padrao", "Precode")  # "Precode" é o valor padrão se não estiver configurado
            
            self.cursor.execute("""
                INSERT INTO clientes (
                    nome, empresa, telefone, email, status, valor_proposta,
                    data_cadastro, origem, responsavel, observacoes,
                    data_proxima_acao, ultima_atualizacao, usuario_id, tipo
                ) VALUES (?, ?, ?, ?, ?, ?, datetime('now'), ?, ?, ?, date('now', '+7 days'), datetime('now'), ?, ?)
            """, (
                nome,
                empresa,
                telefone,
                email,
                status,
                valor_float,
                origem,
                responsavel,
                observacoes,
                self.email_usuario,
                tipo_padrao  # Usa o tipo padrão das configurações
            ))
            
            self.conn.commit()
            self.carregar_clientes()
            self.atualizar_metricas()
            dialog.accept()
            QMessageBox.information(self, "Sucesso", "Cliente cadastrado com sucesso!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar cliente:\n{str(e)}")
        
    def atualizar_cliente(self, cliente_id, nome, empresa, telefone, email, status, valor, origem, responsavel, observacoes, tipo, dialog):
        """Atualiza os dados de um cliente existente"""
        try:
            valor_float = float(valor.replace(",", ".")) if valor else 0
        
            self.cursor.execute("""
                UPDATE clientes SET
                    nome = ?, empresa = ?, telefone = ?, email = ?, status = ?,
                    valor_proposta = ?, origem = ?, responsavel = ?, observacoes = ?,
                    tipo = ?, 
                    ultima_atualizacao = datetime('now')
                WHERE id = ?
            """, (
                nome, empresa, telefone, email, status,
                valor_float, origem, responsavel, observacoes,
                tipo,  # Novo valor
                cliente_id
            ))
            
            self.conn.commit()
            self.carregar_clientes()
            dialog.accept()
            QMessageBox.information(self, "Sucesso", "Cliente atualizado com sucesso!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao atualizar cliente:\n{str(e)}")

    def alterar_status_cliente(self, cliente_id):
        """Altera o status de um cliente"""
        # Obtém o status atual do cliente
        self.cursor.execute("SELECT status FROM clientes WHERE id = ?", (cliente_id,))
        status_atual = self.cursor.fetchone()[0]
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Alterar Status")
        dialog.setFixedSize(300, 150)
        dialog.setStyleSheet("""
            QDialog {
                background: #2c3e50;
                color: white;
            }
            QComboBox {
                background: #34495e;
                color: white;
                padding: 5px;
            }
        """)
        
        layout = QVBoxLayout(dialog)
        
        status_combo = QComboBox()
        status_combo.addItems(["Cliente", "Em atendimento", "Proposta", "Finalizado"])
        status_combo.setCurrentText(status_atual)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(lambda: self.salvar_status_cliente(cliente_id, status_combo.currentText(), dialog))
        buttons.rejected.connect(dialog.reject)
        
        layout.addWidget(QLabel("Novo Status:"))
        layout.addWidget(status_combo)
        layout.addWidget(buttons)
        
        dialog.exec_()

    def salvar_status_cliente(self, cliente_id, novo_status, dialog):
        """Salva o novo status do cliente"""
        try:
            self.cursor.execute("""
                UPDATE clientes SET
                    status = ?,
                    ultima_atualizacao = datetime('now')
                WHERE id = ?
            """, (novo_status, cliente_id))
            
            self.conn.commit()
            self.carregar_clientes()
            self.atualizar_metricas()
            dialog.accept()
            QMessageBox.information(self, "Sucesso", "Status atualizado com sucesso!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao atualizar status:\n{str(e)}")

    def agendar_reuniao(self, cliente_id):
        """Abre o formulário para agendar uma reunião"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Agendar Reunião")
        dialog.setMinimumWidth(500)
        dialog.setStyleSheet("""
            QDialog {
                background: #2c3e50;
                color: white;
            }
            QLineEdit, QComboBox, QDateEdit, QTimeEdit {
                background: #34495e;
                color: white;
                border: 1px solid #2c3e50;
                padding: 5px;
            }
        """)
        
        form = QFormLayout(dialog)
        
        self.cursor.execute("SELECT nome, email FROM clientes WHERE id = ?", (cliente_id,))
        cliente_nome, cliente_email = self.cursor.fetchone()
        
        cliente_label = QLabel(cliente_nome)
        titulo_input = QLineEdit()
        tipo_combo = QComboBox()
        tipo_combo.addItems(["Reunião", "Call", "Apresentação", "Visita"])
        data_input = QDateEdit()
        data_input.setDate(QDate.currentDate())
        data_input.setCalendarPopup(True)
        hora_input = QTimeEdit()
        hora_input.setTime(QTime(10, 0))
        duracao_combo = QComboBox()
        duracao_combo.addItems(["30 minutos", "1 hora", "2 horas"])
        descricao_input = QLineEdit()
        
        form.addRow("Cliente", cliente_label)
        form.addRow("Título *", titulo_input)
        form.addRow("Tipo *", tipo_combo)
        form.addRow("Data *", data_input)
        form.addRow("Hora *", hora_input)
        form.addRow("Duração", duracao_combo)
        form.addRow("Descrição", descricao_input)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(lambda: self.salvar_agendamento(
            cliente_id,
            titulo_input.text().strip(),
            tipo_combo.currentText(),
            data_input.date(),
            hora_input.time(),
            duracao_combo.currentText(),
            descricao_input.text().strip(),
            dialog
        ))
        buttons.rejected.connect(dialog.reject)
        form.addRow(buttons)
        
        dialog.exec_()

    def salvar_agendamento(self, cliente_id, titulo, tipo, data, hora, duracao, descricao, dialog):
        """Salva um novo agendamento e envia para o Google Agenda"""
        if not titulo:
            QMessageBox.warning(self, "Aviso", "O campo Título é obrigatório!")
            return

        try:
            from google.oauth2.credentials import Credentials
            from googleapiclient.discovery import build
            from google.auth.transport.requests import Request
            import os.path
            import smtplib
            from email.mime.text import MIMEText
            from email.mime.multipart import MIMEMultipart

            # Datas formatadas
            data_inicio_str = f"{data.toString('yyyy-MM-dd')}T{hora.toString('HH:mm')}:00"
            data_inicio_obj = datetime.strptime(data_inicio_str, "%Y-%m-%dT%H:%M:%S")

            # Calcular duração
            horas = 1
            if "30" in duracao:
                horas = 0.5
            elif "2" in duracao:
                horas = 2

            data_fim_obj = data_inicio_obj + timedelta(hours=horas)
            data_fim_str = data_fim_obj.strftime("%Y-%m-%dT%H:%M:%S")

            # --- Grava no banco de dados local ---
            self.cursor.execute("""
                INSERT INTO agendamentos (
                    titulo, tipo, data_inicio, data_fim, cliente_id, descricao, status
                ) VALUES (?, ?, ?, ?, ?, ?, 'Confirmado')
            """, (
                titulo,
                tipo,
                data_inicio_str,
                data_fim_str,
                cliente_id,
                descricao
            ))
            self.conn.commit()
            self.carregar_agendamentos()

            # --- Envia para Google Agenda e cria Meet ---
            SCOPES = ['https://www.googleapis.com/auth/calendar']
            creds = None

            if os.path.exists('token.json'):
                creds = Credentials.from_authorized_user_file('token.json', SCOPES)

            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())

            if creds:
                service = build('calendar', 'v3', credentials=creds)

                self.cursor.execute("SELECT nome, email FROM clientes WHERE id = ?", (cliente_id,))
                cliente_nome, cliente_email = self.cursor.fetchone()

                evento = {
                    'summary': titulo,
                    'description': descricao,
                    'start': {
                        'dateTime': data_inicio_str,
                        'timeZone': 'America/Sao_Paulo',
                    },
                    'end': {
                        'dateTime': data_fim_str,
                        'timeZone': 'America/Sao_Paulo',
                    },
                    'attendees': [
                        {'email': cliente_email} if cliente_email else None
                    ],
                    'conferenceData': {
                        'createRequest': {
                            'requestId': f"meet-{cliente_id}-{int(datetime.now().timestamp())}",
                            'conferenceSolutionKey': {
                                'type': 'hangoutsMeet'
                            }
                        }
                    }
                }

                # Remove None dos attendees se não houver email
                evento['attendees'] = [a for a in evento['attendees'] if a is not None]

                evento_criado = service.events().insert(
                    calendarId='primary', 
                    body=evento, 
                    conferenceDataVersion=1
                ).execute()

                meet_link = evento_criado.get('hangoutLink', 'Link não disponível')

                # --- Envia e-mail com o link do Meet ---
                if cliente_email:
                    try:
                        # Configurações do e-mail (substitua com suas configurações)
                        smtp_server = "smtp.gmail.com"
                        smtp_port = 587
                        smtp_user = "seuemail@gmail.com"
                        smtp_password = "suasenha"

                        msg = MIMEMultipart()
                        msg['From'] = smtp_user
                        msg['To'] = cliente_email
                        msg['Subject'] = f"Confirmação de Agendamento: {titulo}"

                        body = f"""
                        <h2>Confirmação de Agendamento</h2>
                        <p>Olá {cliente_nome},</p>
                        <p>Seu agendamento foi confirmado com sucesso:</p>
                        <p><strong>Título:</strong> {titulo}</p>
                        <p><strong>Data:</strong> {data_inicio_obj.strftime('%d/%m/%Y')}</p>
                        <p><strong>Hora:</strong> {data_inicio_obj.strftime('%H:%M')}</p>
                        <p><strong>Duração:</strong> {duracao}</p>
                        <p><strong>Link da Reunião:</strong> <a href="{meet_link}">{meet_link}</a></p>
                        <p>Atenciosamente,<br>Equipe de Atendimento</p>
                        """

                        msg.attach(MIMEText(body, 'html'))

                        with smtplib.SMTP(smtp_server, smtp_port) as server:
                            server.starttls()
                            server.login(smtp_user, smtp_password)
                            server.send_message(msg)

                    except Exception as email_error:
                        QMessageBox.warning(self, "Aviso", 
                            f"Agendamento criado, mas e-mail não enviado:\n{str(email_error)}")

            dialog.accept()
            QMessageBox.information(self, "Sucesso", 
                "Agendamento salvo e enviado para o Google Agenda!\n"
                f"Link da reunião: {meet_link}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar agendamento:\n{str(e)}")

    def excluir_cliente(self, cliente_id):
        """Exclui um cliente do banco de dados"""
        confirm = QMessageBox.question(
            self, "Confirmar Exclusão",
            "Tem certeza que deseja excluir este cliente?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if confirm == QMessageBox.Yes:
            try:
                # Primeiro exclui agendamentos relacionados
                self.cursor.execute("DELETE FROM agendamentos WHERE cliente_id = ?", (cliente_id,))
                # Depois exclui o cliente
                self.cursor.execute("DELETE FROM clientes WHERE id = ?", (cliente_id,))
                self.conn.commit()
                self.carregar_clientes()
                QMessageBox.information(self, "Sucesso", "Cliente excluído com sucesso!")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao excluir cliente:\n{str(e)}")

    def exportar_relatorio(self):
        """Exporta um relatório completo para Excel"""
        try:
            # Obter dados para o relatório
            self.cursor.execute("""
                SELECT 
                    c.nome, c.empresa, c.telefone, c.email, c.status,
                    c.valor_proposta, c.data_cadastro, c.origem, c.responsavel,
                    COUNT(i.id) as interacoes,
                    (SELECT COUNT(*) FROM agendamentos a WHERE a.cliente_id = c.id) as agendamentos
                FROM clientes c
                LEFT JOIN interacoes i ON i.cliente_id = c.id
                GROUP BY c.id
                ORDER BY c.data_cadastro DESC
            """)
            dados = self.cursor.fetchall()
            
            # Criar DataFrame
            from pandas import DataFrame
            df = DataFrame(dados, columns=[
                "Nome", "Empresa", "Telefone", "Email", "Status",
                "Valor Proposta", "Data Cadastro", "Origem", "Responsável",
                "Total Interações", "Total Agendamentos"
            ])
            
            # Formata os valores monetários
            df['Valor Proposta'] = df['Valor Proposta'].apply(
                lambda x: f"R${x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if x else "-"
            )
            
            # Salvar arquivo Excel
            file_name, _ = QFileDialog.getSaveFileName(
                self, "Exportar Relatório", "", "Excel Files (*.xlsx)"
            )
            
            if file_name:
                df.to_excel(file_name, index=False)
                QMessageBox.information(self, "Sucesso", f"Relatório exportado para:\n{file_name}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao exportar relatório:\n{str(e)}")

    def verificar_agendamentos_proximos(self):
        """Verifica se há agendamentos próximos e mostra notificações (apenas uma vez)"""
        try:
            minutos_notificacao = self.notification_minutes
            agora = datetime.now()
            limite = agora + timedelta(minutes=minutos_notificacao)
            
            agora_str = agora.strftime("%Y-%m-%dT%H:%M:%S")
            limite_str = limite.strftime("%Y-%m-%dT%H:%M:%S")
            
            self.cursor.execute("""
                SELECT a.id, a.titulo, a.data_inicio, c.nome 
                FROM agendamentos a
                LEFT JOIN clientes c ON a.cliente_id = c.id
                WHERE a.data_inicio BETWEEN ? AND ?
                AND a.status = 'Confirmado'
            """, (agora_str, limite_str))
            
            for agendamento_id, titulo, data, cliente in self.cursor.fetchall():
                # Verifica se já notificou este agendamento
                if agendamento_id not in self.agendamentos_notificados:
                    try:
                        data_obj = datetime.strptime(data, "%Y-%m-%dT%H:%M:%S")
                        minutos_restantes = int((data_obj - agora).total_seconds() / 60)
                        
                        if minutos_restantes > 0:
                            msg = QMessageBox(self)
                            msg.setWindowTitle("Lembrete de Agendamento")
                            msg.setIcon(QMessageBox.Information)
                            msg.setText(f"Você tem um agendamento em {minutos_restantes} minutos:")
                            msg.setInformativeText(f"<b>{titulo}</b><br>Com: {cliente if cliente else 'Sem cliente'}<br>Horário: {data_obj.strftime('%H:%M')}")
                            msg.setStandardButtons(QMessageBox.Ok)
                            msg.exec_()
                            
                            # Marca como notificado
                            self.agendamentos_notificados.add(agendamento_id)
                            
                    except Exception as e:
                        print(f"Erro ao mostrar notificação: {str(e)}")
                    
            # Limpa agendamentos antigos (opcional)
            self.agendamentos_notificados = {
                id_ for id_ in self.agendamentos_notificados 
                if id_ in [row[0] for row in self.cursor.execute("SELECT id FROM agendamentos").fetchall()]
            }
        
        except Exception as e:
            print(f"Erro ao verificar agendamentos: {str(e)}")

            # Reset diário (opcional)
            if not hasattr(self, 'ultimo_dia_verificado'):
                self.ultimo_dia_verificado = agora.day

            if agora.day != self.ultimo_dia_verificado:
                self.agendamentos_notificados = set()
                self.ultimo_dia_verificado = agora.day

    def show_configuracoes(self):
        """Mostra a janela de configurações"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Configurações")
        dialog.setFixedSize(400, 350)  # Aumentei a altura
        dialog.setStyleSheet("""
            QDialog {
                background: #2c3e50;
                color: white;
            }
            QComboBox {
                background: #34495e;
                color: white;
                padding: 5px;
            }
        """)
        
        layout = QVBoxLayout(dialog)
        
        # Configuração de notificação
        notif_label = QLabel("Notificar agendamentos com antecedência:")
        notif_combo = QComboBox()
        notif_combo.addItems(["10 minutos", "15 minutos", "20 minutos", "30 minutos"])
        notif_combo.setCurrentText(f"{self.notification_minutes} minutos")
        
        # Configuração do tipo padrão
        tipo_label = QLabel("Tipo padrão para novos clientes:")
        tipo_combo = QComboBox()
        tipo_combo.addItems(["Precode", "Allpost"])
        
        # Carrega o valor salvo
        settings = QSettings("GustavoMartins", "CRMApp")
        tipo_padrao = settings.value("tipo_padrao", "Precode")
        tipo_combo.setCurrentText(tipo_padrao)
        
        # Botões
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(lambda: self.salvar_configuracoes(
            int(notif_combo.currentText().split()[0]),
            tipo_combo.currentText(),
            dialog
        ))
        buttons.rejected.connect(dialog.reject)
        
        layout.addWidget(notif_label)
        layout.addWidget(notif_combo)
        layout.addWidget(tipo_label)
        layout.addWidget(tipo_combo)
        layout.addWidget(buttons)
        
        dialog.exec_()

    def salvar_configuracoes(self, minutos, tipo_padrao, dialog):
        """Salva as configurações do usuário"""
        self.notification_minutes = minutos
    
        # Salva o tipo padrão nas configurações
        settings = QSettings("GustavoMartins", "CRMApp")
        settings.setValue("tipo_padrao", tipo_padrao)
        
        dialog.accept()
        QMessageBox.information(self, "Sucesso", "Configurações salvas com sucesso!")

    def get_count(self, table, condition="", params=None):
        """Retorna a contagem de registros em uma tabela de forma segura"""
        query = f"SELECT COUNT(*) FROM {table}"
        if condition and condition.strip():
            # Remove WHERE duplicado se existir
            clean_condition = condition.strip()
            if clean_condition.upper().startswith("AND"):
                clean_condition = clean_condition[3:].strip()
            if not clean_condition.upper().startswith("WHERE"):
                query += " WHERE "
            query += clean_condition
        
        try:
            if params:
                # Converte para tupla se não for
                params_tuple = tuple(params) if isinstance(params, (list, tuple)) else (params,)
                self.cursor.execute(query, params_tuple)
            else:
                self.cursor.execute(query)
                
            return self.cursor.fetchone()[0] or 0
        except sqlite3.Error as e:
            error_msg = f"""
            ERRO NA QUERY:
            Query: {query}
            Params: {params}
            Erro: {str(e)}
            """
            print(error_msg)
            return 0
    
    def get_sum(self, table, column, condition="", params=None):
        """Retorna a soma de uma coluna com tratamento seguro"""
        query = f"SELECT SUM({column}) FROM {table}"
        if condition and condition.strip():
            # Remove WHERE/AND duplicados
            clean_condition = condition.strip()
            if clean_condition.upper().startswith("AND"):
                clean_condition = clean_condition[3:].strip()
            if not clean_condition.upper().startswith("WHERE"):
                query += " WHERE "
            query += clean_condition
        
        try:
            if params:
                params_tuple = tuple(params) if isinstance(params, (list, tuple)) else (params,)
                self.cursor.execute(query, params_tuple)
            else:
                self.cursor.execute(query)
                
            result = self.cursor.fetchone()[0]
            return float(result) if result else 0.0
        except sqlite3.Error as e:
            error_msg = f"""
            ERRO NA QUERY SOMA:
            Query: {query}
            Params: {params}
            Erro: {str(e)}
            """
            print(error_msg)
            return 0.0

    def show_dashboard(self):
        """Mostra a página de dashboard"""
        if hasattr(self, 'content_stack'):
            self.content_stack.setCurrentWidget(self.dashboard_page)
            self.atualizar_metricas()

    def show_clientes(self):
        """Mostra a página de clientes"""
        self.content_stack.setCurrentWidget(self.clientes_page)
        self.carregar_clientes()

    def show_agenda(self):
        """Mostra a página de agenda"""
        self.content_stack.setCurrentWidget(self.agenda_page)
        self.carregar_agendamentos()

    def show_relatorios(self):
        """Mostra a página de relatórios"""
        self.content_stack.setCurrentWidget(self.relatorios_page)
        self.atualizar_metricas()

    def closeEvent(self, event):
        """Fecha a conexão com o banco de dados ao sair"""
        self.conn.close()
        event.accept()

class MetricCard(QFrame):
    """Widget personalizado para cards de métricas com interação"""
    def __init__(self, value, title, color, parent=None):
        super().__init__(parent)
        self.value = value
        self.title = title
        self.color = color
        
        self.setStyleSheet(f"""
            QFrame {{
                background: {color};
                border-radius: 8px;
                padding: 15px;
            }}
            QLabel {{
                color: white;
            }}
        """)
        
        self.setMinimumHeight(120)
        self.setMinimumWidth(180)
        
        self.layout = QVBoxLayout(self)
        self.layout.setAlignment(Qt.AlignCenter)
        
        self.value_label = QLabel(value)
        self.value_label.setStyleSheet("font-size: 28px;")
        self.value_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.value_label)
        
        self.title_label = QLabel(title)
        self.title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        self.title_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.title_label)
    
    def update_value(self, new_value):
        """Atualiza o valor exibido no card"""
        self.value = new_value
        self.value_label.setText(new_value)
    
    def update_status(self, new_status):
        """Atualiza o status exibido no card"""
        self.title = new_status
        self.title_label.setText(new_status)

if __name__ == "__main__":
    try:
        # Configura para mostrar mensagens detalhadas de erro no Windows
        if sys.platform == "win32":
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("GustavoMartins.CRMApp.1.0")
        
        app = QApplication(sys.argv)
        app.setStyle('Fusion')  # Garante um estilo consistente em todas as plataformas
        
        # Configuração para lembrar o último usuário
        settings = QSettings("GustavoMartins", "CRMApp")
        saved_email = settings.value("saved_email", "")
        remember_me = settings.value("remember_me", False, type=bool)
        
        # Verifica se há atualizações pendentes (opcional)
        try:
            from atualizacoes import verificar_atualizacoes
            verificar_atualizacoes()
        except:
            pass
        
        login_dialog = LoginDialog()
        
        if login_dialog.exec_() == QDialog.Accepted:
            email, senha = login_dialog.get_credentials()
            
            # Verifica se o banco existe com tratamento robusto
            db_path = f"crm_database_{email}.db"
            try:
                if not os.path.exists(db_path):
                    QMessageBox.critical(None, "Erro", "Este e-mail ainda não está cadastrado!")
                    sys.exit(1)
                
                # Verifica se o banco é acessível
                try:
                    teste_conn = sqlite3.connect(db_path)
                    teste_conn.close()
                except sqlite3.Error as e:
                    QMessageBox.critical(None, "Erro", f"Banco de dados corrompido:\n{str(e)}")
                    sys.exit(1)
                
                # Inicia o app com tratamento de erros
                try:
                    window = CRMApp(email)
                    window.show()
                    
                    # Configura para fechar corretamente ao clicar no X
                    app.aboutToQuit.connect(window.closeEvent)
                    
                    sys.exit(app.exec_())
                    
                except Exception as app_error:
                    error_msg = (
                        f"Erro crítico na aplicação:\n{str(app_error)}\n\n"
                        "Por favor, reinicie o aplicativo. Se o problema persistir,\n"
                        "entre em contato com o suporte técnico."
                    )
                    QMessageBox.critical(None, "Erro Fatal", error_msg)
                    sys.exit(1)
                    
            except Exception as db_error:
                QMessageBox.critical(None, "Erro", f"Falha ao acessar banco de dados:\n{str(db_error)}")
                sys.exit(1)
    
    except Exception as main_error:
        # Captura qualquer erro não tratado no nível mais alto
        error_msg = (
            f"Erro inesperado:\n{str(main_error)}\n\n"
            "O aplicativo será encerrado. Por favor, tente novamente."
        )
        QMessageBox.critical(None, "Erro Inesperado", error_msg)
        sys.exit(1)
