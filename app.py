from flask import Flask, render_template, request, jsonify, session, redirect, url_for, flash, send_file
import datetime
import logging
import mysql.connector
from mysql.connector import Error
import os
import csv
import io
import json
from werkzeug.security import generate_password_hash, check_password_hash
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from decimal import Decimal

# =============================================================================
# CONFIGURAÇÕES PARA PYTHONANYWHERE
# =============================================================================
import sys

# Configuração para PythonAnywhere
if 'PYTHONANYWHERE_DOMAIN' in os.environ:
    # Desativa debug no ambiente de produção
    DEBUG = False
    # Configurações específicas para PythonAnywhere
    MYSQL_HOST = 'enzzodril.mysql.pythonanywhere-services.com'
    MYSQL_DATABASE = 'enzzodril$default'
    MYSQL_USER = 'enzzodril'
    MYSQL_PASSWORD = '123formato'
else:
    # Configurações para desenvolvimento local
    DEBUG = True
    MYSQL_HOST = 'localhost'
    MYSQL_DATABASE = 'sig_me_db'
    MYSQL_USER = 'root'
    MYSQL_PASSWORD = '123456'

# CONFIGURAR LOGGING
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
app = Flask(__name__)
app.secret_key = 'sig-me-chave-secreta-2024'

# Configuração do app para PythonAnywhere
if 'PYTHONANYWHERE_DOMAIN' in os.environ:
    app.config.update(
        DEBUG=False,
        SESSION_COOKIE_SECURE=True,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE='Lax',
        PERMANENT_SESSION_LIFETIME=3600
    )

# CORES DO SITE FORMATO OH
CORES_FORMATO = {
    'primaria': '#1a365d',      # Azul escuro
    'secundaria': '#2d3748',    # Cinza azulado
    'destaque': '#e53e3e',      # Vermelho
    'sucesso': '#38a169',       # Verde
    'alerta': '#dd6b20',        # Laranja
    'info': '#3182ce',          # Azul
    'claro': '#f7fafc',         # Cinza muito claro
    'branco': '#ffffff',
    'texto': '#2d3748',
    'texto_claro': '#718096'
}

# =============================================================================
# MÓDULO: DATABASE - CONEXÃO MYSQL REAL
# =============================================================================
class DatabaseConnection:
    def __init__(self):
        self.host = MYSQL_HOST
        self.database = MYSQL_DATABASE
        self.user = MYSQL_USER
        self.password = MYSQL_PASSWORD
        self.connection = None

    def connect(self, create_database=False):
        """Conecta ao MySQL real"""
        try:
            if create_database:
                self.connection = mysql.connector.connect(
                    host=self.host,
                    user=self.user,
                    password=self.password
                )
            else:
                self.connection = mysql.connector.connect(
                    host=self.host,
                    database=self.database,
                    user=self.user,
                    password=self.password
                )

            if self.connection.is_connected():
                logging.info("✅ Conectado ao MySQL com sucesso!")
                return True
        except Error as e:
            logging.error(f"❌ Erro de conexão MySQL: {e}")
            return False

    def execute_query(self, query, params=None):
        """Executa queries no MySQL real"""
        if not self.connect():
            return None

        try:
            cursor = self.connection.cursor(dictionary=True)
            cursor.execute(query, params or ())
            result = cursor.fetchall()
            cursor.close()

            # Converter objetos Decimal para float
            for row in result:
                for key, value in row.items():
                    if isinstance(value, Decimal):
                        row[key] = float(value)

            return result
        except Error as e:
            logging.error(f"❌ Erro na query: {e}")
            return None
        finally:
            if self.connection and self.connection.is_connected():
                self.connection.close()

    def execute_update(self, query, params=None):
        """Executa updates no MySQL real"""
        if not self.connect():
            return False

        try:
            cursor = self.connection.cursor()
            cursor.execute(query, params or ())
            self.connection.commit()
            cursor.close()
            logging.info(f"✅ Update executado: {query}")
            return True
        except Error as e:
            logging.error(f"❌ Erro no update: {e}")
            return False
        finally:
            if self.connection and self.connection.is_connected():
                self.connection.close()

# =============================================================================
# INICIALIZAR DATABASE E CRIAR TABELAS SE NÃO EXISTIREM
# =============================================================================
def inicializar_database():
    """Inicializa o banco de dados e cria tabelas se não existirem"""
    try:
        db = DatabaseConnection()

        # Conectar sem database específico para criar se não existir
        conn = mysql.connector.connect(
            host=db.host,
            user=db.user,
            password=db.password
        )
        cursor = conn.cursor()

        # Criar database se não existir
        cursor.execute(f"CREATE DATABASE IF NOT EXISTS {db.database}")
        cursor.execute(f"USE {db.database}")

        # === CRIAR TABELA DE CLIENTES ===
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS clientes (
                id INT AUTO_INCREMENT PRIMARY KEY,
                Nome_Fantasia VARCHAR(255) NOT NULL,
                Nome_Razao_Social VARCHAR(255),
                CNPJ_CPF VARCHAR(20),
                Telefone VARCHAR(20),
                Email VARCHAR(100),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Criar tabela de fornecedores se não existir
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS fornecedores (
                id INT AUTO_INCREMENT PRIMARY KEY,
                razao_social VARCHAR(255) NOT NULL,
                cnpj VARCHAR(20),
                telefone VARCHAR(20),
                email VARCHAR(100),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Criar tabela de contratos
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS contratos (
                id INT AUTO_INCREMENT PRIMARY KEY,
                tipo ENUM('cliente', 'fornecedor') NOT NULL,
                cliente_id INT,
                fornecedor_id INT,
                descricao TEXT NOT NULL,
                valor_mensal DECIMAL(10,2) NOT NULL,
                data_inicio DATE NOT NULL,
                data_fim DATE,
                status ENUM('ativo', 'inativo', 'vencido') DEFAULT 'ativo',
                observacoes TEXT,
                arquivo_pdf VARCHAR(255),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Criar tabela de placas (inventário)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS placas (
                id INT AUTO_INCREMENT PRIMARY KEY,
                Codigo_Ativo VARCHAR(100) NOT NULL,
                Endereco TEXT NOT NULL,
                Regiao VARCHAR(100) NOT NULL,
                Tipo_Placa VARCHAR(100) NOT NULL,
                Status_Atual ENUM('disponível', 'locado', 'reservado', 'manutenção') DEFAULT 'disponível',
                Cliente_Locacao VARCHAR(255),
                Valor_Mensal DECIMAL(10,2) DEFAULT 0,
                Data_Cadastro DATE,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Criar tabela de transações financeiras
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS transacoes_financeiras (
                id INT AUTO_INCREMENT PRIMARY KEY,
                tipo ENUM('receita', 'despesa') NOT NULL,
                descricao VARCHAR(255) NOT NULL,
                valor DECIMAL(10,2) NOT NULL,
                categoria VARCHAR(100) NOT NULL,
                data_transacao DATE NOT NULL,
                data_vencimento DATE,
                status ENUM('pendente', 'pago', 'atrasado') DEFAULT 'pendente',
                cliente_id INT,
                fornecedor_id INT,
                observacoes TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Inserir alguns clientes de exemplo se a tabela estiver vazia
        cursor.execute("SELECT COUNT(*) as count FROM clientes")
        if cursor.fetchone()[0] == 0:
            clientes_exemplo = [
                ('Concessionária XYZ', 'Concessionária XYZ Ltda', '12.345.678/0001-90', '(11) 3333-4444', 'contato@xyz.com.br'),
                ('Rede de Farmácias', 'Rede de Farmácias SA', '98.765.432/0001-10', '(11) 5555-6666', 'vendas@rededefarmacias.com.br'),
                ('Banco Nacional', 'Banco Nacional SA', '55.444.333/0001-20', '(11) 7777-8888', 'contratos@bancacional.com.br')
            ]

            for cliente in clientes_exemplo:
                cursor.execute(
                    "INSERT INTO clientes (Nome_Fantasia, Nome_Razao_Social, CNPJ_CPF, Telefone, Email) VALUES (%s, %s, %s, %s, %s)",
                    cliente
                )

        # Inserir alguns fornecedores de exemplo se a tabela estiver vazia
        cursor.execute("SELECT COUNT(*) as count FROM fornecedores")
        if cursor.fetchone()[0] == 0:
            fornecedores_exemplo = [
                ('Construtora Silva Ltda', '12.345.678/0001-90', '(11) 9999-8888', 'contato@silvaconstrucoes.com.br'),
                ('Iluminação Publica SP', '98.765.432/0001-10', '(11) 7777-6666', 'contato@iluminacaosp.com.br'),
                ('Manutenção Express', '55.444.333/0001-20', '(11) 5555-4444', 'servicos@manutencaoexpress.com.br')
            ]

            for fornecedor in fornecedores_exemplo:
                cursor.execute(
                    "INSERT INTO fornecedores (razao_social, cnpj, telefone, email) VALUES (%s, %s, %s, %s)",
                    fornecedor
                )

        # Inserir alguns contratos de exemplo se a tabela estiver vazia
        cursor.execute("SELECT COUNT(*) as count FROM contratos")
        if cursor.fetchone()[0] == 0:
            contratos_exemplo = [
                ('cliente', 1, None, 'Contrato de Locação - Concessionária XYZ', 3500.00, '2024-01-01', '2024-12-31', 'ativo', 'Contrato anual renovável', 'contrato_xyz.pdf'),
                ('cliente', 2, None, 'Locação Especial - Rede Farmácias', 3200.00, '2024-02-01', '2024-12-31', 'ativo', 'Campanha temporária', 'contrato_farmacias.pdf'),
                ('cliente', 3, None, 'Contrato Banco Nacional', 2800.00, '2024-03-01', '2024-12-31', 'ativo', 'Contrato corporativo', 'contrato_banco.pdf')
            ]

            for contrato in contratos_exemplo:
                cursor.execute(
                    "INSERT INTO contratos (tipo, cliente_id, fornecedor_id, descricao, valor_mensal, data_inicio, data_fim, status, observacoes, arquivo_pdf) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
                    contrato
                )

        # Inserir algumas placas de exemplo se a tabela estiver vazia
        cursor.execute("SELECT COUNT(*) as count FROM placas")
        if cursor.fetchone()[0] == 0:
            placas_exemplo = [
                # São Roque - 8 placas
                ('43', 'Rod Raposo Tavares km S8 - Semi. Centro de São Roque', 'São Roque', 'Outdoor', 'locado', 'Concessionária XYZ', 3500.00, '2024-01-15'),
                ('44-2', 'Rod Raposo Tavares km S8 - Entrada de São Roque', 'São Roque', 'Outdoor', 'locado', 'Rede de Farmácias', 3200.00, '2024-01-15'),
                ('141', 'Av. Principal - Centro', 'São Roque', 'Outdoor', 'locado', 'Banco Nacional', 2800.00, '2024-02-01'),
                ('142', 'Av. Getúlio Vargas, 500 - Centro', 'São Roque', 'Front Light', 'locado', 'Supermercado Preço Baixo', 2200.00, '2024-02-10'),
                ('143', 'Rua das Flores, 123 - Jardim Europa', 'São Roque', 'Back Light', 'reservado', 'Construtora Progresso', 1800.00, '2024-02-15'),
                ('144', 'Rodovia Prefeito Quintino de Lima', 'São Roque', 'Outdoor', 'disponível', None, 2500.00, '2024-03-01'),
                ('145', 'Praça da Matriz - Centro Histórico', 'São Roque', 'Led', 'manutenção', None, 0.00, '2024-03-05'),
                ('146', 'Av. João Pessoa, 789 - Vila Nova', 'São Roque', 'Front Light', 'disponível', None, 1900.00, '2024-03-10'),

                # Mairinque - 6 placas
                ('201', 'Rod. Bunjiro Nakao, km 45', 'Mairinque', 'Outdoor', 'locado', 'Posto Shell', 3100.00, '2024-01-20'),
                ('202', 'Av. São Paulo, 456 - Centro', 'Mairinque', 'Front Light', 'locado', 'Farmácia Popular', 2100.00, '2024-02-05'),
                ('203', 'Rua Amazonas, 321 - Vila Maria', 'Mairinque', 'Back Light', 'disponível', None, 1700.00, '2024-02-12'),
                ('204', 'Entrada da Cidade - BR-376', 'Mairinque', 'Outdoor', 'reservado', 'Concessionária ABC', 2900.00, '2024-02-20'),
                ('205', 'Praça Central - Centro', 'Mairinque', 'Led', 'disponível', None, 2300.00, '2024-03-01'),
                ('206', 'Av. Industrial, 1000 - Distrito Industrial', 'Mairinque', 'Front Light', 'locado', 'Indústria Metalúrgica', 2000.00, '2024-03-08'),

                # Ibuna - 5 placas
                ('301', 'Estrada Municipal Ibuna-São Roque', 'Ibuna', 'Outdoor', 'locado', 'Laticínios Serra', 2700.00, '2024-01-25'),
                ('302', 'Centro Comercial de Ibuna', 'Ibuna', 'Front Light', 'disponível', None, 1600.00, '2024-02-08'),
                ('303', 'Entrada do Bairro Rural', 'Ibuna', 'Back Light', 'disponível', None, 1400.00, '2024-02-14'),
                ('304', 'Praça da Igreja Matriz', 'Ibuna', 'Outdoor', 'locado', 'Mercado Central', 2400.00, '2024-02-25'),
                ('305', 'Rod. Vicinal Principal', 'Ibuna', 'Front Light', 'manutenção', None, 0.00, '2024-03-03'),

                # Araçá - 6 placas
                ('401', 'Rua Principal - Centro do Araçá', 'Araçá', 'Outdoor', 'locado', 'Auto Peças Araçá', 2600.00, '2024-01-30'),
                ('402', 'Av. Comercial, 200', 'Araçá', 'Front Light', 'disponível', None, 1800.00, '2024-02-12'),
                ('403', 'Praça do Mercado Municipal', 'Araçá', 'Back Light', 'reservado', 'Panificadora Doce Pão', 1500.00, '2024-02-18'),
                ('404', 'Entrada da Cidade - SP-250', 'Araçá', 'Outdoor', 'locado', 'Transportadora Rápida', 2800.00, '2024-02-28'),
                ('405', 'Rua da Escola Municipal', 'Araçá', 'Led', 'disponível', None, 2100.00, '2024-03-05'),
                ('406', 'Bairro Novo Horizonte', 'Araçá', 'Front Light', 'disponível', None, 1700.00, '2024-03-12'),

                # Piedade - 7 placas
                ('501', 'Rod. Padre Anchieta, km 78', 'Piedade', 'Outdoor', 'locado', 'Hotel Serra Verde', 3300.00, '2024-02-01'),
                ('502', 'Centro Histórico de Piedade', 'Piedade', 'Front Light', 'locado', 'Restaurante Sabor Caseiro', 2200.00, '2024-02-10'),
                ('503', 'Av. Beira Rio, 150', 'Piedade', 'Back Light', 'disponível', None, 1600.00, '2024-02-15'),
                ('504', 'Entrada do Parque Municipal', 'Piedade', 'Outdoor', 'reservado', 'Agência de Turismo', 3000.00, '2024-02-22'),
                ('505', 'Praça do Santuário', 'Piedade', 'Led', 'locado', 'Loja de Artigos Religiosos', 1900.00, '2024-03-01'),
                ('506', 'Rua Comercial, 345', 'Piedade', 'Front Light', 'disponível', None, 2000.00, '2024-03-08'),
                ('507', 'Bairro dos Pinheiros', 'Piedade', 'Back Light', 'manutenção', None, 0.00, '2024-03-15'),

                # Alumínio - 5 placas
                ('601', 'Rod. dos Metalúrgicos, km 12', 'Alumínio', 'Outdoor', 'locado', 'Indústria de Alumínio', 3200.00, '2024-02-05'),
                ('602', 'Centro de Alumínio', 'Alumínio', 'Front Light', 'disponível', None, 2100.00, '2024-02-12'),
                ('603', 'Av. Industrial, 500', 'Alumínio', 'Back Light', 'locado', 'Metalúrgica Forte', 2300.00, '2024-02-20'),
                ('604', 'Entrada do Complexo Industrial', 'Alumínio', 'Outdoor', 'reservado', 'Cooperativa de Trabalho', 2900.00, '2024-02-28'),
                ('605', 'Praça dos Trabalhadores', 'Alumínio', 'Led', 'disponível', None, 2400.00, '2024-03-06')
            ]

            for placa in placas_exemplo:
                cursor.execute(
                    "INSERT INTO placas (Codigo_Ativo, Endereco, Regiao, Tipo_Placa, Status_Atual, Cliente_Locacao, Valor_Mensal, Data_Cadastro) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)",
                    placa
                )

        # Inserir algumas transações financeiras de exemplo
        cursor.execute("SELECT COUNT(*) as count FROM transacoes_financeiras")
        if cursor.fetchone()[0] == 0:
            transacoes_exemplo = [
                # Receitas
                ('receita', 'Locação Placa 43 - Concessionária XYZ', 3500.00, 'Locação Placas', '2024-03-01', '2024-03-01', 'pago', 1, None, 'Pagamento recebido'),
                ('receita', 'Locação Placa 44-2 - Rede Farmácias', 3200.00, 'Locação Placas', '2024-03-05', '2024-03-05', 'pago', 2, None, 'Pagamento recebido'),
                ('receita', 'Locação Placa 141 - Banco Nacional', 2800.00, 'Locação Placas', '2024-03-10', '2024-03-10', 'pago', 3, None, 'Pagamento recebido'),
                ('receita', 'Locação Placa 142 - Supermercado', 2200.00, 'Locação Placas', '2024-03-15', '2024-03-15', 'pendente', None, None, 'Aguardando pagamento'),

                # Despesas
                ('despesa', 'Manutenção Placa 145', 450.00, 'Manutenção', '2024-03-02', '2024-03-02', 'pago', None, 3, 'Manutenção preventiva'),
                ('despesa', 'Aluguel Escritório', 1200.00, 'Despesas Operacionais', '2024-03-05', '2024-03-05', 'pago', None, None, 'Aluguel mensal'),
                ('despesa', 'Energia Elétrica', 380.00, 'Despesas Operacionais', '2024-03-08', '2024-03-15', 'pendente', None, None, 'Conta a vencer'),
                ('despesa', 'Material de Limpeza', 150.00, 'Despesas Operacionais', '2024-03-10', '2024-03-10', 'pago', None, None, 'Material de escritório')
            ]

            for transacao in transacoes_exemplo:
                cursor.execute(
                    "INSERT INTO transacoes_financeiras (tipo, descricao, valor, categoria, data_transacao, data_vencimento, status, cliente_id, fornecedor_id, observacoes) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
                    transacao
                )

        conn.commit()
        cursor.close()
        conn.close()

        logging.info("✅ Banco de dados inicializado com sucesso!")
        return True

    except Error as e:
        logging.error(f"❌ Erro ao inicializar banco de dados: {e}")
        return False

# =============================================================================
# MÓDULO: RELATÓRIOS
# =============================================================================
class RelatorioManager:
    def __init__(self, db_connection):
        self.db = db_connection

    def gerar_relatorio_ocupacao(self):
        """Gera relatório de ocupação detalhado"""
        try:
            query = """
                SELECT Regiao, Status_Atual, COUNT(*) as quantidade,
                       SUM(Valor_Mensal) as faturamento
                FROM placas
                GROUP BY Regiao, Status_Atual
                ORDER BY Regiao, Status_Atual
            """
            dados = self.db.execute_query(query) or []

            # Agrupar por região
            relatorio = {}
            for item in dados:
                regiao = item['Regiao']
                if regiao not in relatorio:
                    relatorio[regiao] = []
                relatorio[regiao].append(item)

            return relatorio

        except Exception as e:
            logging.error(f"❌ Erro ao gerar relatório ocupação: {e}")
            return {}

    def gerar_relatorio_financeiro(self):
        """Gera relatório financeiro detalhado"""
        try:
            # Faturamento por região
            faturamento_regiao = self.db.execute_query("""
                SELECT Regiao, SUM(Valor_Mensal) as faturamento
                FROM placas
                WHERE Status_Atual IN ('locado', 'reservado')
                GROUP BY Regiao
                ORDER BY faturamento DESC
            """) or []

            # Faturamento por tipo de placa
            faturamento_tipo = self.db.execute_query("""
                SELECT Tipo_Placa, SUM(Valor_Mensal) as faturamento
                FROM placas
                WHERE Status_Atual IN ('locado', 'reservado')
                GROUP BY Tipo_Placa
                ORDER BY faturamento DESC
            """) or []

            # Top clientes
            top_clientes = self.db.execute_query("""
                SELECT Cliente_Locacao, SUM(Valor_Mensal) as faturamento
                FROM placas
                WHERE Status_Atual IN ('locado', 'reservado') AND Cliente_Locacao IS NOT NULL
                GROUP BY Cliente_Locacao
                ORDER BY faturamento DESC
                LIMIT 10
            """) or []

            return {
                'faturamento_regiao': faturamento_regiao,
                'faturamento_tipo': faturamento_tipo,
                'top_clientes': top_clientes,
                'faturamento_total': sum(item['faturamento'] or 0 for item in faturamento_regiao)
            }

        except Exception as e:
            logging.error(f"❌ Erro ao gerar relatório financeiro: {e}")
            return {}

    def gerar_relatorio_regiao(self):
        """Gera relatório por região"""
        try:
            query = """
                SELECT Regiao,
                       COUNT(*) as total_placas,
                       SUM(CASE WHEN Status_Atual = 'disponível' THEN 1 ELSE 0 END) as disponiveis,
                       SUM(CASE WHEN Status_Atual = 'locado' THEN 1 ELSE 0 END) as locadas,
                       SUM(CASE WHEN Status_Atual = 'reservado' THEN 1 ELSE 0 END) as reservadas,
                       SUM(CASE WHEN Status_Atual = 'manutenção' THEN 1 ELSE 0 END) as manutencao,
                       SUM(Valor_Mensal) as faturamento_total
                FROM placas
                GROUP BY Regiao
                ORDER BY faturamento_total DESC
            """
            return self.db.execute_query(query) or []

        except Exception as e:
            logging.error(f"❌ Erro ao gerar relatório região: {e}")
            return []

    def gerar_relatorio_inventario(self):
        """Gera relatório completo do inventário"""
        try:
            query = """
                SELECT Codigo_Ativo, Endereco, Regiao, Tipo_Placa, Status_Atual,
                       Cliente_Locacao, Valor_Mensal, Data_Cadastro
                FROM placas
                ORDER BY Regiao, Codigo_Ativo
            """
            return self.db.execute_query(query) or []

        except Exception as e:
            logging.error(f"❌ Erro ao gerar relatório inventário: {e}")
            return []

# =============================================================================
# INICIALIZAR MANAGERS
# =============================================================================
relatorio_manager = RelatorioManager(DatabaseConnection())

# =============================================================================
# ROTAS PRINCIPAIS
# =============================================================================
@app.route('/')
def index():
    """Página inicial - redireciona para login"""
    return redirect('/login')

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Página de login"""
    if request.method == 'POST':
        email = request.form['email']
        senha = request.form['senha']

        # Simulação de login bem-sucedido
        if email in ['admin@formato.com', 'eduardo@formato.com', 'pamela@formato.com'] and senha == 'teste123':
            session['user'] = {
                'nome': 'Usuário Teste',
                'perfil': 'Administrador'
            }
            session.permanent = True
            return redirect('/inventario')
        else:
            return '''
            <!DOCTYPE html>
            <html>
            <head>
                <title>Login - SIG-ME</title>
                <style>
                    body { font-family: Arial, sans-serif; padding: 40px; text-align: center; }
                    .login-box { max-width: 400px; margin: 0 auto; padding: 40px; border: 1px solid #ddd; border-radius: 10px; }
                    input, button { width: 100%; padding: 10px; margin: 10px 0; }
                    .error { color: red; margin-top: 10px; }
                </style>
            </head>
            <body>
                <div class="login-box">
                    <h1>🚀 SIG-ME</h1>
                    <h3>Login</h3>
                    <div class="error">Login falhou! Verifique suas credenciais.</div>
                    <form method="POST">
                        <input type="email" name="email" placeholder="Email" required>
                        <input type="password" name="senha" placeholder="Senha" required>
                        <button type="submit">Entrar</button>
                    </form>
                    <p><strong>Contas teste:</strong> admin@formato.com / teste123</p>
                </div>
            </body>
            </html>
            '''

    return '''
    <!DOCTYPE html>
    <html>
    <head>
        <title>Login - SIG-ME</title>
        <style>
            body { font-family: Arial, sans-serif; padding: 40px; text-align: center; }
            .login-box { max-width: 400px; margin: 0 auto; padding: 40px; border: 1px solid #ddd; border-radius: 10px; }
            input, button { width: 100%; padding: 10px; margin: 10px 0; }
        </style>
    </head>
    <body>
        <div class="login-box">
            <h1>🚀 SIG-ME</h1>
            <h3>Login</h3>
            <form method="POST">
                <input type="email" name="email" placeholder="Email" required>
                <input type="password" name="senha" placeholder="Senha" required>
                <button type="submit">Entrar</button>
            </form>
            <p><strong>Contas teste:</strong> admin@formato.com / teste123</p>
        </div>
    </body>
    </html>
    '''

@app.route('/logout')
def logout():
    """Logout do usuário"""
    session.clear()
    return redirect('/login')

# =============================================================================
# MÓDULO DE INVENTÁRIO (COMPLETO)
# =============================================================================
@app.route('/inventario')
def inventario():
    """Página principal do inventário"""
    if 'user' not in session:
        return redirect('/login')

    db = DatabaseConnection()

    # Buscar todas as placas
    placas = db.execute_query("""
        SELECT * FROM placas
        ORDER BY Regiao, Codigo_Ativo
    """) or []

    # Métricas do inventário
    metricas = db.execute_query("""
        SELECT
            COUNT(*) as total_placas,
            SUM(CASE WHEN Status_Atual = 'disponível' THEN 1 ELSE 0 END) as disponiveis,
            SUM(CASE WHEN Status_Atual = 'locado' THEN 1 ELSE 0 END) as locadas,
            SUM(CASE WHEN Status_Atual = 'reservado' THEN 1 ELSE 0 END) as reservadas,
            SUM(CASE WHEN Status_Atual = 'manutenção' THEN 1 ELSE 0 END) as manutencao,
            SUM(Valor_Mensal) as faturamento_potencial
        FROM placas
    """)
    metricas = metricas[0] if metricas else {}

    # Distribuição por região
    distribuicao_regiao = db.execute_query("""
        SELECT Regiao, COUNT(*) as quantidade
        FROM placas
        GROUP BY Regiao
        ORDER BY quantidade DESC
    """) or []

    return render_inventario_template(session['user'], placas, metricas, distribuicao_regiao)

@app.route('/inventario/placa/nova', methods=['POST'])
def nova_placa():
    """Cria uma nova placa - CORRIGIDO"""
    if 'user' not in session:
        return jsonify({'success': False, 'message': 'Não autorizado'})

    try:
        dados = request.get_json()
        db = DatabaseConnection()

        # CORREÇÃO: Query atualizada para incluir todos os campos
        query = """
            INSERT INTO placas
            (Codigo_Ativo, Endereco, Regiao, Tipo_Placa, Valor_Mensal, Data_Cadastro, Status_Atual, Cliente_Locacao)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        """
        params = (
            dados['codigo_ativo'],
            dados['endereco'],
            dados['regiao'],
            dados['tipo_placa'],
            dados['valor_mensal'],
            datetime.datetime.now().date(),
            dados['status_atual'],  # CORREÇÃO: Usar status do formulário
            dados.get('cliente_locacao')  # CORREÇÃO: Incluir cliente
        )

        success = db.execute_update(query, params)

        if success:
            return jsonify({'success': True, 'message': 'Placa criada com sucesso!'})
        else:
            return jsonify({'success': False, 'message': 'Erro ao criar placa'})

    except Exception as e:
        logging.error(f"❌ Erro ao criar placa: {e}")
        return jsonify({'success': False, 'message': f'Erro: {str(e)}'})

@app.route('/inventario/placa/editar', methods=['POST'])
def editar_placa():
    """Edita uma placa existente - CORRIGIDO"""
    if 'user' not in session:
        return jsonify({'success': False, 'message': 'Não autorizado'})

    try:
        dados = request.get_json()
        db = DatabaseConnection()

        # CORREÇÃO: Query atualizada para garantir que todos os campos sejam atualizados
        query = """
            UPDATE placas
            SET Endereco = %s, Regiao = %s, Tipo_Placa = %s, Valor_Mensal = %s,
                Status_Atual = %s, Cliente_Locacao = %s
            WHERE Codigo_Ativo = %s
        """
        params = (
            dados['endereco'],
            dados['regiao'],
            dados['tipo_placa'],
            dados['valor_mensal'],
            dados['status_atual'],
            dados.get('cliente_locacao'),  # CORREÇÃO: Usar get() para evitar KeyError
            dados['codigo_ativo']
        )

        success = db.execute_update(query, params)

        if success:
            return jsonify({'success': True, 'message': 'Placa atualizada com sucesso!'})
        else:
            return jsonify({'success': False, 'message': 'Erro ao atualizar placa'})

    except Exception as e:
        logging.error(f"❌ Erro ao editar placa: {e}")
        return jsonify({'success': False, 'message': f'Erro: {str(e)}'})

@app.route('/inventario/placa/excluir/<string:codigo_placa>', methods=['DELETE'])
def excluir_placa(codigo_placa):
    """Exclui uma placa"""
    if 'user' not in session:
        return jsonify({'success': False, 'message': 'Não autorizado'})

    try:
        db = DatabaseConnection()

        query = "DELETE FROM placas WHERE Codigo_Ativo = %s"
        success = db.execute_update(query, (codigo_placa,))

        if success:
            return jsonify({'success': True, 'message': 'Placa excluída com sucesso!'})
        else:
            return jsonify({'success': False, 'message': 'Erro ao excluir placa'})

    except Exception as e:
        logging.error(f"❌ Erro ao excluir placa: {e}")
        return jsonify({'success': False, 'message': f'Erro: {str(e)}'})

@app.route('/inventario/placa/<string:codigo_placa>')
def detalhes_placa(codigo_placa):
    """Detalhes de uma placa específica"""
    if 'user' not in session:
        return redirect('/login')

    db = DatabaseConnection()
    placa = db.execute_query("SELECT * FROM placas WHERE Codigo_Ativo = %s", (codigo_placa,))

    if not placa:
        return "Placa não encontrada", 404

    placa = placa[0]

    # HTML para detalhes da placa
    return f'''
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Detalhes da Placa - SIG-ME</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: {CORES_FORMATO['claro']};
            color: {CORES_FORMATO['texto']};
        }}
        .header {{
            background: {CORES_FORMATO['branco']};
            padding: 20px 30px;
            box-shadow: 0 2px 20px rgba(0, 0, 0, 0.08);
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 4px solid {CORES_FORMATO['destaque']};
        }}
        .logo h1 {{
            color: {CORES_FORMATO['primaria']};
            font-size: 24px;
            font-weight: 700;
        }}
        .user-info {{
            display: flex;
            align-items: center;
            gap: 15px;
            color: {CORES_FORMATO['texto']};
            font-weight: 500;
        }}
        .btn-logout {{
            background: {CORES_FORMATO['destaque']};
            color: {CORES_FORMATO['branco']};
            padding: 10px 20px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            text-decoration: none;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s;
        }}
        .btn-logout:hover {{
            background: {CORES_FORMATO['alerta']};
            transform: translateY(-2px);
        }}
        .container {{
            max-width: 1200px;
            margin: 30px auto;
            padding: 0 25px;
        }}
        .navigation {{
            display: flex;
            gap: 12px;
            margin-bottom: 30px;
            flex-wrap: wrap;
        }}
        .nav-btn {{
            background: {CORES_FORMATO['primaria']};
            color: {CORES_FORMATO['branco']};
            padding: 14px 24px;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            text-decoration: none;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .nav-btn:hover {{
            background: {CORES_FORMATO['secundaria']};
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(26, 54, 93, 0.2);
        }}
        .content {{
            background: {CORES_FORMATO['branco']};
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
        }}
        .placa-detalhes {{
            max-width: 600px;
            margin: 0 auto;
        }}
        .placa-header {{
            text-align: center;
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 2px solid {CORES_FORMATO['claro']};
        }}
        .placa-codigo {{
            font-size: 32px;
            font-weight: 700;
            color: {CORES_FORMATO['primaria']};
            margin-bottom: 10px;
        }}
        .placa-status {{
            display: inline-block;
            padding: 8px 16px;
            border-radius: 20px;
            font-weight: 600;
            font-size: 14px;
        }}
        .status-disponivel {{ background: {CORES_FORMATO['sucesso']}20; color: {CORES_FORMATO['sucesso']}; }}
        .status-locado {{ background: {CORES_FORMATO['destaque']}20; color: {CORES_FORMATO['destaque']}; }}
        .status-reservado {{ background: {CORES_FORMATO['alerta']}20; color: {CORES_FORMATO['alerta']}; }}
        .status-manutencao {{ background: {CORES_FORMATO['info']}20; color: {CORES_FORMATO['info']}; }}
        .placa-info {{
            margin-bottom: 30px;
        }}
        .info-item {{
            margin-bottom: 15px;
            padding: 15px;
            background: {CORES_FORMATO['claro']};
            border-radius: 8px;
        }}
        .info-label {{
            font-weight: 600;
            color: {CORES_FORMATO['primaria']};
            margin-bottom: 5px;
        }}
        .info-value {{
            color: {CORES_FORMATO['texto']};
        }}
        .btn-voltar {{
            background: {CORES_FORMATO['primaria']};
            color: {CORES_FORMATO['branco']};
            padding: 12px 24px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            text-decoration: none;
            display: inline-block;
            font-weight: 600;
        }}
        .btn-salvar {{
            background: {CORES_FORMATO['sucesso']};
            color: {CORES_FORMATO['branco']};
            padding: 12px 24px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            margin-left: 10px;
        }}
        .form-group {{
            margin-bottom: 20px;
        }}
        .form-group label {{
            display: block;
            margin-bottom: 5px;
            font-weight: 600;
            color: {CORES_FORMATO['primaria']};
        }}
        .form-group input, .form-group select {{
            width: 100%;
            padding: 10px;
            border: 1px solid {CORES_FORMATO['claro']};
            border-radius: 5px;
            font-size: 14px;
        }}
    </style>
</head>
<body>
    <div class="header">
        <div class="logo">
            <h1>🚀 SIG-ME - Detalhes da Placa</h1>
        </div>
        <div class="user-info">
            <span>👋 Olá, {session['user']['nome']} <small>({session['user']['perfil']})</small></span>
            <a href="/logout" class="btn-logout">🚪 Sair</a>
        </div>
    </div>

    <div class="container">
        <div class="navigation">
            <a href="/inventario" class="nav-btn">📋 Voltar ao Inventário</a>
        </div>

        <div class="content">
            <div class="placa-detalhes">
                <div class="placa-header">
                    <div class="placa-codigo">{placa['Codigo_Ativo']}</div>
                    <div class="placa-status status-{placa['Status_Atual'].lower().replace('ç', 'c').replace('ã', 'a')}">{placa['Status_Atual']}</div>
                </div>

                <form id="formEditarPlaca">
                    <div class="placa-info">
                        <div class="form-group">
                            <label for="endereco">📍 Endereço</label>
                            <input type="text" id="endereco" name="endereco" value="{placa['Endereco']}" required>
                        </div>
                        <div class="form-group">
                            <label for="regiao">🏙️ Região</label>
                            <select id="regiao" name="regiao" required>
                                <option value="São Roque" {'selected' if placa['Regiao'] == 'São Roque' else ''}>São Roque</option>
                                <option value="Mairinque" {'selected' if placa['Regiao'] == 'Mairinque' else ''}>Mairinque</option>
                                <option value="Ibuna" {'selected' if placa['Regiao'] == 'Ibuna' else ''}>Ibuna</option>
                                <option value="Araçá" {'selected' if placa['Regiao'] == 'Araçá' else ''}>Araçá</option>
                                <option value="Piedade" {'selected' if placa['Regiao'] == 'Piedade' else ''}>Piedade</option>
                                <option value="Alumínio" {'selected' if placa['Regiao'] == 'Alumínio' else ''}>Alumínio</option>
                            </select>
                        </div>
                        <div class="form-group">
                            <label for="tipo_placa">📺 Tipo de Placa</label>
                            <select id="tipo_placa" name="tipo_placa" required>
                                <option value="Outdoor" {'selected' if placa['Tipo_Placa'] == 'Outdoor' else ''}>Outdoor</option>
                                <option value="Front Light" {'selected' if placa['Tipo_Placa'] == 'Front Light' else ''}>Front Light</option>
                                <option value="Back Light" {'selected' if placa['Tipo_Placa'] == 'Back Light' else ''}>Back Light</option>
                                <option value="Led" {'selected' if placa['Tipo_Placa'] == 'Led' else ''}>Led</option>
                            </select>
                        </div>
                        <div class="form-group">
                            <label for="cliente_locacao">👥 Cliente</label>
                            <input type="text" id="cliente_locacao" name="cliente_locacao" value="{placa['Cliente_Locacao'] or ''}" placeholder="Deixe vazio se não houver cliente">
                        </div>
                        <div class="form-group">
                            <label for="valor_mensal">💰 Valor Mensal (R$)</label>
                            <input type="number" id="valor_mensal" name="valor_mensal" value="{placa['Valor_Mensal']}" step="0.01" min="0" required>
                        </div>
                        <div class="form-group">
                            <label for="status_atual">Status</label>
                            <select id="status_atual" name="status_atual" required>
                                <option value="disponível" {'selected' if placa['Status_Atual'] == 'disponível' else ''}>Disponível</option>
                                <option value="locado" {'selected' if placa['Status_Atual'] == 'locado' else ''}>Locado</option>
                                <option value="reservado" {'selected' if placa['Status_Atual'] == 'reservado' else ''}>Reservado</option>
                                <option value="manutenção" {'selected' if placa['Status_Atual'] == 'manutenção' else ''}>Manutenção</option>
                            </select>
                        </div>
                    </div>

                    <div style="text-align: center; margin-top: 30px;">
                        <a href="/inventario" class="btn-voltar">← Voltar ao Inventário</a>
                        <button type="submit" class="btn-salvar">💾 Salvar Alterações</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <script>
        document.getElementById('formEditarPlaca').addEventListener('submit', function(e) {{
            e.preventDefault();

            const formData = new FormData(this);
            const dados = {{
                codigo_ativo: '{placa['Codigo_Ativo']}',
                endereco: formData.get('endereco'),
                regiao: formData.get('regiao'),
                tipo_placa: formData.get('tipo_placa'),
                valor_mensal: parseFloat(formData.get('valor_mensal')),
                status_atual: formData.get('status_atual'),
                cliente_locacao: formData.get('cliente_locacao') || null
            }};

            fetch('/inventario/placa/editar', {{
                method: 'POST',
                headers: {{
                    'Content-Type': 'application/json',
                }},
                body: JSON.stringify(dados)
            }})
            .then(response => response.json())
            .then(data => {{
                if (data.success) {{
                    alert('Placa atualizada com sucesso!');
                    window.location.href = '/inventario';
                }} else {{
                    alert('Erro ao atualizar placa: ' + data.message);
                }}
            }})
            .catch(error => {{
                console.error('Erro:', error);
                alert('Erro ao atualizar placa');
            }});
        }});
    </script>
</body>
</html>
'''

def render_inventario_template(user, placas, metricas, distribuicao_regiao):
    """Renderiza o template de inventário completo"""

    metricas_html = f"""
    <div class="metrics-grid">
        <div class="metric-card">
            <div class="metric-value">{metricas.get('total_placas', 0)}</div>
            <div class="metric-label">Total de Placas</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">{metricas.get('disponiveis', 0)}</div>
            <div class="metric-label">Disponíveis</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">{metricas.get('locadas', 0)}</div>
            <div class="metric-label">Locadas</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">{metricas.get('reservadas', 0)}</div>
            <div class="metric-label">Reservadas</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">{metricas.get('manutencao', 0)}</div>
            <div class="metric-label">Manutenção</div>
        </div>
    </div>
    """

    distribuicao_html = ""
    for regiao in distribuicao_regiao:
        distribuicao_html += f"""
        <div class="distribuicao-item" onclick="filtrarPorRegiao('{regiao['Regiao']}')">
            <div class="regiao-nome">{regiao['Regiao']}</div>
            <div class="regiao-quantidade">{regiao['quantidade']} placas</div>
        </div>
        """

    placas_html = ""
    for placa in placas:
        status_class = placa['Status_Atual'].lower().replace('ç', 'c').replace('ã', 'a')
        cliente = placa['Cliente_Locacao'] or 'Sem locação'
        placas_html += f"""
        <div class="placa-card {status_class}">
            <div class="placa-header">
                <h4>📋 {placa['Codigo_Ativo']}</h4>
                <span class="status-badge {status_class}">{placa['Status_Atual']}</span>
            </div>
            <div class="placa-info">
                <p><strong>📍 Endereço:</strong> {placa['Endereco']}</p>
                <p><strong>🏙️ Região:</strong> {placa['Regiao']}</p>
                <p><strong>📺 Tipo:</strong> {placa['Tipo_Placa']}</p>
                <p><strong>👥 Cliente:</strong> {cliente}</p>
                <p><strong>💰 Valor:</strong> R$ {placa['Valor_Mensal']:,.2f}</p>
                <p><strong>📅 Cadastro:</strong> {placa['Data_Cadastro']}</p>
            </div>
            <div class="placa-actions">
                <button class="btn-action" onclick="editarPlaca('{placa['Codigo_Ativo']}')">✏️ Editar</button>
                <button class="btn-action" onclick="detalhesPlaca('{placa['Codigo_Ativo']}')">👁️ Detalhes</button>
                <button class="btn-action btn-excluir" onclick="excluirPlaca('{placa['Codigo_Ativo']}')">🗑️ Excluir</button>
            </div>
        </div>
        """

    def convert_to_serializable(obj):
        if isinstance(obj, Decimal):
            return float(obj)
        elif isinstance(obj, datetime.date):
            return obj.isoformat()
        elif isinstance(obj, dict):
            return {key: convert_to_serializable(value) for key, value in obj.items()}
        elif isinstance(obj, list):
            return [convert_to_serializable(item) for item in obj]
        else:
            return obj

    placas_serializable = convert_to_serializable(placas)

    # CORREÇÃO: Template string JavaScript corrigida
    template_js = """
    placasFiltradas.forEach(placa => {
        const status_class = placa.Status_Atual.toLowerCase().replace('ç', 'c').replace('ã', 'a');
        const cliente = placa.Cliente_Locacao || 'Sem locação';

        placasGrid.innerHTML += `
        <div class="placa-card ${status_class}">
            <div class="placa-header">
                <h4>📋 ${placa.Codigo_Ativo}</h4>
                <span class="status-badge ${status_class}">${placa.Status_Atual}</span>
            </div>
            <div class="placa-info">
                <p><strong>📍 Endereço:</strong> ${placa.Endereco}</p>
                <p><strong>🏙️ Região:</strong> ${placa.Regiao}</p>
                <p><strong>📺 Tipo:</strong> ${placa.Tipo_Placa}</p>
                <p><strong>👥 Cliente:</strong> ${cliente}</p>
                <p><strong>💰 Valor:</strong> R$ ${placa.Valor_Mensal.toFixed(2).replace('.', ',')}</p>
                <p><strong>📅 Cadastro:</strong> ${placa.Data_Cadastro}</p>
            </div>
            <div class="placa-actions">
                <button class="btn-action" onclick="editarPlaca('${placa.Codigo_Ativo}')">✏️ Editar</button>
                <button class="btn-action" onclick="detalhesPlaca('${placa.Codigo_Ativo}')">👁️ Detalhes</button>
                <button class="btn-action btn-excluir" onclick="excluirPlaca('${placa.Codigo_Ativo}')">🗑️ Excluir</button>
            </div>
        </div>
        `;
    });
    """

    return f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Inventário - SIG-ME</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: {CORES_FORMATO['claro']};
            color: {CORES_FORMATO['texto']};
        }}
        .header {{
            background: {CORES_FORMATO['branco']};
            padding: 20px 30px;
            box-shadow: 0 2px 20px rgba(0, 0, 0, 0.08);
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 4px solid {CORES_FORMATO['destaque']};
        }}
        .logo h1 {{
            color: {CORES_FORMATO['primaria']};
            font-size: 24px;
            font-weight: 700;
        }}
        .user-info {{
            display: flex;
            align-items: center;
            gap: 15px;
            color: {CORES_FORMATO['texto']};
            font-weight: 500;
        }}
        .btn-logout {{
            background: {CORES_FORMATO['destaque']};
            color: {CORES_FORMATO['branco']};
            padding: 10px 20px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            text-decoration: none;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s;
        }}
        .btn-logout:hover {{
            background: {CORES_FORMATO['alerta']};
            transform: translateY(-2px);
        }}
        .container {{
            max-width: 1400px;
            margin: 30px auto;
            padding: 0 25px;
        }}
        .navigation {{
            display: flex;
            gap: 12px;
            margin-bottom: 30px;
            flex-wrap: wrap;
        }}
        .nav-btn {{
            background: {CORES_FORMATO['primaria']};
            color: {CORES_FORMATO['branco']};
            padding: 14px 24px;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            text-decoration: none;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .nav-btn:hover {{
            background: {CORES_FORMATO['secundaria']};
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(26, 54, 93, 0.2);
        }}
        .nav-btn.active {{
            background: {CORES_FORMATO['destaque']};
        }}
        .content {{
            background: {CORES_FORMATO['branco']};
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
        }}
        .section-title {{
            font-size: 24px;
            font-weight: 700;
            color: {CORES_FORMATO['primaria']};
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid {CORES_FORMATO['claro']};
        }}
        .metrics-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 15px;
            margin-bottom: 30px;
        }}
        .metric-card {{
            background: {CORES_FORMATO['claro']};
            padding: 20px;
            border-radius: 8px;
            text-align: center;
            border-left: 4px solid {CORES_FORMATO['info']};
        }}
        .metric-value {{
            font-size: 28px;
            font-weight: 700;
            color: {CORES_FORMATO['primaria']};
            margin-bottom: 5px;
        }}
        .metric-label {{
            color: {CORES_FORMATO['texto_claro']};
            font-size: 14px;
            font-weight: 600;
        }}
        .grid-2col {{
            display: grid;
            grid-template-columns: 1fr 2fr;
            gap: 30px;
            margin-top: 30px;
        }}
        .card {{
            background: {CORES_FORMATO['claro']};
            padding: 25px;
            border-radius: 10px;
        }}
        .card h3 {{
            color: {CORES_FORMATO['primaria']};
            margin-bottom: 20px;
            font-size: 18px;
        }}
        .distribuicao-item {{
            background: {CORES_FORMATO['branco']};
            padding: 15px;
            border-radius: 6px;
            margin-bottom: 10px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            cursor: pointer;
            transition: all 0.3s;
        }}
        .distribuicao-item:hover {{
            background: {CORES_FORMATO['info']}20;
            transform: translateX(5px);
        }}
        .regiao-nome {{
            font-weight: 600;
            color: {CORES_FORMATO['primaria']};
        }}
        .regiao-quantidade {{
            color: {CORES_FORMATO['texto_claro']};
            font-size: 14px;
        }}
        .placas-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
            gap: 20px;
            max-height: 600px;
            overflow-y: auto;
            padding: 10px;
        }}
        .placa-card {{
            background: {CORES_FORMATO['branco']};
            padding: 20px;
            border-radius: 10px;
            border: 1px solid {CORES_FORMATO['claro']};
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        }}
        .placa-card.disponivel {{ border-left: 4px solid {CORES_FORMATO['sucesso']}; }}
        .placa-card.locado {{ border-left: 4px solid {CORES_FORMATO['destaque']}; }}
        .placa-card.reservado {{ border-left: 4px solid {CORES_FORMATO['alerta']}; }}
        .placa-card.manutencao {{ border-left: 4px solid {CORES_FORMATO['info']}; }}
        .placa-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 1px solid {CORES_FORMATO['claro']};
        }}
        .placa-header h4 {{
            color: {CORES_FORMATO['primaria']};
        }}
        .status-badge {{
            padding: 6px 12px;
            border-radius: 15px;
            font-size: 12px;
            font-weight: 600;
        }}
        .status-badge.disponivel {{ background: {CORES_FORMATO['sucesso']}20; color: {CORES_FORMATO['sucesso']}; }}
        .status-badge.locado {{ background: {CORES_FORMATO['destaque']}20; color: {CORES_FORMATO['destaque']}; }}
        .status-badge.reservado {{ background: {CORES_FORMATO['alerta']}20; color: {CORES_FORMATO['alerta']}; }}
        .status-badge.manutencao {{ background: {CORES_FORMATO['info']}20; color: {CORES_FORMATO['info']}; }}
        .placa-info p {{
            margin: 8px 0;
            font-size: 14px;
        }}
        .placa-actions {{
            display: flex;
            gap: 10px;
            margin-top: 15px;
        }}
        .btn-action {{
            background: {CORES_FORMATO['primaria']};
            color: {CORES_FORMATO['branco']};
            border: none;
            padding: 8px 15px;
            border-radius: 5px;
            cursor: pointer;
            font-size: 12px;
            flex: 1;
        }}
        .btn-action:hover {{
            background: {CORES_FORMATO['secundaria']};
        }}
        .btn-excluir {{
            background: {CORES_FORMATO['destaque']};
        }}
        .btn-excluir:hover {{
            background: #c53030;
        }}
        .btn-nova-placa {{
            background: {CORES_FORMATO['sucesso']};
            color: {CORES_FORMATO['branco']};
            padding: 12px 24px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            margin-bottom: 20px;
            transition: all 0.3s;
        }}
        .btn-nova-placa:hover {{
            background: {CORES_FORMATO['info']};
            transform: translateY(-2px);
        }}
        .alert {{
            padding: 15px;
            margin: 15px 0;
            border-radius: 5px;
            font-weight: 500;
        }}
        .alert.success {{
            background: {CORES_FORMATO['sucesso']}20;
            color: {CORES_FORMATO['sucesso']};
            border: 1px solid {CORES_FORMATO['sucesso']};
        }}
        .alert.error {{
            background: {CORES_FORMATO['destaque']}20;
            color: {CORES_FORMATO['destaque']};
            border: 1px solid {CORES_FORMATO['destaque']};
        }}
        .modal {{
            display: none;
            position: fixed;
            z-index: 1000;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0,0,0,0.5);
        }}
        .modal-content {{
            background-color: {CORES_FORMATO['branco']};
            margin: 5% auto;
            padding: 30px;
            border-radius: 10px;
            width: 90%;
            max-width: 500px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.2);
        }}
        .modal-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            padding-bottom: 15px;
            border-bottom: 2px solid {CORES_FORMATO['claro']};
        }}
        .close {{
            color: {CORES_FORMATO['texto_claro']};
            font-size: 28px;
            font-weight: bold;
            cursor: pointer;
        }}
        .close:hover {{
            color: {CORES_FORMATO['destaque']};
        }}
        .form-group {{
            margin-bottom: 20px;
        }}
        .form-group label {{
            display: block;
            margin-bottom: 5px;
            font-weight: 600;
            color: {CORES_FORMATO['primaria']};
        }}
        .form-group input, .form-group select {{
            width: 100%;
            padding: 10px;
            border: 1px solid {CORES_FORMATO['claro']};
            border-radius: 5px;
            font-size: 14px;
        }}
        .form-actions {{
            display: flex;
            gap: 10px;
            justify-content: flex-end;
            margin-top: 20px;
        }}
        .btn-cancelar {{
            background: {CORES_FORMATO['texto_claro']};
            color: {CORES_FORMATO['branco']};
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
        }}
        .btn-salvar {{
            background: {CORES_FORMATO['sucesso']};
            color: {CORES_FORMATO['branco']};
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
        }}
    </style>
</head>
<body>
    <div class="header">
        <div class="logo">
            <h1>🚀 SIG-ME - Inventário</h1>
        </div>
        <div class="user-info">
            <span>👋 Olá, {user['nome']} <small>({user['perfil']})</small></span>
            <a href="/logout" class="btn-logout">🚪 Sair</a>
        </div>
    </div>

    <div class="container">
        <div class="navigation">
            <a href="/inventario" class="nav-btn active">📋 Inventário</a>
            <a href="/relatorios" class="nav-btn">📊 Relatórios</a>
            <a href="/comercial" class="nav-btn">💼 Comercial</a>
            <a href="/contratos" class="nav-btn">📑 Contratos</a>
            <a href="/exportar_dados" class="nav-btn">📥 Exportar Dados</a>
        </div>

        <div class="content">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
                <h2 class="section-title">📋 Módulo de Inventário</h2>
                <button class="btn-nova-placa" onclick="abrirModalNovaPlaca()">➕ Nova Placa</button>
            </div>

            <div id="alert-container"></div>

            {metricas_html}

            <div class="grid-2col">
                <div class="card">
                    <h3>🏙️ Distribuição por Região</h3>
                    {distribuicao_html if distribuicao_html else '<p>Nenhuma placa cadastrada</p>'}
                </div>

                <div class="card">
                    <h3>📋 Lista de Placas</h3>
                    <div class="placas-grid" id="placas-grid">
                        {placas_html if placas_html else '<p>Nenhuma placa cadastrada</p>'}
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal Nova Placa -->
    <div id="modalPlaca" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3 id="modalTituloPlaca">Nova Placa</h3>
                <span class="close" onclick="fecharModalPlaca()">&times;</span>
            </div>
            <form id="formPlaca">
                <input type="hidden" id="placa_codigo_editar" name="placa_codigo_editar">
                <div class="form-group">
                    <label for="codigo_ativo">Código da Placa:</label>
                    <input type="text" id="codigo_ativo" name="codigo_ativo" required>
                </div>
                <div class="form-group">
                    <label for="endereco">Endereço:</label>
                    <input type="text" id="endereco" name="endereco" required>
                </div>
                <div class="form-group">
                    <label for="regiao">Região:</label>
                    <select id="regiao" name="regiao" required>
                        <option value="">Selecione a região</option>
                        <option value="São Roque">São Roque</option>
                        <option value="Mairinque">Mairinque</option>
                        <option value="Ibuna">Ibuna</option>
                        <option value="Araçá">Araçá</option>
                        <option value="Piedade">Piedade</option>
                        <option value="Alumínio">Alumínio</option>
                    </select>
                </div>
                <div class="form-group">
                    <label for="tipo_placa">Tipo de Placa:</label>
                    <select id="tipo_placa" name="tipo_placa" required>
                        <option value="">Selecione o tipo</option>
                        <option value="Outdoor">Outdoor</option>
                        <option value="Front Light">Front Light</option>
                        <option value="Back Light">Back Light</option>
                        <option value="Led">Led</option>
                    </select>
                </div>
                <div class="form-group">
                    <label for="valor_mensal">Valor Mensal (R$):</label>
                    <input type="number" id="valor_mensal" name="valor_mensal" step="0.01" min="0" required>
                </div>
                <div class="form-group">
                    <label for="status_atual">Status:</label>
                    <select id="status_atual" name="status_atual" required>
                        <option value="disponível">Disponível</option>
                        <option value="locado">Locado</option>
                        <option value="reservado">Reservado</option>
                        <option value="manutenção">Manutenção</option>
                    </select>
                </div>
                <div class="form-group">
                    <label for="cliente_locacao">Cliente:</label>
                    <input type="text" id="cliente_locacao" name="cliente_locacao" placeholder="Deixe vazio se não houver cliente">
                </div>
                <div class="form-actions">
                    <button type="button" class="btn-cancelar" onclick="fecharModalPlaca()">Cancelar</button>
                    <button type="submit" class="btn-salvar">Salvar Placa</button>
                </div>
            </form>
        </div>
    </div>

    <script>
        let placaEditando = null;
        const placas = {json.dumps(placas_serializable)};

        function abrirModalNovaPlaca() {{
            document.getElementById('modalTituloPlaca').textContent = 'Nova Placa';
            document.getElementById('formPlaca').reset();
            document.getElementById('placa_codigo_editar').value = '';
            document.getElementById('codigo_ativo').disabled = false;
            placaEditando = null;
            document.getElementById('modalPlaca').style.display = 'block';
        }}

        function fecharModalPlaca() {{
            document.getElementById('modalPlaca').style.display = 'none';
        }}

        function editarPlaca(codigo) {{
            const placa = placas.find(p => p.Codigo_Ativo === codigo);

            if (placa) {{
                document.getElementById('modalTituloPlaca').textContent = 'Editar Placa';
                document.getElementById('placa_codigo_editar').value = placa.Codigo_Ativo;
                document.getElementById('codigo_ativo').value = placa.Codigo_Ativo;
                document.getElementById('codigo_ativo').disabled = true;
                document.getElementById('endereco').value = placa.Endereco;
                document.getElementById('regiao').value = placa.Regiao;
                document.getElementById('tipo_placa').value = placa.Tipo_Placa;
                document.getElementById('valor_mensal').value = placa.Valor_Mensal;
                document.getElementById('status_atual').value = placa.Status_Atual;
                document.getElementById('cliente_locacao').value = placa.Cliente_Locacao || '';

                placaEditando = placa;
                document.getElementById('modalPlaca').style.display = 'block';
            }}
        }}

        function detalhesPlaca(codigo) {{
            window.location.href = '/inventario/placa/' + codigo;
        }}

        function excluirPlaca(codigo) {{
            if (confirm('Tem certeza que deseja excluir a placa ' + codigo + '?')) {{
                fetch('/inventario/placa/excluir/' + codigo, {{
                    method: 'DELETE'
                }})
                .then(response => response.json())
                .then(data => {{
                    if (data.success) {{
                        mostrarAlert(data.message, 'success');
                        setTimeout(() => location.reload(), 1500);
                    }} else {{
                        mostrarAlert(data.message, 'error');
                    }}
                }})
                .catch(error => {{
                    console.error('Erro:', error);
                    mostrarAlert('Erro ao excluir placa', 'error');
                }});
            }}
        }}

        function filtrarPorRegiao(regiao) {{
            const placasGrid = document.getElementById('placas-grid');
            placasGrid.innerHTML = '';

            const placasFiltradas = placas.filter(placa => placa.Regiao === regiao);

            if (placasFiltradas.length === 0) {{
                placasGrid.innerHTML = '<p>Nenhuma placa encontrada para esta região.</p>';
                return;
            }}

            {template_js}
        }}

        function mostrarAlert(mensagem, tipo) {{
            const alertContainer = document.getElementById('alert-container');
            const alert = document.createElement('div');
            alert.className = 'alert ' + tipo;
            alert.textContent = mensagem;
            alertContainer.appendChild(alert);

            setTimeout(() => {{
                alert.remove();
            }}, 5000);
        }}

        // Fechar modal ao clicar fora
        window.onclick = function(event) {{
            const modal = document.getElementById('modalPlaca');
            if (event.target === modal) {{
                fecharModalPlaca();
            }}
        }}

        // Envio do formulário - CORREÇÃO: Garantir que todos os dados sejam enviados
        document.getElementById('formPlaca').addEventListener('submit', function(e) {{
            e.preventDefault();

            const formData = new FormData(this);
            const codigoEditar = formData.get('placa_codigo_editar');

            const dados = {{
                codigo_ativo: formData.get('codigo_ativo'),
                endereco: formData.get('endereco'),
                regiao: formData.get('regiao'),
                tipo_placa: formData.get('tipo_placa'),
                valor_mensal: parseFloat(formData.get('valor_mensal')),
                status_atual: formData.get('status_atual'),
                cliente_locacao: formData.get('cliente_locacao') || null  // CORREÇÃO: Garantir null se vazio
            }};

            const url = codigoEditar ? '/inventario/placa/editar' : '/inventario/placa/nova';
            const method = 'POST';

            fetch(url, {{
                method: method,
                headers: {{
                    'Content-Type': 'application/json',
                }},
                body: JSON.stringify(dados)
            }})
            .then(response => response.json())
            .then(data => {{
                if (data.success) {{
                    mostrarAlert(data.message, 'success');
                    fecharModalPlaca();
                    setTimeout(() => location.reload(), 1500);
                }} else {{
                    mostrarAlert(data.message, 'error');
                }}
            }})
            .catch(error => {{
                console.error('Erro:', error);
                mostrarAlert('Erro ao salvar placa', 'error');
            }});
        }});
    </script>
</body>
</html>
'''

# =============================================================================
# MÓDULO COMERCIAL (CORRIGIDO)
# =============================================================================
@app.route('/comercial')
def comercial():
    """Página comercial"""
    if 'user' not in session:
        return redirect('/login')

    db = DatabaseConnection()

    # CORREÇÃO: Query simplificada para métricas
    metricas_comerciais = db.execute_query("""
        SELECT
            COUNT(*) as total_placas,
            SUM(CASE WHEN Status_Atual IN ('locado', 'reservado') THEN Valor_Mensal ELSE 0 END) as faturamento_mensal,
            SUM(CASE WHEN Status_Atual IN ('locado', 'reservado') THEN 1 ELSE 0 END) as placas_ocupadas,
            SUM(CASE WHEN Status_Atual = 'disponível' THEN 1 ELSE 0 END) as placas_disponiveis,
            COUNT(DISTINCT Cliente_Locacao) as total_clientes
        FROM placas
    """)
    metricas_comerciais = metricas_comerciais[0] if metricas_comerciais else {}

    # CORREÇÃO: Query corrigida para clientes - incluindo ID explicitamente
    clientes = db.execute_query("SELECT id, Nome_Fantasia, Nome_Razao_Social, CNPJ_CPF, Telefone, Email FROM clientes ORDER BY Nome_Fantasia") or []

    return render_comercial_template(session['user'], metricas_comerciais, clientes, [])

@app.route('/comercial/cliente/novo', methods=['POST'])
def novo_cliente():
    """Cria um novo cliente"""
    if 'user' not in session:
        return jsonify({'success': False, 'message': 'Não autorizado'})

    try:
        dados = request.get_json()
        db = DatabaseConnection()

        query = """
            INSERT INTO clientes
            (Nome_Fantasia, Nome_Razao_Social, CNPJ_CPF, Telefone, Email)
            VALUES (%s, %s, %s, %s, %s)
        """
        params = (
            dados['nome_fantasia'],
            dados['razao_social'],
            dados['cnpj_cpf'],
            dados['telefone'],
            dados['email']
        )

        success = db.execute_update(query, params)

        if success:
            return jsonify({'success': True, 'message': 'Cliente criado com sucesso!'})
        else:
            return jsonify({'success': False, 'message': 'Erro ao criar cliente'})

    except Exception as e:
        logging.error(f"❌ Erro ao criar cliente: {e}")
        return jsonify({'success': False, 'message': f'Erro: {str(e)}'})

@app.route('/comercial/cliente/editar', methods=['POST'])
def editar_cliente():
    """Edita um cliente existente"""
    if 'user' not in session:
        return jsonify({'success': False, 'message': 'Não autorizado'})

    try:
        dados = request.get_json()
        db = DatabaseConnection()

        query = """
            UPDATE clientes
            SET Nome_Fantasia = %s, Nome_Razao_Social = %s, CNPJ_CPF = %s, Telefone = %s, Email = %s
            WHERE id = %s
        """
        params = (
            dados['nome_fantasia'],
            dados['razao_social'],
            dados['cnpj_cpf'],
            dados['telefone'],
            dados['email'],
            dados['id']
        )

        success = db.execute_update(query, params)

        if success:
            return jsonify({'success': True, 'message': 'Cliente atualizado com sucesso!'})
        else:
            return jsonify({'success': False, 'message': 'Erro ao atualizar cliente'})

    except Exception as e:
        logging.error(f"❌ Erro ao editar cliente: {e}")
        return jsonify({'success': False, 'message': f'Erro: {str(e)}'})

@app.route('/comercial/cliente/excluir/<int:cliente_id>', methods=['DELETE'])
def excluir_cliente(cliente_id):
    """Exclui um cliente"""
    if 'user' not in session:
        return jsonify({'success': False, 'message': 'Não autorizado'})

    try:
        db = DatabaseConnection()

        query = "DELETE FROM clientes WHERE id = %s"
        success = db.execute_update(query, (cliente_id,))

        if success:
            return jsonify({'success': True, 'message': 'Cliente excluído com sucesso!'})
        else:
            return jsonify({'success': False, 'message': 'Erro ao excluir cliente'})

    except Exception as e:
        logging.error(f"❌ Erro ao excluir cliente: {e}")
        return jsonify({'success': False, 'message': f'Erro: {str(e)}'})

def render_comercial_template(user, metricas, clientes, placas_ativas):
    """Renderiza o template comercial completo"""

    metricas_html = f"""
    <div class="metrics-grid">
        <div class="metric-card">
            <div class="metric-value">{metricas.get('total_clientes', 0)}</div>
            <div class="metric-label">Total de Clientes</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">R$ {metricas.get('faturamento_mensal', 0):,.2f}</div>
            <div class="metric-label">Faturamento Mensal</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">{metricas.get('placas_ocupadas', 0)}</div>
            <div class="metric-label">Placas Ocupadas</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">{metricas.get('placas_disponiveis', 0)}</div>
            <div class="metric-label">Placas Disponíveis</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">
                {((metricas.get('placas_ocupadas', 0) / (metricas.get('total_placas', 1)) * 100) if metricas.get('total_placas', 0) > 0 else 0):.1f}%
            </div>
            <div class="metric-label">Taxa de Ocupação</div>
        </div>
    </div>
    """

    clientes_html = ""
    for cliente in clientes:
        clientes_html += f"""
        <div class="cliente-card">
            <div class="cliente-header">
                <h4>{cliente['Nome_Fantasia']}</h4>
                <div class="cliente-actions">
                    <button class="btn-editar-cliente" onclick="editarCliente({cliente['id']})">✏️</button>
                    <button class="btn-excluir-cliente" onclick="excluirCliente({cliente['id']})">🗑️</button>
                </div>
            </div>
            <div class="cliente-info">
                <p><strong>Razão Social:</strong> {cliente['Nome_Razao_Social']}</p>
                <p><strong>CNPJ/CPF:</strong> {cliente['CNPJ_CPF']}</p>
                <p><strong>Telefone:</strong> {cliente['Telefone']}</p>
                <p><strong>Email:</strong> {cliente['Email']}</p>
            </div>
        </div>
        """

    # Converter clientes para JSON serializable
    def convert_to_serializable(obj):
        if isinstance(obj, Decimal):
            return float(obj)
        elif isinstance(obj, datetime.date):
            return obj.isoformat()
        elif isinstance(obj, dict):
            return {key: convert_to_serializable(value) for key, value in obj.items()}
        elif isinstance(obj, list):
            return [convert_to_serializable(item) for item in obj]
        else:
            return obj

    clientes_serializable = convert_to_serializable(clientes)

    return f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Comercial - SIG-ME</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: {CORES_FORMATO['claro']};
            color: {CORES_FORMATO['texto']};
        }}
        .header {{
            background: {CORES_FORMATO['branco']};
            padding: 20px 30px;
            box-shadow: 0 2px 20px rgba(0, 0, 0, 0.08);
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 4px solid {CORES_FORMATO['destaque']};
        }}
        .logo h1 {{
            color: {CORES_FORMATO['primaria']};
            font-size: 24px;
            font-weight: 700;
        }}
        .user-info {{
            display: flex;
            align-items: center;
            gap: 15px;
            color: {CORES_FORMATO['texto']};
            font-weight: 500;
        }}
        .btn-logout {{
            background: {CORES_FORMATO['destaque']};
            color: {CORES_FORMATO['branco']};
            padding: 10px 20px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            text-decoration: none;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s;
        }}
        .btn-logout:hover {{
            background: {CORES_FORMATO['alerta']};
            transform: translateY(-2px);
        }}
        .container {{
            max-width: 1400px;
            margin: 30px auto;
            padding: 0 25px;
        }}
        .navigation {{
            display: flex;
            gap: 12px;
            margin-bottom: 30px;
            flex-wrap: wrap;
        }}
        .nav-btn {{
            background: {CORES_FORMATO['primaria']};
            color: {CORES_FORMATO['branco']};
            padding: 14px 24px;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            text-decoration: none;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .nav-btn:hover {{
            background: {CORES_FORMATO['secundaria']};
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(26, 54, 93, 0.2);
        }}
        .nav-btn.active {{
            background: {CORES_FORMATO['destaque']};
        }}
        .content {{
            background: {CORES_FORMATO['branco']};
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
        }}
        .section-title {{
            font-size: 24px;
            font-weight: 700;
            color: {CORES_FORMATO['primaria']};
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid {CORES_FORMATO['claro']};
        }}
        .metrics-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        .metric-card {{
            background: {CORES_FORMATO['claro']};
            padding: 25px;
            border-radius: 10px;
            text-align: center;
            border-left: 4px solid {CORES_FORMATO['sucesso']};
        }}
        .metric-value {{
            font-size: 32px;
            font-weight: 700;
            color: {CORES_FORMATO['primaria']};
            margin-bottom: 8px;
        }}
        .metric-label {{
            color: {CORES_FORMATO['texto_claro']};
            font-size: 14px;
            font-weight: 600;
        }}
        .card {{
            background: {CORES_FORMATO['claro']};
            padding: 25px;
            border-radius: 10px;
            margin-top: 30px;
        }}
        .card h3 {{
            color: {CORES_FORMATO['primaria']};
            margin-bottom: 20px;
            font-size: 20px;
        }}
        .clientes-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
            gap: 20px;
            margin-top: 15px;
        }}
        .cliente-card {{
            background: {CORES_FORMATO['branco']};
            padding: 20px;
            border-radius: 8px;
            border: 1px solid {CORES_FORMATO['claro']};
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        }}
        .cliente-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
        }}
        .cliente-actions {{
            display: flex;
            gap: 5px;
        }}
        .btn-editar-cliente {{
            background: {CORES_FORMATO['info']};
            color: {CORES_FORMATO['branco']};
            border: none;
            padding: 5px 10px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
        }}
        .btn-excluir-cliente {{
            background: {CORES_FORMATO['destaque']};
            color: {CORES_FORMATO['branco']};
            border: none;
            padding: 5px 10px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
        }}
        .cliente-card h4 {{
            color: {CORES_FORMATO['primaria']};
            margin-bottom: 10px;
        }}
        .cliente-info p {{
            margin: 5px 0;
            font-size: 14px;
        }}
        .btn-novo-cliente {{
            background: {CORES_FORMATO['sucesso']};
            color: {CORES_FORMATO['branco']};
            padding: 12px 24px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            margin-bottom: 20px;
            transition: all 0.3s;
        }}
        .btn-novo-cliente:hover {{
            background: {CORES_FORMATO['info']};
            transform: translateY(-2px);
        }}
        .modal {{
            display: none;
            position: fixed;
            z-index: 1000;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0,0,0,0.5);
        }}
        .modal-content {{
            background-color: {CORES_FORMATO['branco']};
            margin: 5% auto;
            padding: 30px;
            border-radius: 10px;
            width: 90%;
            max-width: 500px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.2);
        }}
        .modal-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            padding-bottom: 15px;
            border-bottom: 2px solid {CORES_FORMATO['claro']};
        }}
        .close {{
            color: {CORES_FORMATO['texto_claro']};
            font-size: 28px;
            font-weight: bold;
            cursor: pointer;
        }}
        .close:hover {{
            color: {CORES_FORMATO['destaque']};
        }}
        .form-group {{
            margin-bottom: 20px;
        }}
        .form-group label {{
            display: block;
            margin-bottom: 5px;
            font-weight: 600;
            color: {CORES_FORMATO['primaria']};
        }}
        .form-group input {{
            width: 100%;
            padding: 10px;
            border: 1px solid {CORES_FORMATO['claro']};
            border-radius: 5px;
            font-size: 14px;
        }}
        .form-actions {{
            display: flex;
            gap: 10px;
            justify-content: flex-end;
            margin-top: 20px;
        }}
        .btn-cancelar {{
            background: {CORES_FORMATO['texto_claro']};
            color: {CORES_FORMATO['branco']};
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
        }}
        .btn-salvar {{
            background: {CORES_FORMATO['sucesso']};
            color: {CORES_FORMATO['branco']};
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
        }}
        .alert {{
            padding: 15px;
            margin: 15px 0;
            border-radius: 5px;
            font-weight: 500;
        }}
        .alert.success {{
            background: {CORES_FORMATO['sucesso']}20;
            color: {CORES_FORMATO['sucesso']};
            border: 1px solid {CORES_FORMATO['sucesso']};
        }}
        .alert.error {{
            background: {CORES_FORMATO['destaque']}20;
            color: {CORES_FORMATO['destaque']};
            border: 1px solid {CORES_FORMATO['destaque']};
        }}
    </style>
</head>
<body>
    <div class="header">
        <div class="logo">
            <h1>🚀 SIG-ME - Comercial</h1>
        </div>
        <div class="user-info">
            <span>👋 Olá, {user['nome']} <small>({user['perfil']})</small></span>
            <a href="/logout" class="btn-logout">🚪 Sair</a>
        </div>
    </div>

    <div class="container">
        <div class="navigation">
            <a href="/inventario" class="nav-btn">📋 Inventário</a>
            <a href="/relatorios" class="nav-btn">📊 Relatórios</a>
            <a href="/comercial" class="nav-btn active">💼 Comercial</a>
            <a href="/contratos" class="nav-btn">📑 Contratos</a>
            <a href="/exportar_dados" class="nav-btn">📥 Exportar Dados</a>
        </div>

        <div class="content">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
                <h2 class="section-title">💼 Módulo Comercial</h2>
                <button class="btn-novo-cliente" onclick="abrirModalNovoCliente()">➕ Novo Cliente</button>
            </div>
            <div id="alert-container"></div>

            {metricas_html}

            <div class="card">
                <h3>👥 Clientes Cadastrados</h3>
                <div class="clientes-grid">
                    {clientes_html if clientes_html else '<p>Nenhum cliente cadastrado</p>'}
                </div>
            </div>
        </div>
    </div>

    <!-- Modal Cliente -->
    <div id="modalCliente" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3 id="modalTituloCliente">Novo Cliente</h3>
                <span class="close" onclick="fecharModalCliente()">&times;</span>
            </div>
            <form id="formCliente">
                <input type="hidden" id="cliente_id" name="cliente_id">
                <div class="form-group">
                    <label for="nome_fantasia">Nome Fantasia:</label>
                    <input type="text" id="nome_fantasia" name="nome_fantasia" required>
                </div>
                <div class="form-group">
                    <label for="razao_social">Razão Social:</label>
                    <input type="text" id="razao_social" name="razao_social" required>
                </div>
                <div class="form-group">
                    <label for="cnpj_cpf">CNPJ/CPF:</label>
                    <input type="text" id="cnpj_cpf" name="cnpj_cpf" required>
                </div>
                <div class="form-group">
                    <label for="telefone">Telefone:</label>
                    <input type="text" id="telefone" name="telefone" required>
                </div>
                <div class="form-group">
                    <label for="email">Email:</label>
                    <input type="email" id="email" name="email" required>
                </div>
                <div class="form-actions">
                    <button type="button" class="btn-cancelar" onclick="fecharModalCliente()">Cancelar</button>
                    <button type="submit" class="btn-salvar">Salvar Cliente</button>
                </div>
            </form>
        </div>
    </div>
    <script>
        let clienteEditando = null;
        const clientes = {json.dumps(clientes_serializable)};

        function abrirModalNovoCliente() {{
            document.getElementById('modalTituloCliente').textContent = 'Novo Cliente';
            document.getElementById('formCliente').reset();
            document.getElementById('cliente_id').value = '';
            clienteEditando = null;
            document.getElementById('modalCliente').style.display = 'block';
        }}

        function fecharModalCliente() {{
            document.getElementById('modalCliente').style.display = 'none';
        }}

        function editarCliente(clienteId) {{
            const cliente = clientes.find(c => c.id === clienteId);

            if (cliente) {{
                document.getElementById('modalTituloCliente').textContent = 'Editar Cliente';
                document.getElementById('cliente_id').value = cliente.id;
                document.getElementById('nome_fantasia').value = cliente.Nome_Fantasia;
                document.getElementById('razao_social').value = cliente.Nome_Razao_Social;
                document.getElementById('cnpj_cpf').value = cliente.CNPJ_CPF;
                document.getElementById('telefone').value = cliente.Telefone;
                document.getElementById('email').value = cliente.Email;

                clienteEditando = cliente;
                document.getElementById('modalCliente').style.display = 'block';
            }}
        }}

        function excluirCliente(clienteId) {{
            if (confirm('Tem certeza que deseja excluir este cliente?')) {{
                fetch('/comercial/cliente/excluir/' + clienteId, {{ method: 'DELETE' }})
                    .then(response => response.json())
                    .then(data => {{
                        if (data.success) {{
                            mostrarAlert(data.message, 'success');
                            setTimeout(() => location.reload(), 1500);
                        }} else {{
                            mostrarAlert(data.message, 'error');
                        }}
                    }})
                    .catch(error => {{
                        console.error('Erro:', error);
                        mostrarAlert('Erro ao excluir cliente', 'error');
                    }});
            }}
        }}

        function mostrarAlert(mensagem, tipo) {{
            const alertContainer = document.getElementById('alert-container');
            const alert = document.createElement('div');
            alert.className = 'alert ' + tipo;
            alert.textContent = mensagem;
            alertContainer.appendChild(alert);

            setTimeout(() => {{
                alert.remove();
            }}, 5000);
        }}

        document.getElementById('formCliente').addEventListener('submit', function(e) {{
            e.preventDefault();

            const formData = new FormData(this);
            const clienteId = formData.get('cliente_id');

            const dados = {{
                nome_fantasia: formData.get('nome_fantasia'),
                razao_social: formData.get('razao_social'),
                cnpj_cpf: formData.get('cnpj_cpf'),
                telefone: formData.get('telefone'),
                email: formData.get('email')
            }};

            if (clienteId) {{
                // Editar cliente existente
                dados.id = parseInt(clienteId);
                fetch('/comercial/cliente/editar', {{
                    method: 'POST',
                    headers: {{ 'Content-Type': 'application/json' }},
                    body: JSON.stringify(dados)
                }})
                .then(response => response.json())
                .then(data => {{
                    if (data.success) {{
                        mostrarAlert(data.message, 'success');
                        fecharModalCliente();
                        setTimeout(() => location.reload(), 1500);
                    }} else {{
                        mostrarAlert(data.message, 'error');
                    }}
                }})
                .catch(error => {{
                    console.error('Erro:', error);
                    mostrarAlert('Erro ao editar cliente', 'error');
                }});
            }} else {{
                // Novo cliente
                fetch('/comercial/cliente/novo', {{
                    method: 'POST',
                    headers: {{ 'Content-Type': 'application/json' }},
                    body: JSON.stringify(dados)
                }})
                .then(response => response.json())
                .then(data => {{
                    if (data.success) {{
                        mostrarAlert(data.message, 'success');
                        fecharModalCliente();
                        setTimeout(() => location.reload(), 1500);
                    }} else {{
                        mostrarAlert(data.message, 'error');
                    }}
                }})
                .catch(error => {{
                    console.error('Erro:', error);
                    mostrarAlert('Erro ao criar cliente', 'error');
                }});
            }}
        }});

        // Fechar modal ao clicar fora
        window.onclick = function(event) {{
            const modal = document.getElementById('modalCliente');
            if (event.target === modal) {{
                fecharModalCliente();
            }}
        }}
    </script>
</body>
</html>
'''

# =============================================================================
# MÓDULO DE CONTRATOS (CORRIGIDO)
# =============================================================================
@app.route('/contratos')
def contratos():
    """Página de contratos"""
    if 'user' not in session:
        return redirect('/login')

    db = DatabaseConnection()

    # CORREÇÃO: Query corrigida incluindo todos os campos necessários
    contratos = db.execute_query("""
        SELECT id, tipo, cliente_id, fornecedor_id, descricao, valor_mensal,
               data_inicio, data_fim, status, observacoes, arquivo_pdf
        FROM contratos
        ORDER BY data_inicio DESC
    """) or []

    # Buscar nomes dos clientes separadamente
    clientes_map = {}
    clientes = db.execute_query("SELECT id, Nome_Fantasia FROM clientes") or []
    for cliente in clientes:
        clientes_map[cliente['id']] = cliente['Nome_Fantasia']

    # Adicionar nome do cliente aos contratos
    for contrato in contratos:
        if contrato['cliente_id']:
            contrato['cliente_nome'] = clientes_map.get(contrato['cliente_id'], 'Cliente não encontrado')
        else:
            contrato['cliente_nome'] = 'N/A'

    return render_contratos_template(session['user'], contratos)

@app.route('/contratos/novo', methods=['POST'])
def novo_contrato():
    """Cria um novo contrato"""
    if 'user' not in session:
        return jsonify({'success': False, 'message': 'Não autorizado'})

    try:
        dados = request.get_json()
        db = DatabaseConnection()

        query = """
            INSERT INTO contratos
            (tipo, cliente_id, fornecedor_id, descricao, valor_mensal, data_inicio, data_fim, status, observacoes, arquivo_pdf)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        params = (
            dados['tipo'],
            dados.get('cliente_id'),
            dados.get('fornecedor_id'),
            dados['descricao'],
            dados['valor_mensal'],
            dados['data_inicio'],
            dados.get('data_fim'),
            'ativo',
            dados.get('observacoes', ''),
            dados.get('arquivo_pdf', '')
        )

        success = db.execute_update(query, params)

        if success:
            return jsonify({'success': True, 'message': 'Contrato criado com sucesso!'})
        else:
            return jsonify({'success': False, 'message': 'Erro ao criar contrato'})

    except Exception as e:
        logging.error(f"❌ Erro ao criar contrato: {e}")
        return jsonify({'success': False, 'message': f'Erro: {str(e)}'})

@app.route('/contratos/excluir/<int:contrato_id>', methods=['DELETE'])
def excluir_contrato(contrato_id):
    """Exclui um contrato"""
    if 'user' not in session:
        return jsonify({'success': False, 'message': 'Não autorizado'})

    try:
        db = DatabaseConnection()

        query = "DELETE FROM contratos WHERE id = %s"
        success = db.execute_update(query, (contrato_id,))

        if success:
            return jsonify({'success': True, 'message': 'Contrato excluído com sucesso!'})
        else:
            return jsonify({'success': False, 'message': 'Erro ao excluir contrato'})

    except Exception as e:
        logging.error(f"❌ Erro ao excluir contrato: {e}")
        return jsonify({'success': False, 'message': f'Erro: {str(e)}'})

def render_contratos_template(user, contratos):
    """Renderiza o template de contratos simplificado"""

    contratos_html = ""
    for contrato in contratos:
        nome_contratante = contrato.get('cliente_nome', 'N/A')
        tipo_contratante = 'Cliente' if contrato.get('tipo') == 'cliente' else 'Fornecedor'
        data_fim = contrato.get('data_fim') or 'Não definida'
        arquivo_pdf = contrato.get('arquivo_pdf') or 'Nenhum arquivo'
        descricao = contrato.get('descricao', 'Sem descrição')[:30] + '...'
        valor_mensal = contrato.get('valor_mensal', 0)
        data_inicio = contrato.get('data_inicio', 'Não definida')
        status = contrato.get('status', 'ativo')

        contratos_html += f"""
        <div class="contrato-card">
            <div class="contrato-header">
                <h4>📋 {descricao}</h4>
                <div class="contrato-actions">
                    <button class="btn-pdf" onclick="visualizarPDF('{arquivo_pdf}')">📄 PDF</button>
                    <button class="btn-excluir" onclick="excluirContrato({contrato.get('id', 0)})">🗑️</button>
                </div>
            </div>
            <div class="contrato-info">
                <p><strong>{tipo_contratante}:</strong> {nome_contratante}</p>
                <p><strong>Valor Mensal:</strong> R$ {valor_mensal:,.2f}</p>
                <p><strong>Início:</strong> {data_inicio}</p>
                <p><strong>Término:</strong> {data_fim}</p>
                <p><strong>Status:</strong> <span class="status-badge {status}">{status}</span></p>
                <p><strong>Arquivo:</strong> {arquivo_pdf}</p>
            </div>
        </div>
        """

    return f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Contratos - SIG-ME</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: {CORES_FORMATO['claro']};
            color: {CORES_FORMATO['texto']};
        }}
        .header {{
            background: {CORES_FORMATO['branco']};
            padding: 20px 30px;
            box-shadow: 0 2px 20px rgba(0, 0, 0, 0.08);
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 4px solid {CORES_FORMATO['destaque']};
        }}
        .logo h1 {{
            color: {CORES_FORMATO['primaria']};
            font-size: 24px;
            font-weight: 700;
        }}
        .user-info {{
            display: flex;
            align-items: center;
            gap: 15px;
            color: {CORES_FORMATO['texto']};
            font-weight: 500;
        }}
        .btn-logout {{
            background: {CORES_FORMATO['destaque']};
            color: {CORES_FORMATO['branco']};
            padding: 10px 20px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            text-decoration: none;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s;
        }}
        .btn-logout:hover {{
            background: {CORES_FORMATO['alerta']};
            transform: translateY(-2px);
        }}
        .container {{
            max-width: 1400px;
            margin: 30px auto;
            padding: 0 25px;
        }}
        .navigation {{
            display: flex;
            gap: 12px;
            margin-bottom: 30px;
            flex-wrap: wrap;
        }}
        .nav-btn {{
            background: {CORES_FORMATO['primaria']};
            color: {CORES_FORMATO['branco']};
            padding: 14px 24px;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            text-decoration: none;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .nav-btn:hover {{
            background: {CORES_FORMATO['secundaria']};
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(26, 54, 93, 0.2);
        }}
        .nav-btn.active {{
            background: {CORES_FORMATO['destaque']};
        }}
        .content {{
            background: {CORES_FORMATO['branco']};
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
        }}
        .section-title {{
            font-size: 24px;
            font-weight: 700;
            color: {CORES_FORMATO['primaria']};
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid {CORES_FORMATO['claro']};
        }}
        .contratos-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
            gap: 20px;
            margin-top: 20px;
        }}
        .contrato-card {{
            background: {CORES_FORMATO['branco']};
            padding: 20px;
            border-radius: 10px;
            border: 1px solid {CORES_FORMATO['claro']};
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        }}
        .contrato-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 1px solid {CORES_FORMATO['claro']};
        }}
        .contrato-actions {{
            display: flex;
            gap: 8px;
        }}
        .btn-pdf {{
            background: {CORES_FORMATO['info']};
            color: {CORES_FORMATO['branco']};
            border: none;
            padding: 6px 12px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
        }}
        .btn-excluir {{
            background: {CORES_FORMATO['destaque']};
            color: {CORES_FORMATO['branco']};
            border: none;
            padding: 6px 12px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
        }}
        .contrato-info p {{
            margin: 8px 0;
            font-size: 14px;
        }}
        .status-badge {{
            padding: 4px 8px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: 600;
        }}
        .status-badge.ativo {{
            background: {CORES_FORMATO['sucesso']}20;
            color: {CORES_FORMATO['sucesso']};
        }}
        .status-badge.inativo {{
            background: {CORES_FORMATO['texto_claro']}20;
            color: {CORES_FORMATO['texto_claro']};
        }}
        .status-badge.vencido {{
            background: {CORES_FORMATO['destaque']}20;
            color: {CORES_FORMATO['destaque']};
        }}
        .btn-novo-contrato {{
            background: {CORES_FORMATO['sucesso']};
            color: {CORES_FORMATO['branco']};
            padding: 12px 24px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            margin-bottom: 20px;
            transition: all 0.3s;
        }}
        .btn-novo-contrato:hover {{
            background: {CORES_FORMATO['info']};
            transform: translateY(-2px);
        }}
        .alert {{
            padding: 15px;
            margin: 15px 0;
            border-radius: 5px;
            font-weight: 500;
        }}
        .alert.success {{
            background: {CORES_FORMATO['sucesso']}20;
            color: {CORES_FORMATO['sucesso']};
            border: 1px solid {CORES_FORMATO['sucesso']};
        }}
        .alert.error {{
            background: {CORES_FORMATO['destaque']}20;
            color: {CORES_FORMATO['destaque']};
            border: 1px solid {CORES_FORMATO['destaque']};
        }}
    </style>
</head>
<body>
    <div class="header">
        <div class="logo">
            <h1>🚀 SIG-ME - Contratos</h1>
        </div>
        <div class="user-info">
            <span>👋 Olá, {user['nome']} <small>({user['perfil']})</small></span>
            <a href="/logout" class="btn-logout">🚪 Sair</a>
        </div>
    </div>

    <div class="container">
        <div class="navigation">
            <a href="/inventario" class="nav-btn">📋 Inventário</a>
            <a href="/relatorios" class="nav-btn">📊 Relatórios</a>
            <a href="/comercial" class="nav-btn">💼 Comercial</a>
            <a href="/contratos" class="nav-btn active">📑 Contratos</a>
            <a href="/exportar_dados" class="nav-btn">📥 Exportar Dados</a>
        </div>

        <div class="content">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
                <h2 class="section-title">📑 Módulo de Contratos</h2>
                <button class="btn-novo-contrato" onclick="abrirModalNovoContrato()">➕ Novo Contrato</button>
            </div>
            <div id="alert-container"></div>

            <div class="contratos-grid">
                {contratos_html if contratos_html else '<p>Nenhum contrato cadastrado</p>'}
            </div>
        </div>
    </div>
    <script>
        function visualizarPDF(arquivo) {{
            if (arquivo !== 'Nenhum arquivo') {{
                alert('Visualizando PDF: ' + arquivo + '\\n\\nFuncionalidade em desenvolvimento!');
            }} else {{
                alert('Nenhum arquivo PDF disponível para este contrato.');
            }}
        }}
        function excluirContrato(contratoId) {{
            if (confirm('Tem certeza que deseja excluir este contrato?')) {{
                fetch('/contratos/excluir/' + contratoId, {{ method: 'DELETE' }})
                    .then(response => response.json())
                    .then(data => {{
                        if (data.success) {{
                            mostrarAlert(data.message, 'success');
                            setTimeout(() => location.reload(), 1500);
                        }} else {{
                            mostrarAlert(data.message, 'error');
                        }}
                    }})
                    .catch(error => {{
                        console.error('Erro:', error);
                        mostrarAlert('Erro ao excluir contrato', 'error');
                    }});
            }}
        }}
        function abrirModalNovoContrato() {{
            alert('Funcionalidade de novo contrato em desenvolvimento!');
        }}
        function mostrarAlert(mensagem, tipo) {{
            const alertContainer = document.getElementById('alert-container');
            const alert = document.createElement('div');
            alert.className = 'alert ' + tipo;
            alert.textContent = mensagem;
            alertContainer.appendChild(alert);

            setTimeout(() => {{
                alert.remove();
            }}, 5000);
        }}
    </script>
</body>
</html>
'''

# =============================================================================
# MÓDULOS ADICIONAIS (COMPLETADOS)
# =============================================================================
@app.route('/relatorios')
def relatorios():
    """Página de relatórios"""
    if 'user' not in session:
        return redirect('/login')

    db = DatabaseConnection()

    # Gerar relatórios
    relatorio_ocupacao = relatorio_manager.gerar_relatorio_ocupacao()
    relatorio_financeiro = relatorio_manager.gerar_relatorio_financeiro()
    relatorio_regiao = relatorio_manager.gerar_relatorio_regiao()

    return f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatórios - SIG-ME</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: {CORES_FORMATO['claro']};
            color: {CORES_FORMATO['texto']};
        }}
        .header {{
            background: {CORES_FORMATO['branco']};
            padding: 20px 30px;
            box-shadow: 0 2px 20px rgba(0, 0, 0, 0.08);
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 4px solid {CORES_FORMATO['destaque']};
        }}
        .logo h1 {{
            color: {CORES_FORMATO['primaria']};
            font-size: 24px;
            font-weight: 700;
        }}
        .user-info {{
            display: flex;
            align-items: center;
            gap: 15px;
            color: {CORES_FORMATO['texto']};
            font-weight: 500;
        }}
        .btn-logout {{
            background: {CORES_FORMATO['destaque']};
            color: {CORES_FORMATO['branco']};
            padding: 10px 20px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            text-decoration: none;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s;
        }}
        .btn-logout:hover {{
            background: {CORES_FORMATO['alerta']};
            transform: translateY(-2px);
        }}
        .container {{
            max-width: 1400px;
            margin: 30px auto;
            padding: 0 25px;
        }}
        .navigation {{
            display: flex;
            gap: 12px;
            margin-bottom: 30px;
            flex-wrap: wrap;
        }}
        .nav-btn {{
            background: {CORES_FORMATO['primaria']};
            color: {CORES_FORMATO['branco']};
            padding: 14px 24px;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            text-decoration: none;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .nav-btn:hover {{
            background: {CORES_FORMATO['secundaria']};
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(26, 54, 93, 0.2);
        }}
        .nav-btn.active {{
            background: {CORES_FORMATO['destaque']};
        }}
        .content {{
            background: {CORES_FORMATO['branco']};
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
        }}
        .section-title {{
            font-size: 24px;
            font-weight: 700;
            color: {CORES_FORMATO['primaria']};
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid {CORES_FORMATO['claro']};
        }}
        .relatorios-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
            gap: 25px;
            margin-top: 20px;
        }}
        .relatorio-card {{
            background: {CORES_FORMATO['claro']};
            padding: 25px;
            border-radius: 10px;
            border: 1px solid {CORES_FORMATO['claro']};
        }}
        .relatorio-card h3 {{
            color: {CORES_FORMATO['primaria']};
            margin-bottom: 15px;
            font-size: 18px;
        }}
        .relatorio-item {{
            background: {CORES_FORMATO['branco']};
            padding: 15px;
            border-radius: 6px;
            margin-bottom: 10px;
        }}
        .relatorio-item p {{
            margin: 5px 0;
            font-size: 14px;
        }}
        .btn-exportar {{
            background: {CORES_FORMATO['sucesso']};
            color: {CORES_FORMATO['branco']};
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 14px;
            margin-top: 10px;
        }}
    </style>
</head>
<body>
    <div class="header">
        <div class="logo">
            <h1>🚀 SIG-ME - Relatórios</h1>
        </div>
        <div class="user-info">
            <span>👋 Olá, {session['user']['nome']} <small>({session['user']['perfil']})</small></span>
            <a href="/logout" class="btn-logout">🚪 Sair</a>
        </div>
    </div>

    <div class="container">
        <div class="navigation">
            <a href="/inventario" class="nav-btn">📋 Inventário</a>
            <a href="/relatorios" class="nav-btn active">📊 Relatórios</a>
            <a href="/comercial" class="nav-btn">💼 Comercial</a>
            <a href="/contratos" class="nav-btn">📑 Contratos</a>
            <a href="/exportar_dados" class="nav-btn">📥 Exportar Dados</a>
        </div>

        <div class="content">
            <h2 class="section-title">📊 Módulo de Relatórios</h2>

            <div class="relatorios-grid">
                <div class="relatorio-card">
                    <h3>📊 Ocupação por Região</h3>
                    {generate_ocupacao_html(relatorio_ocupacao)}
                </div>

                <div class="relatorio-card">
                    <h3>💰 Relatórios Financeiros</h3>
                    {generate_financeiro_html(relatorio_financeiro)}
                </div>

                <div class="relatorio-card">
                    <h3>🏙️ Resumo por Região</h3>
                    {generate_regiao_html(relatorio_regiao)}
                </div>
            </div>
        </div>
    </div>
</body>
</html>
'''

def generate_ocupacao_html(relatorio_ocupacao):
    """Gera HTML para relatório de ocupação"""
    html = ""
    for regiao, dados in relatorio_ocupacao.items():
        html += f"""
        <div class="relatorio-item">
            <h4>{regiao}</h4>
            <p><strong>Total:</strong> {sum(item['quantidade'] for item in dados)} placas</p>
        """
        for item in dados:
            html += f"""
            <p><strong>{item['Status_Atual']}:</strong> {item['quantidade']} placas
            (Faturamento: R$ {item['faturamento'] or 0:,.2f})</p>
            """
        html += "</div>"
    return html if html else "<p>Nenhum dado disponível</p>"

def generate_financeiro_html(relatorio_financeiro):
    """Gera HTML para relatório financeiro"""
    if not relatorio_financeiro:
        return "<p>Nenhum dado financeiro disponível</p>"

    html = f"""
    <div class="relatorio-item">
        <p><strong>Faturamento Total:</strong> R$ {relatorio_financeiro['faturamento_total']:,.2f}</p>
    </div>
    """

    html += """
    <div class="relatorio-item">
        <h4>Faturamento por Região</h4>
    """
    for item in relatorio_financeiro['faturamento_regiao']:
        html += f"""
        <p><strong>{item['Regiao']}:</strong> R$ {item['faturamento']:,.2f}</p>
        """
    html += "</div>"

    html += """
    <div class="relatorio-item">
        <h4>Faturamento por Tipo de Placa</h4>
    """
    for item in relatorio_financeiro['faturamento_tipo']:
        html += f"""
        <p><strong>{item['Tipo_Placa']}:</strong> R$ {item['faturamento']:,.2f}</p>
        """
    html += "</div>"

    html += """
    <div class="relatorio-item">
        <h4>Top 10 Clientes</h4>
    """
    for item in relatorio_financeiro['top_clientes']:
        html += f"""
        <p><strong>{item['Cliente_Locacao']}:</strong> R$ {item['faturamento']:,.2f}</p>
        """
    html += "</div>"

    return html

def generate_regiao_html(relatorio_regiao):
    """Gera HTML para relatório por região"""
    if not relatorio_regiao:
        return "<p>Nenhum dado por região disponível</p>"

    html = ""
    for item in relatorio_regiao:
        html += f"""
        <div class="relatorio-item">
            <h4>{item['Regiao']}</h4>
            <p><strong>Total de Placas:</strong> {item['total_placas']}</p>
            <p><strong>Disponíveis:</strong> {item['disponiveis']}</p>
            <p><strong>Locadas:</strong> {item['locadas']}</p>
            <p><strong>Reservadas:</strong> {item['reservadas']}</p>
            <p><strong>Manutenção:</strong> {item['manutencao']}</p>
            <p><strong>Faturamento Total:</strong> R$ {item['faturamento_total']:,.2f}</p>
        </div>
        """
    return html

@app.route('/exportar_dados')
def exportar_dados():
    """Página de exportação de dados"""
    if 'user' not in session:
        return redirect('/login')

    return f'''
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exportar Dados - SIG-ME</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: {CORES_FORMATO['claro']};
            color: {CORES_FORMATO['texto']};
        }}
        .header {{
            background: {CORES_FORMATO['branco']};
            padding: 20px 30px;
            box-shadow: 0 2px 20px rgba(0, 0, 0, 0.08);
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 4px solid {CORES_FORMATO['destaque']};
        }}
        .logo h1 {{
            color: {CORES_FORMATO['primaria']};
            font-size: 24px;
            font-weight: 700;
        }}
        .user-info {{
            display: flex;
            align-items: center;
            gap: 15px;
            color: {CORES_FORMATO['texto']};
            font-weight: 500;
        }}
        .btn-logout {{
            background: {CORES_FORMATO['destaque']};
            color: {CORES_FORMATO['branco']};
            padding: 10px 20px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            text-decoration: none;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s;
        }}
        .btn-logout:hover {{
            background: {CORES_FORMATO['alerta']};
            transform: translateY(-2px);
        }}
        .container {{
            max-width: 1400px;
            margin: 30px auto;
            padding: 0 25px;
        }}
        .navigation {{
            display: flex;
            gap: 12px;
            margin-bottom: 30px;
            flex-wrap: wrap;
        }}
        .nav-btn {{
            background: {CORES_FORMATO['primaria']};
            color: {CORES_FORMATO['branco']};
            padding: 14px 24px;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            text-decoration: none;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .nav-btn:hover {{
            background: {CORES_FORMATO['secundaria']};
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(26, 54, 93, 0.2);
        }}
        .nav-btn.active {{
            background: {CORES_FORMATO['destaque']};
        }}
        .content {{
            background: {CORES_FORMATO['branco']};
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
        }}
        .section-title {{
            font-size: 24px;
            font-weight: 700;
            color: {CORES_FORMATO['primaria']};
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid {CORES_FORMATO['claro']};
        }}
        .export-options {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 20px;
            margin-top: 30px;
        }}
        .export-card {{
            background: {CORES_FORMATO['claro']};
            padding: 25px;
            border-radius: 10px;
            text-align: center;
        }}
        .export-card h3 {{
            color: {CORES_FORMATO['primaria']};
            margin-bottom: 15px;
        }}
        .btn-exportar {{
            background: {CORES_FORMATO['sucesso']};
            color: {CORES_FORMATO['branco']};
            padding: 12px 24px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s;
        }}
        .btn-exportar:hover {{
            background: {CORES_FORMATO['info']};
            transform: translateY(-2px);
        }}
    </style>
</head>
<body>
    <div class="header">
        <div class="logo">
            <h1>🚀 SIG-ME - Exportar Dados</h1>
        </div>
        <div class="user-info">
            <span>👋 Olá, {session['user']['nome']} <small>({session['user']['perfil']})</small></span>
            <a href="/logout" class="btn-logout">🚪 Sair</a>
        </div>
    </div>

    <div class="container">
        <div class="navigation">
            <a href="/inventario" class="nav-btn">📋 Inventário</a>
            <a href="/relatorios" class="nav-btn">📊 Relatórios</a>
            <a href="/comercial" class="nav-btn">💼 Comercial</a>
            <a href="/contratos" class="nav-btn">📑 Contratos</a>
            <a href="/exportar_dados" class="nav-btn active">📥 Exportar Dados</a>
        </div>

        <div class="content">
            <h2 class="section-title">📥 Exportar Dados</h2>

            <div class="export-options">
                <div class="export-card">
                    <h3>📋 Exportar Inventário</h3>
                    <p>Exportar todas as placas do inventário em formato Excel</p>
                    <button class="btn-exportar" onclick="window.location.href='/exportar/inventario'">Exportar</button>
                </div>

                <div class="export-card">
                    <h3>👥 Exportar Clientes</h3>
                    <p>Exportar lista de clientes em formato Excel</p>
                    <button class="btn-exportar" onclick="window.location.href='/exportar/clientes'">Exportar</button>
                </div>

                <div class="export-card">
                    <h3>📑 Exportar Contratos</h3>
                    <p>Exportar lista de contratos em formato Excel</p>
                    <button class="btn-exportar" onclick="window.location.href='/exportar/contratos'">Exportar</button>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
'''

@app.route('/exportar/inventario')
def exportar_inventario():
    """Exporta o inventário para Excel"""
    if 'user' not in session:
        return redirect('/login')

    try:
        db = DatabaseConnection()
        placas = db.execute_query("""
            SELECT Codigo_Ativo, Endereco, Regiao, Tipo_Placa, Status_Atual,
                   Cliente_Locacao, Valor_Mensal, Data_Cadastro
            FROM placas
            ORDER BY Regiao, Codigo_Ativo
        """) or []

        # Criar um workbook e adicionar uma planilha
        wb = Workbook()
        ws = wb.active
        ws.title = "Inventário"

        # Cabeçalhos
        headers = ["Código", "Endereço", "Região", "Tipo", "Status", "Cliente", "Valor Mensal", "Data Cadastro"]
        ws.append(headers)

        # Adicionar dados
        for placa in placas:
            ws.append([
                placa['Codigo_Ativo'],
                placa['Endereco'],
                placa['Regiao'],
                placa['Tipo_Placa'],
                placa['Status_Atual'],
                placa['Cliente_Locacao'] or "Sem locação",
                placa['Valor_Mensal'],
                placa['Data_Cadastro']
            ])

        # Formatação
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color=CORES_FORMATO['primaria'], end_color=CORES_FORMATO['primaria'], fill_type="solid")

        # Ajustar largura das colunas
        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width

        # Salvar em um buffer
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        return send_file(
            buffer,
            as_attachment=True,
            download_name=f'inventario_sigme_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        logging.error(f"❌ Erro ao exportar inventário: {e}")
        return "Erro ao exportar dados", 500

@app.route('/exportar/clientes')
def exportar_clientes():
    """Exporta a lista de clientes para Excel"""
    if 'user' not in session:
        return redirect('/login')

    try:
        db = DatabaseConnection()
        clientes = db.execute_query("""
            SELECT Nome_Fantasia, Nome_Razao_Social, CNPJ_CPF, Telefone, Email, created_at
            FROM clientes
            ORDER BY Nome_Fantasia
        """) or []

        # Criar um workbook e adicionar uma planilha
        wb = Workbook()
        ws = wb.active
        ws.title = "Clientes"

        # Cabeçalhos
        headers = ["Nome Fantasia", "Razão Social", "CNPJ/CPF", "Telefone", "Email", "Data Cadastro"]
        ws.append(headers)

        # Adicionar dados
        for cliente in clientes:
            ws.append([
                cliente['Nome_Fantasia'],
                cliente['Nome_Razao_Social'],
                cliente['CNPJ_CPF'],
                cliente['Telefone'],
                cliente['Email'],
                cliente['created_at']
            ])

        # Formatação
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color=CORES_FORMATO['primaria'], end_color=CORES_FORMATO['primaria'], fill_type="solid")

        # Ajustar largura das colunas
        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width

        # Salvar em um buffer
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        return send_file(
            buffer,
            as_attachment=True,
            download_name=f'clientes_sigme_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        logging.error(f"❌ Erro ao exportar clientes: {e}")
        return "Erro ao exportar dados", 500

@app.route('/exportar/contratos')
def exportar_contratos():
    """Exporta a lista de contratos para Excel"""
    if 'user' not in session:
        return redirect('/login')

    try:
        db = DatabaseConnection()
        contratos = db.execute_query("""
            SELECT tipo, cliente_id, fornecedor_id, descricao, valor_mensal,
                   data_inicio, data_fim, status, observacoes, arquivo_pdf, created_at
            FROM contratos
            ORDER BY data_inicio DESC
        """) or []

        # Criar um workbook e adicionar uma planilha
        wb = Workbook()
        ws = wb.active
        ws.title = "Contratos"

        # Cabeçalhos
        headers = ["Tipo", "Cliente ID", "Fornecedor ID", "Descrição", "Valor Mensal",
                   "Data Início", "Data Fim", "Status", "Observações", "Arquivo PDF", "Data Cadastro"]
        ws.append(headers)

        # Adicionar dados
        for contrato in contratos:
            ws.append([
                contrato['tipo'],
                contrato['cliente_id'],
                contrato['fornecedor_id'],
                contrato['descricao'],
                contrato['valor_mensal'],
                contrato['data_inicio'],
                contrato['data_fim'],
                contrato['status'],
                contrato['observacoes'],
                contrato['arquivo_pdf'],
                contrato['created_at']
            ])

        # Formatação
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color=CORES_FORMATO['primaria'], end_color=CORES_FORMATO['primaria'], fill_type="solid")

        # Ajustar largura das colunas
        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width

        # Salvar em um buffer
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        return send_file(
            buffer,
            as_attachment=True,
            download_name=f'contratos_sigme_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        logging.error(f"❌ Erro ao exportar contratos: {e}")
        return "Erro ao exportar dados", 500

if __name__ == '__main__':
    if inicializar_database():
        app.run(debug=DEBUG, host='0.0.0.0', port=5000)
    else:
        logging.error("Falha ao inicializar o sistema!")
