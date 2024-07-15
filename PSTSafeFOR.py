import os
import getpass
from pathlib import Path
import subprocess
import time
from tqdm import tqdm
from datetime import datetime
import openpyxl


# Função para registrar o backup em um arquivo Excel
def registrar_backup(nome_pessoa, arquivo_backup, pasta_destino):
    # Verificar se o arquivo não é um arquivo temporário (.tmp)
    if not arquivo_backup.endswith('.tmp'):
        log_path = os.path.join(pasta_destino, 'log_backup.xlsx')
        try:
            wb = openpyxl.load_workbook(log_path)
            ws = wb.active
        except FileNotFoundError:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Data", "Nome da Pessoa", "Arquivo de Backup"])

        data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        ws.append([data_hora, nome_pessoa, arquivo_backup])

        wb.save(log_path)

# Função para limpar a tela do console
def limpar_tela():
    os.system('cls')

# Função para imprimir texto em verde no console
def verde(texto):
    print("\033[92m" + texto + "\033[0m")

# Função para obter a data e hora atual
def data():
    dia = datetime.now().strftime(" %d/%m/%y")
    hora = datetime.now().strftime("às"" %H:%M:%S")
    verde(" BACKUP CONCLUÍDO EM:")
    print(dia, hora)

# Função para saber o tamanho do arquivo
def tamanho_do_arquivo(caminho_do_arquivo):
    tamanho = os.path.getsize(caminho_do_arquivo)
    return tamanho

# Função para fechar o Outlook
def fechar_outlook():
    os.system("taskkill /f /im outlook.exe 2> nul")
    print("Outlook fechado.\n")

# Função para abrir o Outlook
def abrir_outlook():
    try:
        # Usa o comando "start" para abrir o Outlook sem especificar o caminho
        os.system("start outlook")
        
        # Imprimir uma mensagem indicando que o Outlook foi aberto com sucesso
        print("\nMicrosoft Outlook foi aberto com sucesso.")
        
    except Exception as e:
        print(f"Erro ao abrir o Microsoft Outlook: {e}")
        
# Função para mapear uma unidade de rede no Windows
def mapear_unidade_rede_windows(letra_unidade, caminho_rede, usuarioADM=None, senha=None):
    cmd = ['net', 'use', letra_unidade + ':', caminho_rede]
    if usuarioADM is not None:
        cmd.extend(['/user:' + usuarioADM])
        if senha is not None:
            cmd.extend([senha])
    subprocess.run(cmd, shell=True)

# Função para desmapear uma unidade de rede no Windows
def desmapear_unidade_rede(letra_unidade):
    resultado = subprocess.run(['net', 'use', letra_unidade + ':', '/delete', '/y'], capture_output=True, text=True)
    if resultado.returncode == 0:
        print(f'\nA unidade de rede {letra_unidade} foi desmapeada com sucesso.\n')
    else:
        print('\nOcorreu um erro ao desmapear a unidade de rede.\n')
        print(resultado.stderr)

# Mapear a unidade de rede
mapear_unidade_rede_windows('B', '\\\\fs02\\backup_pst\\FOR', 'Jordan', '030323sJ$')

# Definir origem e destino dos arquivos usando o usuário logado no computador
caminho_pst = Path()
caminho_user = Path.home()
usuario = getpass.getuser()

# Fechar o Outlook para não corromper o PST
fechar_outlook()
time.sleep(1)
limpar_tela()

# Caminho do PST
origem = (f"{caminho_user}\\Documents\\Arquivos do Outlook")

# Pasta de destino dentro do usuário
pasta_destino = os.path.join("B:", usuario)
pasta_destino_log =os.path.join("B:\Log FOR")
if not os.path.exists(pasta_destino):
    os.makedirs(pasta_destino)

# Listar arquivos a serem copiados
arquivos = os.listdir(origem)

# Copiar PST para dentro da pasta de destino e exibir barra de progresso
for arquivo in arquivos:
    caminho_origem = os.path.join(origem, arquivo)
    caminho_destino = os.path.join(pasta_destino, arquivo)
    tamanho_arquivo = os.path.getsize(caminho_origem)
    with open(caminho_origem, 'rb') as f_origem:
        with open(caminho_destino, 'wb') as f_destino:
            with tqdm(total=tamanho_arquivo, desc=f"Copiando {arquivo}", unit="B", unit_scale=True, bar_format="{desc:<30}{percentage:3.0f}%|{bar:30}{r_bar}", dynamic_ncols=True, colour='green') as pbar:
                while True:
                    dados = f_origem.read(64*1024)
                    if not dados:
                        break
                    f_destino.write(dados)
                    pbar.update(len(dados))

# Verificar se o backup foi concluído com sucesso e gerar o log
time.sleep(1.5)
if tamanho_do_arquivo(caminho_destino) == tamanho_do_arquivo(caminho_origem):
    time.sleep(1)
    limpar_tela()
    time.sleep(1)
    data()
    time.sleep(2)
    verde('\nO seu outlook será aberto em segundos...\n')
    time.sleep(7)
    abrir_outlook()
    # Registrar o backup no log Excel
for arquivo in arquivos:
    registrar_backup(usuario, arquivo, pasta_destino_log)
time.sleep(5)   
desmapear_unidade_rede('b')