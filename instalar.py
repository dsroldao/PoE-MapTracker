import os
import shutil
import sys
import subprocess

# --- CONFIGURAÇÃO ---
APP_NAME = "PoE Map Tracker"
EXE_NAME = "PoEMapTracker.exe" # Nome final do arquivo
SOURCE_EXE = os.path.join("dist", "tracker.exe") # Nome que o PyInstaller gera (baseado no tracker.py)

def install():
    print(f"--- INSTALANDO {APP_NAME} ---")

    # 1. Verifica se o executável foi criado
    if not os.path.exists(SOURCE_EXE):
        print(f"[ERRO] Arquivo '{SOURCE_EXE}' não encontrado!")
        print("Você precisa rodar o comando 'pyinstaller' primeiro.")
        input("Pressione ENTER para sair...")
        sys.exit()

    # 2. Caminhos
    # Instala na pasta %localappdata%
    install_dir = os.path.join(os.environ["LOCALAPPDATA"], "PoEMapTracker")
    target_exe = os.path.join(install_dir, EXE_NAME)
    
    # Caminho do Menu Iniciar
    start_menu = os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "Start Menu", "Programs")
    shortcut_path = os.path.join(start_menu, f"{APP_NAME}.lnk")

    # 3. Copiar Arquivo
    print(f"Instalando em: {install_dir}...")
    if not os.path.exists(install_dir):
        os.makedirs(install_dir)

    try:
        shutil.copy2(SOURCE_EXE, target_exe)
        print("Arquivos copiados.")
    except Exception as e:
        print(f"[ERRO] Não foi possível copiar: {e}")
        input("Pressione ENTER para sair...")
        sys.exit()

    # 4. Criar Atalho via PowerShell
    print("Criando atalho no Menu Iniciar...")
    ps_script = f"""
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("{shortcut_path}")
    $Shortcut.TargetPath = "{target_exe}"
    $Shortcut.WorkingDirectory = "{install_dir}"
    $Shortcut.Description = "PoE Map Tracker"
    $Shortcut.Save()
    """
    try:
        subprocess.run(["powershell", "-Command", ps_script], check=True)
        print("Atalho criado com sucesso.")
    except:
        print("Aviso: Não foi possível criar o atalho automaticamente.")

    print("\n--- SUCESSO! ---")
    print(f"Pode fechar esta janela. Procure por '{APP_NAME}' no seu Menu Iniciar.")
    input("Pressione ENTER para fechar.")

if __name__ == "__main__":
    install()