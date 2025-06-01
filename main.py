import requests
from tqdm import tqdm  # precisa instalar: pip install tqdm
import subprocess
import os
import msvcrt
from time import sleep

from rich import print
from rich.panel import Panel

# criar exe
# pyinstaller --onefile --icon=meu_icone.ico nome_do_seu_script.py
# pyinstaller --onefile --exclude-module=difflib main.py

# dimensão icone, 16, 48, 256


def start_download(url, file_name, v):
    os.system('cls')
    print(f"\nIniciando Download do Pacote Office - {v}\n")
    response = requests.get(url, stream=True)
    total = int(response.headers.get('content-length', 0))  # Tamanho total do arquivo em bytes

    with open(file_name, "wb") as arq, tqdm(
        desc=f"Baixando {file_name}",
        total=total,
        unit='B',
        unit_scale=True,
        unit_divisor=1024
    ) as bar:
        for b in response.iter_content(chunk_size=8192):
            if b:
                arq.write(b)
                bar.update(len(b))

    print(f"\nDownload concluído: {file_name}")
    sleep(3)
    start()
    

def install_off(path_img):
    print("\nMontando imagem...")

    # Garante o caminho absoluto
    img_path = os.path.abspath(path_img)

    if not os.path.exists(img_path):
        print(f"Arquivo não encontrado: {img_path}")
        return

    subprocess.run(["PowerShell", "Mount-DiskImage", "-ImagePath", img_path], check=True)

    print("Imagem montada. Buscando unidade...")

    output = subprocess.check_output([
        "PowerShell",
        f"(Get-DiskImage -ImagePath '{img_path}') | Get-Volume | Select-Object -ExpandProperty DriveLetter"
    ], text=True).strip()

    letra_unidade = output + ":\\"
    print(f"Unidade montada em {letra_unidade}")

    setup_path = find_setup(letra_unidade)
    print(f"Caminho do setup.exe: {setup_path}")

    print("Iniciando instalação...")
    # # subprocess.run([setup_path], check=True)
    
    
    pwr_cmd = f"""
    Set-Location -Path {letra_unidade}
    Start-Process .\\Setup.exe -Verb runAs
    """

    subprocess.run(["PowerShell", "-Command", pwr_cmd], check=True)
    
    # subprocess.run(["PowerShell", "Start-Process", setup_path, "-Verb", "runAs"], check=True)


    # print("Instalação iniciada.")
    
    """ 
        if not os.path.exists(path_img):
        print(f"Arquivo não encontrado: {path_img}")
        return
    """

def find_setup(l):
    arq = os.listdir(l)
    
    for name in arq:
        if name.lower() == "setup.exe":
            return os.path.join(l, name)
    
    return None  # Nenhum encontrado


def down_off():
    version_off = {
        "2019": f"https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/pt-br/ProPlus2019Retail.img",
        "2021": f"https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/pt-br/ProPlus2021Retail.img",
    }
    opcao = "0"
    
    os.system('cls')
    while True:
        # menu_text = (
        #     "[bright_yellow]1[/bright_yellow] - Download Office\n"
        #     "[bright_yellow]2[/bright_yellow] - Ativar Office"
        # )
        # print(Panel.fit(menu_text, title="Menu", padding=(4, 5, 4, 5)))
        
        menu_text = (
            "Escolha a versão do Office para instalar: \n\n"
            "[bright_yellow]1[/bright_yellow] - Office 2019\n"
            "[bright_yellow]2[/bright_yellow] - Office 2021\n"
            "[bright_yellow]0[/bright_yellow] - Voltar"
        )
        print(Panel.fit(menu_text, title="Menu", padding=(4, 5, 4, 5)))
        
        opcao = input("Digite o número da versão desejada: ")
        
        if opcao == "1":
            url = version_off["2019"]
            v = "2019"
            file_name = "Office2019.img"
            break
        elif opcao == "2":
            url = version_off["2021"]
            v = "2021"
            file_name = "Office2021.img"
            break
        elif opcao == "0":
            start()
        else:
            print(f"\n[red]Opção ({opcao}) inválida. Tente novamente[/red]")
            print("Precione qualquer tecla para continuar....")
            
            # print("\033[F" + " " * 50 + "\033[F", end='')
            msvcrt.getch()
        
        os.system('cls')
    
    start_download(url, file_name, v)
    


""" def ler_input_interativo(prompt):
    buffer = ""
    sys.stdout.write(prompt)
    sys.stdout.flush()

    while True:
        char = msvcrt.getch()

        # ENTER
        if char == b'\r':
            print()
            return buffer

        # BACKSPACE
        elif char == b'\x08':
            if len(buffer) > 0:
                buffer = buffer[:-1]
                sys.stdout.write('\b \b')
                sys.stdout.flush()

        # Caracter visível
        elif 32 <= ord(char) <= 126:
            buffer += char.decode()
            sys.stdout.write(char.decode())
            sys.stdout.flush() """


def activate_off():
    os.system("cls")
    
    """ 
        menu_text = (
            "[bright_yellow]1[/bright_yellow] - Download Office\n"
            "[bright_yellow]2[/bright_yellow] - Ativar Office"
        )
        print(Panel.fit(menu_text, title="Menu", padding=(4, 5, 4, 5)))
    """
    menu_text = (
        "[bright_red]ATENÇÃO[/bright_red]: Vai abrir uma janela preta chamada PowerShell.\n\n"
        "Caso peça permição permisão do sistema, aperte a opção [bright_green]'Sim'[/bright_green]\n"
        "Nessa janela, você vai ver algumas opções numeradas.\n\n"
        "Quando a janela abrir completamente:\n"
        "\t- Aperte a tecla [bright_green]'2'[/bright_green] e depois [bright_green]'Enter'[/bright_green]. Para opção [bright_green]'Ohook'[/bright_green]\n"
        "\t- Em seguida aperte [bright_green]'1'[/bright_green] e depois [bright_green]'Enter'[/bright_green]. Para opção [bright_green]'Install Ohook Office Activation'[/bright_green]\n\n"
        "Depois disso, espere o processo terminar sozinho, pode demorar um pouco.\n"
        "Não feche a janela até ver que o processo terminou.\n"
        "Mantenha o PC conectado a internet até o processo ser concluido\n"
    )
    print(Panel.fit(menu_text, title="INSTRUÇÕES", padding=(4, 7, 4, 7)))
        
    
    # print("ATENÇÃO: Vai abrir uma janela preta chamada PowerShell.\n")
    # print("\tNessa janela, você vai ver algumas opções numeradas.")
    # print("\tQuando a janela abrir completamente, aperte a tecla '2' e depois 'Enter'")
    # print("\tEm seguida aperte '1' e depois 'Enter'")
    # print("\tDepois disso, espere o processo terminar sozinho, pode demorar um pouco.")
    # print("\tNão feche a janela até ver que o processo terminou.")
    # print("\tMantenha o PC conectado a internet até o processo ser concluido")
    # print()
    print("Para continuar, pressione qualquer tecla agora...")
    msvcrt.getch()

    
    
    print("Iniciando ativação via PowerShell...")

    try:
        subprocess.run(
            ["PowerShell", "-Command", "irm https://get.activated.win | iex"],
            check=True
        )
    except subprocess.CalledProcessError as e:
        print("Erro ao ativar o Office:")
        print(e)



def start():
    os.system('cls')
    # prompt = "Digite um número: "
    # def draw_fancy_square(width, height):
    #     print("╔" + "═" * (width - 2) + "╗")
    #     for _ in range(height - 2):
    #         print("║" + " " * (width - 2) + "║")
    #     print("╚" + "═" * (width - 2) + "╝")

    # draw_fancy_square(30, 10)  # quadrado 30 colunas x 10 linhas
    
    
    opcao = ""
    while True: 
        # print("\n")
        # print(f"{Fore.LIGHTYELLOW_EX}\t1{Style.RESET_ALL} - Download Office")
        # print(f"{Fore.LIGHTYELLOW_EX}\t2{Style.RESET_ALL} - Ativar Office")
        # print("\n")

        # print(Panel.fit(f"[yellow]1[/yellow] - Download Office\n[yellow]2[/yellow] - Ativar Office", title="Menu"))
        
        menu_text = (
            "[bright_yellow]1[/bright_yellow] - Download Office\n"
            "[bright_yellow]2[/bright_yellow] - Ativar Office"
        )
        print(Panel.fit(menu_text, title="Menu", padding=(4, 5, 4, 5)))
        
        # print("Digite um número: ", end="", flush=True)
        opcao = input("Digite um número: ")
        
        if opcao == "1":
            down_off()
            break
        elif opcao == "2":
            activate_off()
            break
        else: 
            # print(f"\n{Fore.RED}Opção ({opcao}) inválida. Tente novamente{Style.RESET_ALL}")
            print(f"\n[red]Opção ({opcao}) inválida. Tente novamente[/red]")
            print("Precione qualquer tecla para continuar....")
            
            # print("\033[F" + " " * 50 + "\033[F", end='')
            msvcrt.getch()
        os.system('cls')
        # draw_fancy_square(30, 10)  # redesenha o quadrado após limpar a tela




    
start()

# install_off("Office2019.img")