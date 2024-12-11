import ttkbootstrap as tb
from tkinter import filedialog, messagebox
import os
import subprocess
import sys

# Variáveis globais para armazenar os caminhos das planilhas
planilha_1_caminho = ""
planilha_2_caminho = ""

def selecionar_primeira_planilha():
    """Abre um diálogo para selecionar a primeira planilha."""
    global planilha_1_caminho
    file_path = filedialog.askopenfilename(
        title="Selecione a Planilha Estoque (opcional)",
        filetypes=[("Planilhas Excel", "*.xlsx *.xls")]
    )
    if file_path:
        planilha_1_caminho = file_path
        label_1.config(text=f"Planilha Estoque: {os.path.basename(file_path)}")
    else:
        messagebox.showinfo("Info", "Nenhuma planilha selecionada para Estoque.")

def selecionar_segunda_planilha():
    """Abre um diálogo para selecionar a segunda planilha."""
    global planilha_2_caminho
    file_path = filedialog.askopenfilename(
        title="Selecione a Planilha Vendido (opcional)",
        filetypes=[("Planilhas Excel", "*.xlsx *.xls")]
    )
    if file_path:
        planilha_2_caminho = file_path
        label_2.config(text=f"Planilha Vendido: {os.path.basename(file_path)}")
    else:
        messagebox.showinfo("Info", "Nenhuma planilha selecionada para Vendido.")

def chamar_subaplicativo(file_path_1=None, file_path_2=None, output_file=None):
    """Chama o script correto para processar as planilhas."""
    try:
        # Caminho fixo para o script
        script_path = r"C:\Users\nicolas.souza\Desktop\Planilhas\dist\RevendaMais\Estoque_Veiculos\estoque_veiculos.py"

        # Validar se o script existe
        if not os.path.exists(script_path):
            raise FileNotFoundError(f"Script não encontrado: {script_path}")

        # Preparar os argumentos
        args = [sys.executable, script_path]
        if file_path_1:
            args.append(file_path_1)
        if file_path_2:
            args.append(file_path_2)
        args.append(output_file)

        # Log de depuração
        print(f"Comando para subprocess: {' '.join(args)}")

        # Executar o script com subprocess
        result = subprocess.run(args, check=True, capture_output=True, text=True)

        # Log de saída do script
        print("Saída do script:", result.stdout)

        # Mensagem de sucesso
        messagebox.showinfo("Concluído", "Conversão da planilha concluída com sucesso!")
        limpar_selecao()

    except FileNotFoundError as e:
        messagebox.showerror("Erro", str(e))
    except subprocess.CalledProcessError as e:
        # Log de erro do script
        print("Erro no script:", e.stderr)
        messagebox.showerror("Erro", f"Erro ao executar o script: {e.stderr}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro inesperado: {e}")

def limpar_selecao():
    """Limpa as seleções de planilhas no formulário."""
    global planilha_1_caminho, planilha_2_caminho
    planilha_1_caminho = ""
    planilha_2_caminho = ""
    label_1.config(text="Nenhuma planilha selecionada")
    label_2.config(text="Nenhuma planilha selecionada")

def converter():
    """Função chamada ao clicar no botão de converter."""
    if not planilha_1_caminho and not planilha_2_caminho:
        messagebox.showwarning("Aviso", "Selecione pelo menos uma planilha para converter.")
        return

    # Perguntar onde salvar o arquivo de saída
    output_file = filedialog.asksaveasfilename(
        title="Selecione o local e o nome do arquivo de saída",
        defaultextension=".xlsx",
        filetypes=[("Planilhas Excel", "*.xlsx"), ("Todos os Arquivos", "*.*")]
    )

    if not output_file:
        messagebox.showwarning("Aviso", "Nenhum arquivo foi selecionado para salvar o resultado.")
        return

    chamar_subaplicativo(planilha_1_caminho, planilha_2_caminho, output_file)

def mostrar_opcoes_de_planilha():
    """Exibe as opções de seleção para planilhas depois de selecionar o sistema e a categoria."""
    if sistema_var.get() and categoria_var.get():
        frame_planilhas.pack(pady=10, fill="both", expand=True)
    else:
        messagebox.showwarning("Erro", "Por favor, selecione o sistema e a funcionalidade primeiro!")

# Configuração da interface
root = tb.Window(themename="cosmo")
root.title("Conversor de Planilhas")
root.geometry("600x500")

# Centraliza a janela na tela
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = 600
window_height = 500
center_x = (screen_width // 2) - (window_width // 2)
center_y = (screen_height // 2) - (window_height // 2)
root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")

# Configuração da interface principal
header_label = tb.Label(
    root,
    text="Conversor de Planilhas",
    font=("Arial", 18, "bold"),
    bootstyle="light"
)
header_label.pack(pady=10)

# Dropdown para o sistema e categoria
frame_main = tb.Frame(root)
frame_main.pack(pady=10)

# Dropdown Sistema
tb.Label(frame_main, text="Escolha o sistema:", font=("Arial", 12)).pack(pady=5)
sistema_var = tb.StringVar()
sistema_dropdown = tb.Combobox(
    frame_main,
    textvariable=sistema_var,
    bootstyle="secondary"
)
sistema_dropdown['values'] = ["RevendaMais", "AutoConf", "Boom Sistemas"]
sistema_dropdown.pack(pady=5, fill="x", padx=10)

# Dropdown Categoria
tb.Label(frame_main, text="Escolha a funcionalidade:", font=("Arial", 12)).pack(pady=5)
categoria_var = tb.StringVar()
categoria_dropdown = tb.Combobox(
    frame_main,
    textvariable=categoria_var,
    bootstyle="secondary"
)
categoria_dropdown['values'] = ["Estoque de veículos", "Clientes", "Oportunidades"]
categoria_dropdown.pack(pady=5, fill="x", padx=10)

# Botão para prosseguir após selecionar sistema e categoria
botao_continuar = tb.Button(
    frame_main,
    text="Continuar",
    bootstyle="success",
    command=mostrar_opcoes_de_planilha
)
botao_continuar.pack(pady=10)

# Frame para seleção das planilhas
frame_planilhas = tb.Frame(root)

# Botão para selecionar planilha 1
botao_1 = tb.Button(
    frame_planilhas,
    text="Selecionar Planilha Estoque",
    bootstyle="secondary",
    command=selecionar_primeira_planilha
)
botao_1.pack(pady=5)

label_1 = tb.Label(frame_planilhas, text="Nenhuma planilha selecionada", font=("Arial", 10))
label_1.pack(pady=5)

# Botão para selecionar planilha 2
botao_2 = tb.Button(
    frame_planilhas,
    text="Selecionar Planilha Vendido",
    bootstyle="secondary",
    command=selecionar_segunda_planilha
)
botao_2.pack(pady=5)

label_2 = tb.Label(frame_planilhas, text="Nenhuma planilha selecionada", font=("Arial", 10))
label_2.pack(pady=5)

# Botão para converter
botao_converter = tb.Button(
    frame_planilhas,
    text="Converter",
    bootstyle="success",
    command=converter
)
botao_converter.pack(pady=10)

root.mainloop()
