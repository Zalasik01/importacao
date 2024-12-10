import ttkbootstrap as tb
from tkinter import filedialog, messagebox
import os
import subprocess


def selecionar_planilha():
    """Abre um diálogo para selecionar a planilha de origem."""
    file_path = filedialog.askopenfilename(
        title="Selecione a Planilha de Origem",
        filetypes=[("Planilhas Excel", "*.xlsx *.xls")]
    )
    return file_path


def chamar_subaplicativo(sistema, categoria, file_path):
    """Chama o aplicativo correto com base no sistema e categoria."""
    try:
        # Caminho do script
        script_path = os.path.join(os.getcwd(), sistema, categoria, f"{categoria.lower()}.py")

        # Debug - Verificar caminho real
        print(f"Tentando acessar o script: {script_path}")
        if not os.path.exists(script_path):
            raise FileNotFoundError(f"Script não encontrado: {script_path}")

        # Executar o script
        subprocess.run(["python", script_path, file_path], check=True)

        # Mensagem no próprio aplicativo avisando que a conversão foi concluída
        messagebox.showinfo("Concluído", "Conversão da planilha concluída com sucesso!")
    except FileNotFoundError as e:
        messagebox.showerror("Erro", str(e))
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Erro", f"Erro ao executar o script: {e}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro inesperado: {e}")


def iniciar_processo():
    """Função chamada ao clicar no botão."""
    sistema = sistema_var.get()
    categoria = categoria_var.get()

    if not sistema or not categoria:
        messagebox.showwarning("Aviso", "Por favor, selecione todas as opções!")
        return

    file_path = selecionar_planilha()
    if not file_path:
        return

    if categoria == "Estoque de veículos":
        chamar_subaplicativo(sistema, "Estoque_Veiculos", file_path)
    else:
        messagebox.showwarning("Aviso", "Essa funcionalidade ainda não está implementada.")


# Configuração da interface
root = tb.Window(themename="cosmo")
root.title("Conversor de Planilhas")
root.geometry("500x400")

# Centraliza a janela na tela
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = 500
window_height = 400
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

frame_main = tb.Frame(root)
frame_main.pack(pady=10, fill="both", expand=True)

# Dropdown para seleção do sistema
tb.Label(frame_main, text="De qual sistema a planilha de origem é:", font=("Arial", 12)).pack(pady=5)
sistema_var = tb.StringVar()
sistema_dropdown = tb.Combobox(
    frame_main,
    textvariable=sistema_var,
    bootstyle="secondary"
)
sistema_dropdown['values'] = ["RevendaMais", "AutoConf", "Boom Sistemas"]
sistema_dropdown.pack(pady=5, fill="x", padx=10)

# Dropdown para seleção da funcionalidade
tb.Label(frame_main, text="Você deseja importar o quê:", font=("Arial", 12)).pack(pady=5)
categoria_var = tb.StringVar()
categoria_dropdown = tb.Combobox(
    frame_main,
    textvariable=categoria_var,
    bootstyle="secondary"
)
categoria_dropdown['values'] = ["Estoque de veículos", "Clientes", "Oportunidades"]
categoria_dropdown.pack(pady=5, fill="x", padx=10)

# Botão para iniciar o processo
button_start = tb.Button(
    frame_main,
    text="Iniciar Processo",
    bootstyle="success",
    command=iniciar_processo
)
button_start.pack(pady=20)

root.mainloop()