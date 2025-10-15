import os
import json
import requests
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox

# 🔢 Versão local atual do sistema
LOCAL_VERSION = "1.0.0"

# 🌐 URL remota para buscar a versão mais recente
REMOTE_VERSION_URL = "https://raw.githubusercontent.com/gabriel-araujo-git/sistema-orcamentos/main/version.json"


# 🧩 Função principal de verificação
def verificar_atualizacao_visual(parent=None):
    """Verifica se há uma nova versão e pergunta ao usuário se deseja atualizar."""
    def verificar():
        try:
            r = requests.get(REMOTE_VERSION_URL, timeout=10)
            if r.status_code != 200:
                return

            dados = json.loads(r.text)
            versao_remota = dados.get("versao")
            url_exe = dados.get("url")

            if versao_remota and url_exe and versao_remota != LOCAL_VERSION:
                resposta = messagebox.askyesno(
                    "Atualização disponível",
                    f"Uma nova versão ({versao_remota}) está disponível.\nDeseja atualizar agora?"
                )
                if resposta:
                    threading.Thread(target=baixar_com_progresso, args=(url_exe,), daemon=True).start()
        except Exception as e:
            print(f"Erro ao verificar atualização: {e}")

    threading.Thread(target=verificar, daemon=True).start()


# 💾 Função de download com barra de progresso
def baixar_com_progresso(url_exe):
    """Baixa a nova versão com feedback visual em tempo real."""
    try:
        janela = tk.Toplevel()
        janela.title("Atualizando EletroFlow")
        janela.geometry("380x140")
        janela.resizable(False, False)
        janela.attributes('-topmost', True)

        tk.Label(janela, text="Baixando nova versão...", font=("Segoe UI", 10)).pack(pady=10)
        barra = ttk.Progressbar(janela, length=300, mode='determinate')
        barra.pack(pady=10)
        progresso_label = tk.Label(janela, text="0%", font=("Segoe UI", 9))
        progresso_label.pack()

        janela.update_idletasks()

        r = requests.get(url_exe, stream=True, timeout=60)
        total = int(r.headers.get('content-length', 0))
        destino = os.path.join(os.getcwd(), os.path.basename(url_exe))

        with open(destino, "wb") as f:
            baixado = 0
            for chunk in r.iter_content(1024):
                if chunk:
                    f.write(chunk)
                    baixado += len(chunk)
                    percentual = int((baixado / total) * 100)
                    barra['value'] = percentual
                    progresso_label.config(text=f"{percentual}%")
                    janela.update_idletasks()

        tk.Label(janela, text="Download concluído!", font=("Segoe UI", 10)).pack(pady=5)
        janela.update()
        janela.after(1500, janela.destroy)

        # Abre o novo executável automaticamente
        subprocess.run(["explorer", destino])

    except Exception as e:
        messagebox.showerror("Erro de atualização", f"Ocorreu um erro: {e}")


# 🔄 Função simples (modo automático, sem interface)
def verificar_atualizacao_silenciosa():
    """Verifica silenciosamente no início da aplicação."""
    try:
        r = requests.get(REMOTE_VERSION_URL, timeout=10)
        if r.status_code != 200:
            return
        dados = json.loads(r.text)
        versao_remota = dados.get("versao")
        url_exe = dados.get("url")
        if versao_remota and url_exe and versao_remota != LOCAL_VERSION:
            print(f"Nova versão detectada: {versao_remota}")
            threading.Thread(target=baixar_com_progresso, args=(url_exe,), daemon=True).start()
    except Exception as e:
        print(f"Erro silencioso ao verificar atualização: {e}")


# 🧭 Execução direta (teste isolado)
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Oculta janela principal
    verificar_atualizacao_visual(root)
    root.mainloop()
