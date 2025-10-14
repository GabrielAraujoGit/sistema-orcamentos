import os
import json
import requests
import subprocess

LOCAL_VERSION = "1.0.0"
REMOTE_VERSION_URL = "https://raw.githubusercontent.com/gabriel-araujo-git/sistema-orcamentos/main/version.json"

def verificar_atualizacao():
    try:
        r = requests.get(REMOTE_VERSION_URL, timeout=10)
        if r.status_code != 200:
            print("Não foi possível verificar atualização.")
            return

        dados = json.loads(r.text)
        versao_remota = dados.get("versao")
        url_exe = dados.get("url")
        if versao_remota and url_exe and versao_remota != LOCAL_VERSION:
            print(f"Nova versão disponível: {versao_remota}")
            baixar_nova_versao(url_exe)
        else:
            print("Aplicativo já está atualizado.")
    except Exception as e:
        print(f"Erro ao verificar atualização: {e}")

def baixar_nova_versao(url_exe):
    try:
        print("Baixando nova versão...")
        r = requests.get(url_exe, timeout=30)
        if r.status_code != 200:
            print("Falha ao baixar a nova versão.")
            return

        nome_arquivo = os.path.basename(url_exe)
        destino = os.path.join(os.getcwd(), nome_arquivo)
        with open(destino, "wb") as f:
            f.write(r.content)
        print(f"Nova versão salva como {nome_arquivo}.")
        subprocess.run(["explorer", destino])
    except Exception as e:
        print(f"Erro durante download da nova versão: {e}")

if __name__ == "__main__":
    verificar_atualizacao()
