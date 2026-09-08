import webview
import subprocess
import sys
import os


class VIX:

    def _init_(self):

        self.pasta_principal = os.path.dirname(
            os.path.abspath(__file__)
        )

        print("Pasta principal:")
        print(self.pasta_principal)


    # ==========================================================
    # EXECUTAR RESERVA
    # ==========================================================

    def executar_reserva(self):

        try:

            caminho_script = os.path.join(
                self.pasta_principal,
                "reserva.py"
            )

            print("Tentando executar:")
            print(caminho_script)

            if not os.path.isfile(caminho_script):

                return {
                    "sucesso": False,
                    "mensagem":
                        f"Arquivo não encontrado:\n{caminho_script}"
                }

            subprocess.Popen(
                [sys.executable, caminho_script],
                cwd=self.pasta_principal
            )

            return {
                "sucesso": True,
                "mensagem":
                    "reserva.py iniciado com sucesso."
            }

        except Exception as erro:

            print("ERRO:", erro)

            return {
                "sucesso": False,
                "mensagem":
                    f"Erro ao iniciar reserva.py: {erro}"
            }


    # ==========================================================
    # EXECUTAR REMESSA TERRESTRE
    # ==========================================================

    def executar_remessa(self):

        try:

            caminho_script = os.path.join(
                self.pasta_principal,
                "remessa-terrestre.py"
            )

            print("=================================")
            print("EXECUTANDO REMESSA TERRESTRE")
            print("=================================")

            print("Caminho do script:")
            print(caminho_script)

            # Verifica se existe
            if not os.path.isfile(caminho_script):

                print("ARQUIVO NÃO ENCONTRADO!")

                return {
                    "sucesso": False,
                    "mensagem":
                        f"Arquivo não encontrado:\n{caminho_script}"
                }

            print("Arquivo encontrado!")

            # Executa o Python
            processo = subprocess.Popen(
                [sys.executable, caminho_script],
                cwd=self.pasta_principal
            )

            print(
                f"remessa-terrestre.py iniciado. "
                f"PID: {processo.pid}"
            )

            return {
                "sucesso": True,
                "mensagem":
                    "remessa-terrestre.py iniciado com sucesso."
            }

        except Exception as erro:

            print("ERRO AO EXECUTAR REMESSA:")
            print(erro)

            return {
                "sucesso": False,
                "mensagem":
                    f"Erro ao iniciar remessa-terrestre.py:\n{erro}"
            }


# ==========================================================
# CAMINHO DA PASTA PRINCIPAL
# ==========================================================

pasta_principal = os.path.dirname(
    os.path.abspath(__file__)
)


# ==========================================================
# CAMINHO DO HTML
# ==========================================================

caminho_html = os.path.join(
    pasta_principal,
    "interface",
    "index.html"
)


# ==========================================================
# VERIFICAR HTML
# ==========================================================

if not os.path.isfile(caminho_html):

    raise FileNotFoundError(
        f"index.html não encontrado:\n{caminho_html}"
    )


print("HTML encontrado:")
print(caminho_html)


# ==========================================================
# CRIAR JANELA
# ==========================================================

webview.create_window(
    "Planejamento e Controle - APOIO",
    caminho_html,
    js_api=VIX(),
    width=700,
    height=600,
    resizable=True
)


# ==========================================================
# INICIAR PYWEBVIEW
# ==========================================================

webview.start(debug=True)