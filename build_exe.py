"""
GERADOR DO EXECUTÁVEL (.exe)
Empacota a interface, o motor de auditoria e a base de exemplo numa pasta
autocontida que pode ser enviada para qualquer pessoa — não exige Python.

Modo "onedir" (pasta): abertura instantânea, pois nada é descompactado em disco
a cada execução. Para distribuir, envie a PASTA inteira (dist/Auditoria_Matrizes)
compactada em .zip — o usuário extrai e roda o .exe de dentro dela.

Uso:  python build_exe.py
Saída: dist/Auditoria_Matrizes/Auditoria_Matrizes.exe (+ arquivos de apoio)
"""

import os
import shutil
import subprocess
import sys

NOME_APP = "Auditoria_Matrizes"
PASTA = os.path.dirname(os.path.abspath(__file__))
SEPARADOR = ";" if sys.platform.startswith("win") else ":"


def main():
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("PyInstaller não encontrado. Instale com:  pip install pyinstaller")
        return 1

    for pasta in ("build", "dist"):
        caminho = os.path.join(PASTA, pasta)
        if os.path.isdir(caminho):
            shutil.rmtree(caminho, ignore_errors=True)

    comando = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onedir",                  # pasta autocontida -> abertura instantânea
        "--windowed",                # sem console preto atrás da janela
        "--name", NOME_APP,
        "--add-data", f"Dados_Exportacao_CSV.csv{SEPARADOR}.",
        "--collect-all", "ortools",  # o CP-SAT traz bibliotecas nativas
        "--hidden-import", "PLANEJADOR",
        "INTERFACE.py",
    ]

    print("Empacotando... (a geração leva alguns minutos)\n")
    resultado = subprocess.run(comando, cwd=PASTA)
    if resultado.returncode != 0:
        print("\nFalha ao gerar o executável.")
        return resultado.returncode

    pasta_app = os.path.join(PASTA, "dist", NOME_APP)
    destino = os.path.join(pasta_app, f"{NOME_APP}.exe")

    # coloca o guia de uso na raiz da pasta, ao lado do .exe
    leia_me = os.path.join(PASTA, "LEIA-ME.txt")
    if os.path.exists(leia_me) and os.path.isdir(pasta_app):
        shutil.copy2(leia_me, os.path.join(pasta_app, "LEIA-ME.txt"))
        print("LEIA-ME.txt copiado para a pasta do aplicativo.")

    if os.path.exists(destino):
        total = sum(os.path.getsize(os.path.join(raiz, f))
                    for raiz, _, arquivos in os.walk(pasta_app) for f in arquivos)
        print(f"\nAplicativo gerado em: {pasta_app}  ({total / (1024 * 1024):.0f} MB no total)")
        print("Abertura instantânea. Para enviar, compacte a PASTA inteira em .zip;")
        print("o destinatário extrai e executa Auditoria_Matrizes.exe de dentro dela.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
