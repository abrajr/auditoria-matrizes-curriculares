"""
GERADOR DO APLICATIVO (Windows .exe / macOS .app)
Empacota a interface, o motor de auditoria e a base de exemplo num aplicativo
autocontido que pode ser enviado para qualquer pessoa — não exige Python.

O MESMO script funciona nos dois sistemas: rode-o NO sistema de destino.
  - No Windows gera:  dist/Auditoria_Matrizes/  (pasta com Auditoria_Matrizes.exe)
  - No macOS   gera:  dist/Auditoria_Matrizes-mac/  (com Auditoria_Matrizes.app)

IMPORTANTE: o PyInstaller NÃO faz compilação cruzada. Um aplicativo de Mac só
pode ser gerado rodando este script em um Mac (localmente ou via GitHub Actions).

Uso:  python build_exe.py
"""

import os
import shutil
import subprocess
import sys

NOME_APP = "Auditoria_Matrizes"
PASTA = os.path.dirname(os.path.abspath(__file__))

EH_MAC = sys.platform == "darwin"
EH_WIN = sys.platform.startswith("win")
SEPARADOR = ";" if EH_WIN else ":"

TEXTO_ABRIR_MAC = """\
COMO ABRIR NO macOS
===================

Como o aplicativo não possui assinatura da Apple, o macOS bloqueia a primeira
abertura (recurso "Gatekeeper"). Faça UMA vez:

  Opção 1 (mais simples):
    - Clique com o BOTÃO DIREITO (ou Control+clique) sobre Auditoria_Matrizes.app
    - Escolha "Abrir"
    - Na janela de aviso, clique novamente em "Abrir"
    (nas próximas vezes, basta um duplo clique normal)

  Opção 2 (se a opção 1 não liberar):
    - Abra o aplicativo Terminal
    - Digite  xattr -cr   (com um espaço no final) e arraste o Auditoria_Matrizes.app
      para a janela do Terminal; tecle Enter
    - Depois é só dar duplo clique no aplicativo

O relatório é gravado na sua pasta Documentos (Relatorio_Auditoria_Matrizes.xlsx).
Diagnóstico, se algo falhar:  abra o Terminal e rode
  /caminho/para/Auditoria_Matrizes.app/Contents/MacOS/Auditoria_Matrizes --autoteste
"""


def _rodar_pyinstaller():
    comando = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onedir",                  # pasta autocontida -> abertura instantânea
        "--windowed",                # sem console; no macOS gera o pacote .app
        "--name", NOME_APP,
        "--add-data", f"Dados_Exportacao_CSV.csv{SEPARADOR}.",
        "--collect-all", "ortools",  # o CP-SAT traz bibliotecas nativas
        "--hidden-import", "PLANEJADOR",
        "INTERFACE.py",
    ]
    print("Empacotando... (a geração leva alguns minutos)\n")
    return subprocess.run(comando, cwd=PASTA).returncode


def _tamanho_mb(caminho):
    total = 0
    for raiz, _, arquivos in os.walk(caminho):
        for f in arquivos:
            try:
                total += os.path.getsize(os.path.join(raiz, f))
            except OSError:
                pass
    return total / (1024 * 1024)


def _finalizar_windows():
    pasta_app = os.path.join(PASTA, "dist", NOME_APP)
    exe = os.path.join(pasta_app, f"{NOME_APP}.exe")
    if not os.path.exists(exe):
        print("Falha: executável não encontrado em", exe)
        return 1
    leia_me = os.path.join(PASTA, "LEIA-ME.txt")
    if os.path.exists(leia_me):
        shutil.copy2(leia_me, os.path.join(pasta_app, "LEIA-ME.txt"))
        print("LEIA-ME.txt copiado para a pasta do aplicativo.")
    print(f"\nAplicativo (Windows) gerado em: {pasta_app}  ({_tamanho_mb(pasta_app):.0f} MB)")
    print("Para enviar, compacte a PASTA inteira em .zip; o destinatário extrai e")
    print("executa Auditoria_Matrizes.exe de dentro dela. Abertura instantânea.")
    return 0


def _finalizar_mac():
    app = os.path.join(PASTA, "dist", f"{NOME_APP}.app")
    if not os.path.exists(app):
        print("Falha: pacote .app não encontrado em", app)
        return 1
    # monta uma pasta de distribuição limpa: o .app + guias ao lado (fora do .app,
    # para não quebrar a estrutura do pacote)
    dist_mac = os.path.join(PASTA, "dist", f"{NOME_APP}-mac")
    if os.path.isdir(dist_mac):
        shutil.rmtree(dist_mac, ignore_errors=True)
    os.makedirs(dist_mac)
    shutil.copytree(app, os.path.join(dist_mac, f"{NOME_APP}.app"), symlinks=True)
    leia_me = os.path.join(PASTA, "LEIA-ME.txt")
    if os.path.exists(leia_me):
        shutil.copy2(leia_me, os.path.join(dist_mac, "LEIA-ME.txt"))
    with open(os.path.join(dist_mac, "COMO-ABRIR-NO-MAC.txt"), "w", encoding="utf-8") as f:
        f.write(TEXTO_ABRIR_MAC)
    print(f"\nAplicativo (macOS) gerado em: {dist_mac}  ({_tamanho_mb(dist_mac):.0f} MB)")
    print("Para enviar, compacte a PASTA inteira em .zip (use o Finder ou 'ditto').")
    print("Leia COMO-ABRIR-NO-MAC.txt: o macOS bloqueia a 1ª abertura de apps não assinados.")
    return 0


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

    if _rodar_pyinstaller() != 0:
        print("\nFalha ao gerar o aplicativo.")
        return 1

    return _finalizar_mac() if EH_MAC else _finalizar_windows()


if __name__ == "__main__":
    sys.exit(main())
