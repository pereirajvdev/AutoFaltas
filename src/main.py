import argparse
import time
from pathlib import Path

import pandas as pd
from pywinauto import Desktop


TITULO_JANELA = (
    "Gerador de Relatórios "
    "( Versão: 4.0 ) - Recursos Humanos e Folha de Pagamento"
)

HANDLE_MATRICULA = 330106


def ler_planilha(caminho):
    df = pd.read_excel(caminho)

    if len(df.columns) < 2:
        raise ValueError(
            "A planilha deve possuir duas colunas: matrícula e nome."
        )

    dados = []

    for _, linha in df.iterrows():
        matricula = str(linha.iloc[0]).strip()
        nome = str(linha.iloc[1]).strip()

        if not matricula:
            continue

        dados.append((matricula, nome))

    return dados


def obter_janela():
    desktop = Desktop(backend="win32")

    janela = desktop.window(title=TITULO_JANELA)

    if not janela.exists(timeout=5):
        raise RuntimeError(
            "Janela do Gerador de Relatórios não encontrada."
        )

    return janela


def preencher_matricula(matricula):
    janela = obter_janela()

    controle = janela.child_window(
        handle=HANDLE_MATRICULA
    )

    if not controle.exists(timeout=2):
        raise RuntimeError(
            f"Controle da matrícula não encontrado. "
            f"Handle: {HANDLE_MATRICULA}"
        )

    print(f"Preenchendo matrícula: {matricula}")

    controle.set_focus()
    controle.type_keys("^a")
    controle.type_keys(str(matricula))

    print(
        "Valor no campo:",
        repr(controle.window_text())
    )


def pressionar_enter():
    janela = obter_janela()

    controle = janela.child_window(
        handle=HANDLE_MATRICULA
    )

    controle.set_focus()
    controle.type_keys("{ENTER}")

    print("Enter pressionado.")


def processar_planilha(caminho):
    dados = ler_planilha(caminho)

    print(f"Total de registros: {len(dados)}")
    print()

    for matricula, nome in dados:
        print(f"Processando: {matricula} - {nome}")

        preencher_matricula(matricula)
        pressionar_enter()

        time.sleep(1)


def main():
    parser = argparse.ArgumentParser(
        description="Bot de automação do sistema de Recursos Humanos."
    )

    parser.add_argument(
        "planilha",
        type=Path,
        help="Caminho da planilha Excel."
    )

    args = parser.parse_args()

    if not args.planilha.exists():
        parser.error(
            f"Planilha não encontrada: {args.planilha}"
        )

    if not args.planilha.is_file():
        parser.error(
            f"O caminho informado não é um arquivo: {args.planilha}"
        )

    processar_planilha(args.planilha)


if __name__ == "__main__":
    main()