import argparse
import time
from pathlib import Path

import pandas as pd
import pyperclip

from pywinauto import Desktop


TITULO_JANELA = (
    "Gerador de Relatórios "
    "( Versão: 4.0 ) - Recursos Humanos e Folha de Pagamento"
)


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


def encontrar_campo_funcionario(janela):
    abas = janela.descendants(
        class_name="TcxTabSheet"
    )

    for aba in abas:
        try:
            if aba.window_text() != "Funcionário":
                continue

            combos = aba.descendants(
                class_name="TelLookupCombo"
            )

            if combos:
                return combos[0]

        except Exception:
            pass

    return None


def encontrar_campo_lancamento(janela):
    abas = janela.descendants(
        class_name="TcxTabSheet"
    )

    # Primeiro encontramos a aba Funcionário
    aba_funcionario = None

    for aba in abas:
        try:
            if aba.window_text() == "Funcionário":
                aba_funcionario = aba
                break
        except Exception:
            pass

    # Procuramos todos os TelLookupCombo da janela
    combos = janela.descendants(
        class_name="TelLookupCombo"
    )

    for combo in combos:
        try:
            # Se o combo estiver dentro da aba Funcionário,
            # não é o campo de lançamento.
            if aba_funcionario is not None:
                try:
                    if combo.is_child(aba_funcionario):
                        continue
                except Exception:
                    pass

            # Verificamos o painel pai imediato
            pais = combo.parent()

            if pais is not None:
                combos_pai = pais.descendants(
                    class_name="TelLookupCombo"
                )

                if len(combos_pai) == 1:
                    return combo

        except Exception:
            pass

    return None


def preencher_matricula(matricula):
    janela = obter_janela()

    controle = encontrar_campo_funcionario(janela)

    if controle is None:
        raise RuntimeError(
            "Campo de matrícula não encontrado."
        )

    print(
        f"Preenchendo matrícula {matricula} "
        f"(handle atual: {controle.handle})"
    )

    controle.set_focus()

    pyperclip.copy(str(matricula))

    controle.type_keys("^a")
    controle.type_keys("^v")


def preencher_lancamento(valor):
    janela = obter_janela()

    controle = encontrar_campo_lancamento(janela)

    if controle is None:
        raise RuntimeError(
            "Campo de lançamento não encontrado."
        )

    print(
        f"Preenchendo lançamento {valor} "
        f"(handle atual: {controle.handle})"
    )

    controle.set_focus()

    pyperclip.copy(str(valor))

    controle.type_keys("^a")
    controle.type_keys("^v")


def pressionar_enter():
    janela = obter_janela()

    controle = encontrar_campo_funcionario(janela)

    if controle is None:
        raise RuntimeError(
            "Campo de matrícula não encontrado."
        )

    controle.set_focus()
    controle.type_keys("{ENTER}")


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

    # Teste inicial
    preencher_lancamento("682")

    # Depois do teste, processa a planilha
    processar_planilha(args.planilha)


if __name__ == "__main__":
    main()