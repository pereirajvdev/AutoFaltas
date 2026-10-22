
import re
import argparse
import time
from pathlib import Path

import pandas as pd
import pyperclip

from pywinauto import Desktop
from pywinauto import mouse


TITULO_JANELA = (
    "Gerador de Relatórios "
    "( Versão: 4.0 ) - Recursos Humanos e Folha de Pagamento"
)

DIRETORIO_SERVIDORES = Path(
    r"C:\Users\Joao Castro\Desktop\JOAO\Servidores"
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

        if matricula == "nan":
            continue

        if matricula.endswith(".0"):
            matricula = matricula[:-2]

        nome = str(linha.iloc[1]).strip()

        if not matricula:
            continue

        print(f"Matrícula lida: {matricula} - {nome}")

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


def encontrar_botao_ok(janela):
    botoes = janela.descendants(
        class_name="TBitBtn"
    )

    for botao in botoes:
        try:
            if botao.window_text() == "&OK":
                return botao
        except Exception:
            pass

    return None


def clicar_ok():
    janela = obter_janela()

    botao = encontrar_botao_ok(janela)

    if botao is None:
        raise RuntimeError(
            "Botão OK não encontrado."
        )

    print(
        f"Clicando em OK (handle atual: {botao.handle})"
    )

    botao.click()


def get_screen_name():
    desktop = Desktop(backend="win32")

    try:
        janelas = desktop.windows()
    except Exception as erro:
        print(
            f"[AVISO] Falha ao enumerar janelas: {erro}"
        )
        return "Desconhecida"

    for janela in janelas:
        try:
            titulo = janela.window_text()
            classe = janela.class_name()

            if titulo == TITULO_JANELA:
                return "Gerador"

            if (
                titulo == "Lancamentos_Especificos__por_Ano"
                and classe == "TfrmPreview"
            ):
                return "Preview"

        except Exception:
            # A janela pode ter sido destruída durante a consulta.
            continue

    return "Desconhecida"


def processar_resultado(
    nome_funcionario,
    ano,
    codigo_lancamento,
    pasta_saida,
):
    print("Aguardando resultado...")

    inicio = time.monotonic()
    timeout = 120

    while time.monotonic() - inicio < timeout:
        tela_atual = get_screen_name()

        if tela_atual == "Preview":
            print("Relatório aberto no preview.")

            preview = None
            desktop = Desktop(backend="win32")

            try:
                for janela in desktop.windows():
                    try:
                        if (
                            janela.window_text()
                            == "Lancamentos_Especificos__por_Ano"
                            and janela.class_name() == "TfrmPreview"
                        ):
                            preview = janela
                            break
                    except Exception:
                        continue

            except Exception as erro:
                print(
                    "[AVISO] Falha ao localizar o preview: "
                    f"{erro}"
                )

            clicar_botao_preview()

            salvo = imprimir_como_pdf(
                nome_funcionario,
                ano,
                codigo_lancamento,
                pasta_saida,
            )

            if salvo:
                print("Fechando preview...")

                if preview is not None:
                    try:
                        if preview.is_valid():
                            preview.close()
                            time.sleep(0.5)

                    except Exception as erro:
                        print(
                            "[AVISO] Não foi possível fechar "
                            f"o preview normalmente: {erro}"
                        )

                # Tenta localizar novamente, pois o handle
                # anterior pode ter deixado de ser válido.
                fechar_preview()

                return True

            print(
                "[ERRO] A impressão não foi concluída "
                "com sucesso."
            )
            return False

        time.sleep(0.2)

    raise TimeoutError(
        f"O relatório da matrícula/ano {ano} "
        f"não apareceu no preview em {timeout} segundos."
    )


def clicar_botao_preview():
    x = 306
    y = 39

    print(f"Clicando no botão em ({x}, {y})...")

    mouse.click(coords=(x, y))

    print("Botão clicado.")


def imprimir_como_pdf(
    nome_funcionario,
    ano,
    codigo_lancamento,
    pasta_saida,
):
    desktop = Desktop(backend="win32")

    print("Aguardando janela de impressão...")

    janela = None

    while janela is None:
        janelas = desktop.windows(
            class_name="TfrPrintForm"
        )

        for item in janelas:
            try:
                if item.is_visible():
                    janela = item
                    break
            except Exception:
                pass

        if janela is None:
            time.sleep(0.2)

    print("Janela de impressão encontrada.")
    print("Handle:", janela.handle)

    combo_impressora = None

    for controle in janela.descendants():
        try:
            if controle.class_name() != "TComboBox":
                continue

            itens = controle.texts()

            if "Microsoft Print to PDF" in itens:
                combo_impressora = controle
                break

        except Exception:
            pass

    if combo_impressora is None:
        print("ComboBox da impressora não encontrada.")
        return False

    print("Selecionando Microsoft Print to PDF...")

    try:
        combo_impressora.select(
            "Microsoft Print to PDF"
        )
    except Exception as e:
        print("Erro ao selecionar impressora:", e)
        return False

    time.sleep(0.5)

    botao_ok = None

    for controle in janela.descendants():
        try:
            if (
                controle.class_name() == "TButton"
                and controle.window_text() == "OK"
            ):
                botao_ok = controle
                break
        except Exception:
            pass

    if botao_ok is None:
        print("Botão OK da impressão não encontrado.")
        return False

    print("Clicando em OK da impressão...")

    botao_ok.click()

    return salvar_pdf(
        nome_funcionario,
        ano,
        codigo_lancamento,
        pasta_saida,
    )


def salvar_pdf(
    nome_funcionario,
    ano,
    codigo_lancamento,
    pasta_saida,
):
    desktop = Desktop(backend="win32")

    print("Aguardando janela para salvar PDF...")

    janela = None

    while janela is None:
        for item in desktop.windows():
            try:
                if (
                    item.is_visible()
                    and item.class_name() == "#32770"
                ):
                    janela = item
                    break
            except Exception:
                pass

        if janela is None:
            time.sleep(0.2)

    print("Janela de salvamento encontrada.")
    print("Título:", janela.window_text())

    pasta_saida.mkdir(parents=True, exist_ok=True)

    nome_arquivo = f"FALTAS {codigo_lancamento} {ano}.pdf"
    caminho_pdf = pasta_saida / nome_arquivo

    print("Pasta de saída:", pasta_saida)
    print("Arquivo:", caminho_pdf)

    campo_nome = None

    for controle in janela.descendants():
        try:
            if controle.class_name() != "Edit":
                continue

            if controle.rectangle().width() > 500:
                campo_nome = controle
                break

        except Exception:
            pass

    if campo_nome is None:
        print("Campo do nome do arquivo não encontrado.")
        return False

    campo_nome.set_focus()
    pyperclip.copy(str(caminho_pdf))
    campo_nome.type_keys("^a")
    campo_nome.type_keys("^v")

    botao_salvar = None

    for controle in janela.descendants():
        try:
            if (
                controle.class_name() == "Button"
                and controle.window_text() == "Sa&lvar"
            ):
                botao_salvar = controle
                break
        except Exception:
            pass

    if botao_salvar is None:
        print("Botão Salvar não encontrado.")
        return False

    botao_salvar.click()

    print("Comando de salvar enviado.")

    confirmar_sobrescrita_pdf()

    time.sleep(0.5)

    if caminho_pdf.exists():
        print("PDF salvo:", caminho_pdf)
        return True

    print("PDF não encontrado após salvar:", caminho_pdf)
    return False


def fechar_preview():
    desktop = Desktop(backend="win32")

    try:
        janelas = desktop.windows()
    except Exception:
        print("Não foi possível listar as janelas no momento.")
        return False

    for janela in janelas:
        try:
            if (
                janela.window_text()
                == "Lancamentos_Especificos__por_Ano"
                and janela.class_name() == "TfrmPreview"
            ):
                print(
                    "Fechando Lancamentos_Especificos__por_Ano..."
                )

                janela.close()

                time.sleep(0.5)

                print("Preview fechado.")

                return True

        except Exception:
            # Ignora janelas que foram destruídas durante a busca
            continue

    print("Preview não encontrado para fechar.")

    return False


def interpretar_intervalo_anos(valor):
    try:
        inicio, fim = valor.split("-")
        inicio = int(inicio)
        fim = int(fim)
    except ValueError:
        raise ValueError(
            "O intervalo de anos deve estar no formato YYYY-YYYY. "
            "Exemplo: 2019-2026"
        )

    if inicio > fim:
        raise ValueError(
            "O primeiro ano deve ser menor ou igual ao segundo."
        )

    return inicio, fim


def encontrar_campo_ano(janela):
    spin_edits = janela.descendants(class_name="TcxSpinEdit")

    for spin in spin_edits:
        try:
            filhos = spin.descendants(
                class_name="TcxCustomInnerTextEdit"
            )

            for filho in filhos:
                texto = filho.window_text().strip()

                if texto.isdigit() and len(texto) == 4:
                    return filho

        except Exception:
            pass

    return None


def preencher_ano(ano):
    janela = obter_janela()
    controle = encontrar_campo_ano(janela)

    if controle is None:
        raise RuntimeError("Campo de ano não encontrado.")

    print(
        f"Preenchendo ano {ano} "
        f"(handle atual: {controle.handle})"
    )

    controle.set_focus()

    pyperclip.copy(str(ano))
    controle.type_keys("^a")
    controle.type_keys("^v")

    controle.type_keys("{ENTER}")


def selecionar_funcionarios(dados):
    print(f"Selecionando {len(dados)} funcionários...")

    for matricula, nome in dados:
        print(f"Selecionando: {matricula} - {nome}")

        preencher_matricula(matricula)
        pressionar_enter()

        time.sleep(0.3)


def processar_anos(
    caminho_saida,
    inicio_ano,
    fim_ano,
    codigo_lancamento,
    nome_funcionario,
):
    for ano in range(inicio_ano, fim_ano + 1):

        print()
        print("=" * 60)
        print(
            f"PROCESSANDO ANO {ano} - "
            f"LANÇAMENTO {codigo_lancamento}"
        )
        print("=" * 60)

        preencher_ano(ano)

        preencher_lancamento(codigo_lancamento)

        clicar_ok()

        salvo = processar_resultado(
            nome_funcionario,
            ano,
            codigo_lancamento,
            caminho_saida,
        )

        caminho_pdf = (
            caminho_saida
            / f"FALTAS {codigo_lancamento} {ano}.pdf"
        )

        if salvo:
            if caminho_pdf.exists():
                print()
                print(
                    f"[OK] PDF CONFIRMADO:"
                )
                print(caminho_pdf)
                print(
                    f"Tamanho: "
                    f"{caminho_pdf.stat().st_size:,} bytes"
                )
            else:
                print()
                print(
                    "[ERRO] O sistema informou que o PDF foi salvo,"
                )
                print(
                    "mas o arquivo NÃO foi encontrado:"
                )
                print(caminho_pdf)

        else:
            print()
            print(
                f"[ERRO] Falha ao gerar "
                f"FALTAS {codigo_lancamento} {ano}.pdf"
            )


def obter_nome_funcionario(janela):
    listboxes = janela.descendants(class_name="TListBox")

    for listbox in listboxes:
        try:
            for item in listbox.texts():
                item = item.strip()

                if " - " not in item:
                    continue

                codigo, nome = item.split(" - ", 1)

                if codigo.strip().isdigit() and nome.strip():
                    return nome.strip()

        except Exception:
            pass

    return None


def confirmar_sobrescrita_pdf():
    print("Verificando confirmação de sobrescrita...")

    time.sleep(0.5)

    desktop = Desktop(backend="win32")

    for janela in desktop.windows():
        try:
            if (
                janela.is_visible()
                and janela.class_name() == "#32770"
                and "confirmar" in janela.window_text().lower()
            ):
                janela.set_focus()
                janela.type_keys("{LEFT}")
                time.sleep(0.2)
                janela.type_keys("{ENTER}")

                print("Confirmação de sobrescrita enviada.")
                time.sleep(0.5)

                return True

        except Exception:
            continue

    print("Janela de confirmação não encontrada.")
    return False


def obter_pasta_servidor(nome_funcionario):
    nome = re.sub(r'[<>:"/\\|?*]', "_", nome_funcionario)
    nome = nome.strip().rstrip(".")

    if not nome:
        raise ValueError("O nome do servidor está vazio.")

    return DIRETORIO_SERVIDORES / nome


def main():
    parser = argparse.ArgumentParser(
        description="Bot de automação do sistema de Recursos Humanos."
    )

    parser.add_argument(
        "anos",
        help=(
            "Intervalo de anos no formato YYYY-YYYY. "
            "Exemplo: 2019-2026"
        )
    )

    parser.add_argument(
        "--mat",
        type=str,
        help="Matrícula do funcionário a ser selecionado."
    )

    parser.add_argument(
        "--planilha",
        type=Path,
        help="Caminho da planilha Excel com matrícula e nome."
    )

    parser.add_argument(
        "--out",
        type=Path,
        help="Pasta de saída dos PDFs gerados."
    )
    
    args = parser.parse_args()

    pasta_saida = (
        args.out
        if args.out is not None
        else DIRETORIO_SERVIDORES
    )

    pasta_saida = pasta_saida.resolve()

    try:
        inicio_ano, fim_ano = interpretar_intervalo_anos(
            args.anos
        )
    except ValueError as erro:
        parser.error(str(erro))

    print(f"Período: {inicio_ano} até {fim_ano}")

    # Seleciona os funcionários conforme os argumentos recebidos.
    if args.mat:
        dados = [(args.mat, "Matrícula informada manualmente")]

    elif args.planilha:
        if not args.planilha.exists():
            parser.error(
                f"Planilha não encontrada: {args.planilha}"
            )

        if not args.planilha.is_file():
            parser.error(
                f"O caminho informado não é um arquivo: {args.planilha}"
            )

        dados = ler_planilha(args.planilha)

    else:
        parser.error(
            "Informe --mat MATRICULA ou --planilha CAMINHO."
        )

    

    # Seleciona todos os funcionários uma única vez.
    selecionar_funcionarios(dados)

    # No modo --mat, identifica o servidor selecionado
    # e direciona os PDFs para a pasta individual dele.
    if args.mat:
        janela = obter_janela()
        nome_funcionario = obter_nome_funcionario(janela)

        if not nome_funcionario:
            raise RuntimeError(
                f"Não foi possível identificar o nome do servidor "
                f"da matrícula {args.mat}. "
                "Nenhum relatório será gerado."
            )

        pasta_saida = obter_pasta_servidor(nome_funcionario)
        pasta_saida.mkdir(parents=True, exist_ok=True)

        print()
        print("=" * 60)
        print(f"Matrícula: {args.mat}")
        print(f"Servidor: {nome_funcionario}")
        print(f"Pasta de saída: {pasta_saida}")
        print("=" * 60)

    # O relatório contém todos os funcionários selecionados.
    for codigo_lancamento in ["682", "619"]:
        print()
        print("#" * 60)
        print(f"INICIANDO LANÇAMENTO {codigo_lancamento}")
        print("#" * 60)

        processar_anos(
            pasta_saida,
            inicio_ano,
            fim_ano,
            codigo_lancamento,
            "",
        )


if __name__ == "__main__":
    main()