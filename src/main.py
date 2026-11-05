import subprocess
import pyautogui
import argparse
import ctypes
import re
import time
from contextlib import contextmanager
from ctypes import wintypes
from functools import wraps
from pathlib import Path


from pywinauto import Desktop, keyboard, mouse
from pywinauto.timings import Timings

import pyperclip


# ======================================================================
# CONSTANTES
# ======================================================================
TITULO_INICIAL = (
    "- Recursos Humanos e Folha de Pagamento "
    "[ Versão 3.02 - 17/06/2026 ] "
    "( 066 - RITA DE CASSIA CARVALHO BRANQUINHO )"
)
TITULO_PRINCIPAL = "PREFEITURA MUNICIPAL DE TRÊS CORAÇÕES - Recursos Humanos e Folha de Pagamento"
TITULO_JANELA_PRINCIPAL = (
    "Gerador de Relatórios "
    "( Versão: 4.0 ) - Recursos Humanos e Folha de Pagamento"
)
TITULO_LANCAMENTOS_ESPECIFICOS = "Lancamentos_Especificos__por_Ano"
TITULO_IMPRIMIR = "Imprimir"
TITULO_SOBRESCREVER = "Confirmar Salvar como"
TITULO_SAIDA_IMPRESSAO = "Salvar saída de impressão como"
TITULO_TESTE = "Teste"

DIRETORIO_SERVIDORES = Path(
    r"C:\Users\Joao Castro\Desktop\JOAO\Servidores"
)

SCRIPT = Path(
    r"C:\Users\Joao Castro\Documents\FolhaInicio\src\main.py"
)

TEMPO_LIMITE = 60
TEMPO_ARQUIVO_PDF = 15

# Intervalo de polling. Agora a verificação é barata (ctypes), então pode ser curto.
INTERVALO_VERIFICACAO = 0.05

TEMPO_FOLHA_INICIO = 180      # segundos para o FolhaInicio terminar
INTERVALO_TECLAS = 0.3        # pausa entre as teclas do atalho
INTERVALO_SETAS = 0.0001        # pausa entre as teclas de setas

ATALHO_FICHA = ("{VK_MENU}", "e", "0")   # Alt (sozinho) -> E -> 0

# Pausas de segurança. A busca antiga de janelas era lenta e, sem querer,
# funcionava como "pausa" entre as ações. Como agora é rápida, estas pausas
# curtas garantem que o Delphi processou a ação anterior.
# Se a automação estiver estável, você pode reduzi-las UMA POR VEZ.
PAUSA_APOS_COLAR_MATRICULA = 0.15      # entre colar a matrícula e dar ENTER
PAUSA_APOS_SELECIONAR_FUNCIONARIO = 0.3  # igual ao original
PAUSA_ENTRE_CAMPOS = 0.1               # ano -> lançamento -> botão OK
PAUSA_ANTES_CLIQUE_PREVIEW = 0.3       # preview aberto -> clique por coordenada
PAUSA_APOS_SELECIONAR_IMPRESSORA = 0.1  # igual ao original


# ======================================================================
# WIN32 (ctypes): enumeração rápida de janelas
# ======================================================================
_user32 = ctypes.windll.user32

_WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

_user32.EnumWindows.argtypes = [_WNDENUMPROC, wintypes.LPARAM]
_user32.EnumWindows.restype = wintypes.BOOL
_user32.IsWindow.argtypes = [wintypes.HWND]
_user32.IsWindow.restype = wintypes.BOOL
_user32.IsWindowVisible.argtypes = [wintypes.HWND]
_user32.IsWindowVisible.restype = wintypes.BOOL
_user32.IsHungAppWindow.argtypes = [wintypes.HWND]
_user32.IsHungAppWindow.restype = wintypes.BOOL
_user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
_user32.GetWindowTextLengthW.restype = ctypes.c_int
_user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
_user32.GetWindowTextW.restype = ctypes.c_int

# Instância única de Desktop, usada só para criar o wrapper da janela escolhida.
_DESKTOP = Desktop(backend="win32")

_AVISOS_TRAVADA = set()


# ======================================================================
# MEDIÇÃO DE DESEMPENHO (opcional: só ativa com --perf)
# ======================================================================
PERF_ATIVO = False
CACHE_ATIVO = True
_TEMPOS = {}


def _registrar_tempo(nome, duracao):
    registro = _TEMPOS.setdefault(nome, [0, 0.0])
    registro[0] += 1
    registro[1] += duracao


@contextmanager
def medir(nome):
    """Mede o tempo de um bloco. Sem efeito se --perf estiver desligado."""
    if not PERF_ATIVO:
        yield
        return

    inicio = time.perf_counter()
    try:
        yield
    finally:
        _registrar_tempo(nome, time.perf_counter() - inicio)


def cronometrar(nome=None):
    """Decorator que mede o tempo total de uma função (só com --perf)."""
    def decorador(func):
        rotulo = nome or func.__name__

        @wraps(func)
        def wrapper(*args, **kwargs):
            if not PERF_ATIVO:
                return func(*args, **kwargs)

            inicio = time.perf_counter()
            try:
                return func(*args, **kwargs)
            finally:
                _registrar_tempo(rotulo, time.perf_counter() - inicio)

        return wrapper
    return decorador


def imprimir_relatorio_tempos():
    if not PERF_ATIVO or not _TEMPOS:
        return

    print("\n" + "=" * 66)
    print(" RELATÓRIO DE DESEMPENHO (--perf)")
    print(" Obs.: tempos aninhados se sobrepõem; não some as linhas.")
    print("=" * 66)
    print(f"{'Etapa':<42}{'Chamadas':>9}{'Total (s)':>11}{'Média (s)':>11}")

    for nome, (chamadas, total) in sorted(
        _TEMPOS.items(), key=lambda item: item[1][1], reverse=True
    ):
        media = total / chamadas if chamadas else 0.0
        print(f"{nome:<42}{chamadas:>9}{total:>11.2f}{media:>11.3f}")

    print("=" * 66)


# ======================================================================
# JANELAS
# ======================================================================
def _listar_janelas_visiveis():
    """Enumera janelas de nível superior visíveis: [(hwnd, titulo), ...]."""
    resultado = []

    def callback(hwnd, _lparam):
        if _user32.IsWindowVisible(hwnd):
            tamanho = _user32.GetWindowTextLengthW(hwnd)
            if tamanho:
                buffer = ctypes.create_unicode_buffer(tamanho + 1)
                _user32.GetWindowTextW(hwnd, buffer, tamanho + 1)
                resultado.append((hwnd, buffer.value))
        return True

    proc = _WNDENUMPROC(callback)
    _user32.EnumWindows(proc, 0)
    return resultado


def _janela_visivel(hwnd):
    return bool(_user32.IsWindow(hwnd)) and bool(_user32.IsWindowVisible(hwnd))


def _buscar_janela(titulos):
    """
    Faz UMA enumeração e retorna (hwnd, titulo_real) da primeira janela
    válida, respeitando a ordem dos títulos informados.
    Comparação: texto idêntico (sem diferenciar maiúsculas/minúsculas),
    como na versão original. Janelas travadas são ignoradas.
    """
    candidatas = {}

    for hwnd, texto in _listar_janelas_visiveis():
        texto = texto.strip()
        candidatas.setdefault(texto.casefold(), []).append((hwnd, texto))

    for titulo in titulos:
        for hwnd, texto in candidatas.get(titulo.strip().casefold(), []):
            if _user32.IsHungAppWindow(hwnd):
                if hwnd not in _AVISOS_TRAVADA:
                    _AVISOS_TRAVADA.add(hwnd)
                    print(f"[AGUARDANDO] Janela sem resposta: {texto}")
                continue

            return hwnd, texto

    return None


def _criar_wrapper(hwnd):
    return _DESKTOP.window(handle=hwnd).wrapper_object()


@cronometrar()
def localizar_janela(titulo_esperado):
    """Verificação única (sem espera). Retorna wrapper ou None."""
    achado = _buscar_janela([titulo_esperado])

    if achado is None:
        return None

    hwnd, titulo = achado
    print(f"[OK] Janela identificada: {titulo}")
    return _criar_wrapper(hwnd)


@cronometrar()
def aguardar_janela(*titulos_esperados, time_limit=TEMPO_LIMITE):
    """
    Aguarda até que uma janela com um dos títulos apareça.
    Retorna o wrapper da primeira janela válida (seguindo a ordem dos
    títulos) ou None se o tempo limite estourar.
    """
    print(
        f"[INFO] Aguardando uma das janelas: "
        f"{', '.join(titulos_esperados)}"
    )

    limite = time.monotonic() + time_limit

    while True:
        with medir("enumerar janelas"):
            achado = _buscar_janela(titulos_esperados)

        if achado is not None:
            hwnd, titulo = achado
            print(f"[OK] Janela identificada: {titulo}")
            return _criar_wrapper(hwnd)

        if time.monotonic() >= limite:
            break

        time.sleep(INTERVALO_VERIFICACAO)

    print(
        f"[INFO] Nenhuma janela válida apareceu em "
        f"{time_limit} segundos: "
        f"{', '.join(titulos_esperados)}"
    )

    return None


def exigir_janela(*titulos_esperados, time_limit=TEMPO_LIMITE):
    janela = aguardar_janela(*titulos_esperados, time_limit=time_limit)

    if janela is None:
        raise RuntimeError(
            "Janela não encontrada: " + ", ".join(titulos_esperados)
        )

    return janela


def aguardar_fechamento(hwnd, time_limit=10):
    """Espera até a janela deixar de existir/ficar visível."""
    limite = time.monotonic() + time_limit

    while _janela_visivel(hwnd) and time.monotonic() < limite:
        time.sleep(INTERVALO_VERIFICACAO)

    return not _janela_visivel(hwnd)


# ======================================================================
# CONTROLES (com cache validado)
# ======================================================================
_CACHE_CONTROLES = {}


def _controle_valido(controle):
    try:
        return _janela_visivel(controle.handle)
    except Exception:
        return False


def obter_controle(chave, janela, buscador):
    """
    Retorna o controle, reutilizando a referência se o handle ainda existir
    e estiver visível. Caso contrário, refaz a busca completa.
    Com --sem-cache, sempre busca de novo (comportamento original).
    """
    chave_cache = (chave, janela.handle)

    if CACHE_ATIVO:
        controle = _CACHE_CONTROLES.get(chave_cache)

        if controle is not None and _controle_valido(controle):
            return controle

    controle = buscador(janela)

    if controle is not None and CACHE_ATIVO:
        _CACHE_CONTROLES[chave_cache] = controle
    else:
        _CACHE_CONTROLES.pop(chave_cache, None)

    return controle


@cronometrar()
def encontrar_campo_funcionario(janela):
    abas = janela.descendants(class_name="TcxTabSheet")

    for aba in abas:
        try:
            if aba.window_text() != "Funcionário":
                continue

            combos = aba.descendants(class_name="TelLookupCombo")

            if combos:
                return combos[0]

        except Exception:
            pass

    return None


@cronometrar()
def encontrar_campo_lancamento(janela):
    abas = janela.descendants(class_name="TcxTabSheet")

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
    combos = janela.descendants(class_name="TelLookupCombo")

    for combo in combos:
        try:
            # Combo dentro da aba Funcionário não é o campo de lançamento.
            if aba_funcionario is not None:
                try:
                    if combo.is_child(aba_funcionario):
                        continue
                except Exception:
                    pass

            # Verificamos o painel pai imediato
            pais = combo.parent()

            if pais is not None:
                combos_pai = pais.descendants(class_name="TelLookupCombo")

                if len(combos_pai) == 1:
                    return combo

        except Exception:
            pass

    return None


@cronometrar()
def encontrar_campo_ano(janela):
    spin_edits = janela.descendants(class_name="TcxSpinEdit")

    for spin in spin_edits:
        try:
            filhos = spin.descendants(class_name="TcxCustomInnerTextEdit")

            for filho in filhos:
                texto = filho.window_text().strip()

                if texto.isdigit() and len(texto) == 4:
                    return filho

        except Exception:
            pass

    return None


@cronometrar()
def encontrar_botao_ok(janela):
    botoes = janela.descendants(class_name="TBitBtn")

    for botao in botoes:
        try:
            if botao.window_text() == "&OK":
                return botao
        except Exception:
            pass

    return None


# ======================================================================
# ENTRADA DE DADOS
# ======================================================================
@cronometrar()
def colar_texto(controle, valor):
    """
    Foca o controle UMA vez, copia o valor e cola substituindo o conteúdo.
    set_foreground=False evita que type_keys refaça set_focus a cada chamada
    (por padrão o pywinauto chama set_focus dentro de cada type_keys).
    """
    controle.set_focus()

    pyperclip.copy(str(valor))

    controle.type_keys("^a", set_foreground=False)
    controle.type_keys("^v", set_foreground=False)


def ler_planilha(caminho):
    import pandas as pd  # importação tardia: só custa tempo quando usado

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


@cronometrar()
def selecionar_funcionario(matricula):
    janela = exigir_janela(TITULO_JANELA_PRINCIPAL)

    controle = obter_controle(
        "campo_funcionario", janela, encontrar_campo_funcionario
    )

    if controle is None:
        raise RuntimeError("Campo de matrícula não encontrado.")

    print(
        f"Preenchendo matrícula {matricula} "
        f"(handle atual: {controle.handle})"
    )

    colar_texto(controle, matricula)

    time.sleep(PAUSA_APOS_COLAR_MATRICULA)

    # Mesmo comportamento do original: foca o campo antes do ENTER.
    controle.set_focus()
    controle.type_keys("{ENTER}", set_foreground=False)


@cronometrar()
def selecionar_funcionarios(dados):
    print(f"Selecionando {len(dados)} funcionários...")

    for matricula, nome in dados:
        print(f"Selecionando: {matricula} - {nome}")

        selecionar_funcionario(matricula)

        time.sleep(PAUSA_APOS_SELECIONAR_FUNCIONARIO)


@cronometrar()
def preencher_ano(ano, janela):
    controle = obter_controle("campo_ano", janela, encontrar_campo_ano)

    if controle is None:
        raise RuntimeError("Campo de ano não encontrado.")

    print(
        f"Preenchendo ano {ano} "
        f"(handle atual: {controle.handle})"
    )

    colar_texto(controle, ano)

    controle.type_keys("{ENTER}", set_foreground=False)


@cronometrar()
def preencher_lancamento(valor, janela):
    controle = obter_controle(
        "campo_lancamento", janela, encontrar_campo_lancamento
    )

    if controle is None:
        raise RuntimeError("Campo de lançamento não encontrado.")

    print(
        f"Preenchendo lançamento {valor} "
        f"(handle atual: {controle.handle})"
    )

    colar_texto(controle, valor)


# ======================================================================
# PREVIEW / IMPRESSÃO / PDF
# ======================================================================
def clicar_botao_preview():
    janela = exigir_janela(TITULO_LANCAMENTOS_ESPECIFICOS)
    janela.set_focus()

    time.sleep(PAUSA_ANTES_CLIQUE_PREVIEW)

    x = 306
    y = 39

    print(f"Clicando no botão em ({x}, {y})...")

    mouse.click(coords=(x, y))

    print("Botão clicado.")


@cronometrar()
def imprimir_como_pdf(ano, codigo_lancamento, pasta_saida):
    janela = exigir_janela(TITULO_IMPRIMIR)
    janela.set_focus()

    print("Janela de impressão encontrada.")

    combo_impressora = None

    for controle in janela.descendants(class_name="TComboBox"):
        try:
            if "Microsoft Print to PDF" in controle.texts():
                combo_impressora = controle
                break
        except Exception:
            pass

    if combo_impressora is None:
        print("ComboBox da impressora não encontrada.")
        return False

    print("Selecionando Microsoft Print to PDF...")

    try:
        combo_impressora.select("Microsoft Print to PDF")
    except Exception as e:
        print("Erro ao selecionar impressora:", e)
        return False

    time.sleep(PAUSA_APOS_SELECIONAR_IMPRESSORA)

    botao_ok = None

    for controle in janela.descendants(class_name="TButton"):
        try:
            if controle.window_text() == "OK":
                botao_ok = controle
                break
        except Exception:
            pass

    if botao_ok is None:
        print("Botão OK da impressão não encontrado.")
        return False

    print("Clicando em OK da impressão...")

    botao_ok.click()

    return salvar_pdf(ano, codigo_lancamento, pasta_saida)


def confirmar_sobrescrita_pdf(janela):
    print("Verificando confirmação de sobrescrita...")

    janela.type_keys("{LEFT}")

    janela.type_keys("{ENTER}")

    print("Confirmação de sobrescrita enviada.")


def aguardar_arquivo_pdf(caminho_pdf, mtime_antes, time_limit=TEMPO_ARQUIVO_PDF):
    """
    Espera o PDF existir, ter conteúdo e (se já existia) ter sido regravado.
    Se o tempo estourar, devolve apenas se o arquivo existe.
    """
    limite = time.monotonic() + time_limit

    while True:
        try:
            info = caminho_pdf.stat()

            if info.st_size > 0 and (
                mtime_antes is None or info.st_mtime_ns != mtime_antes
            ):
                return True
        except OSError:
            pass

        if time.monotonic() >= limite:
            break

        time.sleep(INTERVALO_VERIFICACAO)

    if caminho_pdf.exists():
        print(
            "[AVISO] O arquivo existe, mas não foi atualizado "
            f"em {time_limit}s (pode ser o PDF antigo)."
        )
        return True

    return False


@cronometrar()
def salvar_pdf(ano, codigo_lancamento, pasta_saida):
    janela = exigir_janela(TITULO_SAIDA_IMPRESSAO)
    janela.set_focus()
    hwnd_salvar = janela.handle

    pasta_saida.mkdir(parents=True, exist_ok=True)

    caminho_pdf = pasta_saida / f"FALTAS {codigo_lancamento} {ano}.pdf"

    print("Pasta de saída:", pasta_saida)
    print("Arquivo:", caminho_pdf)

    campo_nome = None

    for controle in janela.descendants(class_name="Edit"):
        try:
            if controle.rectangle().width() > 500:
                campo_nome = controle
                break
        except Exception:
            pass

    if campo_nome is None:
        print("Campo do nome do arquivo não encontrado.")
        return False

    colar_texto(campo_nome, caminho_pdf)

    botao_salvar = None

    for controle in janela.descendants(class_name="Button"):
        try:
            if controle.window_text() == "Sa&lvar":
                botao_salvar = controle
                break
        except Exception:
            pass

    if botao_salvar is None:
        print("Botão Salvar não encontrado.")
        return False

    try:
        mtime_antes = caminho_pdf.stat().st_mtime_ns
    except OSError:
        mtime_antes = None

    botao_salvar.click()

    print("Comando de salvar enviado.")

    # Espera determinística: ou aparece a confirmação de sobrescrita,
    # ou a janela "Salvar como" fecha sozinha.
    limite = time.monotonic() + TEMPO_LIMITE

    while time.monotonic() < limite:
        confirmacao = localizar_janela(TITULO_SOBRESCREVER)

        if confirmacao is not None:
            confirmacao.set_focus()
            confirmar_sobrescrita_pdf(confirmacao)
            break

        if not _janela_visivel(hwnd_salvar):
            break

        time.sleep(INTERVALO_VERIFICACAO)

    aguardar_fechamento(hwnd_salvar, TEMPO_LIMITE)

    if aguardar_arquivo_pdf(caminho_pdf, mtime_antes):
        print("PDF salvo:", caminho_pdf)
        return True

    print("PDF não encontrado após salvar:", caminho_pdf)
    return False


def fechar_preview():
    janela = aguardar_janela(TITULO_LANCAMENTOS_ESPECIFICOS)

    if janela is None:
        print("Preview não encontrado para fechar.")
        return False

    try:
        hwnd = janela.handle
        janela.set_focus()
        janela.close()
    except Exception as erro:
        print(f"Erro ao fechar o preview: {erro}")
        return False

    # Garante que o preview antigo sumiu antes do próximo ano começar.
    aguardar_fechamento(hwnd)

    print("Preview fechado.")
    return True


@cronometrar()
def processar_resultado(ano, codigo_lancamento, pasta_saida):
    print("Aguardando resultado...")

    clicar_botao_preview()

    salvo = imprimir_como_pdf(ano, codigo_lancamento, pasta_saida)

    if salvo:
        if not fechar_preview():
            raise RuntimeError(
                "PDF salvo, mas o preview não pôde ser fechado."
            )

        return True

    return False

# =========================================================================
# SISTEMA DE ABERTURA DA JANELA DE FICHA FINANCEIRA - GERADOR DE RELATORIOS
# =========================================================================
def folhaInicio():

    print()
    print("=" * 80)
    print("EXECUTANDO FOLHA INÍCIO")
    print("=" * 80)

    print(f"Script: {SCRIPT}")
    print(f"Existe: {SCRIPT.exists()}")
    print(f"Diretório: {SCRIPT.parent}")
    print(f"Diretório existe: {SCRIPT.parent.exists()}")

    if not SCRIPT.exists():

        print(
            "ERRO: main.py do FolhaInicio não encontrado."
        )

        return None

    if not SCRIPT.parent.is_dir():

        print(
            "ERRO: diretório do FolhaInicio inválido."
        )

        return None

    processo = subprocess.Popen(
        ["python", str(SCRIPT)],
        cwd=str(SCRIPT.parent)
    )

    return processo

def _construir_arvore_janelas():
    """
    Enumera as janelas uma única vez (Desktop win32), aplica os filtros originais
    e monta a hierarquia (filhas via GA_PARENT, proprietárias via GW_OWNER).

    Retorna (nos, raizes, raiz_de):
      nos     : {hwnd: {"titulo": str, "filhos": [hwnd, ...]}}  (ordem de enumeração)
      raizes  : [hwnd, ...] janelas sem relação com outra janela exibida
      raiz_de : {hwnd: hwnd_da_raiz} árvore à qual cada janela pertence
    """
    import ctypes
    from ctypes import wintypes
    from pywinauto import Desktop

    ignorar = {
        "lidDialog",
        "DataSetValidDialog",
        "Chart editor",
        "Editor de Cross-tab",
        "Dictionary editor",
        "Rich Edit",
        "Task Host Window",
        "CDA Server",
        "RadeonSoftware",
        "WUIconWindow",
        "OneDrive",
        "ad_msg_wnd",
        "Window",
        "Program Manager",
        "Default IME",
        "M",
    }

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    user32.GetWindow.argtypes = [wintypes.HWND, wintypes.UINT]
    user32.GetWindow.restype = wintypes.HWND
    user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
    user32.GetAncestor.restype = wintypes.HWND
    user32.GetDesktopWindow.argtypes = []
    user32.GetDesktopWindow.restype = wintypes.HWND

    GW_OWNER = 4
    GA_PARENT = 1
    desktop_hwnd = user32.GetDesktopWindow()

    # 1) Descoberta + filtros originais
    nos = {}
    ordem = []
    try:
        janelas = Desktop(backend="win32").windows()
    except Exception:
        janelas = []

    for janela in janelas:
        try:
            titulo = janela.window_text().strip()

            if not titulo:
                continue

            if titulo in ignorar:
                continue

            if "GDI+ Window" in titulo:
                continue

            if "cmd.exe" in titulo.lower():
                continue

            if "C:\\Windows" in titulo:
                continue

            if "AMD" in titulo:
                continue

            if "Microsoft" in titulo:
                continue

            if "Realtek" in titulo:
                continue

            if "Bluetooth" in titulo:
                continue

            if "HiddenFax" in titulo:
                continue

            if ".NET-BroadcastEventWindow" in titulo:
                continue

            hwnd = janela.handle
            if hwnd in nos:
                continue

            nos[hwnd] = {"titulo": titulo, "filhos": []}
            ordem.append(hwnd)

        except Exception:
            # janela deixou de existir durante a consulta
            pass

    # 2) Relação direta: filha real (GA_PARENT) > proprietária (GW_OWNER)
    def pai_direto(hwnd):
        pai = user32.GetAncestor(hwnd, GA_PARENT)
        if pai and pai != desktop_hwnd:
            return pai
        return user32.GetWindow(hwnd, GW_OWNER) or None

    # 3) Sobe a cadeia até um ancestral presente entre as janelas descobertas
    def resolver_pai(hwnd):
        visitados = {hwnd}
        candidato = pai_direto(hwnd)
        while candidato and candidato not in visitados:
            if candidato in nos:
                return candidato
            visitados.add(candidato)
            candidato = pai_direto(candidato)
        return None

    pais = {}
    raizes = []
    for hwnd in ordem:
        try:
            pai = resolver_pai(hwnd)
        except Exception:
            pai = None
        pais[hwnd] = pai
        if pai is None:
            raizes.append(hwnd)
        else:
            nos[pai]["filhos"].append(hwnd)

    # 4) Raiz de cada janela (com memoização e proteção contra ciclos)
    raiz_de = {}
    for hwnd in ordem:
        if hwnd in raiz_de:
            continue
        caminho = []
        atual = hwnd
        while atual is not None and atual not in raiz_de and atual not in caminho:
            caminho.append(atual)
            atual = pais[atual]
        raiz = raiz_de[atual] if atual in raiz_de else caminho[-1]
        for h in caminho:
            raiz_de[h] = raiz

    return nos, raizes, raiz_de


def deve_ser_resetado(*titulos):
    """
    True  -> alguma árvore encontrada possui janela não informada.
    False -> nenhuma árvore encontrada possui janela não informada.
    """

    alvo = {str(t).strip() for t in titulos}

    print("\n" + "=" * 70)
    print("[ARVORE] Verificando existência de janelas inesperadas")
    print("=" * 70)

    print(f"[ALVO] {alvo}")

    if not alvo:
        print("[AVISO] Nenhum título informado.")
        return False

    try:
        nos, raizes, raiz_de = _construir_arvore_janelas()

        print(f"[INFO] Janelas identificadas: {len(nos)}")
        print(f"[INFO] Árvores identificadas: {len(raizes)}")

        encontrados = {
            h: no
            for h, no in nos.items()
            if no["titulo"].strip() in alvo
        }

        # Regra obrigatória: nenhum título encontrado -> False imediato.
        # Outras janelas independentes nunca são motivo para retornar True.
        if not encontrados:
            print("\n[RESULTADO] False")
            print("[MOTIVO] Nenhum dos títulos informados foi encontrado.")
            return True

        # A partir daqui pelo menos um título foi confirmado.
        arvores = {raiz_de[h] for h in encontrados}

        print(f"\n[ENCONTRADOS] {len(encontrados)}")

        for hwnd, no in encontrados.items():
            raiz = raiz_de[hwnd]
            print(
                f"  [OK] {no['titulo']} "
                f"| HWND={hwnd} "
                f"| Raiz={nos[raiz]['titulo']}"
            )

        print(f"\n[INFO] Árvores relevantes: {len(arvores)}")

        inesperadas = []

        for hwnd, no in nos.items():
            raiz = raiz_de[hwnd]

            if raiz not in arvores:
                continue

            titulo = no["titulo"].strip()

            if titulo not in alvo:
                inesperadas.append((hwnd, titulo, raiz))

        if inesperadas:
            print(f"\n[ALERTA] {len(inesperadas)} janela(s) inesperada(s):")

            for hwnd, titulo, raiz in inesperadas:
                print(f"  [!] {titulo} | HWND={hwnd}")
                print(f"      Árvore: {nos[raiz]['titulo']}")

            print("\n[RESULTADO] True")
            print("[MOTIVO] Existem janelas fora dos títulos informados.")

            return True

        print("\n[RESULTADO] False")
        print("[MOTIVO] Todas as janelas das árvores correspondentes foram autorizadas.")

        return False

    except Exception as erro:
        print(f"\n[ERRO] Falha ao verificar árvore: {erro}")
        raise

    finally:
        print("=" * 70 + "\n")

@cronometrar()
def ligar_janela_ficha():
    """Abre o sistema (FolhaInicio) e chega ao Gerador pelos atalhos de menu."""
    # verificação rápida se está ligada a outra janela
    if deve_ser_resetado(
        TITULO_INICIAL,
        TITULO_TESTE,
        TITULO_PRINCIPAL
    ):
        folha_iniciada = False
        while not folha_iniciada:
            try:
                processo = folhaInicio()

                if processo is None:
                    raise RuntimeError("Não foi possível iniciar o FolhaInicio.")

                print(f"Aguardando o FolhaInicio terminar (até {TEMPO_FOLHA_INICIO} s)...")

                try:
                    codigo = processo.wait(timeout=TEMPO_FOLHA_INICIO)
                except subprocess.TimeoutExpired:
                    processo.terminate()
                    raise TimeoutError(
                        f"O FolhaInicio não terminou em {TEMPO_FOLHA_INICIO} segundos."
                    )

                if codigo != 0:
                    raise RuntimeError(
                        f"O FolhaInicio terminou com erro (código {codigo})."
                    )
                folha_iniciada = True
            except:
                pass

            print("FolhaInicio concluído.")
    
    janela = exigir_janela(TITULO_INICIAL)
    janela.set_focus()

    for tecla in ATALHO_FICHA:
        keyboard.send_keys(tecla, vk_packet=False)
        time.sleep(INTERVALO_TECLAS)

    # Executar seta para baixo 30 vezes
    pyautogui.press(
        "down",
        presses=30,
        interval=.001
    )

    # Confirmar com Enter
    pyautogui.press("enter")
    
    return exigir_janela(TITULO_JANELA_PRINCIPAL)


# ======================================================================
# FLUXO PRINCIPAL
# ======================================================================
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


def processar_anos(caminho_saida, inicio_ano, fim_ano, codigo_lancamento):
    for ano in range(inicio_ano, fim_ano + 1):

        with medir(f"ano completo (lançamento {codigo_lancamento})"):

            print()
            print("=" * 60)
            print(
                f"PROCESSANDO ANO {ano} - "
                f"LANÇAMENTO {codigo_lancamento}"
            )
            print("=" * 60)

            janela = exigir_janela(TITULO_JANELA_PRINCIPAL)
            janela.set_focus()

            preencher_ano(ano, janela)
            time.sleep(PAUSA_ENTRE_CAMPOS)

            preencher_lancamento(codigo_lancamento, janela)
            time.sleep(PAUSA_ENTRE_CAMPOS)

            # Clicar no botão OK do Gerador
            botao = obter_controle("botao_ok", janela, encontrar_botao_ok)

            if botao is None:
                raise RuntimeError("Botão OK não encontrado.")

            print("Clicando em OK")

            botao.click()

            salvo = processar_resultado(ano, codigo_lancamento, caminho_saida)

            caminho_pdf = (
                caminho_saida / f"FALTAS {codigo_lancamento} {ano}.pdf"
            )

            if salvo:
                if caminho_pdf.exists():
                    print()
                    print("[OK] PDF CONFIRMADO:")
                    print(caminho_pdf)
                    print(
                        f"Tamanho: "
                        f"{caminho_pdf.stat().st_size:,} bytes"
                    )
                else:
                    print()
                    print("[ERRO] O sistema informou que o PDF foi salvo,")
                    print("mas o arquivo NÃO foi encontrado:")
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


def obter_pasta_servidor(nome_funcionario):
    nome = re.sub(r'[<>:"/\\|?*]', "_", nome_funcionario)
    nome = nome.strip().rstrip(".")

    if not nome:
        raise ValueError("O nome do servidor está vazio.")

    return DIRETORIO_SERVIDORES / nome


def main():
    global PERF_ATIVO, CACHE_ATIVO

    inicio_execucao = time.perf_counter()

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

    parser.add_argument(
        "--perf",
        action="store_true",
        help="(Opcional) Exibe relatório detalhado de tempos ao final."
    )

    parser.add_argument(
        "--rapido",
        action="store_true",
        help=(
            "(Opcional) Reduz as esperas internas do pywinauto "
            "(Timings.fast()). Teste antes de usar no dia a dia."
        )
    )

    parser.add_argument(
        "--sem-cache",
        action="store_true",
        help=(
            "(Opcional) Desativa o cache de controles e busca tudo "
            "de novo, como na versão original."
        )
    )

    args = parser.parse_args()

    PERF_ATIVO = args.perf
    CACHE_ATIVO = not args.sem_cache

    ligar_janela_ficha()

    if args.rapido:
        Timings.fast()

    if args.out is not None:
        pasta_saida = args.out

    elif args.planilha is not None:
        pasta_saida = args.planilha.resolve().parent

    else:
        pasta_saida = DIRETORIO_SERVIDORES

    pasta_saida = pasta_saida.resolve()

    try:
        inicio_ano, fim_ano = interpretar_intervalo_anos(args.anos)
    except ValueError as erro:
        parser.error(str(erro))

    print(f"Período: {inicio_ano} até {fim_ano}")

    # Seleciona os funcionários conforme os argumentos recebidos.
    if args.mat:
        dados = [(args.mat, "Matrícula informada manualmente")]

    elif args.planilha:
        if not args.planilha.exists():
            parser.error(f"Planilha não encontrada: {args.planilha}")

        if not args.planilha.is_file():
            parser.error(
                f"O caminho informado não é um arquivo: {args.planilha}"
            )

        dados = ler_planilha(args.planilha)

    else:
        parser.error("Informe --mat MATRICULA ou --planilha CAMINHO.")

    # Seleciona todos os funcionários uma única vez.
    selecionar_funcionarios(dados)

    # No modo --mat, identifica o servidor selecionado
    # e direciona os PDFs para a pasta individual dele.
    if args.mat:
        janela = exigir_janela(TITULO_JANELA_PRINCIPAL)
        janela.set_focus()
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

        with medir(f"TOTAL lançamento {codigo_lancamento}"):
            processar_anos(
                pasta_saida,
                inicio_ano,
                fim_ano,
                codigo_lancamento,
            )

    tempo_total = time.perf_counter() - inicio_execucao

    print("\n" + "=" * 60)
    print(f" TEMPO TOTAL DE EXECUÇÃO: {tempo_total:.2f} segundos")
    print("=" * 60)

    imprimir_relatorio_tempos()


if __name__ == "__main__":
    main()