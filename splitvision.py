import os
import re
import sys
import threading
import ctypes
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageDraw, ImageFont, ImageTk, ImageOps, ImageEnhance
from PyPDF2 import PdfReader, PdfWriter
import pypdfium2 as pdfium
import windnd
import winocr
import winrt.windows.media.ocr as ocr
import conversor_pdf_word

# Determina caminho absoluto para o executável ou script
if getattr(sys, "frozen", False):
    diretorio_base = os.path.dirname(sys.executable)
else:
    diretorio_base = os.path.dirname(os.path.abspath(__file__))


def obter_caminho_recurso(nome_arquivo):
    """ Retorna o caminho absoluto do recurso, seja rodando em script ou empacotado. """
    try:
        # PyInstaller cria a pasta temporária sys._MEIPASS
        base_path = sys._MEIPASS  # type: ignore
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, nome_arquivo)

arquivo_pdf = None  # caminho do PDF selecionado
deve_cancelar = False  # flag global para cancelamento de processamento



def obter_idioma_ocr():
    try:
        idiomas = [lang.language_tag for lang in ocr.OcrEngine.available_recognizer_languages]
        if not idiomas:
            return None

        # Preferir português
        for tag in ("pt-BR", "pt-PT", "pt"):
            for disp in idiomas:
                if disp.lower() == tag.lower() or disp.lower().startswith(tag.lower() + "-"):
                    return disp

        # Preferir inglês
        for tag in ("en-US", "en"):
            for disp in idiomas:
                if disp.lower() == tag.lower() or disp.lower().startswith(tag.lower() + "-"):
                    return disp

        return idiomas[0]
    except Exception:
        return None


def carregar_idiomas_ocr():
    try:
        return [lang.language_tag for lang in ocr.OcrEngine.available_recognizer_languages]
    except Exception:
        return []


def sanitizar_nome_arquivo(nome):
    if not nome:
        return "pagina"
    # Substitui caracteres inválidos do Windows por hífens
    nome_limpo = re.sub(r'[\\/*?:"<>|]', "-", nome)
    # Remove múltiplos hífens ou espaços
    nome_limpo = re.sub(r'-+', "-", nome_limpo)
    nome_limpo = re.sub(r'\s+', " ", nome_limpo)
    return nome_limpo.strip("- ")


def validar_caminho_pdf(caminho):
    if not caminho:
        raise ValueError("Nenhum arquivo ou pasta foi selecionada.")
    caminho_abs = os.path.abspath(caminho)
    if os.path.isdir(caminho_abs):
        return caminho_abs
    if os.path.isfile(caminho_abs):
        if not caminho_abs.lower().endswith(".pdf"):
            raise ValueError("O arquivo selecionado não é um PDF válido.")
        return caminho_abs
    raise FileNotFoundError(f"Caminho não encontrado: {caminho_abs}")



def validar_pasta_saida(pasta):
    if not pasta:
        raise ValueError("Nenhuma pasta de saída foi selecionada.")
    pasta_abs = os.path.abspath(pasta)
    if os.path.exists(pasta_abs) and not os.path.isdir(pasta_abs):
        raise NotADirectoryError(f"O caminho de saída não é uma pasta válida: {pasta_abs}")
    return pasta_abs


def escolher_arquivo():
    global arquivo_pdf
    caminho = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if caminho:
        try:
            arquivo_pdf = validar_caminho_pdf(caminho)
            atualizar_interface_arquivo()
        except Exception as exc:
            messagebox.showerror("Arquivo inválido", str(exc))


def dropped_files(files):
    global arquivo_pdf
    if files:
        caminho = files[0]
        ext = os.path.splitext(caminho)[1].lower()

        aba_atual = None
        if globals().get("notebook"):
            try:
                aba_atual = notebook.select()
            except Exception:
                pass

        if ext in (".docx", ".doc"):
            if globals().get("notebook") and globals().get("pane_conversor"):
                notebook.select(pane_conversor)
            if globals().get("arquivo_conversor_var"):
                arquivo_conversor_var.set(caminho)
            if globals().get("modo_conversor_var"):
                modo_conversor_var.set("Word (.docx) ➔ PDF")
            if globals().get("pasta_saida_conversor_var"):
                pasta_saida_conversor_var.set(os.path.dirname(caminho))
            adicionar_log(f"[INFO] Documento Word detectado: {os.path.basename(caminho)}")
            adicionar_log("[DICA] Modo alterado automaticamente para Conversor Word ➔ PDF.")
            if globals().get("atualizar_interface_conversor"):
                atualizar_interface_conversor()

        elif aba_atual and globals().get("pane_conversor") and str(aba_atual) == str(pane_conversor):
            # Se o usuário estiver na aba Conversor e arrastar um PDF:
            if globals().get("arquivo_conversor_var"):
                arquivo_conversor_var.set(caminho)
            if globals().get("modo_conversor_var"):
                modo_conversor_var.set("PDF ➔ Word (.docx)")
            if globals().get("pasta_saida_conversor_var"):
                pasta_saida_conversor_var.set(os.path.dirname(caminho) if os.path.isfile(caminho) else caminho)
            adicionar_log(f"[INFO] Arquivo selecionado no Conversor: {os.path.basename(caminho)}")
            if globals().get("atualizar_interface_conversor"):
                atualizar_interface_conversor()

        else:
            try:
                arquivo_pdf = validar_caminho_pdf(caminho)
                atualizar_interface_arquivo()
                if globals().get("arquivo_conversor_var"):
                    arquivo_conversor_var.set(caminho)
                if globals().get("pasta_saida_conversor_var"):
                    pasta_saida_conversor_var.set(os.path.dirname(caminho) if os.path.isfile(caminho) else caminho)
                if globals().get("atualizar_interface_conversor"):
                    atualizar_interface_conversor()
            except Exception as exc:
                messagebox.showwarning("Formato incorreto", str(exc))


def escolher_pasta_saida():
    pasta = filedialog.askdirectory(title="Selecione a pasta onde salvar os arquivos split")
    if pasta:
        try:
            pasta_valida = validar_pasta_saida(pasta)
            pasta_saida_var.set(pasta_valida)
        except Exception as exc:
            messagebox.showerror("Pasta inválida", str(exc))


def resetar_selecao():
    global arquivo_pdf
    arquivo_pdf = None
    pasta_saida_var.set("")
    set_progresso(0)
    atualizar_interface_arquivo()


def atualizar_interface_arquivo():
    global arquivo_pdf
    if arquivo_pdf:
        if os.path.isdir(arquivo_pdf):
            nome_curto = os.path.basename(arquivo_pdf)
            if not nome_curto:
                nome_curto = arquivo_pdf
            if len(nome_curto) > 30:
                nome_curto = nome_curto[:27] + "..."
            try:
                pdfs = [f for f in os.listdir(arquivo_pdf) if f.lower().endswith(".pdf")]
                total_pdfs = len(pdfs)
            except Exception:
                total_pdfs = 0
            texto_label = f"Pasta Carregada: {nome_curto} ({total_pdfs} PDFs)"
            cor_fundo = "#E67E22"  # Laranja/Âmbar para Pasta
        else:
            nome_curto = os.path.basename(arquivo_pdf)
            if len(nome_curto) > 35:
                nome_curto = nome_curto[:32] + "..."
            texto_label = f"PDF Carregado: {nome_curto}"
            cor_fundo = "#2B8A3E"  # Verde para Arquivo Único

        lbl_arquivo.config(
            text=texto_label,
            bg=cor_fundo,
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=10,
            pady=8,
            relief="flat",
        )
        lbl_arquivo.pack(pady=10)
        
        # Muda a cor da borda do canvas de arraste e solta se os IDs existirem
        if globals().get('canvas_arraste') and globals().get('rect_arraste_id'):
            canvas_arraste.itemconfig(rect_arraste_id, outline=cor_fundo)
            
        # Ativa o botão Limpar
        if globals().get('btn_limpar'):
            btn_limpar.config(state="normal", bg=CORES["cinza_texto"])
        
        # Define diretório padrão se ainda não estiver configurado
        if not pasta_saida_var.get():
            if os.path.isdir(arquivo_pdf):
                pasta_saida_var.set(arquivo_pdf)
            else:
                pasta_saida_var.set(os.path.dirname(arquivo_pdf))
            
        set_progresso(0) # Resetar progresso ao carregar novo arquivo
        btn_processar.config(state="normal", bg=CORES["ativo"], fg=CORES["texto_ativo"])
    else:
        lbl_arquivo.pack_forget() # Esconder label se nenhum arquivo estiver carregado
        
        # Restaura a cor da borda padrão do canvas
        if globals().get('canvas_arraste') and globals().get('rect_arraste_id'):
            canvas_arraste.itemconfig(rect_arraste_id, outline=CORES["ativo"])
            
        # Desativa o botão Limpar
        if globals().get('btn_limpar'):
            btn_limpar.config(state="disabled", bg=CORES["desativado"])
            
        btn_processar.config(state="disabled", bg=CORES["desativado"], fg=CORES["texto_desativado"], disabledforeground=CORES["texto_desativado"])



def set_progresso(porcentagem, total_paginas=None, pagina_atual=None, status_extra=None):
    if not globals().get("canvas_progresso"):
        return
    canvas_progresso.delete("all")
    porcentagem = min(100, max(0, porcentagem))
    
    # Fundo da barra
    canvas_progresso.create_rectangle(0, 0, 380, 25, fill="#E5E5E5", outline="", width=0)
    
    # Progresso preenchido (Muda para verde no final)
    cor_barra = "#2B8A3E" if porcentagem >= 100 else "#0067C0"
    largura = int(380 * (porcentagem / 100))
    if largura > 0:
        canvas_progresso.create_rectangle(0, 0, largura, 25, fill=cor_barra, outline="", width=0)
    
    # Formatação limpa do texto
    if porcentagem >= 100:
        texto = "100% - Concluído com Sucesso!"
    elif total_paginas and pagina_atual:
        texto = f"{porcentagem}% ({pagina_atual}/{total_paginas})"
    else:
        texto = f"{porcentagem}%"
    
    cor_texto = "#1A1A1A" if porcentagem < 50 else "#FFFFFF"
    canvas_progresso.create_text(190, 12, text=texto, font=("Segoe UI", 10, "bold"), fill=cor_texto)


def set_progresso_conversor(porcentagem, total_paginas=None, pagina_atual=None, status_extra=None):
    if not globals().get("canvas_progresso_conv"):
        return
    canvas_progresso_conv.delete("all")
    porcentagem = min(100, max(0, porcentagem))
    
    # Fundo da barra
    canvas_progresso_conv.create_rectangle(0, 0, 380, 25, fill="#E5E5E5", outline="", width=0)
    
    # Progresso preenchido (Muda para verde no final)
    cor_barra = "#2B8A3E" if porcentagem >= 100 else "#0067C0"
    largura = int(380 * (porcentagem / 100))
    if largura > 0:
        canvas_progresso_conv.create_rectangle(0, 0, largura, 25, fill=cor_barra, outline="", width=0)
    
    # Formatação limpa do texto
    if porcentagem >= 100:
        texto = "100% - Concluído com Sucesso!"
    elif total_paginas and pagina_atual:
        texto = f"{porcentagem}% ({pagina_atual}/{total_paginas})"
    else:
        texto = f"{porcentagem}%"
    
    cor_texto = "#1A1A1A" if porcentagem < 50 else "#FFFFFF"
    canvas_progresso_conv.create_text(190, 12, text=texto, font=("Segoe UI", 10, "bold"), fill=cor_texto)



def limpar_texto(texto):
    if not texto:
        return ""
    return re.sub(r"\s+", " ", texto).strip()


def extrair_dados_nome_codigo_data(texto):
    if not texto:
        return None

    linhas = texto.splitlines()
    for idx, linha in enumerate(linhas):
        match_codigo = re.search(r'C[óoí]{1,2}d[oií]?go\s*[:\-]?\s*(\d+)', linha, re.IGNORECASE)
        match_nome = re.search(r'\bNome\s*[:\-]?\s*', linha, re.IGNORECASE)
        
        if match_nome and match_codigo:
            start_nome_val = match_nome.end()
            end_nome_val = match_codigo.start()
            
            if start_nome_val < end_nome_val:
                nome = limpar_texto(linha[start_nome_val:end_nome_val])
                codigo = match_codigo.group(1)
                
                # Procura a data nas linhas seguintes
                data = None
                for j in range(idx + 1, len(linhas)):
                    linha_data = linhas[j]
                    match_data = re.search(r'(?:Data|Emiss[ãa]o|Dt\.?)\s*[:\-]?\s*(\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4})', linha_data, re.IGNORECASE)
                    if not match_data:
                        match_data = re.search(r'(\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4})', linha_data)
                    
                    if match_data:
                        data = match_data.group(1).replace("/", "-").replace(".", "-")
                        break # Encontrou a data, não precisa ler mais
                
                if nome and codigo:
                    return nome, codigo, data

    # Se não encontrou no formato específico da mesma linha, faz o fallback para a busca genérica
    nome = None
    codigo = None
    data = None

    padrao_nome = re.search(r"Nome\s*[:\-]\s*(.+?)(?=\n|$|C[óoí]{1,2}d[oií]?go)", texto, re.IGNORECASE | re.DOTALL)
    if padrao_nome:
        nome = limpar_texto(padrao_nome.group(1))
        if nome and (nome.lower().startswith("código") or nome.lower().startswith("codigo")):
            nome = None

    padrao_codigo = re.search(r"C[óoí]{1,2}d[oií]?go\s*[:\-]\s*(\d+)", texto, re.IGNORECASE)
    if padrao_codigo:
        codigo = padrao_codigo.group(1)

    padrao_data = re.search(r"(?:Data|Emiss[ãa]o|Dt\.?)\s*[:\-]?\s*(\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4})", texto, re.IGNORECASE)
    if not padrao_data:
        padrao_data = re.search(r"(\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4})", texto)
    if padrao_data:
        data = padrao_data.group(1).replace("/", "-").replace(".", "-")

    if nome and codigo:
        return nome, codigo, data
    return None


def obter_info_pagina(texto):
    if not texto:
        return None
    # Procura por: pág 1/3, página 1 de 3, pg. 1 de 3, etc.
    match_total = re.search(r"\b(?:p.{0,1}g(?:[\w\uFFFD]*|\.)?|page|folha|fl\.?)\s*[:#\-\.]?\s*(\d+)\s*(?:de|/)\s*(\d+)\b", texto, re.IGNORECASE)
    if match_total:
        try:
            curr = int(match_total.group(1))
            total = int(match_total.group(2))
            return {"atual": curr, "total": total}
        except ValueError:
            pass
            
    # Se não achar com o total, procura apenas o número da página
    match_simples = re.search(r"\b(?:p.{0,1}g(?:[\w\uFFFD]*|\.)?|page|folha|fl\.?)\s*[:#\-\.]?\s*(\d+)\b", texto, re.IGNORECASE)
    if match_simples:
        try:
            curr = int(match_simples.group(1))
            return {"atual": curr, "total": None}
        except ValueError:
            pass
            
    return None


def montar_nome_personalizado(texto):
    dados = extrair_dados_nome_codigo_data(texto)
    if not dados:
        return None
    nome, codigo, data = dados
    if data:
        return f"{codigo} - {nome} - {data}"
    return f"{codigo} - {nome}"


def extrair_codigo_item(texto):
    """
    Localiza a linha de item abaixo do cabeçalho da tabela e retorna o código.
    Exemplo: "1 72534 TABLET..." -> retorna "72534"
    Exemplo: "1 27 MOTOR..." -> retorna "27"
    """
    linhas = texto.splitlines()
    idx_cabecalho = -1
    for idx, linha in enumerate(linhas):
        linha_lower = linha.lower()
        if "quantidade" in linha_lower and ("descrição" in linha_lower or "descri" in linha_lower or "ferramenta" in linha_lower):
            idx_cabecalho = idx
            break
            
    if idx_cabecalho != -1:
        # Olha as linhas logo após o cabeçalho para achar a linha do item
        for idx_item in range(idx_cabecalho + 1, min(idx_cabecalho + 5, len(linhas))):
            linha_item = linhas[idx_item].strip()
            if not linha_item:
                continue
            # Busca por: [Quantidade] [Código] [Descrição...]
            match = re.match(r"^(\d+)\s+(\d{1,10})\b", linha_item)
            if match:
                return match.group(2)
                
    return None


def extrair_padrao_os_ou_doc(texto, nome_arquivo_origem=""):
    """
    Identifica se no texto (ou no nome do arquivo original) há um padrão de OS ou Documento.
    Retorna uma tupla (prefixo, numero) ou None.
    Exemplo: ("OS", "616175") ou ("Documento", "132539")
    """
    if not texto:
        texto = ""

    # 1. Busca por Ordem de Serviço explícita (ex: O.S:. 650523, OS: 123456, O.S - 123456, OS 616175)
    match_os = re.search(r"\b[O0]\s*\.?\s*[S5]\s*\.?\s*[:#\-]?\s*(\d{6})\b", texto, re.IGNORECASE)
    if match_os:
        return ("OS", match_os.group(1))

    # 2. Busca por Documento explícito (ex: Documento: 123456, Empréstimo 134332, Nota 134332)
    match_doc = re.search(r"\b(?:Doc(?:umento|urnento|umeuto)?|empréstimo|nota)\b[\s.:\-;,]*(\d{5,7})\b", texto, re.IGNORECASE)
    if match_doc:
        return ("Documento", match_doc.group(1))

    # 3. Fallback: Busca qualquer número de 6 dígitos
    match_num = re.search(r"\b(\d{6})\b", texto)
    if match_num:
        num = match_num.group(1)
        # Se o arquivo original era OS (ou continha OS_...), ou se o número começa com 6 ou se tem 'OS' no texto:
        if (nome_arquivo_origem and (nome_arquivo_origem.startswith("OS_") or "OS_" in nome_arquivo_origem)) \
           or re.search(r"\bO\.?S\.?\b", texto, re.IGNORECASE) \
           or num.startswith("6"):
            return ("OS", num)
        else:
            return ("Documento", num)

    return None


def identificar_tipo_documento(texto, nome_arquivo=""):
    """
    Identifica se o documento é 'Ferramentas', 'Pneus', 'Baterias', 'Pneus e Baterias' ou 'Agregados'.
    """
    texto_lower = texto.lower() if texto else ""
    
    # 1. Verifica se contém ferramentas específicas (ex: "marcador", "caneta", "pneumática")
    eh_ferramenta_especifica = bool(re.search(r"\b(?:marcador|marcadores|caneta|pneum[áa]tic[ao]s?)\b", texto_lower))

    # 2. Busca por fabricante na coluna/linha de ferramenta (para garantir identificação caso o código esteja apagado/borrado):
    # CHIAPERINI, BELZER, GEDORE, TRAM / TRAMONTINA / TRAMOTINA, ROBUST, PADO, SCANIA, ROYCE CONNECT, EVANELI, ROCAST, BOSCH, SANSUNG / SAMSUNG
    tem_fabricante_ferramenta = bool(re.search(
        r"\b(?:CHIAPERINI|BELZER|GEDORE|TRAM(?:ONTINA|OTINA)?|ROBUST|PADO|SCANIA|ROYCE(?:\s+CONNECT)?|EVANELI|ROCAST|BOSCH|SANSUNG|SAMSUNG)\b",
        texto_lower,
        re.IGNORECASE
    ))

    # 3. Extrai e valida código de item de ferramenta (iniciando com 72 ou 75 e 5 dígitos)
    codigo = extrair_codigo_item(texto)
    tem_codigo_72_75 = False
    if codigo and (codigo.startswith("72") or codigo.startswith("75")) and len(codigo) == 5:
        tem_codigo_72_75 = True
    elif bool(re.search(r"\b(72\d{3}|75\d{3})\b", texto)):
        tem_codigo_72_75 = True

    # É Ferramentas se tiver termo de ferramenta, código 72/75 OU fabricante de ferramenta (garantindo documentos borrados)
    eh_ferramenta = eh_ferramenta_especifica or tem_codigo_72_75 or tem_fabricante_ferramenta

    # Regras de Pneus, Baterias e Pneus e Baterias (aplicadas somente se NÃO for item de ferramenta):
    if not eh_ferramenta:
        tem_pneu = bool(re.search(r"\bpneue?s?\b", texto_lower))
        tem_bateria = bool(re.search(r"\bbaterias?\b", texto_lower))
        tem_fogo = bool(re.search(r"\bfogo\b", texto_lower))
        tem_amp = bool(re.search(r"\b(100|150)\b", texto_lower))

        if tem_pneu:
            return "Pneus"
        elif tem_bateria and tem_fogo and tem_amp:
            return "Baterias"
        elif tem_fogo:
            return "Pneus e Baterias"
        
    # Se o arquivo for uma OS mas não contiver pneus/bateria/fogo legítimos, classifica como Agregados
    if nome_arquivo.startswith("OS_") or "OS_" in nome_arquivo:
        return "Agregados"
        
    eh_documento = nome_arquivo.startswith("Documento_") or "Documento_" in nome_arquivo

    # Se contém indicação de página e NÃO é Documento, é Caixa de Ferramentas (Ferramentas)
    if not eh_documento:
        if re.search(r"\b(?:p.{0,1}g(?:[\w\uFFFD]*|\.)?|page|folha|fl\.?)\s*[:#\-\.]?\s*\d+\b", texto_lower):
            return "Ferramentas"
        
    if codigo:
        # Regra do Usuário: Ferramentas sempre começam com 72 ou 75 e têm 5 dígitos (ou possuem fabricante/termo de ferramenta)
        if eh_ferramenta:
            return "Ferramentas"
        else:
            return "Agregados"
            
    # Fallback por palavra-chave se não conseguirmos extrair o código da tabela:
    if eh_ferramenta:
        return "Ferramentas"
        
    # Se for Documento e não encontrou nenhum código de ferramenta legítimo, salva como Agregados
    if eh_documento:
        return "Agregados"

    # Se encontrar "agr" ou "agregado" como palavras inteiras
    if re.search(r"\b(agr|agregado)s?\b", texto_lower):
        return "Agregados"
        
    return "Ferramentas" # Caso padrão


MESES_NOME_PARA_NUM = {
    "jan": 1, "janeiro": 1,
    "fev": 2, "fevereiro": 2,
    "mar": 3, "março": 3, "marco": 3,
    "abr": 4, "abril": 4,
    "mai": 5, "maio": 5,
    "jun": 6, "junho": 6,
    "jul": 7, "julho": 7,
    "ago": 8, "agosto": 8,
    "set": 9, "setembro": 9,
    "out": 10, "outubro": 10,
    "nov": 11, "novembro": 11,
    "dez": 12, "dezembro": 12
}


def extrair_data_direta(texto):
    r"""
    Busca uma data no texto de forma direta, sem depender de Nome ou Código.
    Suporta:
    - Separadores /, -, ., \, ou espaços em branco em volta dos separadores (ex: 14 / 07 / 2026, 14. 07. 2026)
    - Rótulos: Data, Emissão, Dt. Emissão, Dt, Emissao, Abertura, Entrada, etc.
    - Formato ISO (AAAA-MM-DD ou AAAA/MM/DD)
    - Nomes de meses por extenso ou abreviados (ex: 14 de Julho de 2026, 14/JUL/2026)
    - Formato MM/AAAA direto se não houver dia (ex: 07/2026)
    - Substituição inteligente de OCR para O/0 e I/l/1 em posições numéricas
    """
    if not texto:
        return None

    def sanitizar_numeros_ocr(s):
        res = []
        for char in s:
            if char in "Oo":
                res.append("0")
            elif char in "lI":
                res.append("1")
            else:
                res.append(char)
        return "".join(res)

    # 1. Tenta buscar datas com rótulos de data explícitos (ex: Data:, Emissão:, Dt. Emissão:, Abertura:, Entrada:)
    match_lbl = re.search(
        r"\b(?:Data|Emiss[ãa]o|Dt\.?\s*Emiss[ãa]o|Dt\.?|Abertura|Entrada)\b[\s:\-]*(\d{1,2}|[OlI]{1,2})\s*[/.\-\\]\s*(\d{1,2}|[OlI]{1,2})\s*[/.\-\\]\s*(\d{2,4})",
        texto,
        re.IGNORECASE
    )
    if match_lbl:
        dia = sanitizar_numeros_ocr(match_lbl.group(1))
        mes = sanitizar_numeros_ocr(match_lbl.group(2))
        ano = sanitizar_numeros_ocr(match_lbl.group(3))
        return f"{dia}-{mes}-{ano}"

    # 2. Busca qualquer data numérica no formato DD/MM/AAAA, DD.MM.AAAA, DD-MM-AAAA (sem exigir \b rígido)
    match_num = re.search(
        r"(?:^|[^\d])(\d{1,2}|[OlI]{1,2})\s*[/.\-\\]\s*(\d{1,2}|[OlI]{1,2})\s*[/.\-\\]\s*(20\d{2}|\d{2})(?:[^\d]|$)",
        texto
    )
    if match_num:
        dia = sanitizar_numeros_ocr(match_num.group(1))
        mes = sanitizar_numeros_ocr(match_num.group(2))
        ano = sanitizar_numeros_ocr(match_num.group(3))
        return f"{dia}-{mes}-{ano}"

    # 3. Busca formato ISO (AAAA-MM-DD, AAAA/MM/DD, AAAA.MM.DD)
    match_iso = re.search(
        r"(?:^|[^\d])(20\d{2})\s*[/.\-\\]\s*(\d{1,2}|[OlI]{1,2})\s*[/.\-\\]\s*(\d{1,2}|[OlI]{1,2})(?:[^\d]|$)",
        texto
    )
    if match_iso:
        ano = match_iso.group(1)
        mes = sanitizar_numeros_ocr(match_iso.group(2))
        dia = sanitizar_numeros_ocr(match_iso.group(3))
        return f"{dia}-{mes}-{ano}"

    # 4. Busca datas com mês por extenso/abreviado (ex: 14 de Julho de 2026, 14/JUL/2026)
    match_ext = re.search(
        r"\b(\d{1,2})\s*(?:de|[/.\-\\])\s*([a-zç]{3,9})\s*(?:de|[/.\-\\])\s*(\d{2,4})\b",
        texto,
        re.IGNORECASE
    )
    if match_ext:
        dia = match_ext.group(1)
        nome_mes = match_ext.group(2).lower()
        ano = match_ext.group(3)
        num_mes = MESES_NOME_PARA_NUM.get(nome_mes)
        if num_mes:
            return f"{dia}-{num_mes}-{ano}"

    # 5. Fallback: Busca mês e ano no formato MM/AAAA ou MM-AAAA (ex: 07/2026, 07.2026)
    match_mes_ano = re.search(
        r"(?:^|[^\d])(0[1-9]|1[0-2])\s*[/.\-\\]\s*(20\d{2})(?:[^\d]|$)",
        texto
    )
    if match_mes_ano:
        mes = match_mes_ano.group(1)
        ano = match_mes_ano.group(2)
        return f"01-{mes}-{ano}"

    return None


def extrair_mes_ano(data_str):
    if not data_str:
        return "Sem_Data"
    partes = data_str.split("-")
    if len(partes) == 3:
        # Se for formato DD-MM-AAAA ou DD-MM-AA
        if len(partes[0]) <= 2 and len(partes[2]) >= 2:
            dia, mes, ano = partes[0], partes[1], partes[2]
            if len(ano) == 2:
                ano = f"20{ano}"
            try:
                m_int = int(mes)
                if 1 <= m_int <= 12 and len(ano) == 4:
                    return f"{str(m_int).zfill(2)}-{ano}"
            except ValueError:
                pass
        # Se for formato AAAA-MM-DD
        elif len(partes[0]) == 4:
            ano, mes = partes[0], partes[1]
            try:
                m_int = int(mes)
                if 1 <= m_int <= 12:
                    return f"{str(m_int).zfill(2)}-{ano}"
            except ValueError:
                pass
    return "Sem_Data"


def reconstruir_layout_ocr(res):
    if not res or not res.get("lines"):
        return ""
    words = []
    for line in res.get("lines", []):
        for w in line.get("words", []):
            rect = w.get("bounding_rect", {})
            words.append({
                "text": w.get("text", ""),
                "x": rect.get("x", 0),
                "y": rect.get("y", 0)
            })
    if not words:
        return ""
    reconstructed_lines = []
    tolerance = 15
    words_sorted = sorted(words, key=lambda wd: wd["y"])
    current_line_words = []
    for w in words_sorted:
        if not current_line_words:
            current_line_words.append(w)
        else:
            avg_y = sum(x["y"] for x in current_line_words) / len(current_line_words)
            if abs(w["y"] - avg_y) <= tolerance:
                current_line_words.append(w)
            else:
                current_line_words.sort(key=lambda wd: wd["x"])
                reconstructed_lines.append(" ".join(x["text"] for x in current_line_words))
                current_line_words = [w]
    if current_line_words:
        current_line_words.sort(key=lambda wd: wd["x"])
        reconstructed_lines.append(" ".join(x["text"] for x in current_line_words))
    return "\n".join(reconstructed_lines)


def pre_processar_imagem_ocr(img):
    """Aplica melhorias de contraste, escala de cinza e nitidez para melhorar OCR de textos claros."""
    try:
        from PIL import ImageChops, ImageOps, ImageEnhance
        
        # 1. Converter para RGB se necessário
        if img.mode != "RGB":
            img = img.convert("RGB")
        r, g, b = img.split()
        
        # 2. Obter a imagem em tons de cinza com base no canal mais escuro de cada pixel
        # Isso destaca textos de qualquer cor (amarelo, azul claro, verde, cinza desbotado, etc.)
        img_min = ImageChops.darker(ImageChops.darker(r, g), b)
        
        # 3. Aplicar autocontraste para garantir que o fundo vire branco puro (255)
        # e o texto mais escuro vire preto puro (0), esticando a faixa de contraste.
        img_calibrada = ImageOps.autocontrast(img_min, cutoff=1)
        
        # 4. Aumentar o contraste de forma agressiva (fator 3.0) para tornar
        # textos cinzas claros bem escuros/pretos, mantendo o fundo branco e a suavidade das bordas
        enhancer_c = ImageEnhance.Contrast(img_calibrada)
        img_contraste = enhancer_c.enhance(3.0)
        
        # 5. Aumentar a nitidez (sharpness) para definir melhor as bordas dos caracteres
        enhancer_s = ImageEnhance.Sharpness(img_contraste)
        img_final = enhancer_s.enhance(2.0)
        
        # 6. Converter de volta para RGB para manter total compatibilidade com o winocr
        return img_final.convert("RGB")
    except Exception:
        # Se ocorrer algum erro inesperado no processamento, retorna a imagem original intacta
        return img


def corrigir_rotacao_pagina(pdf_path, page_num):
    """Renderiza página para OCR com qualidade."""
    try:
        doc = pdfium.PdfDocument(pdf_path)
        page = doc[page_num]
        bitmap = page.render(scale=5)
        img = bitmap.to_pil()
        page.close()
        doc.close()
        return img
    except Exception:
        try:
            with Image.open(pdf_path) as img:
                if hasattr(img, "n_frames") and img.n_frames > page_num:
                    img.seek(page_num)
                elif page_num != 0:
                    return None
                if img.mode != "RGB":
                    img = img.convert("RGB")
                return img.copy()
        except Exception:
            return None


def limpar_logs():
    txt_log.config(state="normal")
    txt_log.delete("1.0", tk.END)
    txt_log.config(state="disabled")


def adicionar_log(mensagem):
    txt_log.config(state="normal")
    
    # Determina a tag de cor com base nas palavras-chave do log
    if "[ERRO]" in mensagem or "[ERRO CRÍTICO]" in mensagem or "falhou" in mensagem or "FALHA" in mensagem:
        tag = "erro"
    elif "[AVISO]" in mensagem or "Alerta" in mensagem:
        tag = "aviso"
    elif "[SUCESSO]" in mensagem or "PROCESSAMENTO CONCLUÍDO" in mensagem:
        tag = "sucesso"
    elif "=== " in mensagem:
        tag = "titulo"
    elif "[INFO]" in mensagem:
        tag = "info"
    else:
        tag = "normal"
        
    txt_log.insert(tk.END, mensagem + "\n", tag)
    txt_log.see(tk.END)
    txt_log.config(state="disabled")


# --- Constantes de Estilo da Interface ---
CORES = {
    "ativo": "#0067C0",
    "hover": "#164FAF",
    "desativado": "#0380ee",
    "texto_ativo": "#FFFFFF",
    "texto_desativado": "#FFFFFF",
    "fundo": "#F1F3F5",
    "fundo_card": "#FFFFFF",
    "fundo_log": "#1A1D20",
    "texto_log": "#E9ECEF",
    "sucesso": "#51CF66",
    "aviso": "#FFD25A",
    "erro": "#FF6B6B",
    "info": "#9ECBFF",
    "titulo_log": "#339AF0",
    "cinza_texto": "#495057",
}

def cancelar_processamento():
    global deve_cancelar
    deve_cancelar = True
    adicionar_log("[AVISO] Solicitando cancelamento do processamento...")
    btn_processar.config(state="disabled", bg=CORES["desativado"], fg=CORES["texto_desativado"])


def desativar_botoes():
    btn_escolher.config(state="disabled", bg=CORES["desativado"], fg=CORES["texto_desativado"])
    btn_alterar_destino.config(state="disabled")
    combo_idioma.config(state="disabled")
    btn_limpar_log.config(state="disabled", bg=CORES["desativado"], fg=CORES["texto_desativado"])
    # Transforma o botão principal em Cancelar
    btn_processar.config(
        text="Cancelar",
        command=cancelar_processamento,
        state="normal",
        bg="#E03131",
        fg="#FFFFFF"
    )


def restaurar_botoes():
    btn_escolher.config(state="normal", bg=CORES["ativo"], fg=CORES["texto_ativo"])
    btn_alterar_destino.config(state="normal")
    if idiomas_disponiveis:
        combo_idioma.config(state="readonly")
    else:
        combo_idioma.config(state="disabled")
    btn_limpar_log.config(state="normal", bg=CORES["cinza_texto"], fg=CORES["texto_ativo"])
    # Transforma o botão principal de volta para Separar e Renomear
    if arquivo_pdf:
        btn_processar.config(
            state="normal",
            text="Separar e Renomear",
            command=processar_pdf,
            bg=CORES["ativo"],
            fg=CORES["texto_ativo"]
        )
    else:
        btn_processar.config(
            state="disabled",
            text="Separar e Renomear",
            command=processar_pdf,
            bg=CORES["desativado"],
            fg=CORES["texto_desativado"],
            disabledforeground=CORES["texto_desativado"]
        )




PALAVRAS_CHAVE_ORIENTACAO = [
    "ordem", "serviço", "servico", "cliente", "data", "emissão", "emissao", 
    "documento", "código", "codigo", "nome", "total", "valor", "quantidade", 
    "descrição", "descricao", "oficina", "empresa", "cnpj", "telefone", "endereço", 
    "endereco", "assina", "assinatura", "empréstimo", "emprestimo", "nota"
]


def texto_contem_palavra_chave(texto):
    if not texto:
        return False
    texto_lower = texto.lower()
    return any(p in texto_lower for p in PALAVRAS_CHAVE_ORIENTACAO)


def processar_pdf_thread(caminho_entrada, pasta_saida, idioma_ocr_selecionado):
    global deve_cancelar
    deve_cancelar = False
    log_caminho = os.path.join(pasta_saida, "log_processamento.txt")
    try:
        caminho_entrada = validar_caminho_pdf(caminho_entrada)
        pasta_saida = validar_pasta_saida(pasta_saida)
        os.makedirs(pasta_saida, exist_ok=True)

        if os.path.isdir(caminho_entrada):
            lista_arquivos = [os.path.join(caminho_entrada, f) for f in os.listdir(caminho_entrada) if f.lower().endswith(".pdf")]
            lista_arquivos.sort()
        else:
            lista_arquivos = [caminho_entrada]

        if not lista_arquivos:
            raise ValueError("Nenhum arquivo PDF encontrado na pasta selecionada.")

        total_arquivos = len(lista_arquivos)
        tempo_inicio = datetime.now()
        total_paginas_processadas = 0
        total_arquivos_gerados = 0

        # Limpa console de logs da interface
        root.after(0, limpar_logs)
        root.after(0, lambda: set_progresso(0))

        with open(log_caminho, "w", encoding="utf-8") as log_f:
            def registrar_log(msg):
                timestamp = datetime.now().strftime("%H:%M:%S")
                msg_formatada = f"[{timestamp}] {msg}"
                log_f.write(msg_formatada + "\n")
                log_f.flush()
                root.after(0, lambda m=msg_formatada: adicionar_log(m))

            if total_arquivos > 1:
                registrar_log(f"=== INICIANDO PROCESSAMENTO EM LOTE DE {total_arquivos} ARQUIVOS ===")
            else:
                registrar_log("=== INICIANDO EXTRAÇÃO E ORGANIZAÇÃO - SPLITVISION ===")

            if idioma_ocr_selecionado and idioma_ocr_selecionado != "Nenhum disponível":
                registrar_log(f"[INFO] OCR ativo com idioma: {idioma_ocr_selecionado}")
            else:
                registrar_log("[AVISO] OCR do Windows inativo ou nenhum idioma disponível.")

            regex_padrao = re.compile(
                r"(\b(?:Doc(?:umento|urnento|umeuto)?|empréstimo|nota)\b[\s.:\-;,]*(\d{5,7})|\b[O0]\s*\.\s*[S5]\s*\.\s*:\s*(\d{6}))",
                re.IGNORECASE,
            )

            def analisar_pagina(page, pdf_path, page_index, idioma_ocr, nome_arquivo_origem=""):
                """Analisa uma única página, extraindo texto via digital ou OCR, e retorna os resultados."""
                texto_digital = ""
                try:
                    texto_digital = page.extract_text() or ""
                except Exception as e:
                    registrar_log(f"[AVISO] Não foi possível extrair texto digital da Página {page_index + 1}: {str(e)}")

                texto_limpo = texto_digital.strip()
                padrao_encontrado = None
                texto_analisado = ""
                origem = "Texto Digital"
                pagina_rotacionada = False

                if texto_limpo:
                    # 1. Primeiro verifica o padrão de caixa de ferramentas (Nome, Código e Data)
                    if montar_nome_personalizado(texto_limpo):
                        padrao_encontrado = "toolbox"
                        texto_analisado = texto_limpo
                    else:
                        # 2. Caso contrário, procura por OS ou Documento
                        res_os_doc = extrair_padrao_os_ou_doc(texto_limpo, nome_arquivo_origem)
                        if res_os_doc:
                            padrao_encontrado = res_os_doc
                            texto_analisado = texto_limpo

                if not padrao_encontrado and idioma_ocr and idioma_ocr != "Nenhum disponível":
                    origem = f"OCR Windows ({idioma_ocr})"
                    try:
                        img_pil = corrigir_rotacao_pagina(pdf_path, page_index)
                        if img_pil is None:
                            raise ValueError("Falha ao renderizar imagem para OCR.")

                        # Mapeamento: (Pillow Transpose, PDF Rotate Clockwise, Descrição de Log)
                        rotacoes = [
                            (None, 0, "padrão"),  # Tenta primeiro sem rotação
                            (Image.ROTATE_180, 180, "180°"),
                            (Image.ROTATE_90, 270, "90° anti-horário"),
                            (Image.ROTATE_270, 90, "90° horário")
                        ]

                        for pil_rot, pdf_rot, desc in rotacoes:
                            img_para_ocr = img_pil.transpose(pil_rot) if pil_rot is not None else img_pil
                            
                            # TENTATIVA 1: OCR na imagem original/limpa (evita distorções em páginas perfeitas)
                            res = winocr.recognize_pil_sync(img_para_ocr, lang=idioma_ocr)
                            texto_ocr = reconstruir_layout_ocr(res)
                            
                            padrao_ocr = None
                            if montar_nome_personalizado(texto_ocr):
                                padrao_ocr = "toolbox"
                            else:
                                res_os_doc = extrair_padrao_os_ou_doc(texto_ocr, nome_arquivo_origem)
                                if res_os_doc:
                                    padrao_ocr = res_os_doc

                            # Valida se a detecção na rotação 0 é legítima (contém palavras-chave em português)
                            # Se não contiver, rejeitamos na rotação 0 para forçar o loop a tentar rotacionar a imagem.
                            if padrao_ocr and pdf_rot == 0:
                                if not texto_contem_palavra_chave(texto_ocr):
                                    padrao_ocr = None

                            # TENTATIVA 2: OCR na imagem pré-processada (caso a primeira tenha falhado)
                            if not padrao_ocr:
                                img_tratada = pre_processar_imagem_ocr(img_para_ocr)
                                res = winocr.recognize_pil_sync(img_tratada, lang=idioma_ocr)
                                texto_ocr = reconstruir_layout_ocr(res)
                                
                                if montar_nome_personalizado(texto_ocr):
                                    padrao_ocr = "toolbox"
                                else:
                                    res_os_doc = extrair_padrao_os_ou_doc(texto_ocr, nome_arquivo_origem)
                                    if res_os_doc:
                                        padrao_ocr = res_os_doc

                                # Valida se a detecção na rotação 0 é legítima (contém palavras-chave em português)
                                if padrao_ocr and pdf_rot == 0:
                                    if not texto_contem_palavra_chave(texto_ocr):
                                        padrao_ocr = None

                            # Determina se a leitura nesta rotação foi bem-sucedida (achou padrão OU identificou número de página)
                            leitura_valida = False
                            if padrao_ocr:
                                leitura_valida = True
                            else:
                                info_pag_ocr = obter_info_pagina(texto_ocr)
                                if info_pag_ocr is not None:
                                    leitura_valida = True

                            if leitura_valida:
                                padrao_encontrado = padrao_ocr
                                texto_analisado = texto_ocr
                                
                                # Determina a rotação final a ser aplicada no PDF
                                rotação_final = pdf_rot
                                angulo_ocr = res.get("text_angle")
                                if pdf_rot == 0 and angulo_ocr is not None:
                                    # Se na rotação padrão (0) o OCR leu o texto inclinado (ex: de ponta cabeça)
                                    if abs(angulo_ocr) > 20:
                                        # Arredonda para o múltiplo de 90° mais próximo
                                        rotação_final = int(round(angulo_ocr / 90.0) * 90) % 360
                                
                                if rotação_final != 0:
                                    page.rotate(rotação_final)
                                    pagina_rotacionada = True
                                    registrar_log(f"[SUCESSO] Página {page_index + 1} corrigida: rotacionada em {rotação_final}°.")
                                break # Encontrou, sai do loop de rotações

                        if not padrao_encontrado:
                            try:
                                img_para_ocr = img_pil
                                img_tratada = pre_processar_imagem_ocr(img_para_ocr)
                                res = winocr.recognize_pil_sync(img_tratada, lang=idioma_ocr)
                                texto_ocr = reconstruir_layout_ocr(res)
                                
                                debug_txt_path = os.path.join(pasta_saida, f"debug_ocr_texto_p{page_index + 1}.txt")
                                with open(debug_txt_path, "w", encoding="utf-8") as df:
                                    df.write(texto_ocr)
                                
                                debug_img_path = os.path.join(pasta_saida, f"debug_ocr_imagem_p{page_index + 1}.png")
                                img_tratada.save(debug_img_path)
                                
                                registrar_log(f"[INFO] Página {page_index + 1}: Depuração salva em: {os.path.basename(debug_txt_path)} e {os.path.basename(debug_img_path)}")
                            except Exception:
                                pass

                    except Exception as exc:
                        texto_analisado = f"[ERRO OCR]: {str(exc)}"
                        registrar_log(f"[ERRO] Página {page_index + 1}: Erro durante OCR - {str(exc)}")
                
                if not texto_analisado:
                    origem = "Sem OCR" if not idioma_ocr or idioma_ocr == "Nenhum disponível" else "OCR Inconclusivo"
                    texto_analisado = texto_digital # Usa o texto digital como fallback se houver

                return {
                    "padrao": padrao_encontrado,
                    "texto": texto_analisado,
                    "origem": origem,
                    "pagina_rotacionada": pagina_rotacionada
                }


            def salvar_grupo(grupo):
                nonlocal total_arquivos_gerados
                if not grupo or not grupo["paginas"]:
                    return
                writer = PdfWriter()
                for pagina in grupo["paginas"]:
                    writer.add_page(pagina)
                
                texto_ocr = grupo.get("texto_ocr", "")
                nome_final = grupo['nome_arquivo']
                nome_seguro = sanitizar_nome_arquivo(nome_final)

                tipo = identificar_tipo_documento(texto_ocr, grupo['nome_arquivo'])
                
                # Extrai a data e formata o mês-ano
                data_doc = extrair_data_direta(texto_ocr)
                mes_ano = extrair_mes_ano(data_doc)
                
                if mes_ano == "Sem_Data":
                    linhas_com_digitos = [l.strip() for l in texto_ocr.splitlines() if re.search(r'\d', l)]
                    amostra_ocr = " | ".join(linhas_com_digitos[:3]) if linhas_com_digitos else (texto_ocr[:120].replace('\n', ' ') if texto_ocr else "sem texto")
                    registrar_log(f"[AVISO] Data não identificada para '{nome_seguro}.pdf'. Trecho de linhas com números: '{amostra_ocr}'")

                # Pasta de destino: pasta_saida / Tipo / Mês-Ano (Ex: pasta/Agregados/05-2026)
                pasta_destino_final = os.path.join(pasta_saida, tipo, mes_ano)
                os.makedirs(pasta_destino_final, exist_ok=True)
                
                caminho_saida = os.path.join(pasta_destino_final, f"{nome_seguro}.pdf")
                with open(caminho_saida, "wb") as f:
                    writer.write(f)
                total_arquivos_gerados += 1
                
                registrar_log(
                    f"[SUCESSO] Salvo em {tipo}/{mes_ano}: {nome_seguro}.pdf ({grupo['origem']}) -> {len(grupo['paginas'])} página(s)"
                )

            for idx_arq, arq_atual in enumerate(lista_arquivos):
                if deve_cancelar:
                    break
                nome_arq_base = os.path.basename(arq_atual)
                status_txt = f"PDF {idx_arq + 1}/{total_arquivos}" if total_arquivos > 1 else ""
                
                if total_arquivos > 1:
                    registrar_log(f"\n[PROCESSO] Arquivo {idx_arq + 1} de {total_arquivos}: {nome_arq_base}")
                else:
                    registrar_log(f"[INFO] Processando: {nome_arq_base}")

                try:
                    reader = PdfReader(arq_atual)
                    total_paginas = len(reader.pages)
                    total_paginas_processadas += total_paginas
                    registrar_log(f"[INFO] ({total_paginas} páginas)")

                    root.after(0, lambda sa=status_txt: set_progresso(0, total_paginas, 0, sa))

                    grupo_atual = None

                    for i in range(total_paginas):
                        if deve_cancelar:
                            break
                        page = reader.pages[i]
                        
                        resultado_analise = analisar_pagina(page, arq_atual, i, idioma_ocr_selecionado, nome_arq_base)
                        padrao = resultado_analise["padrao"]
                        texto_analisado = resultado_analise["texto"]
                        origem = resultado_analise["origem"]

                        info_pag = obter_info_pagina(texto_analisado)

                        # Determina o tipo de grupo que esta página iniciaria
                        tipo_grupo_detectado = None
                        nome_identificador = None
                        nome_arquivo = f"pagina_{i+1}"
                        
                        if padrao:
                            if padrao == "toolbox":
                                nome_personalizado = montar_nome_personalizado(texto_analisado)
                                if nome_personalizado:
                                    nome_identificador = nome_personalizado
                                    nome_arquivo = nome_personalizado
                                    tipo_grupo_detectado = "toolbox"
                            elif isinstance(padrao, (tuple, list)):
                                prefixo, numero = padrao
                                nome_identificador = f"{prefixo}_{numero}"
                                nome_arquivo = nome_identificador
                                tipo_grupo_detectado = "os_ou_doc"

                        # 1. Determina se é continuação do grupo ativo
                        eh_continuacao = False
                        if grupo_atual is not None:
                            tipo_ativo = grupo_atual.get("tipo_grupo")
                            if tipo_ativo == "toolbox":
                                # Para caixa de ferramentas, agrupa se a página atual for explicitamente > 1
                                if info_pag and info_pag["atual"] > 1:
                                    eh_continuacao = True
                                # Ou se não tem info de página mas ainda estamos no total esperado do grupo
                                elif not info_pag and grupo_atual.get("total_esperado"):
                                    if len(grupo_atual["paginas"]) < grupo_atual["total_esperado"]:
                                        eh_continuacao = True
                            elif tipo_ativo == "os_ou_doc":
                                # Para OS ou Documento, agrupa se tiver o mesmo identificador
                                if nome_identificador and nome_identificador == grupo_atual.get("nome_arquivo"):
                                    eh_continuacao = True
                                # Ou se a página atual for explicitamente > 1 (comportamento original)
                                elif info_pag and info_pag["atual"] > 1:
                                    eh_continuacao = True

                        # 2. Determina se é o início de um novo documento
                        eh_inicio = False
                        if not eh_continuacao:
                            if tipo_grupo_detectado is not None:
                                eh_inicio = True

                        # Decisão de Agrupamento
                        if eh_continuacao:
                            grupo_atual["paginas"].append(page)
                            registrar_log(
                                f"[INFO] Página {i+1}: Sem padrão novo. Agrupada ao documento '{sanitizar_nome_arquivo(grupo_atual['nome_arquivo'])}.pdf' ({origem})."
                            )
                        elif eh_inicio:
                            if grupo_atual is not None:
                                salvar_grupo(grupo_atual)
                            
                            # Inicializa o novo grupo e armazena o total esperado de páginas se detectado (ex: de 3)
                            total_esperado = info_pag["total"] if info_pag else None
                            grupo_atual = {
                                "nome_arquivo": nome_arquivo,
                                "paginas": [page],
                                "origem": origem,
                                "texto_ocr": texto_analisado,
                                "total_esperado": total_esperado,
                                "tipo_grupo": tipo_grupo_detectado
                            }
                            
                            if nome_identificador:
                                nome_exibicao = sanitizar_nome_arquivo(nome_arquivo)
                                registrar_log(
                                    f"[INFO] Página {i+1}: Padrão detectado! Renomeando para '{nome_exibicao}.pdf' ({origem})."
                                )
                            else:
                                registrar_log(
                                    f"[INFO] Página {i+1}: Início de documento (Página 1). Criando grupo '{sanitizar_nome_arquivo(nome_arquivo)}.pdf' ({origem})."
                                )
                        elif grupo_atual is not None:
                            salvar_grupo(grupo_atual)
                            grupo_atual = None
                            salvar_grupo({"nome_arquivo": nome_arquivo, "paginas": [page], "origem": origem, "texto_ocr": texto_analisado})
                            registrar_log(
                                f"[INFO] Página {i+1}: Padrão não detectado e não é continuação. Salva como '{sanitizar_nome_arquivo(nome_arquivo)}.pdf' ({origem})."
                            )
                        else:
                            salvar_grupo({"nome_arquivo": nome_arquivo, "paginas": [page], "origem": origem, "texto_ocr": texto_analisado})
                            registrar_log(
                                f"[INFO] Página {i+1}: Padrão não detectado. Salva como '{sanitizar_nome_arquivo(nome_arquivo)}.pdf' ({origem})."
                            )

                        porcentagem = int((i + 1) / total_paginas * 100)
                        root.after(0, lambda p=porcentagem, t=total_paginas, curr=i + 1, sa=status_txt: set_progresso(p, t, curr, sa))

                    if grupo_atual is not None and not deve_cancelar:
                        salvar_grupo(grupo_atual)

                except Exception as file_exc:
                    registrar_log(f"[ERRO] Falha ao processar o arquivo {nome_arq_base}: {str(file_exc)}")

            if deve_cancelar:
                registrar_log("\n=== PROCESSAMENTO CANCELADO PELO USUÁRIO ===")
            else:
                tempo_fim = datetime.now()
                duracao = tempo_fim - tempo_inicio
                segundos = duracao.total_seconds()
                if segundos < 60:
                    tempo_formatado = f"{segundos:.1f} segundos"
                else:
                    minutos = int(segundos // 60)
                    restante = int(segundos % 60)
                    tempo_formatado = f"{minutos}m {restante}s"
                    
                registrar_log("\n=== PROCESSAMENTO CONCLUÍDO COM SUCESSO ===")
                registrar_log(f"[RESUMO] Tempo total gasto: {tempo_formatado}")
                registrar_log(f"[RESUMO] Páginas de entrada processadas: {total_paginas_processadas}")
                registrar_log(f"[RESUMO] Arquivos PDFs finais gerados: {total_arquivos_gerados}")

        def finalizar_sucesso():
            global deve_cancelar
            if deve_cancelar:
                messagebox.showwarning(
                    "Cancelado",
                    f"⚠️ O processamento foi cancelado pelo usuário!\nArquivos salvos até o momento estão em:\n{pasta_saida}",
                )
            else:
                if total_arquivos > 1:
                    msg = f"✅ {total_arquivos} arquivos PDF processados com sucesso!\nSalvo em:\n{pasta_saida}"
                else:
                    msg = f"✅ Arquivo processado!\nSalvo em:\n{pasta_saida}"
                messagebox.showinfo(
                    "Sucesso",
                    f"{msg}\nConsulte o console de logs ou o arquivo 'log_processamento.txt' para mais detalhes.",
                )
            root.after(0, restaurar_botoes)

        root.after(0, finalizar_sucesso)

    except Exception as exc:
        def finalizar_erro(err_msg):
            adicionar_log(f"[ERRO CRÍTICO] Falha no processamento: {err_msg}")
            messagebox.showerror("Erro", f"Ocorreu um erro crítico durante o processamento:\n{err_msg}")
            root.after(0, restaurar_botoes)

        root.after(0, finalizar_erro, str(exc))



def processar_pdf():
    global arquivo_pdf
    if not arquivo_pdf:
        messagebox.showwarning("Aviso", "Selecione ou arraste primeiro um arquivo PDF!")
        return

    pasta_saida = pasta_saida_var.get()
    try:
        pasta_saida = validar_pasta_saida(pasta_saida)
    except Exception as exc:
        messagebox.showerror("Pasta inválida", f"A pasta de saída configurada é inválida:\n{str(exc)}")
        return

    desativar_botoes()
    set_progresso(0)

    idioma_ocr_selecionado = idioma_ocr_var.get()
    
    threading.Thread(
        target=processar_pdf_thread, 
        args=(arquivo_pdf, pasta_saida, idioma_ocr_selecionado), 
        daemon=True
    ).start()


# --- Funções do Conversor PDF / Word (Aba Conversor) ---
def escolher_arquivo_conversor():
    caminho = filedialog.askopenfilename(filetypes=[
        ("Todos os Suportados", "*.pdf *.docx *.doc"),
        ("Arquivos PDF (*.pdf)", "*.pdf"),
        ("Documentos Word (*.docx, *.doc)", "*.docx *.doc"),
        ("Todos os Arquivos", "*.*")
    ])
    if caminho:
        arquivo_conversor_var.set(caminho)
        ext = os.path.splitext(caminho)[1].lower()
        if ext in (".docx", ".doc"):
            modo_conversor_var.set("Word (.docx) ➔ PDF")
        elif ext == ".pdf":
            modo_conversor_var.set("PDF ➔ Word (.docx)")
        if not pasta_saida_conversor_var.get():
            pasta_saida_conversor_var.set(os.path.dirname(caminho))
        adicionar_log(f"[INFO] Arquivo selecionado para conversão: {os.path.basename(caminho)}")
        atualizar_interface_conversor()

def resetar_selecao_conversor():
    arquivo_conversor_var.set("")
    set_progresso_conversor(0)
    atualizar_interface_conversor()

def atualizar_interface_conversor():
    caminho = arquivo_conversor_var.get().strip() if globals().get("arquivo_conversor_var") else ""
    if caminho and os.path.exists(caminho):
        nome_curto = os.path.basename(caminho)
        if len(nome_curto) > 35:
            nome_curto = nome_curto[:32] + "..."
        ext = os.path.splitext(caminho)[1].lower()
        if ext in (".docx", ".doc"):
            texto_label = f"Word Carregado: {nome_curto}"
            cor_fundo = "#1971C2"  # Azul para Word
        else:
            texto_label = f"PDF Carregado: {nome_curto}"
            cor_fundo = "#2B8A3E"  # Verde para PDF
            
        if globals().get("lbl_arquivo_conv"):
            lbl_arquivo_conv.config(
                text=texto_label,
                bg=cor_fundo,
                fg="white",
                font=("Segoe UI", 10, "bold"),
                padx=10,
                pady=8,
                relief="flat",
            )
            lbl_arquivo_conv.pack(pady=10)
        
        if globals().get("canvas_arraste_conv") and globals().get("rect_arraste_conv_id"):
            canvas_arraste_conv.itemconfig(rect_arraste_conv_id, outline=cor_fundo)
            
        if globals().get("btn_limpar_conv"):
            btn_limpar_conv.config(state="normal", bg=CORES["cinza_texto"])

        if globals().get("btn_iniciar_conv"):
            btn_iniciar_conv.config(state="normal", bg=CORES["ativo"], fg=CORES["texto_ativo"])
            
        if globals().get("pasta_saida_conversor_var") and not pasta_saida_conversor_var.get():
            pasta_saida_conversor_var.set(os.path.dirname(caminho))
    else:
        if globals().get("lbl_arquivo_conv"):
            lbl_arquivo_conv.pack_forget()
            
        if globals().get("canvas_arraste_conv") and globals().get("rect_arraste_conv_id"):
            canvas_arraste_conv.itemconfig(rect_arraste_conv_id, outline=CORES["ativo"])
            
        if globals().get("btn_limpar_conv"):
            btn_limpar_conv.config(state="disabled", bg=CORES["desativado"])

        if globals().get("btn_iniciar_conv"):
            btn_iniciar_conv.config(
                state="disabled",
                bg=CORES["desativado"],
                fg=CORES["texto_desativado"],
                disabledforeground=CORES["texto_desativado"]
            )

def escolher_pasta_saida_conversor():
    pasta = filedialog.askdirectory(title="Selecione a pasta de saída para a conversão")
    if pasta:
        pasta_saida_conversor_var.set(pasta)

def processar_conversao_thread(caminho_in, modo, pasta_out):
    global deve_cancelar
    deve_cancelar = False
    try:
        root.after(0, lambda: set_progresso_conversor(0))
        adicionar_log(f"=== INICIANDO OPERAÇÃO: {modo} ===")
        if caminho_in:
            adicionar_log(f"[INFO] Entrada: {os.path.basename(caminho_in)}")
        adicionar_log(f"[INFO] Pasta destino: {pasta_out}")
        
        def callback_prog(pct, msg=""):
            root.after(0, lambda: set_progresso_conversor(pct, status_extra=msg))
            if msg:
                root.after(0, lambda m=msg: adicionar_log(f"[CONVERSÃO] {m}"))

        nome_base = os.path.splitext(os.path.basename(caminho_in))[0] if caminho_in else "documento"
        
        if "PDF ➔ Word" in modo:
            out_file = os.path.join(pasta_out, f"{nome_base}.docx")
            idioma_ocr_sel = idioma_ocr_var.get() if globals().get("idioma_ocr_var") else None
            res = conversor_pdf_word.converter_pdf_para_word(caminho_in, out_file, callback_prog, idioma_ocr=idioma_ocr_sel)
            adicionar_log(f"[SUCESSO] Documento Word editável gerado com sucesso: {os.path.basename(res)}")
            root.after(0, lambda: messagebox.showinfo("Sucesso", f"✅ PDF convertido para Word com sucesso!\nSalvo em:\n{res}"))
            
        elif "Word (.docx) ➔ PDF" in modo:
            out_file = os.path.join(pasta_out, f"{nome_base}.pdf")
            res = conversor_pdf_word.converter_word_para_pdf(caminho_in, out_file, callback_prog)
            adicionar_log(f"[SUCESSO] Arquivo PDF gerado com sucesso: {os.path.basename(res)}")
            root.after(0, lambda: messagebox.showinfo("Sucesso", f"✅ Documento Word convertido para PDF com sucesso!\nSalvo em:\n{res}"))
            
        elif "Juntar Múltiplos PDFs" in modo:
            arquivos_pdf = filedialog.askopenfilenames(
                title="Selecione os PDFs para unir",
                filetypes=[("Arquivos PDF", "*.pdf")]
            )
            if arquivos_pdf:
                out_file = os.path.join(pasta_out, f"PDF_Unificado_{datetime.now().strftime('%H%M%S')}.pdf")
                res = conversor_pdf_word.juntar_pdfs(list(arquivos_pdf), out_file, callback_prog)
                adicionar_log(f"[SUCESSO] {len(arquivos_pdf)} arquivos PDF unidos em: {os.path.basename(res)}")
                root.after(0, lambda: messagebox.showinfo("Sucesso", f"✅ {len(arquivos_pdf)} PDFs unidos com sucesso!\nSalvo em:\n{res}"))
            else:
                adicionar_log("[AVISO] Seleção de PDFs para junção foi cancelada.")

        elif "Exportar Páginas como PNG" in modo:
            pasta_imgs = os.path.join(pasta_out, f"{nome_base}_imagens")
            imgs = conversor_pdf_word.extrair_imagens_pdf(caminho_in, pasta_imgs, callback_progresso=callback_prog)
            adicionar_log(f"[SUCESSO] {len(imgs)} páginas exportadas em alta resolução em: {pasta_imgs}")
            root.after(0, lambda: messagebox.showinfo("Sucesso", f"✅ {len(imgs)} páginas exportadas como PNG em:\n{pasta_imgs}"))

    except Exception as exc:
        msg_err = str(exc)
        adicionar_log(f"[ERRO CRÍTICO] Falha na operação: {msg_err}")
        root.after(0, lambda: messagebox.showerror("Erro de Conversão", f"Falha ao executar a conversão:\n{msg_err}"))
    finally:
        root.after(0, restaurar_botoes)

def iniciar_conversao():
    modo = modo_conversor_var.get()
    if "Juntar Múltiplos PDFs" not in modo:
        caminho_in = arquivo_conversor_var.get().strip()
        if not caminho_in or not os.path.exists(caminho_in):
            messagebox.showwarning("Aviso", "Selecione primeiro um arquivo válido para converter!")
            return
    else:
        caminho_in = ""

    pasta_out = pasta_saida_conversor_var.get().strip()
    if not pasta_out or not os.path.exists(pasta_out):
        if caminho_in:
            pasta_out = os.path.dirname(caminho_in)
            pasta_saida_conversor_var.set(pasta_out)
        else:
            messagebox.showwarning("Aviso", "Selecione uma pasta de saída válida!")
            return

    desativar_botoes()
    set_progresso_conversor(0)
    threading.Thread(
        target=processar_conversao_thread,
        args=(caminho_in, modo, pasta_out),
        daemon=True
    ).start()


# --- Interface Gráfica ---
# Configura o AppUserModelID para mostrar o ícone na barra de tarefas no Windows
try:
    myappid = "tonivanjsilva-ux.splitvisionpdf.1-3-0"
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass

root = tk.Tk()
root.title("SplitVision PDF v1.3.0 - Organizador Inteligente")
root.geometry("1000x720")
root.minsize(850, 600)
root.configure(bg=CORES["fundo"])

try:
    caminho_icone = obter_caminho_recurso("splitvision.ico")
    if os.path.exists(caminho_icone):
        # 1. Tenta o iconphoto (PIL -> PhotoImage) - Altamente compatível e resolve no Windows 11
        try:
            from PIL import Image, ImageTk
            img_icon = Image.open(caminho_icone)
            photo_icon = ImageTk.PhotoImage(img_icon)
            root.iconphoto(True, photo_icon)
            root._icon_photo_ref = photo_icon
        except Exception:
            pass
            
        # 2. Tenta também o iconbitmap tradicional do Windows (para a barra de tarefas)
        try:
            root.iconbitmap(caminho_icone)
        except Exception:
            pass
except Exception:
    pass

# Variáveis de Controle
pasta_saida_var = tk.StringVar()
idioma_ocr_var = tk.StringVar()
arquivo_conversor_var = tk.StringVar()
modo_conversor_var = tk.StringVar(value="PDF ➔ Word (.docx)")
pasta_saida_conversor_var = tk.StringVar()

def gerar_icone_divisao_tab(tamanho=(22, 22)):
    """ Gera um ícone colorido dinâmico para a aba Divisão & OCR """
    img = Image.new("RGBA", tamanho, (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    # Documento PDF Base (Azul)
    draw.rounded_rectangle([2, 2, 16, 21], radius=2, fill="#0067C0", outline="#004080", width=1)
    # Linhas de Texto no PDF
    draw.line([5, 5, 13, 5], fill="#E6F0FA", width=2)
    draw.line([5, 9, 11, 9], fill="#E6F0FA", width=2)
    draw.line([5, 13, 9, 13], fill="#E6F0FA", width=2)
    # Ícone Badge Tesoura/Divisor (Laranja Vibrante)
    draw.ellipse([11, 11, 21, 21], fill="#FF922B", outline="#E67E22", width=1)
    draw.line([14, 14, 18, 18], fill="white", width=2)
    draw.line([18, 14, 14, 18], fill="white", width=2)
    return ImageTk.PhotoImage(img)

def gerar_icone_conversor_tab(tamanho=(22, 22)):
    """ Gera um ícone colorido dinâmico para a aba Conversor PDF ⇄ Word """
    img = Image.new("RGBA", tamanho, (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    # PDF Miniatura (Vermelho)
    draw.rounded_rectangle([1, 2, 9, 19], radius=2, fill="#E03131", outline="#C92A2A", width=1)
    draw.line([3, 6, 7, 6], fill="white", width=1)
    draw.line([3, 10, 7, 10], fill="white", width=1)
    draw.line([3, 14, 6, 14], fill="white", width=1)
    # Word Miniatura (Azul)
    draw.rounded_rectangle([12, 2, 20, 19], radius=2, fill="#1971C2", outline="#1864AB", width=1)
    draw.line([14, 6, 18, 6], fill="white", width=1)
    draw.line([14, 10, 18, 10], fill="white", width=1)
    draw.line([14, 14, 17, 14], fill="white", width=1)
    # Seta Verde de Conversão no centro
    draw.polygon([(7, 9), (14, 9), (14, 7), (17, 10), (14, 13), (14, 11), (7, 11)], fill="#40C057")
    return ImageTk.PhotoImage(img)

# Estilização TTK
style = ttk.Style()
style.theme_use("clam")
style.configure("TFrame", background=CORES["fundo"])
style.configure("TLabelframe", background=CORES["fundo_card"], bordercolor="#E9ECEF", borderwidth=1)
style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground=CORES["cinza_texto"], background=CORES["fundo_card"])
style.configure("TEntry", fieldbackground="#F8F9FA", bordercolor="#CED4DA", padding=5)
style.configure("TCombobox", fieldbackground="#F8F9FA", bordercolor="#CED4DA")
style.configure("TNotebook", background=CORES["fundo"], borderwidth=0)
style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=[14, 8], background="#E9ECEF", foreground="#495057")
style.map("TNotebook.Tab",
    background=[("selected", CORES["fundo_card"])],
    foreground=[("selected", CORES["ativo"])]
)

# --- Cabeçalho (Header Banner) ---
header_frame = tk.Frame(root, bg=CORES["ativo"])
header_frame.pack(fill="x", side="top")

lbl_titulo = tk.Label(
    header_frame, 
    text="SplitVision PDF Pro v1.3.0", 
    font=("Segoe UI", 20, "bold"), 
    fg="white", 
    bg=CORES["ativo"]
)
lbl_titulo.pack(anchor="w", padx=25, pady=(10, 2))

lbl_subtitulo = tk.Label(
    header_frame, 
    text="Organização inteligente de PDFs com OCR e Conversor Bidirecional PDF ⇄ Word (.docx)", 
    font=("Segoe UI", 10), 
    fg="#D0E1FD", 
    bg=CORES["ativo"]
)
lbl_subtitulo.pack(anchor="w", padx=25, pady=(2, 10))


# --- Rodapé (Footer) ---
footer_frame = tk.Frame(root, bg=CORES["fundo"])
footer_frame.pack(side="bottom", pady=10)

lbl_rodape_p1 = tk.Label(
    footer_frame,
    text="SplitVision PDF v1.3.0 | Desenvolvido por ",
    font=("Segoe UI", 8),
    fg=CORES["cinza_texto"],
    bg=CORES["fundo"]
)
lbl_rodape_p1.pack(side="left")

lbl_rodape_p2 = tk.Label(
    footer_frame,
    text="Tonivan Joseph",
    font=("Segoe UI", 8, "bold"),
    fg=CORES["cinza_texto"],
    bg=CORES["fundo"]
)
lbl_rodape_p2.pack(side="left")

lbl_rodape_p3 = tk.Label(
    footer_frame,
    text=" | Todos os direitos reservados",
    font=("Segoe UI", 8),
    fg=CORES["cinza_texto"],
    bg=CORES["fundo"]
)
lbl_rodape_p3.pack(side="left")


# --- Corpo Principal (Two Panes com Abas) ---
body_frame = tk.Frame(root, bg=CORES["fundo"], padx=15, pady=10)
body_frame.pack(fill="both", expand=True)

# Grid Layout: 2 colunas
body_frame.columnconfigure(0, weight=4, minsize=400) # Coluna de Controles (Esquerda)
body_frame.columnconfigure(1, weight=5, minsize=400) # Coluna de Logs (Direita)
body_frame.rowconfigure(0, weight=1)

# Notebook de Abas na Coluna Esquerda
notebook = ttk.Notebook(body_frame)
notebook.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

# Gera os ícones coloridos das abas
try:
    img_tab_divisao = gerar_icone_divisao_tab((22, 22))
    img_tab_conversor = gerar_icone_conversor_tab((22, 22))
except Exception:
    img_tab_divisao = None
    img_tab_conversor = None

# ABA 1: DIVISÃO & OCR INTELIGENTE
pane_esquerda = ttk.Frame(notebook, style="TFrame")
if img_tab_divisao:
    notebook.add(pane_esquerda, text="  Divisão & OCR ", image=img_tab_divisao, compound="left")
    notebook._img_tab1 = img_tab_divisao
else:
    notebook.add(pane_esquerda, text=" ✂️ Divisão & OCR ")

# ABA 2: CONVERSOR PDF ⇄ WORD
pane_conversor = ttk.Frame(notebook, style="TFrame")
if img_tab_conversor:
    notebook.add(pane_conversor, text="  Conversor PDF ⇄ Word ", image=img_tab_conversor, compound="left")
    notebook._img_tab2 = img_tab_conversor
else:
    notebook.add(pane_conversor, text=" 🔄 Conversor PDF ⇄ Word ")

# Card de Arquivo e Arraste
card_arquivo = ttk.Frame(pane_esquerda, style="TLabelframe")
card_arquivo.pack(fill="x", pady=(0, 8))

# Cabeçalho do painel de arquivo
header_arquivo_panel = tk.Frame(card_arquivo, bg=CORES["fundo_card"])
header_arquivo_panel.pack(fill="x", padx=15, pady=(10, 5))

lbl_titulo_arquivo = tk.Label(
    header_arquivo_panel, 
    text="1. Selecionar Arquivo PDF", 
    font=("Segoe UI", 10, "bold"), 
    foreground=CORES["cinza_texto"], 
    bg=CORES["fundo_card"]
)
lbl_titulo_arquivo.pack(side="left")

canvas_arraste = tk.Canvas(card_arquivo, height=115, bg="#F8F9FA", highlightthickness=0)
canvas_arraste.pack(fill="x", padx=15, pady=(5, 10))

rect_arraste_id = canvas_arraste.create_rectangle(4, 4, 380, 111, outline=CORES["ativo"], dash=(4, 4), width=1)
text1_id = canvas_arraste.create_text(190, 25, text="Arraste PDF(s) ou Pasta aqui", font=("Segoe UI", 11, "bold"), fill="#212529")
text2_id = canvas_arraste.create_text(190, 48, text="ou se preferir clique no botão abaixo", font=("Segoe UI", 9), fill="#6C757D")

# --- Função Auxiliar para Efeito Hover ---
def aplicar_efeito_hover(botao, cor_normal, cor_hover):
    def on_enter(e):
        if botao["state"] == "normal":
            botao.config(bg=cor_hover)
    def on_leave(e):
        if botao["state"] == "normal":
            botao.config(bg=cor_normal)
    botao.bind("<Enter>", on_enter)
    botao.bind("<Leave>", on_leave)

# Botão Limpar no cabeçalho
btn_limpar = tk.Button(
    header_arquivo_panel,
    text="Limpar Seleção",
    command=resetar_selecao,
    font=("Segoe UI", 8, "bold"),
    fg="#FFFFFF",
    bg=CORES["cinza_texto"],
    activebackground="#343A40",
    activeforeground="#FFFFFF",
    disabledforeground=CORES["texto_desativado"],
    bd=0,
    relief="flat",
    state="disabled",
    cursor="hand2",
    padx=8,
    pady=3
)
btn_limpar.pack(side="right")
aplicar_efeito_hover(btn_limpar, CORES["cinza_texto"], "#343A40")

# --- Botões Estilizados Sem Imagens (Opção 2) ---
btn_escolher = tk.Button(
    canvas_arraste, 
    text="Escolher Arquivo",
    command=escolher_arquivo,
    font=("Segoe UI", 10, "bold"),
    fg=CORES["texto_ativo"],
    bg=CORES["ativo"],
    activebackground=CORES["hover"],
    activeforeground=CORES["texto_ativo"],
    disabledforeground=CORES["texto_desativado"],
    bd=0,
    relief="flat",
    cursor="hand2",
    padx=15,
    pady=6
)
btn_escolher_window_id = canvas_arraste.create_window(190, 82, window=btn_escolher)

aplicar_efeito_hover(btn_escolher, CORES["ativo"], CORES["hover"])

# Evento para centralizar os elementos do canvas dinamicamente ao redimensionar
def on_canvas_configure(event):
    largura = event.width
    altura = event.height
    canvas_arraste.coords(rect_arraste_id, 4, 4, largura - 4, altura - 4)
    x_centro = largura // 2
    canvas_arraste.coords(text1_id, x_centro, 25)
    canvas_arraste.coords(text2_id, x_centro, 48)
    canvas_arraste.coords(btn_escolher_window_id, x_centro, 82)

canvas_arraste.bind("<Configure>", on_canvas_configure)

lbl_arquivo = tk.Label(card_arquivo, font=("Segoe UI", 10, "bold"), bg=CORES["fundo_card"], borderwidth=0)


# Card de Configuração de Destino
card_destino = ttk.LabelFrame(pane_esquerda, text=" 2. Pasta de Saída ")
card_destino.pack(fill="x", pady=(0, 8))

frame_destino_sub = tk.Frame(card_destino, bg=CORES["fundo_card"])
frame_destino_sub.pack(fill="x", padx=15, pady=10)

entry_destino = ttk.Entry(frame_destino_sub, textvariable=pasta_saida_var, font=("Segoe UI", 9))
entry_destino.pack(side="left", fill="x", expand=True, padx=(0, 10))

btn_alterar_destino = ttk.Button(frame_destino_sub, text="Alterar", command=escolher_pasta_saida)
btn_alterar_destino.pack(side="right")


# Card de Opções do OCR
card_ocr = ttk.LabelFrame(pane_esquerda, text=" 3. Configurações de OCR ")
card_ocr.pack(fill="x", pady=(0, 8))

frame_ocr_sub = tk.Frame(card_ocr, bg=CORES["fundo_card"])
frame_ocr_sub.pack(fill="x", padx=15, pady=10)

lbl_ocr_info = tk.Label(
    frame_ocr_sub, 
    text="Idioma do OCR do Windows (Native OcrEngine):", 
    font=("Segoe UI", 9), 
    bg=CORES["fundo_card"], 
    fg=CORES["cinza_texto"]
)
lbl_ocr_info.pack(anchor="w", pady=(0, 5))

idiomas_disponiveis = carregar_idiomas_ocr()
combo_idioma = ttk.Combobox(frame_ocr_sub, textvariable=idioma_ocr_var, state="readonly", font=("Segoe UI", 9))

if idiomas_disponiveis:
    combo_idioma['values'] = idiomas_disponiveis
    default_lang = obter_idioma_ocr()
    if default_lang in idiomas_disponiveis:
        combo_idioma.set(default_lang)
    else:
        combo_idioma.set(idiomas_disponiveis[0])
else:
    combo_idioma['values'] = ["Nenhum disponível"]
    combo_idioma.set("Nenhum disponível")
    combo_idioma.config(state="disabled")

combo_idioma.pack(fill="x")


# Progresso e Botão de Ação
card_acao = tk.Frame(pane_esquerda, bg=CORES["fundo"])
card_acao.pack(fill="x", pady=(5, 0))

canvas_progresso = tk.Canvas(card_acao, width=380, height=25, bg=CORES["fundo"], highlightthickness=0)
canvas_progresso.pack(pady=(0, 5))
set_progresso(0)

btn_processar = tk.Button(
    card_acao,
    text="Separar e Renomear",
    command=processar_pdf,
    font=("Segoe UI", 11, "bold"),
    fg=CORES["texto_desativado"],
    bg=CORES["desativado"],
    activebackground=CORES["hover"],
    activeforeground=CORES["texto_ativo"],
    disabledforeground=CORES["texto_desativado"],
    bd=0,
    relief="flat",
    state="disabled",
    cursor="hand2",
    padx=30,
    pady=10
)
btn_processar.pack(pady=5)

def on_enter_processar(e):
    if btn_processar["state"] == "normal":
        if btn_processar["text"] == "Cancelar":
            btn_processar.config(bg="#C92A2A")
        else:
            btn_processar.config(bg=CORES["hover"])

def on_leave_processar(e):
    if btn_processar["state"] == "normal":
        if btn_processar["text"] == "Cancelar":
            btn_processar.config(bg="#E03131")
        else:
            btn_processar.config(bg=CORES["ativo"])

btn_processar.bind("<Enter>", on_enter_processar)
btn_processar.bind("<Leave>", on_leave_processar)


# --- CONTEÚDO DA ABA 2: CONVERSOR (pane_conversor) ---
card_arquivo_conv = ttk.Frame(pane_conversor, style="TLabelframe")
card_arquivo_conv.pack(fill="x", pady=(10, 8))

# Cabeçalho do painel de arquivo do conversor
header_conv_panel = tk.Frame(card_arquivo_conv, bg=CORES["fundo_card"])
header_conv_panel.pack(fill="x", padx=15, pady=(10, 5))

lbl_tit_conv1 = tk.Label(
    header_conv_panel, 
    text="1. Selecionar Arquivo (PDF ou Word)", 
    font=("Segoe UI", 10, "bold"), 
    foreground=CORES["cinza_texto"], 
    bg=CORES["fundo_card"]
)
lbl_tit_conv1.pack(side="left")

btn_limpar_conv = tk.Button(
    header_conv_panel,
    text="Limpar Seleção",
    command=resetar_selecao_conversor,
    font=("Segoe UI", 8, "bold"),
    fg="#FFFFFF",
    bg=CORES["cinza_texto"],
    activebackground="#343A40",
    activeforeground="#FFFFFF",
    disabledforeground=CORES["texto_desativado"],
    bd=0,
    relief="flat",
    state="disabled",
    cursor="hand2",
    padx=8,
    pady=3
)
btn_limpar_conv.pack(side="right")
aplicar_efeito_hover(btn_limpar_conv, CORES["cinza_texto"], "#343A40")

# Área de arrastar e soltar (Canvas) no conversor
canvas_arraste_conv = tk.Canvas(card_arquivo_conv, height=115, bg="#F8F9FA", highlightthickness=0)
canvas_arraste_conv.pack(fill="x", padx=15, pady=(5, 10))

rect_arraste_conv_id = canvas_arraste_conv.create_rectangle(4, 4, 380, 111, outline=CORES["ativo"], dash=(4, 4), width=1)
text1_conv_id = canvas_arraste_conv.create_text(190, 25, text="Arraste PDF ou Word (.docx) aqui", font=("Segoe UI", 11, "bold"), fill="#212529")
text2_conv_id = canvas_arraste_conv.create_text(190, 48, text="ou se preferir clique no botão abaixo", font=("Segoe UI", 9), fill="#6C757D")

btn_escolher_conv = tk.Button(
    canvas_arraste_conv, 
    text="Escolher Arquivo",
    command=escolher_arquivo_conversor,
    font=("Segoe UI", 10, "bold"),
    fg=CORES["texto_ativo"],
    bg=CORES["ativo"],
    activebackground=CORES["hover"],
    activeforeground=CORES["texto_ativo"],
    disabledforeground=CORES["texto_desativado"],
    bd=0,
    relief="flat",
    cursor="hand2",
    padx=15,
    pady=6
)
btn_escolher_conv_window_id = canvas_arraste_conv.create_window(190, 82, window=btn_escolher_conv)
aplicar_efeito_hover(btn_escolher_conv, CORES["ativo"], CORES["hover"])

def on_canvas_conv_configure(event):
    largura = event.width
    altura = event.height
    canvas_arraste_conv.coords(rect_arraste_conv_id, 4, 4, largura - 4, altura - 4)
    x_centro = largura // 2
    canvas_arraste_conv.coords(text1_conv_id, x_centro, 25)
    canvas_arraste_conv.coords(text2_conv_id, x_centro, 48)
    canvas_arraste_conv.coords(btn_escolher_conv_window_id, x_centro, 82)

canvas_arraste_conv.bind("<Configure>", on_canvas_conv_configure)

lbl_arquivo_conv = tk.Label(card_arquivo_conv, font=("Segoe UI", 10, "bold"), bg=CORES["fundo_card"], borderwidth=0)

# Card de Seleção da Operação
card_operacao = ttk.Frame(pane_conversor, style="TLabelframe")
card_operacao.pack(fill="x", pady=(0, 8))

lbl_tit_conv2 = tk.Label(
    card_operacao, 
    text="2. Tipo de Operação / Conversão", 
    font=("Segoe UI", 10, "bold"), 
    foreground=CORES["cinza_texto"], 
    bg=CORES["fundo_card"]
)
lbl_tit_conv2.pack(anchor="w", padx=15, pady=(10, 5))

combo_operacao = ttk.Combobox(
    card_operacao,
    textvariable=modo_conversor_var,
    values=[
        "PDF ➔ Word (.docx)",
        "Word (.docx) ➔ PDF",
        "Juntar Múltiplos PDFs (Merge)",
        "Exportar Páginas como PNG"
    ],
    state="readonly",
    font=("Segoe UI", 10)
)
combo_operacao.pack(fill="x", padx=15, pady=(0, 10))

# Card de Pasta de Saída Conversor
card_dest_conv = ttk.Frame(pane_conversor, style="TLabelframe")
card_dest_conv.pack(fill="x", pady=(0, 8))

lbl_tit_conv3 = tk.Label(
    card_dest_conv, 
    text="3. Pasta de Destino para Salvar", 
    font=("Segoe UI", 10, "bold"), 
    foreground=CORES["cinza_texto"], 
    bg=CORES["fundo_card"]
)
lbl_tit_conv3.pack(anchor="w", padx=15, pady=(10, 5))

frame_dest_conv = tk.Frame(card_dest_conv, bg=CORES["fundo_card"])
frame_dest_conv.pack(fill="x", padx=15, pady=(0, 10))

entry_dest_conv = ttk.Entry(frame_dest_conv, textvariable=pasta_saida_conversor_var, font=("Segoe UI", 9))
entry_dest_conv.pack(side="left", fill="x", expand=True, padx=(0, 5))

btn_dest_conv = tk.Button(
    frame_dest_conv,
    text="Alterar",
    command=escolher_pasta_saida_conversor,
    font=("Segoe UI", 9, "bold"),
    fg="#FFFFFF",
    bg=CORES["cinza_texto"],
    activebackground="#343A40",
    activeforeground="#FFFFFF",
    bd=0,
    relief="flat",
    cursor="hand2",
    padx=12,
    pady=4
)
btn_dest_conv.pack(side="right")
aplicar_efeito_hover(btn_dest_conv, CORES["cinza_texto"], "#343A40")

# Container para Barra de Progresso e Botão Iniciar Conversão (Estilo idêntico à Aba Divisão)
card_acao_conv = tk.Frame(pane_conversor, bg=CORES["fundo"])
card_acao_conv.pack(fill="x", pady=(5, 0))

canvas_progresso_conv = tk.Canvas(card_acao_conv, width=380, height=25, bg=CORES["fundo"], highlightthickness=0)
canvas_progresso_conv.pack(pady=(0, 5))
set_progresso_conversor(0)

btn_iniciar_conv = tk.Button(
    card_acao_conv,
    text="Iniciar Conversão",
    command=iniciar_conversao,
    font=("Segoe UI", 11, "bold"),
    fg=CORES["texto_desativado"],
    bg=CORES["desativado"],
    activebackground=CORES["hover"],
    activeforeground=CORES["texto_ativo"],
    disabledforeground=CORES["texto_desativado"],
    bd=0,
    relief="flat",
    state="disabled",
    cursor="hand2",
    padx=30,
    pady=10
)
btn_iniciar_conv.pack(pady=5)

def on_enter_iniciar_conv(e):
    if btn_iniciar_conv["state"] == "normal":
        btn_iniciar_conv.config(bg=CORES["hover"])

def on_leave_iniciar_conv(e):
    if btn_iniciar_conv["state"] == "normal":
        btn_iniciar_conv.config(bg=CORES["ativo"])

btn_iniciar_conv.bind("<Enter>", on_enter_iniciar_conv)
btn_iniciar_conv.bind("<Leave>", on_leave_iniciar_conv)




# PANE DIREITA (Logs)
pane_direita = ttk.Frame(body_frame, style="TLabelframe")
pane_direita.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

# Cabeçalho do painel de logs
header_log_panel = tk.Frame(pane_direita, bg=CORES["fundo_card"])
header_log_panel.pack(fill="x", padx=15, pady=(15, 5))

lbl_titulo_log = tk.Label(
    header_log_panel, 
    text="Log de Processamento (Tempo Real)", 
    font=("Segoe UI", 10, "bold"), 
    foreground=CORES["cinza_texto"], 
    bg=CORES["fundo_card"]
)
lbl_titulo_log.pack(side="left")

btn_limpar_log = tk.Button(
    header_log_panel,
    text="Limpar Log",
    command=limpar_logs,
    font=("Segoe UI", 8, "bold"),
    fg="#FFFFFF",
    bg=CORES["cinza_texto"],
    activebackground="#343A40",
    activeforeground="#FFFFFF",
    disabledforeground=CORES["texto_desativado"],
    bd=0,
    relief="flat",
    state="normal",
    cursor="hand2",
    padx=8,
    pady=3
)
btn_limpar_log.pack(side="right")

aplicar_efeito_hover(btn_limpar_log, CORES["cinza_texto"], "#343A40")

# Container para text box e scrollbar
log_container = tk.Frame(pane_direita, bg=CORES["fundo_log"])
log_container.pack(fill="both", expand=True, padx=15, pady=(5, 15))

txt_log = tk.Text(
    log_container, 
    bg=CORES["fundo_log"], 
    fg=CORES["texto_log"], 
    insertbackground="white",
    font=("Consolas", 9), 
    relief="flat",
    state="disabled",
    wrap="word"
)
txt_log.pack(side="left", fill="both", expand=True)

# Custom tags coloridas para o console de log
txt_log.tag_config("erro", foreground=CORES["erro"])
txt_log.tag_config("aviso", foreground=CORES["aviso"])
txt_log.tag_config("sucesso", foreground=CORES["sucesso"])
txt_log.tag_config("titulo", foreground=CORES["titulo_log"], font=("Consolas", 9, "bold"))
txt_log.tag_config("info", foreground=CORES["info"])
txt_log.tag_config("normal", foreground=CORES["texto_log"])

scrollbar_log = ttk.Scrollbar(log_container, orient="vertical", command=txt_log.yview)
scrollbar_log.pack(side="right", fill="y")
txt_log.config(yscrollcommand=scrollbar_log.set)





# Registra Drop global de arquivos (com fallback para compatibilidade)
if hasattr(windnd, "hook_dropfiles"):
    windnd.hook_dropfiles(root, func=dropped_files, force_unicode=True)
elif hasattr(windnd, "drop_files"):
    windnd.drop_files(root, func=dropped_files, force_unicode=True)

# Inicializa exibindo uma mensagem de boas-vindas no painel de log
adicionar_log("=== BEM-VINDO AO SPLITVISION PDF ===")
adicionar_log("[INFO] Aguardando seleção de arquivo PDF para iniciar...")

root.mainloop()
