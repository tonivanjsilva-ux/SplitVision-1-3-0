# Próximas Implementações: Busca e Nomenclatura Personalizada por Palavra-Chave

Este documento serve como rascunho e guia de referência para a futura implementação de regras personalizadas de busca e nomenclatura no **SplitVision**.

---

## 💡 Objetivo
Permitir que o usuário defina dinamicamente qual termo (palavra-chave) deseja buscar no PDF e qual tipo de informação deve ser capturada logo em seguida para renomear os arquivos divididos. Isso elimina o acoplamento rígido com as palavras padrão (`Documento`, `O.S`, `Nome`, `Código` e `Data`).

---

## 🎨 1. Interface Gráfica (GUI) Proposta
Na coluna esquerda (`pane_esquerda`), entre as **Configurações de OCR** e o **Progresso**, adicionaremos um novo bloco chamado `4. Regra de Renomeação`:

* **Modo de Busca** (Dropdown):
  * `Padrão` (Comportamento atual: Código/Nome/Data e Doc/O.S).
  * `Personalizado` (Ativa as opções abaixo).
* **Palavra-Chave a Buscar** (Campo de texto):
  * O termo exato a ser localizado (ex: `Fatura`, `Matrícula`, `Pedido`, `Contrato`).
* **Tipo de Informação a Capturar** (Dropdown):
  * `Número (Apenas dígitos)` -> Captura números após a palavra-chave.
  * `Texto (Até o final da linha)` -> Captura nomes ou frases curtas.
  * `Data (Formatos DD/MM/AAAA)` -> Captura e padroniza datas.
* **Prefixo de Saída** (Campo de texto - Opcional):
  * Texto fixo a ser inserido antes do valor capturado (ex: `NF_`).

---

## ⚙️ 2. Lógica de Expressões Regulares (RegEx) Dinâmicas
Se o modo selecionado for o **Personalizado**, a expressão regular utilizada para varredura e busca será montada dinamicamente com base nos campos acima:

```python
import re

# Exemplo de variáveis capturadas do formulário
termo_usuario = "Fatura" # palavra-chave fornecida pelo usuário
tipo_informacao = "Número" # ou "Texto" ou "Data"

# Escapar caracteres especiais no termo do usuário
termo_escapado = re.escape(termo_usuario.strip())

if tipo_informacao == "Número":
    # Captura dígitos numéricos após o termo (com ou sem dois pontos/hífen de separação)
    regex_dinamica = re.compile(rf"\b{termo_escapado}\b\s*[:\-]?\s*(\d+)", re.IGNORECASE)

elif tipo_informacao == "Texto":
    # Captura todo o texto restante na linha após o termo
    regex_dinamica = re.compile(rf"\b{termo_escapado}\b\s*[:\-]?\s*(.+)", re.IGNORECASE)

elif tipo_informacao == "Data":
    # Captura formatos padrão de data (ex: 08/07/2026 ou 08-07-2026)
    regex_dinamica = re.compile(rf"\b{termo_escapado}\b\s*[:\-]?\s*(\d{{1,2}}[/-]\d{{1,2}}[/-]\d{{2,4}})", re.IGNORECASE)
```

---

## 🛠️ 3. Onde Mudar no Código (`splitvision.py`)

### A. Definição das Variáveis do Tkinter
Perto da linha 738, criaremos as variáveis de controle correspondentes:
```python
modo_busca_var = tk.StringVar(value="Padrao")
palavra_chave_var = tk.StringVar()
tipo_valor_var = tk.StringVar(value="Numero")
prefixo_saida_var = tk.StringVar()
```

### B. Montagem da GUI
Inserir um novo `LabelFrame` no painel esquerdo perto da linha 970:
```python
card_renomeacao = ttk.LabelFrame(pane_esquerda, text=" 4. Regra de Renomeação ")
card_renomeacao.pack(fill="x", pady=(0, 10))

# Adicionar combos e entries empilhados verticalmente
```

### C. Ajuste na Lógica de Processamento
Dentro da função `processar_pdf_thread` (linha 424), o parâmetro `regex_padrao` (linha 466) deve ser substituído pelo gerador condicional de regex dinâmico:

```python
if modo_busca_var.get() == "Personalizado":
    # Define a regex de acordo com o tipo selecionado
    ...
```

E na definição do nome de saída (linha 596), se o padrão for encontrado, o arquivo será renomeado usando a regra:
```python
valor_capturado = padrao.group(1)
prefixo = prefixo_saida_var.get().strip()
nome_arquivo = f"{prefixo}{valor_capturado}" if prefixo else valor_capturado
```
