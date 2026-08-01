Aqui está uma proposta completa de descrição para a **Loja de Aplicativos (App Store / Microsoft Store)**, estruturada para atrair a atenção do usuário e explicar o funcionamento de forma clara e profissional:

---

# 📄 SplitVision: Divisor Inteligente de PDFs com OCR

### **Breve Resumo (Chamada Rápida)**
Chega de separar PDFs manualmente! O **SplitVision** é uma ferramenta inteligente que divide PDFs gigantes em arquivos menores de forma totalmente automatizada. Ele lê o conteúdo de cada página (mesmo em documentos escaneados) e organiza tudo para você em segundos.

---

### **Descrição Completa**
Você costuma receber arquivos PDF enormes contendo dezenas de notas fiscais, contratos ou Ordens de Serviço (OS) misturados? Separar e renomear esses documentos um por um é exaustivo e toma muito tempo.

O **SplitVision** resolve isso de forma simples e inteligente. Equipado com tecnologia de **reconhecimento visual (OCR)** e inteligência de busca, ele analisa o conteúdo das páginas, identifica onde cada documento começa e termina, e salva tudo separado e renomeado automaticamente.

---

### **✨ Principais Recursos e Destaques:**

* 📁 **Arrastar e Soltar (Drag & Drop):** Interface extremamente simples. Basta arrastar o seu PDF para dentro do aplicativo para começar.
* 👁️ **Tecnologia OCR Integrada:** Consegue ler e identificar textos mesmo em documentos digitalizados, papéis escaneados ou fotos de baixa qualidade.
* 🔄 **Correção Automática de Rotação:** Se alguma página foi escaneada de ponta-cabeça ou de lado, o app detecta o erro, lê o texto e **corrige a orientação** da página no PDF final.
* 🧠 **Agrupamento Inteligente:** Se um contrato ou documento possui mais de uma página, o SplitVision entende que elas fazem parte do mesmo arquivo e não as separa incorretamente.
* 🏷️ **Nomenclatura Automatizada:** Extrai informações do próprio texto (como Código, Nome do Cliente, Data ou número da OS) e usa esses dados para nomear os arquivos de saída de forma organizada.
* 🔒 **100% Seguro e Offline:** O processamento é feito localmente no seu computador. Seus documentos confidenciais nunca são enviados para a internet.

---

### **⚙️ Como funciona em 3 passos simples:**

1. **Importe o arquivo:** Abra o SplitVision e arraste o seu PDF principal.
2. **Escolha o destino:** Defina em qual pasta do seu computador você deseja salvar os novos arquivos divididos.
3. **Clique em Processar:** O aplicativo fará a leitura inteligente de todas as páginas, separando-as em arquivos individuais estruturados e nomeando-os de forma automática (ex: `1024 - João Silva - 04-07-2026.pdf`).

As principais novidades e implementações da nova versão do **SplitVision (1.2.0)** trazem melhorias focadas em **otimização de desempenho**, **redução de tamanho** e **segurança (fim de falsos positivos de antivírus)**.

Abaixo estão os detalhes das novas implementações:

### 1. 🚀 Substituição de Dependências Externas por Soluções In-Process
* **Fim do Poppler:** A renderização de páginas de PDF agora é executada diretamente em processo usando a biblioteca (Google PDFium wrapper). O utilitário externo `pdftoppm.exe` e a biblioteca `pdf2image` foram removidos.
* **Fim do Tesseract OCR:** A engine do Tesseract executável e o wrapper `pytesseract` foram removidos. Agora, o aplicativo utiliza o **Windows Native OCR** (`Windows.Media.Ocr`) por meio das bibliotecas `winocr` e `winrt`, acessando a API nativa do próprio Windows 10/11.

### 2. 🛡️ Segurança e Fim de Falsos Positivos
* **Sem Subprocessos (`subprocess.Popen`):** Com as novas bibliotecas rodando de forma 100% interna, o aplicativo não precisa mais disparar processos filhos em segundo plano. Isso elimina o comportamento que gerava falsos positivos de vírus/malware por heurística e impede que janelas ocultas de console fiquem abrindo.

### 3. 📉 Redução Drástica de Tamanho
Ao utilizar recursos nativos do sistema e pacotes internos do Python, conseguimos realizar uma grande limpeza no workspace:
* A pasta `poppler/` (~24 MB) foi removida.
* A pasta `tesseract/` (~83 MB) foi removida.
* **Tamanho total da distribuição standalone (`.dist`):** Reduzido de **156.20 MB** para **64.16 MB** (uma redução de aproximadamente **59%** no tamanho final do programa compilado).

### 4. 🧠 Idioma Inteligente e Logs Aprimorados
* **Detecção Dinâmica de Idioma:** A função `obter_idioma_ocr()` no arquivo  verifica as linguagens de OCR ativas no Windows do usuário, priorizando o português (`pt-BR`).
* **Melhoria nos Logs:** O arquivo `log_processamento.txt` agora detalha se cada página lida utilizou o texto nativo digitalizado ou se passou pelo `OCR Windows (pt-BR)`.

---

### **Ideal para:**
* Escritórios de contabilidade e advocacia.
* Setores administrativos, financeiros e de logística.
* Qualquer profissional que lide diariamente com digitalização de lotes de documentos.