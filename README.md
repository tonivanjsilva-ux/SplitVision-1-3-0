# SplitVision 1.3.0

SplitVision é uma ferramenta para Windows que permite trabalhar com PDF, OCR e conversão entre PDF e Word.

## Funcionalidades

- Abrir e processar arquivos PDF
- Separar e salvar páginas de PDFs
- OCR para extração de texto
- Conversão entre PDF e Word (.docx)
- Arrastar e soltar arquivos na interface

## Requisitos

- Python 3.13
- Windows 10 ou superior

## Dependências

As dependências do projeto estão listadas em `requirements.txt`:

- PyPDF2
- windnd
- pypdfium2
- winocr
- Pillow

## Como rodar

1. Abra o terminal na pasta do projeto.
2. Instale as dependências:
   ```powershell
   python -m pip install -r requirements.txt
   ```
3. Execute o aplicativo:
   ```powershell
   python splitvision.py
   ```

## Como gerar o executável

Use o script `build_windows.py` para empacotar o projeto com PyInstaller:

```powershell
python build_windows.py
```

O executável gerado será salvo na pasta `dist`.

## Empacotamento MSIX

Use o script `build_msix_folder.py` após gerar o executável para preparar a pasta MSIX com o manifesto e assets:

```powershell
python build_msix_folder.py
```

## Estrutura de arquivos

- `splitvision.py` - código principal do aplicativo
- `build_windows.py` - script de criação do executável com PyInstaller
- `build_msix_folder.py` - script para preparar a pasta MSIX
- `splitvision.spec` - especificação do PyInstaller
- `requirements.txt` - dependências Python

## Contato

Desenvolvido por Tonivan Silva.
