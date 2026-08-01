import os
import shutil
from pathlib import Path
from PIL import Image

ROOT = Path(__file__).resolve().parent
DIST = ROOT / "dist"
MSIX_DIR = DIST / "SplitVisionMSIX"
ASSETS_DIR = MSIX_DIR / "Assets"


def main():
    print("Iniciando a preparação dos arquivos para empacotamento MSIX...")

    # Criar diretórios
    MSIX_DIR.mkdir(parents=True, exist_ok=True)
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)

    # Copiar executável
    exe_source = DIST / "splitvision.exe"
    if not exe_source.exists():
        print(f"Erro: O executável {exe_source} não foi encontrado. Compile-o primeiro.")
        return
        
    shutil.copy2(exe_source, MSIX_DIR / "splitvision.exe")
    print(f"Executável copiado para: {MSIX_DIR / 'splitvision.exe'}")

    # Extrair e redimensionar ícones
    ico_path = ROOT / "splitvision.ico"
    if ico_path.exists():
        try:
            icon = Image.open(ico_path)
            
            # Nós precisamos de:
            # - Square150x150Logo.png (Tile médio do menu iniciar)
            # - Square44x44Logo.png (Ícone da lista de aplicativos)
            # - StoreLogo.png (Ícone para exibição na Loja)
            sizes = {
                "Square150x150Logo.png": (150, 150),
                "Square44x44Logo.png": (44, 44),
                "StoreLogo.png": (50, 50),
            }
            
            for name, size in sizes.items():
                resized = icon.resize(size, Image.Resampling.LANCZOS)
                resized.save(ASSETS_DIR / name, "PNG")
                
            # Para a Splash Screen (tela de carregamento), criamos uma tela 620x300
            # com fundo cinza claro e o ícone centralizado
            splash = Image.new("RGBA", (620, 300), color=(240, 240, 240, 255))
            icon_for_splash = icon.resize((150, 150), Image.Resampling.LANCZOS)
            
            # Centralizar o ícone
            offset = ((620 - 150) // 2, (300 - 150) // 2)
            splash.paste(icon_for_splash, offset, icon_for_splash if icon_for_splash.mode == 'RGBA' else None)
            splash.save(ASSETS_DIR / "SplashScreen.png", "PNG")
            
            print("Assets de imagens (PNG) gerados com sucesso na pasta Assets.")
        except Exception as e:
            print(f"Erro ao processar as imagens do ícone: {e}")
            return
    else:
        print("Aviso: splitvision.ico não encontrado. Imagens não foram geradas.")
        return

    # Gerar o arquivo AppxManifest.xml
    manifest_content = """<?xml version="1.0" encoding="utf-8"?>
<Package
  xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10"
  xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10"
  xmlns:rescap="http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities"
  IgnorableNamespaces="uap rescap">

  <Identity
    Name="TJDeveloper.SplitVisionPDF"
    Publisher="CN=E1237A45-3226-4503-8DDE-1C635C7A45F6"
    Version="1.3.0.0"
    ProcessorArchitecture="x64" />

  <Properties>
    <DisplayName>SplitVision</DisplayName>
    <PublisherDisplayName>TJ Developer</PublisherDisplayName>
    <Logo>Assets\\StoreLogo.png</Logo>
  </Properties>

  <Dependencies>
    <TargetDeviceFamily Name="Windows.Desktop" MinVersion="10.0.17763.0" MaxVersionTested="10.0.22000.0" />
  </Dependencies>

  <Resources>
    <Resource Language="pt-BR" />
    <Resource Language="en-US" />
  </Resources>

  <Applications>
    <Application Id="SplitVision"
      Executable="splitvision.exe"
      EntryPoint="Windows.FullTrustApplication">
      <uap:VisualElements
        DisplayName="SplitVision"
        Description="SplitVision PDF Reader and OCR Tool"
        BackgroundColor="#F0F0F0"
        Square150x150Logo="Assets\\Square150x150Logo.png"
        Square44x44Logo="Assets\\Square44x44Logo.png">
        <uap:SplashScreen Image="Assets\\SplashScreen.png" />
      </uap:VisualElements>
    </Application>
  </Applications>

  <Capabilities>
    <rescap:Capability Name="runFullTrust" />
  </Capabilities>
</Package>
"""

    manifest_path = MSIX_DIR / "AppxManifest.xml"
    with open(manifest_path, "w", encoding="utf-8") as f:
        f.write(manifest_content)

    print(f"\n[SUCESSO] AppxManifest.xml gerado em: {manifest_path}")
    print("-----------------------------------------------------------------")
    print(f"A pasta com os arquivos prontos está em:\n{MSIX_DIR}")
    print("-----------------------------------------------------------------")


if __name__ == "__main__":
    main()
