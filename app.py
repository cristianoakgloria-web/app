name: Gerar Executável macOS (Versão Intel 10.15)

on:
  push:
    branches: [ main ]

jobs:
  build:
    # macos-13 é um servidor Intel, perfeito para o teu Mac Catalina
    runs-on: macos-13 

    steps:
    - name: Checkout código
      uses: actions/checkout@v4

    - name: Configurar Python
      uses: actions/setup-python@v5
      with:
        python-version: '3.10'

    - name: Instalar Dependências
      run: |
        # Removemos o pandas da lista para máxima compatibilidade
        pip install customtkinter openpyxl pyinstaller

    - name: Gerar Executável com PyInstaller
      env:
        MACOSX_DEPLOYMENT_TARGET: 10.15
      run: |
        pyinstaller --windowed --noconsole --onefile \
        --collect-all customtkinter \
        --icon=app.icns \
        app.py

    - name: Guardar Executável (Artifact)
      uses: actions/upload-artifact@v4
      with:
        name: Sistema-Mac-Catalina-Intel
        path: dist/
