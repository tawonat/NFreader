# LeitorFiscal

Versão integrada do backend com OCR local e leitura estruturada de NF-e, NFS-e e boletos.

## Backend

`main.py` coordena classificação, extração e correlação com o cadastro.

`ocr_engine.py` concentra Tesseract + OpenCV e retorna texto, confiança e coordenadas das palavras.

O caminho do Tesseract pode ser definido por `TESSERACT_CMD`. O sistema também procura instalações comuns no Windows.

## Instalação

No Windows:

```powershell
cd back
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```

Se o PowerShell bloquear `npm.ps1` no frontend, use `npm.cmd`.


## Ajustes desta versão
- NF-e: o melhor candidato continua preenchendo código/descrição; outras correspondências plausíveis são listadas em `Observações` como `Outras peças viáveis: ...`.
- NFS-e: leitura de rótulos de PDF com palavras coladas (ex.: `DescriçãodoServiço`) corrigida.
- NFS-e: quando a descrição declarada é genérica, o texto do bloco `SERVIÇO PRESTADO` pode ser usado como contexto apenas para classificação, sem fingir que ele é a descrição original.
