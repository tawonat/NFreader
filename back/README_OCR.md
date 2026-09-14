# Leitor Fiscal - OCR e cadastros fixos

O backend processa somente **NF-e** e **NFS-e**.

## Cadastros fixos

Os dois arquivos ficam em `back/cadastro/` e são carregados automaticamente:

- `Produtos Cadastrados no TOTVS.xlsx` -> usado nas NF-e
- `Serviços Cadastrados no TOTVS.xlsx` -> usado nas NFS-e

Eles não precisam mais ser enviados pela interface.

## Saída

O Excel final possui:

- `Produtos (NF-e)`
- `Serviços (NFS-e)`
- `Outros`

Boletos, CT-e e documentos não identificados vão para `Outros` com o motivo.

## Quantidade

- NF-e: extrai somente quantidade do item.
- NFS-e: extrai quantidade de horas quando aparecer no texto (`6 horas`, `6 h`, etc.); quando não houver indicação, usa `1`.

## Confiança

- 81-100%: sem preenchimento
- 70-80%: amarelo
- 50-69%: laranja
- abaixo de 50%: vermelho

Correspondências muito próximas entre dois itens também são marcadas para revisão.

## Tesseract

O `ocr_engine.py` usa Tesseract local + OpenCV e retorna texto, confiança e coordenadas. A variável `TESSERACT_CMD` pode ser usada para informar o caminho do executável.
