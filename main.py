from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import pdfplumber
import pandas as pd
import io
import re
import json

app = FastAPI()

# Permite que o portal do Lovable acesse esta API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# A sua lista oficial Lecheta integrada
CONTEUDO_MAPEAMENTO = """
COMPRA
1102
1113
1403
2102
2401
2403
1556
1655
2556

DEVOLUÇÃO VENDA
1202
1411
2202
2411

FRETE
1353
1352
1360
1932
2352
2353
2932

BONIFICAÇÃO RECEITA
1901
1910
1911
1912
2910
2911

IMOBILIZADO
1551
2551

VENDAS
5102
5119
5403
5405
5502
5922
5923
6102
6108
6120
6123
6403
6404
5927
6922

BONIFICAÇÃO DESPESA
5910
6910
5911
6911

DEVOLUÇÃO COMPRA
5202
5411
5556
6202
6411
6556
"""

def get_mapeamento():
    dict_reverso = {}
    categoria_atual = ""
    for linha in CONTEUDO_MAPEAMENTO.strip().split('\n'):
        linha = linha.strip()
        if not linha: continue
        if not (linha.isdigit() and len(linha) == 4):
            categoria_atual = linha.upper()
        else:
            dict_reverso[linha] = categoria_atual
    return dict_reverso

DICT_REVERSO = get_mapeamento()

@app.post("/processar")
async def processar_pdf(file: UploadFile = File(...)):
    try:
        content = await file.read()
        achados = []
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            palavras = pdf.pages[0].extract_words()
            for i, p in enumerate(palavras):
                txt = p['text'].strip()
                if txt in DICT_REVERSO:
                    valor = 0.0
                    for j in range(1, 12):
                        if i + j < len(palavras):
                            p2 = palavras[i+j]
                            if abs(p2['top'] - p['top']) < 5:
                                if re.search(r'[\d\.,]+', p2['text']) and (',' in p2['text'] or '.' in p2['text']):
                                    valor = float(p2['text'].replace('.', '').replace(',', '.'))
                                    break
                    achados.append({"categoria": DICT_REVERSO[txt], "cfop": txt, "valor": valor})
        return {"dados": achados}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/exportar")
async def exportar_excel(dados: list):
    # Transforma o JSON recebido do Lovable em Excel usando a sua lógica de blocos
    df = pd.DataFrame(dados)
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    # ... (Aqui entra a lógica do write_block que já criamos para organizar o Excel)
    writer.close()
    output.seek(0)
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=Fechamento_Lecheta.xlsx"})