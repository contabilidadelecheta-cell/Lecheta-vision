from fastapi import FastAPI, UploadFile, File, HTTPException, Body
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from thefuzz import fuzz, process
from fastapi import Body
import pdfplumber
import pandas as pd
import io
import re

app = FastAPI()

# Liberação de acesso para o Lovable
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- LISTA OFICIAL LECHETA ---
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
async def exportar_excel(dados: list = Body(...)):
    try:
        df_temp = pd.DataFrame(dados)
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        workbook = writer.book
        worksheet = workbook.add_worksheet("Fechamento")
        
        # Formatos Lecheta
        fmt_h = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
        fmt_t = workbook.add_format({'bold': True, 'bg_color': '#F4CCCC', 'border': 1, 'num_format': '#,##0.00'})
        fmt_m = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
        fmt_l = workbook.add_format({'border': 1, 'align': 'center'})

        def write_block(title, category, start_row, col_idx):
            data = df_temp[df_temp["categoria"] == category]
            worksheet.write(start_row, col_idx, title, fmt_h)
            worksheet.write(start_row, col_idx + 1, "VALOR", fmt_h)
            curr = start_row + 1; total = 0
            for _, r in data.iterrows():
                worksheet.write(curr, col_idx, r["cfop"], fmt_l)
                worksheet.write(curr, col_idx + 1, r["valor"], fmt_m)
                total += r["valor"]; curr += 1
            worksheet.write(curr, col_idx, "TOTAL " + title, fmt_t)
            worksheet.write(curr, col_idx + 1, total, fmt_t)
            return curr, total

        # Construção do Layout em Colunas
        r_c, t_c = write_block("COMPRA", "COMPRA", 1, 0)
        r_dv, t_dv = write_block("DEVOLUÇÃO VENDA", "DEVOLUÇÃO VENDA", 1, 3)
        r_br, t_br = write_block("BONIFICAÇÃO RECEITA", "BONIFICAÇÃO RECEITA", r_dv + 2, 3)
        r_f, t_f = write_block("FRETE", "FRETE", 1, 6)
        r_im, t_im = write_block("IMOBILIZADO", "IMOBILIZADO", r_br, 6)
        
        st_row = max(r_c, r_br, r_im) + 4
        r_v, t_v = write_block("VENDAS", "VENDAS", st_row, 0)
        r_bd, t_bd = write_block("BONIFICAÇÃO DESPESA", "BONIFICAÇÃO DESPESA", st_row, 3)
        r_dc, t_dc = write_block("DEVOLUÇÃO COMPRA", "DEVOLUÇÃO COMPRA", st_row, 6)

        # Totais de Estoque (A14 e A15)
        worksheet.write(13, 0, "D. ESTOQUE", fmt_h)
        worksheet.write(13, 1, t_c + t_br, fmt_m)
        worksheet.write(14, 0, "C. ESTOQUE", fmt_h)
        worksheet.write(14, 1, t_bd + t_dc, fmt_m)

        worksheet.set_column('A:H', 18)
        writer.close()
        output.seek(0)
        
        return StreamingResponse(
            output, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=Fechamento_Lecheta.xlsx"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# --- NOVA ROTA: CONFERÊNCIA DE VALORES (CARD NOVO) ---
@app.post("/conferir")
async def conferir_valores(pdf_file: UploadFile = File(...), xlsx_file: UploadFile = File(...)):
    try:
        # 1. LER PDF
        pdf_content = await pdf_file.read()
        dados_pdf = []
        with pdfplumber.open(io.BytesIO(pdf_content)) as pdf:
            # Extrai o texto da primeira página
            text = pdf.pages[0].extract_text()
            # Regex que busca: Nome (letras) + Espaço + Valor (números com vírgula/ponto)
            matches = re.findall(r'([A-Za-zÀ-ÿ\s]+)\s+([\d\.,]+)', text)
            for m in matches:
                nome = m[0].strip()
                # Limpa o valor para float (remove ponto de milhar e troca vírgula por ponto)
                valor_str = m[1].replace('.', '').replace(',', '.')
                if len(nome) > 3: 
                    try: dados_pdf.append({"nome": nome, "valor": float(valor_str)})
                    except: continue

        # 2. LER EXCEL
        xlsx_content = await xlsx_file.read()
        df_xlsx = pd.read_excel(io.BytesIO(xlsx_content))
        # Padroniza nomes das colunas para maiúsculo para evitar erro de digitação
        df_xlsx.columns = [str(c).upper().strip() for c in df_xlsx.columns]
        
        # 3. CRUZAMENTO INTELIGENTE (FUZZY MATCHING)
        relatorio = []
        # Tenta encontrar colunas que contenham 'NOME' ou 'RAZAO' e 'VALOR'
        col_nome = [c for c in df_xlsx.columns if 'NOME' in c or 'RAZAO' in c or 'CLIENTE' in c][0]
        col_valor = [c for c in df_xlsx.columns if 'VALOR' in c or 'TOTAL' in c or 'SALDO' in c][0]
        
        nomes_excel = df_xlsx[col_nome].astype(str).tolist()

        for item in dados_pdf:
            # Compara o nome do PDF com todos os nomes do Excel (Semelhança de 75%)
            match = process.extractOne(item['nome'], nomes_excel, scorer=fuzz.token_set_ratio)
            
            if match and match[1] >= 75:
                nome_encontrado = match[0]
                # Pega o valor correspondente no Excel
                valor_excel = float(df_xlsx.loc[df_xlsx[col_nome] == nome_encontrado, col_valor].values[0])
                diferenca = round(item['valor'] - valor_excel, 2)
                
                relatorio.append({
                    "pdf_nome": item['nome'],
                    "excel_nome": nome_encontrado,
                    "pdf_valor": item['valor'],
                    "excel_valor": valor_excel,
                    "diferenca": diferenca,
                    "status": "OK" if abs(diferenca) < 0.05 else "ERRO"
                })
            else:
                relatorio.append({
                    "pdf_nome": item['nome'],
                    "excel_nome": "NÃO ENCONTRADO",
                    "pdf_valor": item['valor'],
                    "excel_valor": 0,
                    "diferenca": item['valor'],
                    "status": "AUSENTE"
                })

        return {"comparativo": relatorio}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro na conferência: {str(e)}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
