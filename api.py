from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import numpy as np
import datetime
import io
import os

# --- NOVAS IMPORTAÇÕES PARA O FIREBASE ---
import firebase_admin
from firebase_admin import credentials, firestore

app = FastAPI(title="API Torre de Controle DRE - Fotus")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

CAMINHO_BASE = "Base_DRE_Unificada.xlsx"

# --- INICIALIZAÇÃO DO FIREBASE (BACKEND) ---
# Tenta conectar usando a chave que você baixou
db = None
try:
    if os.path.exists("firebase-key.json"):
        cred = credentials.Certificate("firebase-key.json")
        if not firebase_admin._apps:
            firebase_admin.initialize_app(cred)
        db = firestore.client()
        print("✅ Conectado ao Firestore com Sucesso!")
    else:
        print("⚠️ Aviso: Arquivo 'firebase-key.json' não encontrado. O sistema lerá apenas do Excel.")
except Exception as e:
    print(f"❌ Erro ao conectar no Firebase: {e}")


@app.post("/api/upload")
async def upload_planilha(file: UploadFile = File(...)):
    try:
        conteudo = await file.read()
        with open(CAMINHO_BASE, "wb") as f:
            f.write(conteudo)
        return {"status": "sucesso", "mensagem": "Planilha importada com sucesso!"}
    except Exception as e:
        return {"status": "erro", "mensagem": str(e)}


@app.get("/api/modelo")
def baixar_modelo():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame(columns=['Empresa', 'Mês', 'Natureza', 'Fornecedor', 'Vlr. Nota', 'Tipo']).to_excel(writer, index=False, sheet_name='BD_Cust.Desp')
        pd.DataFrame(columns=['Empresa', 'Mês', 'Dt. do Faturamento', 'Vlr. Nota', 'Potencia_kWp']).to_excel(writer, index=False, sheet_name='BD_Fat')
        pd.DataFrame(columns=['Empresa', 'Mês', 'TIPO', 'CARGO', 'Quantidade Ativos']).to_excel(writer, index=False, sheet_name='HC')
    output.seek(0)
    headers = {'Content-Disposition': 'attachment; filename="Modelo_DRE_Fotus.xlsx"'}
    return StreamingResponse(output, headers=headers, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def limpar_moeda(val):
    if pd.isna(val): return 0
    if isinstance(val, (int, float)): return float(val)
    v_str = str(val).strip().upper().replace('R$', '').replace(' ', '')
    if v_str == '' or v_str == '-': return 0
    if ',' in v_str and '.' in v_str:
        v_str = v_str.replace('.', '').replace(',', '.')
    elif ',' in v_str:
        v_str = v_str.replace(',', '.')
    try: return float(v_str)
    except: return 0

# --- FUNÇÃO NOVA: BUSCAR DA NUVEM E TRANSFORMAR EM PLANILHA VIRTUAL ---
def buscar_dados_nuvem():
    df_c, df_f, df_h = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    filiais_m2 = {}

    if not db: return df_c, df_f, df_h, filiais_m2

    try:
        # 1. Puxa Custos
        custos_docs = db.collection('lancamentos_custos').stream()
        custos_list = [d.to_dict() for d in custos_docs]
        if custos_list:
            df_c = pd.DataFrame(custos_list)
            df_c.rename(columns={
                'empresa': 'Empresa', 'mes': 'Mês', 'nro_unico': 'Nro. Único',
                'nro_contrato': 'Contrato', 'nro_nota': 'Nro. Nota',
                'parceiro': 'Fornecedor', 'natureza': 'Natureza',
                'classificacao': 'Tipo', 'valor': 'Vlr. Nota'
            }, inplace=True, errors='ignore')

        # 2. Puxa Faturamento
        fat_docs = db.collection('lancamentos_faturamentos').stream()
        fat_list = [d.to_dict() for d in fat_docs]
        if fat_list:
            df_f = pd.DataFrame(fat_list)
            df_f.rename(columns={
                'empresa': 'Empresa', 'data': 'Dt. do Faturamento',
                'nro_fotus': 'Nro. Fotus', 'nro_nota': 'Nro. Nota',
                'tipo_operacao': 'Tipo Operação', 'descricao': 'Descrição (Tipo de Operação)',
                'valor': 'Vlr. Nota', 'potencia': 'Potencia_kWp',
                'peso': 'Peso bruto', 'transportadora': 'Nome Parceiro (Transportadora)',
                'modalidade': 'CIF / FOB', 'parceiro': 'Nome Parceiro (Parceiro)'
            }, inplace=True, errors='ignore')
            # Cria a coluna Mês baseada na Data para o dashboard cruzar
            if 'Dt. do Faturamento' in df_f.columns:
                df_f['Mês'] = pd.to_datetime(df_f['Dt. do Faturamento'], errors='coerce').dt.strftime('%Y-%m')

        # 3. Puxa Headcount
        hc_docs = db.collection('lancamentos_headcount').stream()
        hc_list = [d.to_dict() for d in hc_docs]
        if hc_list:
            df_h = pd.DataFrame(hc_list)
            df_h.rename(columns={
                'empresa': 'Empresa', 'cod_emp': 'COD. EMP', 'mes': 'Mês',
                'cargo': 'CARGO', 'tipo': 'TIPO', 'quantidade': 'Quantidade Ativos'
            }, inplace=True, errors='ignore')

        # 4. Puxa Filiais (Para atualizar os M2 das Facilities)
        filiais_docs = db.collection('filiais').stream()
        for f in filiais_docs:
            dados_f = f.to_dict()
            if 'nome' in dados_f and 'm2' in dados_f:
                filiais_m2[str(dados_f['nome']).upper().strip()] = float(dados_f['m2'])

    except Exception as e:
        print(f"Erro ao buscar dados do Firestore: {e}")

    return df_c, df_f, df_h, filiais_m2

def carregar_e_limpar_dados():
    df_custos_ex, df_fat_ex, df_hc_ex = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    
    # LER DADOS DO EXCEL HISTÓRICO
    if os.path.exists(CAMINHO_BASE):
        try:
            xls = pd.ExcelFile(CAMINHO_BASE, engine='openpyxl')
            sheet_map = {str(s).strip().upper(): s for s in xls.sheet_names}
            aba_custos_real = sheet_map.get('BD_CUST.DESP')
            aba_fat_real = sheet_map.get('BD_FAT')
            aba_hc_real = sheet_map.get('HC')

            if aba_custos_real: df_custos_ex = pd.read_excel(xls, sheet_name=aba_custos_real)
            if aba_fat_real: df_fat_ex = pd.read_excel(xls, sheet_name=aba_fat_real)
            if aba_hc_real: df_hc_ex = pd.read_excel(xls, sheet_name=aba_hc_real)
        except Exception:
            pass # Segue o jogo se não tiver Excel

    # LER DADOS DA NUVEM (FIRESTORE)
    df_custos_fs, df_fat_fs, df_hc_fs, filiais_m2 = buscar_dados_nuvem()

    # FUNDIR (EXCEL + NUVEM)
    df_custos = pd.concat([df_custos_ex, df_custos_fs], ignore_index=True) if not df_custos_ex.empty or not df_custos_fs.empty else pd.DataFrame()
    df_faturamento = pd.concat([df_fat_ex, df_fat_fs], ignore_index=True) if not df_fat_ex.empty or not df_fat_fs.empty else pd.DataFrame()
    df_hc = pd.concat([df_hc_ex, df_hc_fs], ignore_index=True) if not df_hc_ex.empty or not df_hc_fs.empty else pd.DataFrame()

    if df_custos.empty and df_faturamento.empty:
        raise Exception("Nenhum dado encontrado na Nuvem nem no Excel.")

    # --- LIMPEZA E FORMATAÇÃO GLOBAL (Aplica nos dados unificados) ---
    for df in [df_custos, df_faturamento, df_hc]:
        if not df.empty:
            df.columns = df.columns.astype(str).str.strip()
            col_emp = next((col for col in df.columns if 'EMPRESA' in col.upper() or 'FILIAL' in col.upper()), None)
            if col_emp:
                df.rename(columns={col_emp: 'Empresa'}, inplace=True)
                df.dropna(subset=['Empresa'], inplace=True)
                df['Empresa'] = df['Empresa'].astype(str).str.strip().str.upper()

    colunas_mes = {'Mês1': 'Mês', 'MÊS': 'Mês', 'mês': 'Mês', 'Mes': 'Mês', 'MES': 'Mês', 'Competência': 'Mês'}
    if not df_custos.empty: 
        df_custos.rename(columns=colunas_mes, inplace=True, errors='ignore')
        df_custos.rename(columns={'Descrição (Natureza)': 'Natureza', 'Nome Parceiro (Parceiro)': 'Fornecedor'}, inplace=True, errors='ignore')
        df_custos = df_custos.loc[:, ~df_custos.columns.duplicated()].copy()
    if not df_faturamento.empty: 
        df_faturamento.rename(columns=colunas_mes, inplace=True, errors='ignore')
        df_faturamento = df_faturamento.loc[:, ~df_faturamento.columns.duplicated()].copy()
    if not df_hc.empty: 
        df_hc.rename(columns=colunas_mes, inplace=True, errors='ignore')
        df_hc = df_hc.loc[:, ~df_hc.columns.duplicated()].copy()

    def formatar_mes(valor):
        if pd.isna(valor) or str(valor).strip().lower() in ['nan', 'none', 'nat', '']: return 'desconhecido'
        if isinstance(valor, datetime.datetime) or type(valor) is pd.Timestamp: return valor.strftime('%Y-%m')
        val_str = str(valor).strip().lower()
        try: return pd.to_datetime(val_str, errors='raise').strftime('%Y-%m')
        except: return val_str

    for df in [df_custos, df_faturamento, df_hc]:
        if not df.empty and 'Mês' in df.columns:
            df['Mês'] = df['Mês'].apply(formatar_mes)

    # Tratamento Custos
    if not df_custos.empty:
        col_vlr_c = next((col for col in df_custos.columns if col.upper() in ['VLR. NOTA', 'VLR NOTA', 'VALOR DA NOTA']), 'Vlr. Nota')
        if col_vlr_c in df_custos.columns: 
            df_custos['Vlr. Nota'] = df_custos[col_vlr_c].apply(limpar_moeda)
        
        def definir_gestao(row):
            if 'Tipo' in row.index and pd.notna(row['Tipo']):
                t = str(row['Tipo']).upper()
                if 'VARI' in t: return 'VARIÁVEL'
                if 'FIX' in t: return 'FIXA'
                if 'CAPEX' in t: return 'CAPEX'
            nat = str(row.get('Natureza', '')).upper()
            if any(c in nat for c in ['MOVEIS', 'MAQUINAS', 'EMPILHADEIRAS', 'FERRAMENTAS']): return 'CAPEX'
            if any(v in nat for v in ['FRETE', 'COMISSAO', 'COMISSÃO', 'IMPOSTO', 'TAXA ENTREGA']): return 'VARIÁVEL'
            return 'FIXA'

        df_custos['Tipo_Gestao'] = df_custos.apply(definir_gestao, axis=1)
        df_custos['Categoria_Despesa'] = df_custos['Tipo_Gestao'].apply(lambda x: 'CAPEX (Investimento)' if x == 'CAPEX' else 'OPEX (Operacional)')

    # Tratamento Faturamento
    if not df_faturamento.empty:
        col_vlr_f = next((col for col in df_faturamento.columns if col.upper() in ['VLR. NOTA', 'VLR NOTA', 'VALOR DA NOTA', 'FATURAMENTO']), 'Vlr. Nota')
        if col_vlr_f in df_faturamento.columns: 
            df_faturamento['Vlr. Nota'] = df_faturamento[col_vlr_f].apply(limpar_moeda)
        
        col_pot = next((col for col in df_faturamento.columns if 'pot' in str(col).lower() or 'kwp' in str(col).lower()), None)
        if col_pot:
            df_faturamento['Potencia_kWp'] = df_faturamento[col_pot].apply(limpar_moeda)
        else:
            df_faturamento['Potencia_kWp'] = 0
        
        if 'Dt. do Faturamento' in df_faturamento.columns:
            df_faturamento['Data_Real'] = pd.to_datetime(df_faturamento['Dt. do Faturamento'], errors='coerce', dayfirst=True)
            df_faturamento['Data_Formatada'] = df_faturamento['Data_Real'].dt.strftime('%Y-%m-%d').fillna('Sem Data')
        else:
            df_faturamento['Data_Formatada'] = 'Sem Data'
            df_faturamento['Data_Real'] = None

    # Tratamento HC
    if not df_hc.empty:
        col_qtd = next((col for col in df_hc.columns if 'qtd' in str(col).lower() or 'quant' in str(col).lower() or 'ativos' in str(col).lower()), None)
        if col_qtd: 
            df_hc['Quantidade Ativos'] = pd.to_numeric(df_hc[col_qtd], errors='coerce').fillna(0)
        else:
            df_hc['Quantidade Ativos'] = 0

    df_custos = df_custos.replace([np.nan, np.inf, -np.inf], None) if not df_custos.empty else pd.DataFrame()
    df_faturamento = df_faturamento.replace([np.nan, np.inf, -np.inf], None) if not df_faturamento.empty else pd.DataFrame()
    df_hc = df_hc.replace([np.nan, np.inf, -np.inf], None) if not df_hc.empty else pd.DataFrame()

    return df_custos, df_faturamento, df_hc, filiais_m2

@app.get("/api/dashboard-dados")
def get_dashboard_dados(mes: str = "Todos", data_inicio: str = None, data_fim: str = None):
    try:
        df_custos, df_faturamento, df_hc, filiais_m2 = carregar_e_limpar_dados()

        todos_meses = set()
        for df in [df_custos, df_faturamento]:
            if not df.empty and 'Mês' in df.columns: todos_meses.update(df['Mês'].dropna().unique())
        meses_disponiveis = sorted([m for m in todos_meses if m != 'desconhecido'])

        datas_disponiveis = sorted([d for d in df_faturamento['Data_Formatada'].dropna().unique() if d != 'Sem Data']) if not df_faturamento.empty and 'Data_Formatada' in df_faturamento.columns else []

        if data_inicio and data_fim and not df_faturamento.empty and 'Data_Real' in df_faturamento.columns:
            data_ini_dt = pd.to_datetime(data_inicio)
            data_fim_dt = pd.to_datetime(data_fim)
            df_faturamento = df_faturamento[(df_faturamento['Data_Real'] >= data_ini_dt) & (df_faturamento['Data_Real'] <= data_fim_dt)]

        if mes and mes != "Todos":
            mes_lower = mes.lower().strip()
            if not df_custos.empty and 'Mês' in df_custos.columns: df_custos = df_custos[df_custos['Mês'] == mes_lower]
            if not df_hc.empty and 'Mês' in df_hc.columns: df_hc = df_hc[df_hc['Mês'] == mes_lower]
            if not (data_inicio and data_fim) and not df_faturamento.empty and 'Mês' in df_faturamento.columns:
                df_faturamento = df_faturamento[df_faturamento['Mês'] == mes_lower]

        metricas_dinamicas = []
        colunas_padrao_c = ['Empresa', 'Mês', 'Natureza', 'Fornecedor', 'Vlr. Nota', 'Tipo', 'Tipo_Gestao', 'Categoria_Despesa', 'Nro. Único', 'Contrato', 'Nro. Nota']
        if not df_custos.empty:
            for col in df_custos.select_dtypes(include=[np.number]).columns:
                if col not in colunas_padrao_c and not str(col).startswith('Unnamed'):
                    t = float(df_custos[col].sum())
                    if t > 0: metricas_dinamicas.append({"origem": "Custos", "nome": col, "valor": t})

        colunas_padrao_f = ['Empresa', 'Mês', 'Dt. do Faturamento', 'Vlr. Nota', 'Potencia_kWp', 'Data_Real', 'Data_Formatada', 'Nro. Nota', 'Nro. Fotus', 'Nro. Único', 'Ordem de carga', 'Número Pedido']
        if not df_faturamento.empty:
            for col in df_faturamento.select_dtypes(include=[np.number]).columns:
                if col not in colunas_padrao_f and not str(col).startswith('Unnamed'):
                    t = float(df_faturamento[col].sum())
                    if t > 0: metricas_dinamicas.append({"origem": "Faturamento", "nome": col, "valor": t})

        vlr_custos = df_custos['Vlr. Nota'].sum() if not df_custos.empty and 'Vlr. Nota' in df_custos.columns else 0
        vlr_fat = df_faturamento['Vlr. Nota'].sum() if not df_faturamento.empty and 'Vlr. Nota' in df_faturamento.columns else 0
        c_fixo = df_custos[df_custos['Tipo_Gestao'] == 'FIXA']['Vlr. Nota'].sum() if not df_custos.empty and 'Tipo_Gestao' in df_custos.columns else 0
        c_var = df_custos[df_custos['Tipo_Gestao'] == 'VARIÁVEL']['Vlr. Nota'].sum() if not df_custos.empty and 'Tipo_Gestao' in df_custos.columns else 0

        kpis = {
            "custo_total": vlr_custos,
            "faturamento_total": vlr_fat,
            "headcount_total": int(df_hc['Quantidade Ativos'].sum()) if not df_hc.empty and 'Quantidade Ativos' in df_hc.columns else 0,
            "despesa_fixa_total": c_fixo,
            "despesa_variavel_total": c_var,
            "potencia_total": df_faturamento['Potencia_kWp'].sum() if not df_faturamento.empty and 'Potencia_kWp' in df_faturamento.columns else 0,
            "total_pedidos": len(df_faturamento) if not df_faturamento.empty else 0
        }

        custo_por_filial = df_custos.groupby('Empresa')['Vlr. Nota'].sum().reset_index().to_dict(orient='records') if not df_custos.empty and 'Empresa' in df_custos.columns else []
        top_despesas = df_custos.groupby('Natureza')['Vlr. Nota'].sum().nlargest(10).reset_index().to_dict(orient='records') if not df_custos.empty and 'Natureza' in df_custos.columns else []
        producao_diaria = df_faturamento.groupby(['Empresa', 'Mês', 'Data_Formatada']).agg(pedidos=('Vlr. Nota', 'count'), receita=('Vlr. Nota', 'sum')).reset_index().to_dict(orient='records') if not df_faturamento.empty and 'Data_Formatada' in df_faturamento.columns else []
        
        hc_detalhes = []
        if not df_hc.empty:
            col_cargo = next((col for col in df_hc.columns if 'CARGO' in str(col).upper()), None)
            col_tipo = next((col for col in df_hc.columns if 'TIPO' in str(col).upper()), None)
            if col_cargo and col_tipo:
                hc_detalhes = df_hc.groupby(['Empresa', 'Mês', col_tipo, col_cargo])['Quantidade Ativos'].sum().reset_index().rename(columns={col_tipo: 'TIPO', col_cargo: 'CARGO'}).to_dict(orient='records')

        CADASTRO_M2 = {
            "FOTUS PE": 10000, "FOTUS SP": 6500, "FOTUS PA": 2600,
            "FOTUS ES": 5000, "FOTUS SC": 4000, "FOTUS GO": 3000  
        }
        # Substitui pela Metragem Real extraída da nuvem
        if filiais_m2:
            CADASTRO_M2.update(filiais_m2)

        desempenho_filial = []
        empresas_unicas = set()
        if not df_custos.empty and 'Empresa' in df_custos.columns: empresas_unicas.update(df_custos['Empresa'].unique())
        if not df_faturamento.empty and 'Empresa' in df_faturamento.columns: empresas_unicas.update(df_faturamento['Empresa'].unique())
        
        for emp in empresas_unicas:
            cf = df_custos[(df_custos['Empresa'] == emp) & (df_custos['Tipo_Gestao'] == 'FIXA')]['Vlr. Nota'].sum() if not df_custos.empty and 'Tipo_Gestao' in df_custos.columns else 0
            cv = df_custos[(df_custos['Empresa'] == emp) & (df_custos['Tipo_Gestao'] == 'VARIÁVEL')]['Vlr. Nota'].sum() if not df_custos.empty and 'Tipo_Gestao' in df_custos.columns else 0
            ct = cf + cv 
            fat_unidade = df_faturamento[df_faturamento['Empresa'] == emp]['Vlr. Nota'].sum() if not df_faturamento.empty and 'Vlr. Nota' in df_faturamento.columns else 0
            nfs = len(df_faturamento[df_faturamento['Empresa'] == emp]) if not df_faturamento.empty else 0
            kwp = df_faturamento[df_faturamento['Empresa'] == emp]['Potencia_kWp'].sum() if not df_faturamento.empty and 'Potencia_kWp' in df_faturamento.columns else 0
            
            if not df_custos.empty and 'Natureza' in df_custos.columns:
                aluguel = df_custos[(df_custos['Empresa'] == emp) & (df_custos['Natureza'].astype(str).str.upper().str.contains('ALUGUEL'))]['Vlr. Nota'].sum()
            else:
                aluguel = 0
                
            m2_cd = CADASTRO_M2.get(emp, 1000)
            
            desempenho_filial.append({
                "Empresa": emp,
                "Faturamento_Total": fat_unidade, 
                "Custo_Fixo_Total": cf,
                "Custo_por_NF": float(cf / nfs) if nfs > 0 else 0,
                "Custo_por_kWp": float(cf / kwp) if kwp > 0 else 0,
                "Custo_Var_por_NF": float(cv / nfs) if nfs > 0 else 0,
                "Custo_Var_por_kWp": float(cv / kwp) if kwp > 0 else 0,
                "Custo_Total_por_NF": float(ct / nfs) if nfs > 0 else 0,
                "Custo_Total_por_kWp": float(ct / kwp) if kwp > 0 else 0,
                "Tamanho_M2": m2_cd,
                "Custo_Aluguel": aluguel,
                "Custo_por_M2": float(aluguel / m2_cd) if m2_cd > 0 else 0
            })

        custo_filial_mes = []
        if not df_custos.empty and 'Mês' in df_custos.columns and 'Empresa' in df_custos.columns:
            pivot_cfm = df_custos.groupby(['Mês', 'Empresa'])['Vlr. Nota'].sum().reset_index().pivot(index='Mês', columns='Empresa', values='Vlr. Nota').fillna(0).reset_index()
            custo_filial_mes = pivot_cfm.to_dict(orient='records')

        df_custos_clean = df_custos.dropna(axis=1, how='all') if not df_custos.empty else pd.DataFrame()
        detalhes_custos = df_custos_clean[df_custos_clean['Vlr. Nota'] > 0].fillna("").to_dict(orient='records') if not df_custos_clean.empty and 'Vlr. Nota' in df_custos_clean.columns else []
        
        return {
            "status": "sucesso",
            "filtros_disponiveis": {"meses": meses_disponiveis, "datas": datas_disponiveis},
            "kpis": kpis,
            "graficos": {
                "custo_por_filial": custo_por_filial,
                "top_despesas": top_despesas,
                "detalhes_custos": detalhes_custos,
                "producao_diaria": producao_diaria,
                "hc_detalhes": hc_detalhes,
                "desempenho_filial": desempenho_filial,
                "custo_filial_mes": custo_filial_mes,
                "metricas_dinamicas": metricas_dinamicas
            }
        }
    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return {"status": "erro", "mensagem": str(e)}