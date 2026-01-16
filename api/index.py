from flask import Flask, request, send_file, jsonify
import os
import re
import io
import json
from collections import defaultdict
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)

# Configuração de Credenciais
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = 'credentials.json'

def get_services():
    creds = None
    # 1. Tenta carregar do arquivo (Local)
    if os.path.exists(SERVICE_ACCOUNT_FILE):
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    
    # 2. Se não existir, tenta ler da Variável de Ambiente (Vercel)
    elif os.getenv('GOOGLE_CREDENTIALS'):
        import json
        creds_info = json.loads(os.getenv('GOOGLE_CREDENTIALS'))
        creds = service_account.Credentials.from_service_account_info(
            creds_info, scopes=SCOPES)
            
    if not creds:
        raise Exception("Credenciais não encontradas! (Arquivo credentials.json ou ENV GOOGLE_CREDENTIALS)")

    sheets_service = build('sheets', 'v4', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)
    return sheets_service, drive_service

def extrair_id(link):
    m = re.search(r"/d/([a-zA-Z0-9-]+)", link)
    return m.group(1) if m else link.strip()

@app.route('/')
def home():
    # Caminho robusto para o Vercel vs Local
    try:
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        template_path = os.path.join(base_dir, 'public', 'index.html')
        if not os.path.exists(template_path):
             return f"Erro: Arquivo não encontrado em {template_path}. Base: {base_dir}, Files: {str(os.listdir(base_dir))}", 404
        return send_file(template_path)
    except Exception as e:
        return f"Erro interno no Home: {str(e)}", 500

@app.route('/api/ler-colunas', methods=['POST'])
def ler_colunas():
    try:
        data = request.json
        link = data.get('link')
        aba = data.get('aba')
        
        sheets_service, _ = get_services()
        sid = extrair_id(link)
        
        # Pega o cabeçalho na linha 4 (conforme script original)
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=sid, range=f"'{aba}'!A4:ZZ4").execute()
        valores = result.get('values', [[]])
        
        if not valores or not valores[0]:
            return {"error": "Linha 4 vazia ou aba não encontrada"}, 400
            
        cabecalho = valores[0]
        # Retorna lista formatada para o frontend
        colunas = [f"{i}|{n.strip().replace('\"', '')}" for i, n in enumerate(cabecalho) if n.strip()]
        
        return {"colunas": colunas}
    except Exception as e:
        return {"error": str(e)}, 500

@app.route('/api/processar', methods=['POST'])
def processar():
    try:
        data = request.json
        link = data.get('link')
        aba = data.get('aba')
        letra_escola = data.get('letra_escola')
        filtro_excluir = data.get('filtro_excluir')
        colunas_remover_str = data.get('colunas_remover', '')
        formato = data.get('formato')

        sheets_service, drive_service = get_services()
        sid = extrair_id(link)

        # Lendo todos os dados
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=sid, range=f"'{aba}'!A:ZZ").execute()
        linhas_todas = result.get('values', [])

        if len(linhas_todas) < 4:
            return {"error": "Planilha curta demais (menos de 4 linhas)"}, 400

        cabecalho_original = linhas_todas[3] # Linha 4 (índice 3)

        # Índices
        idx_esc = ord(letra_escola.upper()) - ord('A')
        idx_ah = 33 # Coluna AH fixada no script original

        idx_remover = [int(x) for x in colunas_remover_str.split(',') if x.strip().isdigit()]
        colunas_que_ficam = [(i, c.strip()) for i, c in enumerate(cabecalho_original) if i not in idx_remover and c.strip()]

        grupos = defaultdict(list)
        termo_proibido = filtro_excluir.strip().upper() if filtro_excluir else ""

        for linha in linhas_todas[4:]: # Dados começam na linha 5
            if not linha: continue

            # Validações do script original
            valor_ah = str(linha[idx_ah]) if idx_ah < len(linha) else ""
            if not valor_ah.strip(): continue

            linha_texto = " ".join([str(c).strip().upper() for c in linha])
            if termo_proibido and termo_proibido in linha_texto: continue

            escola = str(linha[idx_esc]).strip() if idx_esc < len(linha) else "GERAL"
            # Monta linha apenas com colunas selecionadas
            dados_linha = []
            for i, _ in colunas_que_ficam:
                val = str(linha[i]).strip() if i < len(linha) else ""
                dados_linha.append(val)
            
            grupos[escola].append(dados_linha)

        if not grupos:
             return {"error": "Nenhum dado encontrado após filtros."}, 400

        # --- GERAÇÃO WORD ---
        doc = Document()
        
        # Margens
        for section in doc.sections:
            section.left_margin = Inches(0.6)
            section.right_margin = Inches(0.6)
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)

        # Cabeçalho
        h_txt = ("SECRETARIA DE ESTADO DE EDUCAÇÃO\nSECRETARIA ADJUNTA DE GESTÃO DE PESSOAS\n"
                 "DIRETORIA DE ORGANIZAÇÃO DE PESSOAL\nCOORDENADORIA DE CONTROLE E MOVIMENTAÇÃO")
        p_h = doc.add_paragraph()
        p_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_h = p_h.add_run(h_txt)
        run_h.bold = True
        run_h.font.size = Pt(11)

        # Portaria
        p_p = doc.add_paragraph()
        p_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_p = p_p.add_run("\nPORTARIA N° 000006-2026 - SAGEP")
        run_p.bold = True
        run_p.font.size = Pt(12)

        # Corpo
        texto_corpo = (
            "\nA Secretária Adjunta de Gestão de Pessoas da Secretaria de Estado de Educação, no uso das "
            "atribuições que lhe foram legalmente conferidas, especialmente aquelas previstas no art. 1º da "
            "Portaria n° 53/2025-GS/SEDUC, publicada no Diário Oficial do Estado n° 36.186, de 03 de abril de 2025, "
            "solicitado por meio do processo nº 2026/2033250, RESOLVE:\n\n"
            "Art. 1º Ficam concedidas Férias Regulamentares, nos termos da legislação vigente, aos servidores "
            "constantes nos Anexos desta Portaria, observados os respectivos períodos aquisitivos e datas de "
            "gozo estabelecidos pela IN nº 10/2025.\n\n"
            "Art. 2º Esta Portaria entra em vigor na data de publicação, produzindo os efeitos legais para cada "
            "servidor a partir da data fixada para o início das férias estabelecido nos Anexos abaixo.\n\n"
            "Publique-se. Cumpra-se.\n\n"
            "Belém (PA), 08 de janeiro de 2026."
        )
        p_c = doc.add_paragraph()
        p_c.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_c.add_run(texto_corpo).font.size = Pt(11)

        # Assinatura
        p_a = doc.add_paragraph()
        p_a.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_a = p_a.add_run("\n\nHellen Nyde da Silva e Souza\nID Funcional: 57209554-1\nSecretária Adjunta de Gestão de Pessoas")
        run_a.bold = True

        doc.add_page_break()

        # Anexos
        for escola, lista_dados in grupos.items():
            doc.add_paragraph().add_run(f"ANEXO - {escola}").bold = True

            tabela = doc.add_table(rows=1, cols=len(colunas_que_ficam))
            tabela.style = 'Table Grid'

            # Header Tabela
            hdr_cells = tabela.rows[0].cells
            for j, (_, nome) in enumerate(colunas_que_ficam):
                hdr_cells[j].text = nome
                for p in hdr_cells[j].paragraphs:
                   for r in p.runs: r.font.size = Pt(7); r.bold = True

            # Rows
            for reg in lista_dados:
                row_cells = tabela.add_row().cells
                for j, val in enumerate(reg):
                    row_cells[j].text = str(val)
                    for p in row_cells[j].paragraphs:
                        for r in p.runs: r.font.size = Pt(7)
            doc.add_paragraph()

        # --- ARMAZENAMENTO ---
        # Define caminhos ABSOLUTOS para evitar erros (especialmente no OneDrive)
        base_dir = os.getcwd()
        temp_docx_name = "temp_portaria.docx"
        temp_pdf_name = "temp_portaria.pdf"
        
        temp_docx_path = os.path.join(base_dir, temp_docx_name)
        temp_pdf_path = os.path.join(base_dir, temp_pdf_name)

        # Salva o arquivo DOCX temporário
        doc.save(temp_docx_path)
        
        # SE FOR APENAS DOCS, BAIXA DIRETO
        if formato == 'DOCS':
            return send_file(
                temp_docx_path,
                as_attachment=True,
                download_name="Portaria_2026.docx",
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        # SE FOR PDF, CONVERTE LOCALMENTE (USANDO WORD DO USUÁRIO)
        if formato == 'PDF':
            try:
                # Verifica se está rodando no Windows (Local)
                if os.name == 'nt':
                    from docx2pdf import convert
                    
                    if os.path.exists(temp_pdf_path):
                        os.remove(temp_pdf_path)

                    convert(temp_docx_path, temp_pdf_path)
                    
                    if not os.path.exists(temp_pdf_path):
                         return {"error": "Falha na conversão offline."}, 500

                    return send_file(
                        temp_pdf_path,
                        as_attachment=True,
                        download_name="Portaria_Final.pdf",
                        mimetype="application/pdf"
                    )
                else:
                    # No Vercel (Linux), não tem Word. 
                    return {"error": "A geração de PDF requer Microsoft Word e só funciona rodando no computador local. No servidor online, use a opção 'GERAR NO DOCS'."}, 400

            except Exception as e:
                return {"error": f"Erro PDF: {str(e)}"}, 500

        return {"error": "Formato inválido"}, 400

    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"error": str(e)}, 500

# Execução local
if __name__ == '__main__':
    app.run(debug=True)
