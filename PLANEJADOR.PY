"""
Script de Auditoria e Validação de Matrizes Curriculares
Projeto: Semester Planning
Tecnologias: Google OR-Tools (CP-SAT), Pandas, OpenPyXL
"""

import pandas as pd
from ortools.sat.python import cp_model
import sys
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Configurações de Estilos do Excel (Paleta de Cores Executiva)
COR_CABECALHO_NAVY = "1F3864"
COR_ALERTA_VERMELHO = "FCE4D6"
COR_SUCESSO_VERDE = "E2EFDA"
COR_ZEBRA_CINZA = "F2F2F2"

_FILL_CABECALHO = PatternFill("solid", fgColor=COR_CABECALHO_NAVY)
_FILL_VERMELHO = PatternFill("solid", fgColor="FADBD8")
_FILL_VERDE = PatternFill("solid", fgColor="D5F5E3")
_FILL_ZEBRA = PatternFill("solid", fgColor=COR_ZEBRA_CINZA)

_FONTE_TITULO = Font(name="Arial", size=14, bold=True, color=COR_CABECALHO_NAVY)
_FONTE_SUBTITULO = Font(name="Arial", size=10, italic=True, color="595959")
_FONTE_CABECALHO = Font(name="Arial", size=10, bold=True, color="FFFFFF")
_FONTE_NORMAL = Font(name="Arial", size=10)
_FONTE_NEGRITO = Font(name="Arial", size=10, bold=True)
_FONTE_VERMELHA = Font(name="Arial", size=10, bold=True, color="9C0006")
_FONTE_VERDE = Font(name="Arial", size=10, bold=True, color="006100")

_BORDA_FINA = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9")
)

# ---------------------------------------------------------------------------
# EXIGÊNCIA DE CARGA HORÁRIA DE PRÁTICA CONFORME AS DCNs (consulta em jul/2026)
#
# Regra: a DCN estabelece PISO de formação prática. Cursos cuja DCN não prevê
# componente prático não deveriam alocar CH de prática — essa CH é excedente.
#
# Padrão conservador: todo curso é tratado como EXIGINDO prática. Só entram na
# lista abaixo os cursos com evidência normativa de que a prática NÃO é exigida,
# para não gerar acusação indevida de não conformidade.
#
# Cursos com exigência confirmada (não listados abaixo, portanto):
#   - Enfermagem ....... Res. CNE/CES 1/2026, art. 8º §5º (teórico-práticas >= 20%)
#                        e art. 14 (estágio supervisionado >= 500h)
#   - Engenharias ...... Res. CNE/CES 2/2019, art. 8º (práticas e laboratoriais
#                        obrigatórias, essenciais em Física, Química e Informática)
#   - Licenciaturas .... Res. CNE/CP 2/2019 (400h de prática como componente curricular)
#   - Administração .... Res. CNE/CES 5/2021, art. 4º, V ("atividades práticas
#                        supervisionadas obrigatórias" em todas as IES)
#   - Ciências Contábeis  Res. CNE/CES 1/2024, art. 5º (estágio supervisionado OU
#                        laboratório de simulações em práticas contábeis)
CURSOS_SEM_EXIGENCIA_PRATICA = {
    'Economia',               # Res. CNE/CES 4/2007: estágio facultativo, sem exigência de prática
    'Ciência da Computação',  # Res. CNE/CES 5/2016, art. 6º §4º: práticas "poderão" compor a CH;
    'Sistemas de Informação',  # art. 7º §1º: obrigatoriedade do estágio fica a critério da IES
}


def exige_pratica(curso):
    """Indica se a DCN do curso prevê componente de formação prática."""
    return curso not in CURSOS_SEM_EXIGENCIA_PRATICA


def carregar_dados(nome_arquivo='Dados_Exportacao_CSV.csv'):
    """Lê os dados granulares do arquivo CSV de forma robusta."""
    diretorio_script = os.path.dirname(os.path.abspath(__file__))
    caminho_arquivo = os.path.join(diretorio_script, nome_arquivo)
    
    try:
        try:
            df = pd.read_csv(caminho_arquivo, sep=';', encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(caminho_arquivo, sep=';', encoding='latin-1')
        
        colunas_inteiras = [
            'Carga_Horaria', 'CH_Teorica', 'CH_Pratica_Sala', 'CH_Pratica_Lab', 
            'CH_EAD', 'CH_ASM', 'CH_Estagio', 'CH_Extensao'
        ]
        for col in colunas_inteiras:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
            else:
                df[col] = 0
            
        return df
    except FileNotFoundError:
        print(f"  [Erro]: O arquivo '{caminho_arquivo}' não foi encontrado.")
        return None
    except Exception as e:
        print(f"  [Erro] ao carregar o arquivo: {e}")
        return None

def validar_todas_matrizes(df, indice_inicial=1, total_geral=None):
    """Aplica as regras do MEC sobre a carga horária usando CP-SAT, exibindo o passo a passo com análise 100%.

    indice_inicial/total_geral permitem numerar corretamente o passo a passo quando
    a validação é chamada em lotes (usado pela interface gráfica para exibir progresso).
    """
    lista_matrizes = df['ID_Matriz'].unique()
    total_exibicao = total_geral if total_geral is not None else len(lista_matrizes)
    resultados = []
    
    # Parâmetros Regulatórios Gerais do MEC
    LIMITE_TETO_EAD = 40      # Máximo de 40% de EAD para cursos presenciais
    LIMITE_MIN_EXTENSAO = 10  # Mínimo de 10% de Extensão para todos os cursos
    LIMITE_TETO_ASM = 20      # Máximo de 20% de ASM (Atividade de Sala Multidisciplinar)

    # Parâmetros das DCNs específicas por área
    MIN_ESTAGIO_MEDICINA = 35   # % mínimo de Estágio (Medicina)
    MIN_ESTAGIO_SAUDE = 20      # % mínimo de Estágio (demais cursos da Saúde)
    MIN_ESTAGIO_ENGENHARIAS = 160  # horas absolutas de Estágio (Engenharias)
    MIN_ESTAGIO_LICENCIATURAS = 400  # horas absolutas de Estágio (Licenciaturas)
    MIN_PRATICA_LICENCIATURAS = 400  # horas absolutas de Prática Pedagógica (Licenciaturas)
    
    for idx, id_matriz in enumerate(lista_matrizes, indice_inicial):
        dados_matriz = df[df['ID_Matriz'] == id_matriz]
        modalidade = dados_matriz['Modalidade'].iloc[0] if 'Modalidade' in dados_matriz.columns else 'Não Identificada'
        curso = dados_matriz['Curso'].iloc[0] if 'Curso' in dados_matriz.columns else 'Curso Não Especificado'
        area = dados_matriz['Area'].iloc[0] if 'Area' in dados_matriz.columns else 'Área Não Especificada'
        
        print(f"\n({idx}/{total_exibicao}) Analisando Matriz: {id_matriz} | Curso: {curso} | Modalidade: {modalidade}")
        
        # Calcular os totais a partir dos dados do dataframe
        total_ch_alvo = int(dados_matriz['Carga_Horaria'].sum())
        soma_teorica = int(dados_matriz['CH_Teorica'].sum())
        soma_pratica_sala = int(dados_matriz['CH_Pratica_Sala'].sum())
        soma_pratica_lab = int(dados_matriz['CH_Pratica_Lab'].sum())
        soma_ead = int(dados_matriz['CH_EAD'].sum())
        soma_asm = int(dados_matriz['CH_ASM'].sum())
        soma_estagio = int(dados_matriz['CH_Estagio'].sum())
        soma_extensao = int(dados_matriz['CH_Extensao'].sum())
        
        # Conferência matemática total
        soma_componentes = (soma_teorica + soma_pratica_sala + soma_pratica_lab + 
                            soma_ead + soma_asm + soma_estagio + soma_extensao)
        diferenca_soma = abs(total_ch_alvo - soma_componentes)
        
        # Percentuais reais
        pct_teorica = (soma_teorica / total_ch_alvo * 100) if total_ch_alvo > 0 else 0.0
        pct_pratica_sala = (soma_pratica_sala / total_ch_alvo * 100) if total_ch_alvo > 0 else 0.0
        pct_pratica_lab = (soma_pratica_lab / total_ch_alvo * 100) if total_ch_alvo > 0 else 0.0
        pct_ead = (soma_ead / total_ch_alvo * 100) if total_ch_alvo > 0 else 0.0
        pct_asm = (soma_asm / total_ch_alvo * 100) if total_ch_alvo > 0 else 0.0
        pct_estagio = (soma_estagio / total_ch_alvo * 100) if total_ch_alvo > 0 else 0.0
        pct_extensao = (soma_extensao / total_ch_alvo * 100) if total_ch_alvo > 0 else 0.0
        pct_soma_total = (soma_componentes / total_ch_alvo * 100) if total_ch_alvo > 0 else 0.0
        
        print(f"  * Distribuição da Carga Horária:")
        print(f"    - SOMA DA GRADE:      {soma_componentes: >5}h ({pct_soma_total: >6.2f}%) | Total Alvo: {total_ch_alvo}h")
        
        # Configurar modelo de Programação por Restrições (CP-SAT)
        print("  * Inicializando Solver CP-SAT...")
        model = cp_model.CpModel()
        
        vars_total, vars_teorica, vars_pratica_sala = [], [], []
        vars_pratica_lab, vars_ead, vars_asm = [], [], []
        vars_estagio, vars_extensao = [], []
        
        # Declaração de variáveis
        for _, row in dados_matriz.iterrows():
            id_disc = row['ID_Disciplina']
            vars_total.append(model.NewIntVar(row['Carga_Horaria'], row['Carga_Horaria'], f'tot_{id_disc}'))
            vars_teorica.append(model.NewIntVar(row['CH_Teorica'], row['CH_Teorica'], f'teo_{id_disc}'))
            vars_pratica_sala.append(model.NewIntVar(row['CH_Pratica_Sala'], row['CH_Pratica_Sala'], f'ps_{id_disc}'))
            vars_pratica_lab.append(model.NewIntVar(row['CH_Pratica_Lab'], row['CH_Pratica_Lab'], f'pl_{id_disc}'))
            vars_ead.append(model.NewIntVar(row['CH_EAD'], row['CH_EAD'], f'ead_{id_disc}'))
            vars_asm.append(model.NewIntVar(row['CH_ASM'], row['CH_ASM'], f'asm_{id_disc}'))
            vars_estagio.append(model.NewIntVar(row['CH_Estagio'], row['CH_Estagio'], f'est_{id_disc}'))
            vars_extensao.append(model.NewIntVar(row['CH_Extensao'], row['CH_Extensao'], f'ext_{id_disc}'))

        print("  * Inserindo restrições regulatórias e matemáticas no modelo:")
        
        # R1 e R2: Integridade e Consistência
        model.Add(sum(vars_total) == total_ch_alvo)
        model.Add(sum(vars_teorica) + sum(vars_pratica_sala) + sum(vars_pratica_lab) + 
                  sum(vars_ead) + sum(vars_asm) + sum(vars_estagio) + sum(vars_extensao) == total_ch_alvo)
        
        # R3: Curricularização da Extensão
        soma_extensao_var = sum(vars_extensao) if vars_extensao else 0
        model.Add(soma_extensao_var * 100 >= LIMITE_MIN_EXTENSAO * total_ch_alvo)
        print(f"    - Extensão: Mínimo de {LIMITE_MIN_EXTENSAO}% requerido.")
        
        # R4: Limite EAD
        if modalidade == 'Presencial':
            soma_ead_var = sum(vars_ead) if vars_ead else 0
            model.Add(soma_ead_var * 100 <= LIMITE_TETO_EAD * total_ch_alvo)
            print(f"    - EAD: Teto de {LIMITE_TETO_EAD}% para cursos Presenciais.")
            
        # R5: Regras de Estágio e Práticas (DCNs Específicas)
        soma_estagio_var = sum(vars_estagio) if vars_estagio else 0
        soma_pratica_sala_var = sum(vars_pratica_sala) if vars_pratica_sala else 0
        
        if area == 'Saúde':
            if curso == 'Medicina':
                model.Add(soma_estagio_var * 100 >= MIN_ESTAGIO_MEDICINA * total_ch_alvo)
                print(f"    - DCN Saúde/Medicina: Estágio >= {MIN_ESTAGIO_MEDICINA}%.")
            else:
                model.Add(soma_estagio_var * 100 >= MIN_ESTAGIO_SAUDE * total_ch_alvo)
                print(f"    - DCN Saúde: Estágio >= {MIN_ESTAGIO_SAUDE}%.")
        elif area == 'Engenharias':
            model.Add(soma_estagio_var >= MIN_ESTAGIO_ENGENHARIAS)
            print(f"    - DCN Engenharias: Estágio >= {MIN_ESTAGIO_ENGENHARIAS}h absolutas.")
        elif area == 'Licenciaturas':
            model.Add(soma_estagio_var >= MIN_ESTAGIO_LICENCIATURAS)
            model.Add(soma_pratica_sala_var >= MIN_PRATICA_LICENCIATURAS)
            print(f"    - DCN Licenciaturas: Estágio >= {MIN_ESTAGIO_LICENCIATURAS}h e "
                  f"Prática Pedagógica >= {MIN_PRATICA_LICENCIATURAS}h absolutas.")

        # R7: CH de prática indevida — curso cuja DCN não prevê formação prática
        # não deve alocar CH de prática (sala ou laboratório)
        soma_pratica_total_var = sum(vars_pratica_sala) + sum(vars_pratica_lab)
        if not exige_pratica(curso):
            model.Add(soma_pratica_total_var == 0)
            print(f"    - Prática: DCN de {curso} não prevê formação prática; "
                  f"CH de prática deve ser zero.")
            
        # R6: Limite ASM (Máximo de 20% apenas para cursos que exigem ASM, ou seja, soma_asm > 0)
        soma_asm_var = sum(vars_asm) if vars_asm else 0
        if soma_asm > 0:
            model.Add(soma_asm_var * 100 <= LIMITE_TETO_ASM * total_ch_alvo)
            print(f"    - ASM: Teto de {LIMITE_TETO_ASM}% requerido (Curso exige ASM).")
        else:
            print("    - ASM: Ignorada (Curso não exige ASM).")

        # Resolver o modelo
        solver = cp_model.CpSolver()
        status = solver.Solve(model)
        
        # Geração de Relatório Analítico (Verificações pós-solver para apontar os motivos exatos)
        motivos = []
        status_ead, status_ext, status_estagio, status_asm = "Ok", "Ok", "Ok", "Ok"
        # Prática: piso de 400h nas Licenciaturas; nos cursos cuja DCN não prevê
        # formação prática, qualquer CH de prática é excedente
        soma_pratica_total = soma_pratica_sala + soma_pratica_lab
        status_pratica = "Ok"
        
        if modalidade == 'Presencial' and pct_ead > LIMITE_TETO_EAD:
            status_ead = f"Violada ({pct_ead:.1f}%)"
            motivos.append(f"CH EAD ({pct_ead:.1f}%) excede o limite de {LIMITE_TETO_EAD}%.")
            
        if pct_extensao < LIMITE_MIN_EXTENSAO:
            status_ext = f"Violada ({pct_extensao:.1f}%)"
            motivos.append(f"CH Extensão ({pct_extensao:.1f}%) inferior ao mínimo de {LIMITE_MIN_EXTENSAO}%.")
            
        if soma_asm > 0:
            if pct_asm > LIMITE_TETO_ASM:
                status_asm = f"Violada ({pct_asm:.1f}%)"
                motivos.append(f"CH ASM ({pct_asm:.1f}%) excede o teto de {LIMITE_TETO_ASM}%.")
            else:
                status_asm = f"Ok ({pct_asm:.1f}%)"
        else:
            status_asm = "Não Exigido (0%)"
            
        if diferenca_soma > 0:
            motivos.append(f"Soma dos componentes diverge da CH Total (Diferença: {total_ch_alvo - soma_componentes}h).")
            
        # Avaliação detalhada de Estágio e Prática
        if area == 'Saúde':
            if curso == 'Medicina' and pct_estagio < MIN_ESTAGIO_MEDICINA:
                status_estagio = f"Violada ({pct_estagio:.1f}%)"
                motivos.append(f"CH Estágio ({pct_estagio:.1f}%) inferior ao mínimo de "
                               f"{MIN_ESTAGIO_MEDICINA}% (Medicina).")
            elif curso != 'Medicina' and pct_estagio < MIN_ESTAGIO_SAUDE:
                status_estagio = f"Violada ({pct_estagio:.1f}%)"
                motivos.append(f"CH Estágio ({pct_estagio:.1f}%) inferior ao mínimo de "
                               f"{MIN_ESTAGIO_SAUDE}% (Saúde).")
        elif area == 'Engenharias' and soma_estagio < MIN_ESTAGIO_ENGENHARIAS:
            status_estagio = f"Violada ({soma_estagio}h)"
            motivos.append(f"CH Estágio ({soma_estagio}h) inferior ao mínimo de "
                           f"{MIN_ESTAGIO_ENGENHARIAS}h (Engenharias).")
        elif area == 'Licenciaturas':
            if soma_estagio < MIN_ESTAGIO_LICENCIATURAS:
                status_estagio = f"Violada ({soma_estagio}h)"
                motivos.append(f"CH Estágio ({soma_estagio}h) inferior ao mínimo de "
                               f"{MIN_ESTAGIO_LICENCIATURAS}h (Licenciaturas).")
            if soma_pratica_sala < MIN_PRATICA_LICENCIATURAS:
                status_pratica = f"Violada ({soma_pratica_sala}h)"
                motivos.append(f"CH Prática Pedagógica ({soma_pratica_sala}h) inferior ao mínimo de "
                               f"{MIN_PRATICA_LICENCIATURAS}h (Licenciaturas).")
            else:
                status_pratica = f"Ok ({soma_pratica_sala}h)"

        # Prática indevida: DCN do curso não prevê formação prática
        if not exige_pratica(curso):
            if soma_pratica_total > 0:
                status_pratica = f"Violada ({soma_pratica_total}h excedente)"
                motivos.append(f"CH de Prática ({soma_pratica_total}h) alocada em curso cuja DCN não "
                               f"prevê formação prática ({curso}) — carga horária excedente.")
            else:
                status_pratica = "Não Exigido"

        status_soma = "Erro" if diferenca_soma > 0 else "Ok"
        
        # Avaliação final
        if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
            if diferenca_soma == 0:
                resultado = "VÁLIDA (Regras MEC Atendidas)"
                parecer = "Matriz em conformidade regulatória plena."
                print(f"  --> RESULTADO: VÁLIDA")
            else:
                resultado = "INVIÁVEL (Não Conforme)"
                parecer = " | ".join(motivos)
                print(f"  --> RESULTADO: INVIÁVEL (Soma Matemática)")
        else:
            resultado = "INVIÁVEL (Não Conforme)"
            parecer = " | ".join(motivos)
            print(f"  --> RESULTADO: INVIÁVEL")
            for m in motivos:
                print(f"      [!] NÃO CONFORMIDADE: {m}")
            
        resultados.append({
            'ID_Matriz': id_matriz,
            'Curso': curso,
            'Area': area,
            'Modalidade': modalidade,
            'CH_Total': total_ch_alvo,
            'CH_Teorica': soma_teorica, 'Pct_Teorica': pct_teorica,
            'CH_Pratica_Sala': soma_pratica_sala, 'Pct_Pratica_Sala': pct_pratica_sala,
            'CH_Pratica_Lab': soma_pratica_lab, 'Pct_Pratica_Lab': pct_pratica_lab,
            'CH_Pratica_Total': soma_pratica_total,
            'Pct_Pratica_Total': (soma_pratica_total / total_ch_alvo * 100) if total_ch_alvo > 0 else 0.0,
            'Exige_Pratica': 'Sim' if exige_pratica(curso) else 'Não',
            'CH_EAD': soma_ead, 'Pct_EAD': pct_ead,
            'CH_ASM': soma_asm, 'Pct_ASM': pct_asm,
            'CH_Estagio': soma_estagio, 'Pct_Estagio': pct_estagio,
            'CH_Extensao': soma_extensao, 'Pct_Extensao': pct_extensao,
            'Pct_Soma_Total': pct_soma_total,
            'Status_EAD': status_ead,
            'Status_Extensao': status_ext,
            'Status_Estagio': status_estagio,
            'Status_Pratica': status_pratica,
            'Status_ASM': status_asm,
            'Status_Soma': status_soma,
            'Status_Geral': resultado,
            'Parecer': parecer
        })
        
    return resultados

def _rotulo_status(status_interno):
    """Converte o status interno da regra no rótulo exibido no relatório."""
    if "Violada" in status_interno:
        return "REPROVADO"
    if "Não Exigido" in status_interno:
        return "N/A"
    return "OK"


def gerar_relatorio_excel(resultados, caminho_saida):
    """Gera um arquivo Excel estruturado com a análise 100% das cargas horárias e novas DCNs."""
    wb = Workbook()
    
    # --- ABA 1: RESUMO EXECUTIVO ---
    ws_resumo = wb.active
    ws_resumo.title = "Resumo Executivo"
    ws_resumo.views.sheetView[0].showGridLines = True
    
    ws_resumo.cell(row=2, column=2, value="Relatório de Auditoria Regulatória de Matrizes Curriculares").font = _FONTE_TITULO
    ws_resumo.cell(row=3, column=2, value="Análise percentual completa de todos os componentes da grade e DCNs").font = _FONTE_SUBTITULO
    
    total_matrizes = len(resultados)
    validas = sum(1 for r in resultados if "VÁLIDA" in r['Status_Geral'])
    invalidas = total_matrizes - validas
    pct_conformidade = (validas / total_matrizes) if total_matrizes > 0 else 0.0
    
    kpis = [
        ("Total de Matrizes", total_matrizes, "1F3864", "FFFFFF"),
        ("Matrizes Conformes", validas, "2E7D32", "FFFFFF"),
        ("Matrizes Não Conformes", invalidas, "C62828", "FFFFFF"),
        ("Taxa de Conformidade", f"{pct_conformidade*100:.1f}%", "D1A110", "000000")
    ]
    
    for idx, (label, valor, bg_color, fg_color) in enumerate(kpis):
        col_start = 2 + (idx * 2)
        fill_card = PatternFill("solid", fgColor=bg_color)
        font_label = Font(name="Arial", size=9, bold=True, color=fg_color)
        font_value = Font(name="Arial", size=16, bold=True, color=fg_color)
        
        ws_resumo.cell(row=5, column=col_start, value=label).font = font_label
        ws_resumo.cell(row=5, column=col_start).fill = fill_card
        ws_resumo.cell(row=5, column=col_start).alignment = Alignment(horizontal="center", vertical="center")
        
        ws_resumo.cell(row=6, column=col_start, value=valor).font = font_value
        ws_resumo.cell(row=6, column=col_start).fill = fill_card
        ws_resumo.cell(row=6, column=col_start).alignment = Alignment(horizontal="center", vertical="center")
        
        for r in [5, 6]:
            for c in [col_start, col_start + 1]:
                cell = ws_resumo.cell(row=r, column=c)
                cell.fill = fill_card
                cell.border = _BORDA_FINA
                
        ws_resumo.merge_cells(start_row=5, start_column=col_start, end_row=5, end_column=col_start + 1)
        ws_resumo.merge_cells(start_row=6, start_column=col_start, end_row=6, end_column=col_start + 1)
        
    ws_resumo.cell(row=8, column=2, value="Visão Geral de Todas as Matrizes").font = Font(name="Arial", size=11, bold=True, color=COR_CABECALHO_NAVY)
    
    cabecalhos = [
        "ID Matriz", "Área", "Curso", "Modalidade", "CH Total (h)", 
        "% Teórica", "% Prática Sala", "% Prática Lab", "% EAD", "% ASM", "% Estágio", "% Extensão", "Soma %", "Status Geral"
    ]
    linha_iniciar = 9
    
    for col_idx, cabecalho in enumerate(cabecalhos, start=2):
        cell = ws_resumo.cell(row=linha_iniciar, column=col_idx, value=cabecalho)
        cell.font = _FONTE_CABECALHO
        cell.fill = _FILL_CABECALHO
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _BORDA_FINA
    
    ws_resumo.row_dimensions[linha_iniciar].height = 28
    
    r_atual = linha_iniciar + 1
    for item in resultados:
        ws_resumo.row_dimensions[r_atual].height = 20
        row_values = [
            item['ID_Matriz'], item['Area'], item['Curso'], item['Modalidade'], item['CH_Total'],
            item['Pct_Teorica'] / 100.0, item['Pct_Pratica_Sala'] / 100.0, item['Pct_Pratica_Lab'] / 100.0,
            item['Pct_EAD'] / 100.0, item['Pct_ASM'] / 100.0, item['Pct_Estagio'] / 100.0,
            item['Pct_Extensao'] / 100.0, item['Pct_Soma_Total'] / 100.0, item['Status_Geral']
        ]
        
        bg_row = _FILL_ZEBRA if r_atual % 2 == 0 else PatternFill(fill_type=None)
        
        for col_idx, val in enumerate(row_values, start=2):
            cell = ws_resumo.cell(row=r_atual, column=col_idx, value=val)
            cell.border = _BORDA_FINA
            cell.font = _FONTE_NORMAL
            cell.fill = bg_row
            
            if col_idx in [2, 5]:  
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [3, 4]: 
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col_idx == 6: 
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '#,##0'
            elif col_idx in range(7, 14): 
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '0.0%'
            elif col_idx == 14: 
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '0.0%'
                if abs(val - 1.0) > 0.001: 
                    cell.fill = _FILL_VERMELHO
                    cell.font = _FONTE_VERMELHA
            elif col_idx == 15: 
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if "VÁLIDA" in val:
                    cell.fill = _FILL_VERDE
                    cell.font = _FONTE_VERDE
                else:
                    cell.fill = _FILL_VERMELHO
                    cell.font = _FONTE_VERMELHA
                    
        r_atual += 1
        
    for col in range(2, 2 + len(cabecalhos)):
        max_len = 0
        col_letter = get_column_letter(col)
        for row in range(linha_iniciar, r_atual):
            cell = ws_resumo.cell(row=row, column=col)
            val_str = str(cell.value or '')
            if cell.number_format == '0.0%' and isinstance(cell.value, (float, int)):
                val_str = f"{cell.value*100:.1f}%"
            max_len = max(max_len, len(val_str))
        ws_resumo.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    # --- ABA 2: DETALHAMENTO DE NÃO CONFORMIDADES ---
    ws_ncs = wb.create_sheet(title="Não Conformidades")
    ws_ncs.views.sheetView[0].showGridLines = True
    
    ws_ncs.cell(row=2, column=2, value="Detalhamento de Não Conformidades Regulatórias").font = _FONTE_TITULO
    ws_ncs.cell(row=3, column=2, value="Lista de matrizes reprovadas contendo análise detalhada das regras do MEC, Estágio (DCNs), ASM e verificação matemática").font = _FONTE_SUBTITULO
    
    # Colunas de status por regra: EAD, Extensão, Estágio, Prática Pedagógica, ASM e soma matemática
    cabecalhos_nc = [
        "ID Matriz", "Curso", "Modalidade", "CH Total (h)",
        "% EAD", "Status EAD",
        "% Extensão", "Status Extensão",
        "% Estágio", "Status Estágio",
        "% Prática", "Status Prática",
        "% ASM", "Status ASM", "Soma %", "Status Soma",
        "Parecer Técnico de Auditoria"
    ]
    COL_PARECER = 2 + len(cabecalhos_nc) - 1
    
    linha_iniciar_nc = 5
    for col_idx, cabecalho in enumerate(cabecalhos_nc, start=2):
        cell = ws_ncs.cell(row=linha_iniciar_nc, column=col_idx, value=cabecalho)
        cell.font = _FONTE_CABECALHO
        cell.fill = _FILL_CABECALHO
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _BORDA_FINA
        
    ws_ncs.row_dimensions[linha_iniciar_nc].height = 28
    
    r_nc_atual = linha_iniciar_nc + 1
    matrizes_invalidas = [item for item in resultados if "INVIÁVEL" in item['Status_Geral']]
    
    if not matrizes_invalidas:
        ws_ncs.cell(row=r_nc_atual, column=2, value="Nenhuma não conformidade regulatória ou matemática foi encontrada.").font = _FONTE_VERDE
        ws_ncs.merge_cells(start_row=r_nc_atual, start_column=2, end_row=r_nc_atual, end_column=COL_PARECER)
        ws_ncs.row_dimensions[r_nc_atual].height = 22
    else:
        for item in matrizes_invalidas:
            ws_ncs.row_dimensions[r_nc_atual].height = 22

            is_soma_violado = "Erro" in item['Status_Soma']
            rot_ead = _rotulo_status(item['Status_EAD'])
            rot_ext = _rotulo_status(item['Status_Extensao'])
            rot_estagio = _rotulo_status(item['Status_Estagio'])
            rot_pratica = _rotulo_status(item['Status_Pratica'])
            rot_asm = _rotulo_status(item['Status_ASM'])

            # (valor, tipo, violado) — o tipo define alinhamento/formato e evita
            # depender de índices fixos de coluna ao incluir novas regras
            celulas_nc = [
                (item['ID_Matriz'], 'centro', False),
                (item['Curso'], 'esquerda', False),
                (item['Modalidade'], 'centro', False),
                (item['CH_Total'], 'inteiro', False),
                (item['Pct_EAD'] / 100.0, 'percentual', False),
                (rot_ead, 'status', rot_ead == "REPROVADO"),
                (item['Pct_Extensao'] / 100.0, 'percentual', False),
                (rot_ext, 'status', rot_ext == "REPROVADO"),
                (item['Pct_Estagio'] / 100.0, 'percentual', False),
                (rot_estagio, 'status', rot_estagio == "REPROVADO"),
                (item['Pct_Pratica_Total'] / 100.0, 'percentual', False),
                (rot_pratica, 'status', rot_pratica == "REPROVADO"),
                (item['Pct_ASM'] / 100.0, 'percentual', False),
                (rot_asm, 'status', rot_asm == "REPROVADO"),
                (item['Pct_Soma_Total'] / 100.0, 'percentual', is_soma_violado),
                ("DIVERGENTE" if is_soma_violado else "OK", 'status', is_soma_violado),
                (item['Parecer'], 'parecer', False),
            ]

            bg_nc = _FILL_ZEBRA if r_nc_atual % 2 == 0 else PatternFill(fill_type=None)

            for offset, (val, tipo, violado) in enumerate(celulas_nc):
                cell = ws_ncs.cell(row=r_nc_atual, column=2 + offset, value=val)
                cell.border = _BORDA_FINA
                cell.font = _FONTE_NORMAL
                cell.fill = bg_nc

                if tipo == 'esquerda':
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                elif tipo == 'inteiro':
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = '#,##0'
                elif tipo == 'percentual':
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = '0.0%'
                    if violado:
                        cell.fill = _FILL_VERMELHO
                        cell.font = _FONTE_VERMELHA
                elif tipo == 'status':
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if violado:
                        cell.fill = _FILL_VERMELHO
                        cell.font = _FONTE_VERMELHA
                    elif val != "N/A":
                        cell.font = _FONTE_VERDE
                elif tipo == 'parecer':
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                    cell.font = _FONTE_NEGRITO
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")

            r_nc_atual += 1
            
    for col in range(2, 2 + len(cabecalhos_nc)):
        max_len = 0
        col_letter = get_column_letter(col)
        for row in range(linha_iniciar_nc, r_nc_atual):
            cell = ws_ncs.cell(row=row, column=col)
            val_str = str(cell.value or '')
            if cell.number_format == '0.0%' and isinstance(cell.value, (float, int)):
                val_str = f"{cell.value*100:.1f}%"
            if col == COL_PARECER:
                max_len = 50
                break
            max_len = max(max_len, len(val_str))
        ws_ncs.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    wb.save(caminho_saida)

if __name__ == "__main__":
    print("=" * 80)
    print("      SISTEMA DE AUDITORIA REGULATÓRIA DE MATRIZES CURRICULARES - V2")
    print("=" * 80)
    
    print("\n[ETAPA 1/5] Ingestão de Dados...")
    df_input = carregar_dados('Dados_Exportacao_CSV.csv')
    
    if df_input is not None and not df_input.empty:
        total_linhas = len(df_input)
        total_matrizes = len(df_input['ID_Matriz'].unique())
        print(f"  * Arquivo carregado com sucesso!")
        print(f"  * Total de disciplinas importadas: {total_linhas}")
        print(f"  * Total de matrizes curriculares: {total_matrizes}")
        
        print("\n[ETAPA 2/5] Configuração dos Parâmetros (MEC & DCNs)...")
        print("  * EAD: Máximo 40% (Presencial).")
        print("  * Extensão: Mínimo 10%.")
        print("  * Saúde/Medicina: Estágio >= 35%; Demais Saúde >= 20%.")
        print("  * Engenharias: Estágio >= 160h.")
        print("  * Licenciaturas: Estágio >= 400h | Prática >= 400h.")
        print("  * Prática indevida: CH de prática = 0 em cursos cuja DCN não prevê prática")
        print(f"    ({', '.join(sorted(CURSOS_SEM_EXIGENCIA_PRATICA))}).")
        print("  * ASM: Máximo 20%.")
        print("  * Regra Matemática: Soma da grade deve fechar 100%.")
        
        print("\n[ETAPA 3/5] Processamento com CP-SAT Solver...")
        resultados = validar_todas_matrizes(df_input)
        
        print("\n[ETAPA 4/5] Consolidação dos Resultados...")
        total_validas = sum(1 for r in resultados if "VÁLIDA" in r['Status_Geral'])
        total_invalidas = total_matrizes - total_validas
        pct_conf = (total_validas / total_matrizes * 100) if total_matrizes > 0 else 0.0
        print(f"  * Matrizes em Conformidade: {total_validas}")
        print(f"  * Matrizes Inválidas: {total_invalidas}")
        print(f"  * Taxa Geral de Conformidade: {pct_conf:.1f}%")
        
        print("\n[ETAPA 5/5] Geração de Relatório de Auditoria...")
        diretorio_script = os.path.dirname(os.path.abspath(__file__))
        caminho_excel = os.path.join(diretorio_script, 'Relatorio_Auditoria_Matrizes.xlsx')
        try:
            gerar_relatorio_excel(resultados, caminho_excel)
            print(f"  * Relatório gerado em: {caminho_excel}")
        except PermissionError:
            print(f"  * [Erro]: Permissão negada ao gravar o arquivo Excel.")
            print(f"    --> Verifique se o arquivo '{caminho_excel}' está aberto e feche-o antes de executar novamente o script.")
        except Exception as e:
            print(f"  * [Erro] ao gerar Excel: {e}")
            
        print("\n" + "=" * 80)
        print("AUDITORIA FINALIZADA")
        print("=" * 80)
    else:
        print("\nExecução abortada: Dados não encontrados.")
        sys.exit(1)