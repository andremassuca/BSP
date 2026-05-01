#!/usr/bin/env python3
# Copyright (c) 2024-2026 Andre O. Massuca, Pedro Aleixo, Luis M. Massuca
# All rights reserved. Unauthorised copying or distribution is prohibited.
"""
BSP i18n - Internationalisation Module
Suporte para PT / EN / ES / DE
Andre Massuca, P. Aleixo & Luis M. Massuca | 2026
"""

LINGUAS_DISPONIVEIS = {
    'PT': '🇵🇹 PT',
    'EN': '🇬🇧 EN',
    'ES': '🇪🇸 ES',
    'DE': '🇩🇪 DE',
}

_LANG_ATUAL = ['PT']   # lista mutável - estado global

def definir_lingua(lang: str):
    if lang in LINGUAS_DISPONIVEIS:
        _LANG_ATUAL[0] = lang

def lingua_atual() -> str:
    return _LANG_ATUAL[0]

def T(key: str, **kw) -> str:
    """Devolve a string traduzida para a língua actual.
    Aceita format kwargs, e.g. T('versao_label', v='22')
    """
    lang = _LANG_ATUAL[0]
    tbl = _STRINGS.get(lang, _STRINGS['PT'])
    s = tbl.get(key) or _STRINGS['PT'].get(key) or key
    return s.format(**kw) if kw else s


# ═══════════════════════════════════════════════════════════════════════════
#  DICIONÁRIOS DE STRINGS
# ═══════════════════════════════════════════════════════════════════════════

_STRINGS = {

# ─────────────────────────────────────────────────────────────────────────
'PT': {
    # ── Geral ──────────────────────────────────────────────────────────────
    'app_title':           'BSP  |  Biomechanical Stability Program  v{v}',
    'prog_nome':           'Biomechanical Stability Program',
    'autor_label':         'Autor',
    'individuo_label':     'Indivíduo',
    'data_label':          'Data',
    'versao_label':        'v{v}',
    'lang_tooltip':        'Mudar idioma da interface e dos relatórios',

    # ── Menus / Secções UI ─────────────────────────────────────────────────
    'sec_ficheiros_entrada':  'FICHEIROS DE ENTRADA',
    'sec_ficheiros_saida':    'FICHEIROS DE SAÍDA',
    'sec_opcoes_gerais':      'OPÇÕES GERAIS',
    'sec_protocolo_tiro':     'PROTOCOLO TIRO',
    'sec_opcoes_estacao':     'OPÇÕES DE ESTAÇÃO',
    'sec_analise_bipodal':    'ANÁLISE BIPODAL',
    'sec_analise_unipodal':   'ANÁLISE UNIPODAL',
    'sec_resultados':         'RESULTADOS',
    'sec_log':                'REGISTO DE EXECUÇÃO',

    # ── Etiquetas de campos ────────────────────────────────────────────────
    'pasta_individuos':    'Pasta de indivíduos *',
    'pasta_individuos_s':  'Indivíduos *',
    'pasta_individuos_tip':'Pasta raiz com uma subpasta por indivíduo',
    'fich_tempos':         'Ficheiro Inicio_fim',
    'fich_tempos_s':       'Inicio_fim',
    'fich_tempos_tip':     'Ficheiro Excel com os intervalos de inicio/fim por ensaio',
    'fich_tempos_tiro':    'Ficheiro de tempos (tiro) *',
    'fich_tempos_tiro_s':  'Tempos (tiro) *',
    'fich_tempos_tiro_tip':'Ficheiro Excel com os tempos de toque, pontaria e disparo',
    'fich_atletas_ref':    'Referência demográfica (142 atletas)',
    'fich_atletas_ref_s':  'Ref. demográfica',
    'fich_atletas_ref_tip':'Ficheiro Excel com peso, altura, género, estilo e scores dos atletas de referência',
    'fich_scores':         'Ficheiro de scores (opcional)',
    'fich_scores_s':       'Scores',
    'fich_scores_tip':     'Ficheiro Excel com pontuações de precisão por indivíduo (opcional)',
    'saida_excel':         'Resumo Excel *',
    'saida_excel_s':       'Resumo Excel *',
    'saida_excel_tip':     'Ficheiro Excel de saída com o resumo dos resultados',
    'pasta_individuais':   'Pasta ficheiros individuais',
    'pasta_individuais_s': 'Individuais',
    'pasta_individuais_tip':'Pasta onde serão guardados os ficheiros Excel por indivíduo',
    'relatorio_pdf':       'Relatório PDF (vazio = não gerar)',
    'relatorio_pdf_s':     'PDF',
    'relatorio_pdf_tip':   'Caminho do relatório PDF (deixar vazio para não gerar)',
    'relatorio_word':      'Relatório Word (opcional)',
    'n_ensaios_label':     'N.º ensaios (vazio = protocolo default)',
    'associar_individuo':  'Associar indivíduo por:',
    'id_pasta':            'ID da pasta (ID_Nome)',
    'posicao_lista':       'Posição na lista',
    'intervalos_calcular': 'Intervalos a calcular:',

    # ── Checkboxes ─────────────────────────────────────────────────────────
    'ck_embedded':   'Usar intervalo embedded nos ficheiros .xls',
    'ck_elipse':     'Gerar aba de elipse 95% com gráfico',
    'ck_estab':      'Gerar aba de estabilograma (COF X e Y vs tempo)',
    'ck_individuais':'Gerar ficheiros individuais por indivíduo',
    'ck_pdf':        'Gerar relatório PDF',
    'ck_word':       'Gerar relatório Word',
    'ck_bipodal':    'Incluir análise bipodal (Hurdle Step)',

    # ── Botões ─────────────────────────────────────────────────────────────
    'btn_executar':     '▶  Executar Análise',
    'btn_cancelar':     '✖  Cancelar',
    'btn_guardar':      '💾 Guardar',
    'btn_historico':    '📋 Histórico',
    'btn_poder':        '🔬 Poder',
    'btn_protocolo':    '⇄ Protocolo',
    'btn_tema':         '🌙 Tema',
    'btn_abrir_pasta':  '📂 Abrir Pasta de Resultados',
    'btn_demografia':       'Análise Demográfica',
    'demo_titulo':          'Análise Demográfica (Tiro com Arco)',
    'demo_intro':           'Escolhe métrica e factor/variável e clica num botão abaixo.',
    'demo_metrica':         'Métrica',
    'demo_factor':          'Factor',
    'demo_var_dem':         'Variável demográfica',
    'demo_btn_comparar':    'Comparar grupos',
    'demo_btn_corr_dem':    'Correlação demográfica',
    'demo_btn_corr_score':  'Score vs CoP',
    'btn_aceitar':      '  ✓  Li e Aceito os Termos  ',
    'btn_recusar':      'Recusar e Sair',
    'btn_entrar':       '  Entrar  ',
    'btn_abrir_github': '🔗  Abrir GitHub',
    'btn_aceitar_lic':  '  ✓  Aceitar / Accept  ',
    'btn_recusar_lic':  '  ✗  Recusar / Decline  ',

    # ── Status bar / progresso ────────────────────────────────────────────
    'status_pronto':    '● Pronto',
    'status_prog':      '▶ A processar…',
    'acesso_pass':      'Palavra-passe de acesso',
    'acesso_pass_err':  'Palavra-passe incorrecta.',
    'acesso_github':    'Palavra-passe disponível em',
    'acesso_entrar':    '  Entrar  ',
    'log_bsp_header':   'BSP  -  {prog} v{versao}  -  {proto}',
    'log_atalhos':      'Atalhos: {mod}+Enter executar  |  {mod}+S guardar  |  {mod}+H histórico',

    # ── Ecrã de selecção de protocolo ─────────────────────────────────────
    'proto_titulo':     'Seleciona o protocolo de análise:',
    'proto_confirmar':  'Confirmar e avançar',
    'proto_fechar':     'Fechar',
    'proto_func_titulo':'Tarefa Funcional',
    'proto_func_sub':   'Seleciona o protocolo específico',
    'proto_func_label': 'Protocolos de Tarefa Funcional disponíveis:',
    'proto_confirmar2': 'Confirmar',
    'proto_voltar':     'Voltar',
    'proto_fms_nome':   'FMS Bipodal',
    'proto_fms_descr':  '5 ensaios por pé, bipodal\nElipse 95%, assimetria Dir/Esq, todas as métricas',
    'proto_uni_nome':   'Apoio Unipodal',
    'proto_uni_descr':  '5 ensaios por pé, unipodal\nElipse 95%, métricas de oscilação lateral',
    'proto_func_nome':  'Tarefa Funcional',
    'proto_func_descr': 'Análise de tarefa funcional específica\nSubmenu: Tiro e outras tarefas futuras',
    'proto_tiro_nome':  'Tiro',
    'proto_tiro_descr': '5 ensaios bipodal, 2 janelas por ensaio\nComparação posição vs disparo, correlação com precisão',
    'proto_arco_nome':  'Tiro com Arco',
    'proto_arco_descr': 'Até 30 ensaios, janela única (Confirmação 1 > Confirmação 2)\nAnálise demográfica: género, categoria, estilo, precisão',
    # ── Curso e identificação ──────────────────────────────────────────────
    'curso_label':         '',
    # ── Histórico ─────────────────────────────────────────────────────────
    'hist_titulo':         'Histórico de Sessões',
    'hist_subtitulo':      'Últimas {n} análises executadas neste computador.',
    'hist_col_data':       'Data',
    'hist_col_proto':      'Protocolo',
    'hist_col_atletas':    'Atletas',
    'hist_col_excel':      'Excel',
    'hist_col_pdf':        'PDF',
    'hist_sem':            'Sem histórico',
    'hist_btn_abrir':      '📂  Abrir pasta',
    'hist_btn_reutilizar': '↺  Reutilizar configuração',
    'hist_btn_fechar':     'Fechar',
    'hist_cfg_recarregada':'Configuração recarregada da sessão {data}',
    'hist_pasta_nao_enc':  'Pasta não encontrada.',
    # ── Secções ───────────────────────────────────────────────────────────
    'sec_ficheiros_ent':   'FICHEIROS DE ENTRADA',
    'sec_ficheiros_sai':   'FICHEIROS DE SAÍDA',
    'sec_opcoes_gerais':   'OPÇÕES GERAIS',
    'sec_exportacao':      'EXPORTAÇÃO',
    'sec_proto_tiro':      'PROTOCOLO TIRO',
    'sec_dist_teste':      'DISTÂNCIAS DO TESTE',
    'sec_anal_estat':      'ANÁLISE ESTATÍSTICA',
    'tiro_stat_pos_disp':  'Posição vs Disparo + IP',
    'export_csv_label':    'Exportar CSV (grupo + trials)',
    'export_docx_label':   'Exportar relatório Word (.docx)',
    # ── Tiro sidebar ──────────────────────────────────────────────────────
    'tiro_assoc_por':      'Associar indivíduo por:',
    'tiro_id_pasta':       'ID da pasta (ID_Nome)',
    'tiro_pos_lista':      'Posição na lista',
    'tiro_itvs_calc':      'Intervalos a calcular:',
    'tiro_bipodal':        'Incluir análise bipodal (Hurdle Step)',
    'tiro_n_ensaios':      '  N.ensaios:',
    'tiro_p_pe':           'p/pé',
    'tiro_dist_info':      'Distâncias no ficheiro de tempos (detetadas automaticamente).\nPodes acrescentar distâncias adicionais em baixo:',
    'tiro_add_dist':       '+ Acrescentar distância',
    # ── Intervalos de Tiro ─────────────────────────────────────────────────
    'tiro_itv_pont':       'Toque a Pontaria',
    'tiro_itv_disp':       'Toque a Disparo',
    'tiro_itv_pont_disp':  'Pontaria a Disparo',
    'tiro_itv_disp_fim':   'Disparo ao Fim',
    'tiro_itv_total':      'Ensaio Total',
    # ── PDF capa ──────────────────────────────────────────────────────────
    'pdf_capa_data':       'Data: {data}',
    'pdf_capa_n_indiv':    'N.º indivíduos: {n}',
    'pdf_capa_analisados': 'INDIVÍDUOS ANALISADOS',
    'pdf_capa_analisados2':'INDIVÍDUOS ANALISADOS (cont.)',
    'pdf_capa_distancias': 'Distâncias: {dists}',
    'pdf_capa_intervalos': 'Intervalos: {itvs}',
    'pdf_capa_colab': '',
    'pdf_capa_colab2': '',
    'pdf_capa_mais':       '... e mais {n} indivíduos (ver página seguinte)',
    'meses': ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'],

    # ── Mensagens de progresso / log ───────────────────────────────────────
    'log_inicio':       '▶  A iniciar análise…',
    'log_concluido':    '✔  Análise concluída',
    'log_erro':         '✖  Erro: {msg}',
    'log_aviso':        '⚠  {msg}',
    'log_update':       '⬆  Nova versão disponível: v{nova}  (atual: v{atual})',
    'log_individuo':    '→  A processar: {nome}',
    'log_sem_dados':    '⚠  Sem dados válidos: {nome}',

    # ── PDF - Secções ──────────────────────────────────────────────────────
    'pdf_metrica':       'Métrica',
    'pdf_pe_direito':    'PÉ DIREITO',
    'pdf_pe_esquerdo':   'PÉ ESQUERDO',
    'pdf_posicao':       'POSIÇÃO',
    'pdf_disparo':       'DISPARO',
    'pdf_ai':            'AI (%)',
    'pdf_max':           'máx',
    'pdf_med':           'méd',
    'pdf_dp':            'dp',
    'pdf_resumo_grupo':  'RESUMO DO GRUPO',
    'pdf_estatisticas':  'ESTATÍSTICAS DESCRITIVAS',
    'pdf_normalidade':   'NORMALIDADE',
    'pdf_comparacao':    'COMPARAÇÃO',
    'pdf_correlacoes':   'CORRELAÇÕES',
    'pdf_relatorio':     'RELATÓRIO',
    'pdf_bilateral':     'Análise Bilateral',
    'pdf_unilateral':    'Análise Unilateral',
    'pdf_estabilidade':  'Estabilidade Postural',
    'pdf_individuo':     'Indivíduo',
    'pdf_grupo':         'Grupo',
    'pdf_n':             'n',
    'pdf_media':         'Média',
    'pdf_desvio_padrao': 'DP',
    'pdf_cv':            'CV (%)',
    'pdf_min':           'Mín.',
    'pdf_ic95':          'IC 95%',
    'pdf_p_valor':       'p',
    'pdf_shapiro':       'Shapiro-Wilk',
    'pdf_normal':        'Normal',
    'pdf_nao_normal':    'Não normal',
    'pdf_teste':         'Teste',
    'pdf_cohen_d':       "d Cohen",
    'pdf_efeito':        'Efeito',
    'pdf_grande':        'Grande',
    'pdf_medio':         'Médio',
    'pdf_pequeno':       'Pequeno',
    'pdf_sem_efeito':    'Sem efeito',
    'pdf_assimetria':    'Índice de Assimetria (%)',
    'pdf_direito':       'Dir.',
    'pdf_esquerdo':      'Esq.',
    'pdf_diferenca':     'Diferença',
    'pdf_nota_clinica':  'Nota: Valores acima de 10% de assimetria podem indicar desequilíbrio clinicamente relevante.',
    'pdf_gerado_por':    'Gerado por {prog} v{v} | {autor} | {univ}',
    'pdf_data_relatorio':'Relatório gerado em {data}',
    'pdf_pag':           'Pág. {n}',

    # ── Métricas - labels ──────────────────────────────────────────────────
    'met_amp_x':      'Amplitude X / ML (mm)',
    'met_amp_y':      'Amplitude Y / AP (mm)',
    'met_vel_x':      'Vel. média X (mm/s)',
    'met_vel_y':      'Vel. média Y (mm/s)',
    'met_vel_med':    'Vel. média CoP (mm/s)',
    'met_vel_pico_x': 'Pico vel. X (mm/s)',
    'met_vel_pico_y': 'Pico vel. Y (mm/s)',
    'met_desl':       'Deslocamento (mm)',
    'met_time':       'Tempo (s)',
    'met_ea95':       'Área elipse 95% (mm²)',
    'met_leng_a':     'Semi-eixo a (mm)',
    'met_leng_b':     'Semi-eixo b (mm)',
    'met_ratio_ml_ap':'Rel. amp. ML/AP',
    'met_ratio_vel':  'Rel. vel. ML/AP',
    'met_stiff_x':    'Stiffness X (1/s)',
    'met_stiff_y':    'Stiffness Y (1/s)',
    'met_cov_xx':     'Var CoP X (mm²)',
    'met_cov_yy':     'Var CoP Y (mm²)',
    'met_cov_xy':     'Cov CoP XY (mm²)',
    'met_rms_x':      'RMS ML (mm)',
    'met_rms_y':      'RMS AP (mm)',
    'met_rms_r':      'RMS Radius (mm)',

    # ── Excel - aba labels ─────────────────────────────────────────────────
    'aba_dados':   'DADOS',
    'aba_grupo':   'GRUPO',
    'aba_spss':    'SPSS',
    'aba_elipse':  'elipse',
    'aba_estab':   'estab',

    # ── Word ───────────────────────────────────────────────────────────────
    'word_titulo':          'Relatório de Estabilidade Postural',
    'word_subtitulo':       'Análise Biomecânica - {protocolo}',
    'word_velocidades':     'Velocidades',
    'word_elipse':          'Elipse de Confiança 95%',
    'word_estabilograma':   'Estabilograma',

    # ── Ecrã de licença ────────────────────────────────────────────────────
    'licenca_titulo':   'Acordo de Licença de Utilizador Final (EULA)',
    'licenca_subtitulo':'Leia atentamente antes de prosseguir.',
    'licenca_texto':    '',             # preenchida abaixo por licenca_texto()

    # ── Ecrã de password ───────────────────────────────────────────────────
    'pass_titulo':      'Palavra-passe de acesso',
    'pass_hint':        'Disponível em github.com/andremassuca',
    'pass_erro':        'Palavra-passe incorrecta. Tente novamente.',

    # ── Calculadora de poder ───────────────────────────────────────────────
    'poder_titulo':     'Calculadora de Tamanho Amostral',
    'poder_subtitulo':  'Baseada em Cohen (1988)  |  Teste t-pareado / Wilcoxon',
    'poder_calcular':   'Calcular',
    'poder_fechar':     'Fechar',
    'poder_presets':    'Presets rápidos:',

    # ── Tooltips chave ─────────────────────────────────────────────────────
    'tip_embedded':     'Usa o intervalo de tempo gravado no próprio ficheiro .xls\nem vez do ficheiro Inicio_fim.',
    'tip_elipse':       'Cria a aba "elipse_..." em cada ficheiro individual com o scatter plot do COF e a elipse de 95%.',
    'tip_estab':        'Cria a aba "estab_..." com o gráfico do deslocamento do COF ao longo do tempo.',
    'tip_individuais':  'Cria um ficheiro Excel individual por indivíduo com todas as abas de ensaio, elipse e estabilograma.',
    'tip_n_ensaios':    'Substitui o número de ensaios definido pelo protocolo.',

    # ── Botões gerais ──────────────────────────────────────────────────────
    'btn_stop':         '■  Parar',
    'btn_limpar':       '⊘ Limpar',
    'btn_fechar':       'Fechar',
    'btn_calcular':     'Calcular',
    'btn_download':     'Descarregar →',
    'btn_exportar_xl':  '💾 Exportar Excel',

    # ── Tema ───────────────────────────────────────────────────────────────
    'tema_claro':       'Mudar para Modo Claro',
    'tema_escuro':      'Mudar para Modo Escuro',
    'lbl_idioma':       'Idioma / Language',

    # ── Atalhos / shortcuts overlay ────────────────────────────────────────
    'shortcuts_titulo': '⌨  Atalhos de Teclado',

    # ── Perfis ─────────────────────────────────────────────────────────────
    'perfis_titulo':    '📌  Perfis de Configuração',
    'perfis_descr':     'Guarda e carrega conjuntos de configuração com nome.',
    'perfis_nome_lbl':  'Nome:',
    'perfis_guardar':   '💾 Guardar actual',
    'perfis_carregar':  '↑ Carregar',
    'perfis_apagar':    '✕ Apagar',
    'perfil_guardado':  '✓ Perfil guardado: "{nm}"',
    'perfil_carregado': '✓ Perfil carregado: "{nm}"',

    # ── Análise rápida ─────────────────────────────────────────────────────
    'quick_titulo':     '⚡  Análise Rápida - Ficheiro Único',
    'quick_exportar':   '💾 Exportar Excel',

    # ── Banner de update ────────────────────────────────────────────────────
    'upd_banner_txt':   '  ⬆  Versão v{nova} disponível',

    # ── Histórico ──────────────────────────────────────────────────────────
    'hist_filtrar':     '  filtrar por protocolo/data',

    # ── Calculadora de poder (janela) ──────────────────────────────────────
    'poder_titulo_win': 'Calculadora de Poder  |  BSP',

    # ── Labels de UI ───────────────────────────────────────────────────────
    'lbl_sep':          'Sep:',
    'lbl_dec':          'Dec:',
    'lbl_dist_m':       'Distância (m):',
    'iscpsi_colab': '',

    # ── Checkboxes de ESTATS ───────────────────────────────────────────────
    'ck_estats':        'Testes estatísticos automáticos (aba ESTATS)',
    'ck_estats_gr':     'Descritivos + SW + IC95%',
    'ck_estats_de':     'Dir vs Esq + Cohen\'s d',
    'ck_estats_pd':     'Posição vs Disparo + IP',
    'ck_estats_at':     'Variab. intra-atleta (CV)',
    'ck_estats_fr':     'Friedman entre intervalos',
    'ck_estats_ph':     '  + pos-hoc Bonferroni',
    'ck_estats_co':     'Correl. EA95 vs Score',

    # ── HTML export checkbox ────────────────────────────────────────────────
    'ck_html':          'Gerar relatório HTML interactivo (Chart.js)',

    # ── Tooltips de ESTATS e exportação ────────────────────────────────────
    'tip_posthoc':      'Pos-hoc de Wilcoxon com correcção de Bonferroni\nse Friedman p < 0.05.',
    'tip_csv':          'Gera ficheiros .csv prontos para R / SPSS / Excel:\n  *_grupo.csv   - médias e DP por atleta/lado\n  *_trials.csv  - um trial por linha (trial-level)\nSeparador e decimal configuráveis abaixo.',
    'tip_csv_sep':      'Separador de colunas: ; para Excel PT, , para Excel EN, Tab para R',
    'tip_csv_dec':      'Separador decimal: , para PT/ES, . para EN',
    'tip_docx':         'Gera relatório Word com tabelas estatísticas (ESTATS).\nRequer: pip install python-docx\nInclui descritivos + SW, Dir vs Esq, variabilidade.',
    'tip_html':         'Gera um ficheiro .html standalone com gráficos interactivos.\nAbre no browser sem instalar nada.\nIdeal para partilhar resultados.',
    'tip_dist_extra':   'Distância adicional a processar (ex: 10, 25).\nOs ficheiros devem seguir o formato trial[dist]_[ensaio].',
    'tip_n_ensaios_tiro':'Substitui o número de ensaios definido pelo protocolo.\nÚtil para sessões com menos ou mais ensaios que o padrão.',
    'tip_match':        'Define como cada pasta de indivíduo é associada ao ficheiro de tempos.\n\nID da pasta (ID_Nome): usa o número no início do nome da pasta\n  Ex: "66_Leonor" -> ID=66.\n  Robusto mesmo se a ordem das pastas não corresponder ao ficheiro.\n\nPosição na lista: usa a ordem alfabética das pastas.',
    'tip_rb_id':        'Extrai o número do início do nome da pasta (ex: "66_Leonor" -> ID 66)\ne procura esse ID no ficheiro de tempos.\nRecomendado.',
    'tip_rb_idx':       'Associa por ordem: 1.ª pasta -> linha 1, 2.ª pasta -> linha 2.\nUsar apenas se as pastas não tiverem ID numérico.',
    'tip_bipodal':      'Analisa os ficheiros dir_N e esq_N de cada indivíduo\nusando os tempos da folha "inicio_fim (Hurdle Step)".',
    'tip_n_ens_hs':     'Número de ensaios por pé (dir e esq) para o Hurdle Step.\nDefault: 5.',

    # ── Mensagens de log do thread de análise ──────────────────────────────
    'log_tempos':       'Tempos: {nome}',
    'log_dist_det':     '  Distâncias detectadas: {dists}',
    'log_hs_n':         '  Hurdle Step: {n} indivíduos',
    'log_n_ind':        '  {n} indivíduos.',
    'log_tempo_nao_enc':'Ficheiro de tempos não encontrado.',
    'log_scores':       'Scores: {n} indivíduos.',
    'log_sem_subs':     'Nenhuma subpasta encontrada.',
    'log_n_total':      '{n} indivíduo(s)  [{proto}]',
    'log_ens_override': '  N.º ensaios override: {n}',
    'log_gerar_excel':  '\nA gerar Excel (DADOS + GRUPO + SPSS {estats})...',
    'log_excel_ok':     'Excel: {f}',
    'log_csv':          'CSV: {f}',
    'log_aviso_csv':    'Aviso CSV: {e}',
    'log_word_ok':      'Word: {f}',
    'log_word_nao':     'Word não gerado: {msg}',
    'log_aviso_docx':   'Aviso DOCX: {e}',
    'log_word_sem_estats': 'Word: activa os testes estatísticos para gerar o relatório.',
    'log_html_ok':      'HTML: {f}',
    'log_aviso_html':   'Aviso HTML: {e}',
    'log_fich_ind':     '\nFicheiros individuais: {pasta}',
    'log_tiro_dist':    '  {nome} (tiro, {n} dist.)',
    'log_hs_ok':        '  {fn} (hurdle step)',
    'log_sel_ok':       '  {nome} – {lado} (sel ficheiros)',
    'log_aviso_sel':    '  aviso sel {lado}: {e}',
    'log_gerar_pdf':    '\nA gerar PDF...',
    'log_concluido_n':  '\nConcluído  -  {n} indivíduo(s).',
    'log_cancelado':    'Cancelado.',
    'log_cfg_guardada': '✓ Configuração guardada  ({mod}+S)',
    'log_cfg_erro':     'Erro ao guardar: {e}',
    'log_sem_dados':    'Sem dados.',
    'log_erro_tb':      '\nErro:\n{tb}',

    # ── Validação ──────────────────────────────────────────────────────────
    'val_titulo':       'Verificação de entradas',
    'val_intro':        'Corrija os seguintes problemas antes de executar:\n\n',
    'val_pasta_vazia':  'Pasta de indivíduos não definida.',
    'val_pasta_nao_enc':'Pasta não encontrada:\n{p}',
    'val_pasta_sem_sub':'Nenhuma subpasta na pasta de indivíduos.',
    'val_saida_vazia':  'Ficheiro de saída Excel não definido.',
    'val_ifd_nao_enc':  'Ficheiro de tempos não encontrado:\n{f}',

    # ── Instalador / Desinstalador ──────────────────────────────────────────
    'inst_pasta':       'Pasta de instalação:',
    'inst_opcoes':      'Opções:',
    'inst_btn_inst':    '  Instalar  ',
    'inst_btn_cancel':  'Cancelar',
    'inst_btn_instalar':'A instalar...',
    'inst_sucesso':     '{app} instalado com sucesso!',
    'inst_btn_abrir':   '  Abrir {app}  ',
    'inst_btn_fechar':  '  Fechar  ',
    'desinst_confirm':  'Tens a certeza que queres remover\n{app}?',
    'desinst_btn':      '  Desinstalar  ',
    'desinst_removendo':'A remover...',
    'desinst_sucesso':  '{app} foi removido com sucesso.',
    'desinst_ficheiros':'Os ficheiros são eliminados ao fechar.',
},

# ─────────────────────────────────────────────────────────────────────────
'EN': {
    'app_title':           'BSP  |  Biomechanical Stability Program  v{v}',
    'prog_nome':           'Biomechanical Stability Program',
    'autor_label':         'Author',
    'individuo_label':     'Subject',
    'data_label':          'Date',
    'versao_label':        'v{v}',
    'lang_tooltip':        'Change interface and report language',

    'sec_ficheiros_entrada':  'INPUT FILES',
    'sec_ficheiros_saida':    'OUTPUT FILES',
    'sec_opcoes_gerais':      'GENERAL OPTIONS',
    'sec_protocolo_tiro':     'SHOOTING PROTOCOL',
    'sec_opcoes_estacao':     'STATION OPTIONS',
    'sec_analise_bipodal':    'BIPODAL ANALYSIS',
    'sec_analise_unipodal':   'UNIPODAL ANALYSIS',
    'sec_resultados':         'RESULTS',
    'sec_log':                'EXECUTION LOG',

    'pasta_individuos':    'Subjects folder *',
    'pasta_individuos_s':  'Subjects *',
    'pasta_individuos_tip':'Root folder with one subfolder per subject',
    'fich_tempos':         'Start/end file',
    'fich_tempos_s':       'Start/end',
    'fich_tempos_tip':     'Excel file with start/end intervals per trial',
    'fich_tempos_tiro':    'Timing file (shooting) *',
    'fich_tempos_tiro_s':  'Timing *',
    'fich_tempos_tiro_tip':'Excel file with touch, aim and shot timestamps',
    'fich_atletas_ref':    'Demographic reference (142 athletes)',
    'fich_atletas_ref_s':  'Demographic ref.',
    'fich_atletas_ref_tip':'Excel file with weight, height, gender, style and scores of reference athletes',
    'fich_scores':         'Scores file (optional)',
    'fich_scores_s':       'Scores',
    'fich_scores_tip':     'Excel file with accuracy scores per subject (optional)',
    'saida_excel':         'Excel summary *',
    'saida_excel_s':       'Excel out *',
    'saida_excel_tip':     'Output Excel file with results summary',
    'pasta_individuais':   'Individual files folder',
    'pasta_individuais_s': 'Individuals',
    'pasta_individuais_tip':'Folder where individual Excel files will be saved',
    'relatorio_pdf':       'PDF report (empty = skip)',
    'relatorio_pdf_s':     'PDF',
    'relatorio_pdf_tip':   'PDF report path (leave empty to skip)',
    'relatorio_word':      'Word report (optional)',
    'n_ensaios_label':     'No. of trials (empty = protocol default)',
    'associar_individuo':  'Match subject by:',
    'id_pasta':            'Folder ID (ID_Name)',
    'posicao_lista':       'List position',
    'intervalos_calcular': 'Intervals to compute:',

    'ck_embedded':   'Use embedded interval from .xls files',
    'ck_elipse':     'Generate 95% ellipse tab with chart',
    'ck_estab':      'Generate stabilogram tab (COP X and Y vs time)',
    'ck_individuais':'Generate individual Excel files per subject',
    'ck_pdf':        'Generate PDF report',
    'ck_word':       'Generate Word report',
    'ck_bipodal':    'Include bipodal analysis (Hurdle Step)',

    'btn_executar':     '▶  Run Analysis',
    'btn_cancelar':     '✖  Cancel',
    'btn_guardar':      '💾 Save',
    'btn_historico':    '📋 History',
    'btn_poder':        '🔬 Power',
    'btn_protocolo':    '⇄ Protocol',
    'btn_tema':         '🌙 Theme',
    'btn_abrir_pasta':  '📂 Open Results Folder',
    'btn_demografia':       'Demographic Analysis',
    'demo_titulo':          'Demographic Analysis (Archery)',
    'demo_intro':           'Choose a metric and factor/variable then click a button below.',
    'demo_metrica':         'Metric',
    'demo_factor':          'Factor',
    'demo_var_dem':         'Demographic variable',
    'demo_btn_comparar':    'Compare groups',
    'demo_btn_corr_dem':    'Demographic correlation',
    'demo_btn_corr_score':  'Score vs CoP',
    'btn_aceitar':      '  ✓  I Have Read and Accept  ',
    'btn_recusar':      'Decline and Exit',
    'btn_entrar':       '  Log In  ',
    'btn_abrir_github': '🔗  Open GitHub',
    'btn_aceitar_lic':  '  ✓  Accept  ',
    'btn_recusar_lic':  '  ✗  Decline  ',

    # ── Status bar / progresso ────────────────────────────────────────────
    'status_pronto':    '● Ready',
    'status_prog':      '▶ Processing…',
    'acesso_pass':      'Access password',
    'acesso_pass_err':  'Incorrect password.',
    'acesso_github':    'Password available at',
    'acesso_entrar':    '  Login  ',
    'log_bsp_header':   'BSP  -  {prog} v{versao}  -  {proto}',
    'log_atalhos':      'Shortcuts: {mod}+Enter run  |  {mod}+S save  |  {mod}+H history',

    # ── Protocol selection screen ─────────────────────────────────────────
    'proto_titulo':     'Select the analysis protocol:',
    'proto_confirmar':  'Confirm and continue',
    'proto_fechar':     'Close',
    'proto_func_titulo':'Functional Task',
    'proto_func_sub':   'Select the specific protocol',
    'proto_func_label': 'Available Functional Task protocols:',
    'proto_confirmar2': 'Confirm',
    'proto_voltar':     'Back',
    'proto_fms_nome':   'FMS Bipodal',
    'proto_fms_descr':  '5 trials per foot, bipodal\n95% Ellipse, Dir/Left asymmetry, all metrics',
    'proto_uni_nome':   'Unipodal Support',
    'proto_uni_descr':  '5 trials per foot, unipodal\n95% Ellipse, lateral sway metrics',
    'proto_func_nome':  'Functional Task',
    'proto_func_descr': 'Specific functional task analysis\nSubmenu: Shooting and future tasks',
    'proto_tiro_nome':  'Shooting',
    'proto_tiro_descr': '5 bipodal trials, 2 windows per trial\nPosition vs shot comparison, accuracy correlation',
    'proto_arco_nome':  'Archery',
    'proto_arco_descr': 'Up to 30 trials, single window (Confirmation 1 > Confirmation 2)\nDemographic analysis: gender, category, style, accuracy',
    'curso_label':         '',
    'hist_titulo':         'Session History',
    'hist_subtitulo':      'Last {n} analyses run on this computer.',
    'hist_col_data':       'Date',
    'hist_col_proto':      'Protocol',
    'hist_col_atletas':    'Athletes',
    'hist_col_excel':      'Excel',
    'hist_col_pdf':        'PDF',
    'hist_sem':            'No history',
    'hist_btn_abrir':      '📂  Open folder',
    'hist_btn_reutilizar': '↺  Reuse configuration',
    'hist_btn_fechar':     'Close',
    'hist_cfg_recarregada':'Configuration reloaded from session {data}',
    'hist_pasta_nao_enc':  'Folder not found.',
    'sec_ficheiros_ent':   'INPUT FILES',
    'sec_ficheiros_sai':   'OUTPUT FILES',
    'sec_opcoes_gerais':   'GENERAL OPTIONS',
    'sec_exportacao':      'EXPORT',
    'sec_proto_tiro':      'SHOOTING PROTOCOL',
    'sec_dist_teste':      'TEST DISTANCES',
    'sec_anal_estat':      'STATISTICAL ANALYSIS',
    'tiro_stat_pos_disp':  'Position vs Shot + IP',
    'export_csv_label':    'Export CSV (group + trials)',
    'export_docx_label':   'Export Word report (.docx)',
    'tiro_assoc_por':      'Associate individual by:',
    'tiro_id_pasta':       'Folder ID (ID_Name)',
    'tiro_pos_lista':      'Position in list',
    'tiro_itvs_calc':      'Intervals to calculate:',
    'tiro_bipodal':        'Include bipodal analysis (Hurdle Step)',
    'tiro_n_ensaios':      '  N.trials:',
    'tiro_p_pe':           'p/foot',
    'tiro_dist_info':      'Distances in timing file (detected automatically).\nYou can add extra distances below:',
    'tiro_add_dist':       '+ Add distance',
    'tiro_itv_pont':       'Touch to Aim',
    'tiro_itv_disp':       'Touch to Shot',
    'tiro_itv_pont_disp':  'Aim to Shot',
    'tiro_itv_disp_fim':   'Shot to End',
    'tiro_itv_total':      'Full Trial',
    'pdf_capa_data':       'Date: {data}',
    'pdf_capa_n_indiv':    'No. of individuals: {n}',
    'pdf_capa_analisados': 'INDIVIDUALS ANALYSED',
    'pdf_capa_analisados2':'INDIVIDUALS ANALYSED (cont.)',
    'pdf_capa_distancias': 'Distances: {dists}',
    'pdf_capa_intervalos': 'Intervals: {itvs}',
    'pdf_capa_colab': '',
    'pdf_capa_colab2': '',
    'pdf_capa_mais':       '... and {n} more individuals (see next page)',
    'meses': ['January','February','March','April','May','June','July','August','September','October','November','December'],

    'log_inicio':       '▶  Starting analysis…',
    'log_concluido':    '✔  Analysis complete',
    'log_erro':         '✖  Error: {msg}',
    'log_aviso':        '⚠  {msg}',
    'log_update':       '⬆  New version available: v{nova}  (current: v{atual})',
    'log_individuo':    '→  Processing: {nome}',
    'log_sem_dados':    '⚠  No valid data: {nome}',

    'pdf_metrica':       'Metric',
    'pdf_pe_direito':    'RIGHT FOOT',
    'pdf_pe_esquerdo':   'LEFT FOOT',
    'pdf_posicao':       'POSITION',
    'pdf_disparo':       'SHOT',
    'pdf_ai':            'AI (%)',
    'pdf_max':           'max',
    'pdf_med':           'mean',
    'pdf_dp':            'sd',
    'pdf_resumo_grupo':  'GROUP SUMMARY',
    'pdf_estatisticas':  'DESCRIPTIVE STATISTICS',
    'pdf_normalidade':   'NORMALITY',
    'pdf_comparacao':    'COMPARISON',
    'pdf_correlacoes':   'CORRELATIONS',
    'pdf_relatorio':     'REPORT',
    'pdf_bilateral':     'Bilateral Analysis',
    'pdf_unilateral':    'Unilateral Analysis',
    'pdf_estabilidade':  'Postural Stability',
    'pdf_individuo':     'Subject',
    'pdf_grupo':         'Group',
    'pdf_n':             'n',
    'pdf_media':         'Mean',
    'pdf_desvio_padrao': 'SD',
    'pdf_cv':            'CV (%)',
    'pdf_min':           'Min.',
    'pdf_ic95':          '95% CI',
    'pdf_p_valor':       'p',
    'pdf_shapiro':       'Shapiro-Wilk',
    'pdf_normal':        'Normal',
    'pdf_nao_normal':    'Non-normal',
    'pdf_teste':         'Test',
    'pdf_cohen_d':       "Cohen's d",
    'pdf_efeito':        'Effect',
    'pdf_grande':        'Large',
    'pdf_medio':         'Medium',
    'pdf_pequeno':       'Small',
    'pdf_sem_efeito':    'Negligible',
    'pdf_assimetria':    'Asymmetry Index (%)',
    'pdf_direito':       'Right',
    'pdf_esquerdo':      'Left',
    'pdf_diferenca':     'Difference',
    'pdf_nota_clinica':  'Note: Values above 10% asymmetry may indicate clinically relevant imbalance.',
    'pdf_gerado_por':    'Generated by {prog} v{v} | {autor} | {univ}',
    'pdf_data_relatorio':'Report generated on {data}',
    'pdf_pag':           'p. {n}',

    'met_amp_x':      'Amplitude X / ML (mm)',
    'met_amp_y':      'Amplitude Y / AP (mm)',
    'met_vel_x':      'Mean velocity X (mm/s)',
    'met_vel_y':      'Mean velocity Y (mm/s)',
    'met_vel_med':    'Mean CoP velocity (mm/s)',
    'met_vel_pico_x': 'Peak velocity X (mm/s)',
    'met_vel_pico_y': 'Peak velocity Y (mm/s)',
    'met_desl':       'Displacement (mm)',
    'met_time':       'Time (s)',
    'met_ea95':       '95% Ellipse area (mm²)',
    'met_leng_a':     'Semi-axis a (mm)',
    'met_leng_b':     'Semi-axis b (mm)',
    'met_ratio_ml_ap':'ML/AP amplitude ratio',
    'met_ratio_vel':  'ML/AP velocity ratio',
    'met_stiff_x':    'Stiffness X (1/s)',
    'met_stiff_y':    'Stiffness Y (1/s)',
    'met_cov_xx':     'CoP X variance (mm²)',
    'met_cov_yy':     'CoP Y variance (mm²)',
    'met_cov_xy':     'CoP XY covariance (mm²)',
    'met_rms_x':      'RMS ML (mm)',
    'met_rms_y':      'RMS AP (mm)',
    'met_rms_r':      'RMS Radius (mm)',

    'aba_dados':   'DATA',
    'aba_grupo':   'GROUP',
    'aba_spss':    'SPSS',
    'aba_elipse':  'ellipse',
    'aba_estab':   'stabilogram',

    'word_titulo':          'Postural Stability Report',
    'word_subtitulo':       'Biomechanical Analysis - {protocolo}',
    'word_velocidades':     'Velocities',
    'word_elipse':          '95% Confidence Ellipse',
    'word_estabilograma':   'Stabilogram',

    'licenca_titulo':   'End User Licence Agreement (EULA)',
    'licenca_subtitulo':'Please read carefully before proceeding.',

    'pass_titulo':      'Access password',
    'pass_hint':        'Available at github.com/andremassuca',
    'pass_erro':        'Incorrect password. Please try again.',

    'poder_titulo':     'Sample Size Calculator',
    'poder_subtitulo':  'Based on Cohen (1988)  |  Paired t-test / Wilcoxon',
    'poder_calcular':   'Calculate',
    'poder_fechar':     'Close',
    'poder_presets':    'Quick presets:',

    'tip_embedded':     'Uses the time interval embedded in the .xls file\ninstead of the Start/end file.',
    'tip_elipse':       'Creates the "ellipse_..." tab in each individual file with the CoP scatter plot and 95% ellipse.',
    'tip_estab':        'Creates the "stabilogram_..." tab with the CoP displacement chart over time.',
    'tip_individuais':  'Creates one Excel file per subject with all trial, ellipse and stabilogram tabs.',
    'tip_n_ensaios':    'Overrides the number of trials defined by the protocol.',

    # ── General buttons ────────────────────────────────────────────────────
    'btn_stop':         '■  Stop',
    'btn_limpar':       '⊘ Clear',
    'btn_fechar':       'Close',
    'btn_calcular':     'Calculate',
    'btn_download':     'Download →',
    'btn_exportar_xl':  '💾 Export Excel',

    # ── Theme ──────────────────────────────────────────────────────────────
    'tema_claro':       'Switch to Light Mode',
    'tema_escuro':      'Switch to Dark Mode',
    'lbl_idioma':       'Language',

    # ── Shortcuts overlay ──────────────────────────────────────────────────
    'shortcuts_titulo': '⌨  Keyboard Shortcuts',

    # ── Profiles ───────────────────────────────────────────────────────────
    'perfis_titulo':    '📌  Configuration Profiles',
    'perfis_descr':     'Save and load named configuration sets.',
    'perfis_nome_lbl':  'Name:',
    'perfis_guardar':   '💾 Save current',
    'perfis_carregar':  '↑ Load',
    'perfis_apagar':    '✕ Delete',
    'perfil_guardado':  '✓ Profile saved: "{nm}"',
    'perfil_carregado': '✓ Profile loaded: "{nm}"',

    # ── Quick analysis ─────────────────────────────────────────────────────
    'quick_titulo':     '⚡  Quick Analysis - Single File',
    'quick_exportar':   '💾 Export Excel',

    # ── Update banner ──────────────────────────────────────────────────────
    'upd_banner_txt':   '  ⬆  Version v{nova} available',

    # ── History ────────────────────────────────────────────────────────────
    'hist_filtrar':     '  filter by protocol / date',

    # ── Power calculator ───────────────────────────────────────────────────
    'poder_titulo_win': 'Power Calculator  |  BSP',

    # ── UI labels ──────────────────────────────────────────────────────────
    'lbl_sep':          'Sep:',
    'lbl_dec':          'Dec:',
    'lbl_dist_m':       'Distance (m):',
    'iscpsi_colab': '',

    # ── ESTATS checkboxes ──────────────────────────────────────────────────
    'ck_estats':        'Automatic statistical tests (ESTATS tab)',
    'ck_estats_gr':     'Descriptives + SW + 95% CI',
    'ck_estats_de':     'Right vs Left + Cohen\'s d',
    'ck_estats_pd':     'Position vs Shot + IP',
    'ck_estats_at':     'Intra-subject variability (CV)',
    'ck_estats_fr':     'Friedman between intervals',
    'ck_estats_ph':     '  + post-hoc Bonferroni',
    'ck_estats_co':     'Correl. EA95 vs Score',

    # ── HTML export checkbox ────────────────────────────────────────────────
    'ck_html':          'Generate interactive HTML report (Chart.js)',

    # ── ESTATS and export tooltips ─────────────────────────────────────────
    'tip_posthoc':      'Wilcoxon post-hoc with Bonferroni correction\nif Friedman p < 0.05.',
    'tip_csv':          'Generates .csv files ready for R / SPSS / Excel:\n  *_group.csv   - means and SD per subject/side\n  *_trials.csv  - one trial per row (trial-level)\nSeparator and decimal configurable below.',
    'tip_csv_sep':      'Column separator: ; for Excel PT, , for Excel EN, Tab for R',
    'tip_csv_dec':      'Decimal separator: , for PT/ES, . for EN',
    'tip_docx':         'Generates a Word report with statistical tables (ESTATS).\nRequires: pip install python-docx\nIncludes descriptives + SW, Right vs Left, variability.',
    'tip_html':         'Generates a standalone .html file with interactive charts.\nOpens in any browser without installation.\nIdeal for sharing results.',
    'tip_dist_extra':   'Additional distance to process (e.g. 10, 25).\nFiles must follow the format trial[dist]_[trial].',
    'tip_n_ensaios_tiro':'Overrides the number of trials defined by the protocol.\nUseful for sessions with fewer or more trials than the default.',
    'tip_match':        'Defines how each subject folder is matched to the timing file.\n\nFolder ID (ID_Name): uses the number at the start of the folder name\n  E.g.: "66_Leonor" -> ID=66.\n  Robust even if folder order does not match the file.\n\nList position: uses alphabetical order of folders.',
    'tip_rb_id':        'Extracts the number at the start of the folder name (e.g. "66_Leonor" -> ID 66)\nand looks up that ID in the timing file.\nRecommended.',
    'tip_rb_idx':       'Associates by order: 1st folder -> row 1, 2nd folder -> row 2.\nUse only if folders have no numeric ID.',
    'tip_bipodal':      'Analyses the right_N and left_N files for each subject\nusing times from the "inicio_fim (Hurdle Step)" sheet.',
    'tip_n_ens_hs':     'Number of trials per foot (right and left) for Hurdle Step.\nDefault: 5.',

    # ── Analysis log messages ──────────────────────────────────────────────
    'log_tempos':       'Timing file: {nome}',
    'log_dist_det':     '  Distances detected: {dists}',
    'log_hs_n':         '  Hurdle Step: {n} subjects',
    'log_n_ind':        '  {n} subjects.',
    'log_tempo_nao_enc':'Timing file not found.',
    'log_scores':       'Scores: {n} subjects.',
    'log_sem_subs':     'No subfolders found.',
    'log_n_total':      '{n} subject(s)  [{proto}]',
    'log_ens_override': '  No. trials override: {n}',
    'log_gerar_excel':  '\nGenerating Excel (DATA + GROUP + SPSS {estats})...',
    'log_excel_ok':     'Excel: {f}',
    'log_csv':          'CSV: {f}',
    'log_aviso_csv':    'CSV warning: {e}',
    'log_word_ok':      'Word: {f}',
    'log_word_nao':     'Word not generated: {msg}',
    'log_aviso_docx':   'DOCX warning: {e}',
    'log_word_sem_estats': 'Word: enable statistical tests to generate the report.',
    'log_html_ok':      'HTML: {f}',
    'log_aviso_html':   'HTML warning: {e}',
    'log_fich_ind':     '\nIndividual files: {pasta}',
    'log_tiro_dist':    '  {nome} (shooting, {n} dist.)',
    'log_hs_ok':        '  {fn} (hurdle step)',
    'log_sel_ok':       '  {nome} – {lado} (sel files)',
    'log_aviso_sel':    '  sel warning {lado}: {e}',
    'log_gerar_pdf':    '\nGenerating PDF...',
    'log_concluido_n':  '\nComplete  -  {n} subject(s).',
    'log_cancelado':    'Cancelled.',
    'log_cfg_guardada': '✓ Configuration saved  ({mod}+S)',
    'log_cfg_erro':     'Error saving: {e}',
    'log_sem_dados':    'No data.',
    'log_erro_tb':      '\nError:\n{tb}',

    # ── Validation ─────────────────────────────────────────────────────────
    'val_titulo':       'Input validation',
    'val_intro':        'Fix the following before running:\n\n',
    'val_pasta_vazia':  'Subjects folder not defined.',
    'val_pasta_nao_enc':'Folder not found:\n{p}',
    'val_pasta_sem_sub':'No subfolders in the subjects folder.',
    'val_saida_vazia':  'Excel output file not defined.',
    'val_ifd_nao_enc':  'Timing file not found:\n{f}',

    # ── Installer / Uninstaller ────────────────────────────────────────────
    'inst_pasta':       'Installation folder:',
    'inst_opcoes':      'Options:',
    'inst_btn_inst':    '  Install  ',
    'inst_btn_cancel':  'Cancel',
    'inst_btn_instalar':'Installing...',
    'inst_sucesso':     '{app} installed successfully!',
    'inst_btn_abrir':   '  Open {app}  ',
    'inst_btn_fechar':  '  Close  ',
    'desinst_confirm':  'Are you sure you want to remove\n{app}?',
    'desinst_btn':      '  Uninstall  ',
    'desinst_removendo':'Removing...',
    'desinst_sucesso':  '{app} has been removed successfully.',
    'desinst_ficheiros':'Files will be deleted on close.',
},

# ─────────────────────────────────────────────────────────────────────────
'ES': {
    'app_title':           'BSP  |  Biomechanical Stability Program  v{v}',
    'prog_nome':           'Biomechanical Stability Program',
    'autor_label':         'Autor',
    'individuo_label':     'Sujeto',
    'data_label':          'Fecha',
    'versao_label':        'v{v}',
    'lang_tooltip':        'Cambiar idioma de la interfaz y los informes',

    'sec_ficheiros_entrada':  'ARCHIVOS DE ENTRADA',
    'sec_ficheiros_saida':    'ARCHIVOS DE SALIDA',
    'sec_opcoes_gerais':      'OPCIONES GENERALES',
    'sec_protocolo_tiro':     'PROTOCOLO DE TIRO',
    'sec_opcoes_estacao':     'OPCIONES DE ESTACIÓN',
    'sec_analise_bipodal':    'ANÁLISIS BIPODAL',
    'sec_analise_unipodal':   'ANÁLISIS UNIPODAL',
    'sec_resultados':         'RESULTADOS',
    'sec_log':                'REGISTRO DE EJECUCIÓN',

    'pasta_individuos':    'Carpeta de sujetos *',
    'pasta_individuos_s':  'Sujetos *',
    'pasta_individuos_tip':'Carpeta raíz con una subcarpeta por sujeto',
    'fich_tempos':         'Archivo Inicio_fin',
    'fich_tempos_s':       'Inicio_fin',
    'fich_tempos_tip':     'Archivo Excel con intervalos inicio/fin por ensayo',
    'fich_tempos_tiro':    'Archivo de tiempos (tiro) *',
    'fich_tempos_tiro_s':  'Tiempos *',
    'fich_tempos_tiro_tip':'Archivo Excel con tiempos de toque, puntería y disparo',
    'fich_atletas_ref':    'Referencia demográfica (142 atletas)',
    'fich_atletas_ref_s':  'Ref. demográfica',
    'fich_atletas_ref_tip':'Archivo Excel con peso, altura, género, estilo y puntuaciones de atletas de referencia',
    'fich_scores':         'Archivo de puntuaciones (opcional)',
    'fich_scores_s':       'Puntuaciones',
    'fich_scores_tip':     'Archivo Excel con puntuaciones de precisión (opcional)',
    'saida_excel':         'Resumen Excel *',
    'saida_excel_s':       'Excel salida *',
    'saida_excel_tip':     'Archivo Excel de salida con el resumen de resultados',
    'pasta_individuais':   'Carpeta de archivos individuales',
    'pasta_individuais_s': 'Individuales',
    'pasta_individuais_tip':'Carpeta para los archivos Excel individuales',
    'relatorio_pdf':       'Informe PDF (vacío = omitir)',
    'relatorio_pdf_s':     'PDF',
    'relatorio_pdf_tip':   'Ruta del informe PDF (vacío para omitir)',
    'relatorio_word':      'Informe Word (opcional)',
    'n_ensaios_label':     'N.º ensayos (vacío = protocolo por defecto)',
    'associar_individuo':  'Asociar sujeto por:',
    'id_pasta':            'ID de carpeta (ID_Nombre)',
    'posicao_lista':       'Posición en lista',
    'intervalos_calcular': 'Intervalos a calcular:',

    'ck_embedded':   'Usar intervalo incrustado en archivos .xls',
    'ck_elipse':     'Generar pestaña de elipse 95% con gráfico',
    'ck_estab':      'Generar pestaña de estabilograma (COP X e Y vs tiempo)',
    'ck_individuais':'Generar archivos individuales por sujeto',
    'ck_pdf':        'Generar informe PDF',
    'ck_word':       'Generar informe Word',
    'ck_bipodal':    'Incluir análisis bipodal (Hurdle Step)',

    'btn_executar':     '▶  Ejecutar Análisis',
    'btn_cancelar':     '✖  Cancelar',
    'btn_guardar':      '💾 Guardar',
    'btn_historico':    '📋 Historial',
    'btn_poder':        '🔬 Potencia',
    'btn_protocolo':    '⇄ Protocolo',
    'btn_tema':         '🌙 Tema',
    'btn_abrir_pasta':  '📂 Abrir Carpeta de Resultados',
    'btn_demografia':       'Análisis Demográfico',
    'demo_titulo':          'Análisis Demográfico (Tiro con Arco)',
    'demo_intro':           'Elige métrica y factor/variable y pulsa un botón abajo.',
    'demo_metrica':         'Métrica',
    'demo_factor':          'Factor',
    'demo_var_dem':         'Variable demográfica',
    'demo_btn_comparar':    'Comparar grupos',
    'demo_btn_corr_dem':    'Correlación demográfica',
    'demo_btn_corr_score':  'Puntuación vs CoP',
    'btn_aceitar':      '  ✓  He Leído y Acepto  ',
    'btn_recusar':      'Rechazar y Salir',
    'btn_entrar':       '  Acceder  ',
    'btn_abrir_github': '🔗  Abrir GitHub',
    'btn_aceitar_lic':  '  ✓  Aceptar  ',
    'btn_recusar_lic':  '  ✗  Rechazar  ',

    # ── Status bar / progresso ────────────────────────────────────────────
    'status_pronto':    '● Listo',
    'status_prog':      '▶ Procesando…',
    'acesso_pass':      'Contraseña de acceso',
    'acesso_pass_err':  'Contraseña incorrecta.',
    'acesso_github':    'Contraseña disponible en',
    'acesso_entrar':    '  Acceder  ',
    'log_bsp_header':   'BSP  -  {prog} v{versao}  -  {proto}',
    'log_atalhos':      'Atajos: {mod}+Enter ejecutar  |  {mod}+S guardar  |  {mod}+H historial',

    # ── Pantalla de selección de protocolo ───────────────────────────────
    'proto_titulo':     'Selecciona el protocolo de análisis:',
    'proto_confirmar':  'Confirmar y continuar',
    'proto_fechar':     'Cerrar',
    'proto_func_titulo':'Tarea Funcional',
    'proto_func_sub':   'Selecciona el protocolo específico',
    'proto_func_label': 'Protocolos de Tarea Funcional disponibles:',
    'proto_confirmar2': 'Confirmar',
    'proto_voltar':     'Volver',
    'proto_fms_nome':   'FMS Bipodal',
    'proto_fms_descr':  '5 ensayos por pie, bipodal\nElipse 95%, asimetría Der/Izq, todas las métricas',
    'proto_uni_nome':   'Apoyo Unipodal',
    'proto_uni_descr':  '5 ensayos por pie, unipodal\nElipse 95%, métricas de oscilación lateral',
    'proto_func_nome':  'Tarea Funcional',
    'proto_func_descr': 'Análisis de tarea funcional específica\nSubmenú: Tiro y otras tareas futuras',
    'proto_tiro_nome':  'Tiro',
    'proto_tiro_descr': '5 ensayos bipodal, 2 ventanas por ensayo\nComparación posición vs disparo, correlación con precisión',
    'proto_arco_nome':  'Tiro con Arco',
    'proto_arco_descr': 'Hasta 30 ensayos, ventana única (Confirmación 1 > Confirmación 2)\nAnálisis demográfico: género, categoría, estilo, precisión',
    'curso_label':         '',
    'hist_titulo':         'Historial de Sesiones',
    'hist_subtitulo':      'Últimos {n} análisis ejecutados en este ordenador.',
    'hist_col_data':       'Fecha',
    'hist_col_proto':      'Protocolo',
    'hist_col_atletas':    'Atletas',
    'hist_col_excel':      'Excel',
    'hist_col_pdf':        'PDF',
    'hist_sem':            'Sin historial',
    'hist_btn_abrir':      '📂  Abrir carpeta',
    'hist_btn_reutilizar': '↺  Reutilizar configuración',
    'hist_btn_fechar':     'Cerrar',
    'hist_cfg_recarregada':'Configuración recargada de la sesión {data}',
    'hist_pasta_nao_enc':  'Carpeta no encontrada.',
    'sec_ficheiros_ent':   'ARCHIVOS DE ENTRADA',
    'sec_ficheiros_sai':   'ARCHIVOS DE SALIDA',
    'sec_opcoes_gerais':   'OPCIONES GENERALES',
    'sec_exportacao':      'EXPORTACIÓN',
    'sec_proto_tiro':      'PROTOCOLO DE TIRO',
    'sec_dist_teste':      'DISTANCIAS DEL TEST',
    'sec_anal_estat':      'ANÁLISIS ESTADÍSTICO',
    'tiro_stat_pos_disp':  'Posición vs Disparo + IP',
    'export_csv_label':    'Exportar CSV (grupo + ensayos)',
    'export_docx_label':   'Exportar informe Word (.docx)',
    'tiro_assoc_por':      'Asociar individuo por:',
    'tiro_id_pasta':       'ID de carpeta (ID_Nombre)',
    'tiro_pos_lista':      'Posición en la lista',
    'tiro_itvs_calc':      'Intervalos a calcular:',
    'tiro_bipodal':        'Incluir análisis bipodal (Hurdle Step)',
    'tiro_n_ensaios':      '  N.ensayos:',
    'tiro_p_pe':           'p/pie',
    'tiro_dist_info':      'Distancias en el archivo de tiempos (detectadas automáticamente).\nPuedes añadir distancias adicionales abajo:',
    'tiro_add_dist':       '+ Añadir distancia',
    'tiro_itv_pont':       'Toque a Puntería',
    'tiro_itv_disp':       'Toque a Disparo',
    'tiro_itv_pont_disp':  'Puntería a Disparo',
    'tiro_itv_disp_fim':   'Disparo al Final',
    'tiro_itv_total':      'Ensayo Total',
    'pdf_capa_data':       'Fecha: {data}',
    'pdf_capa_n_indiv':    'N.º individuos: {n}',
    'pdf_capa_analisados': 'INDIVIDUOS ANALIZADOS',
    'pdf_capa_analisados2':'INDIVIDUOS ANALIZADOS (cont.)',
    'pdf_capa_distancias': 'Distancias: {dists}',
    'pdf_capa_intervalos': 'Intervalos: {itvs}',
    'pdf_capa_colab': '',
    'pdf_capa_colab2': '',
    'pdf_capa_mais':       '... y {n} individuos más (ver página siguiente)',
    'meses': ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'],

    'log_inicio':       '▶  Iniciando análisis…',
    'log_concluido':    '✔  Análisis completado',
    'log_erro':         '✖  Error: {msg}',
    'log_aviso':        '⚠  {msg}',
    'log_update':       '⬆  Nueva versión disponible: v{nova}  (actual: v{atual})',
    'log_individuo':    '→  Procesando: {nome}',
    'log_sem_dados':    '⚠  Sin datos válidos: {nome}',

    'pdf_metrica':       'Métrica',
    'pdf_pe_direito':    'PIE DERECHO',
    'pdf_pe_esquerdo':   'PIE IZQUIERDO',
    'pdf_posicao':       'POSICIÓN',
    'pdf_disparo':       'DISPARO',
    'pdf_ai':            'IA (%)',
    'pdf_max':           'máx',
    'pdf_med':           'media',
    'pdf_dp':            'dt',
    'pdf_resumo_grupo':  'RESUMEN DEL GRUPO',
    'pdf_estatisticas':  'ESTADÍSTICAS DESCRIPTIVAS',
    'pdf_normalidade':   'NORMALIDAD',
    'pdf_comparacao':    'COMPARACIÓN',
    'pdf_correlacoes':   'CORRELACIONES',
    'pdf_relatorio':     'INFORME',
    'pdf_bilateral':     'Análisis Bilateral',
    'pdf_unilateral':    'Análisis Unilateral',
    'pdf_estabilidade':  'Estabilidad Postural',
    'pdf_individuo':     'Sujeto',
    'pdf_grupo':         'Grupo',
    'pdf_n':             'n',
    'pdf_media':         'Media',
    'pdf_desvio_padrao': 'DT',
    'pdf_cv':            'CV (%)',
    'pdf_min':           'Mín.',
    'pdf_ic95':          'IC 95%',
    'pdf_p_valor':       'p',
    'pdf_shapiro':       'Shapiro-Wilk',
    'pdf_normal':        'Normal',
    'pdf_nao_normal':    'No normal',
    'pdf_teste':         'Prueba',
    'pdf_cohen_d':       'd de Cohen',
    'pdf_efeito':        'Efecto',
    'pdf_grande':        'Grande',
    'pdf_medio':         'Mediano',
    'pdf_pequeno':       'Pequeño',
    'pdf_sem_efeito':    'Sin efecto',
    'pdf_assimetria':    'Índice de Asimetría (%)',
    'pdf_direito':       'Der.',
    'pdf_esquerdo':      'Izq.',
    'pdf_diferenca':     'Diferencia',
    'pdf_nota_clinica':  'Nota: Valores superiores al 10% de asimetría pueden indicar un desequilibrio clínicamente relevante.',
    'pdf_gerado_por':    'Generado por {prog} v{v} | {autor} | {univ}',
    'pdf_data_relatorio':'Informe generado el {data}',
    'pdf_pag':           'p. {n}',

    'met_amp_x':      'Amplitud X / ML (mm)',
    'met_amp_y':      'Amplitud Y / AP (mm)',
    'met_vel_x':      'Vel. media X (mm/s)',
    'met_vel_y':      'Vel. media Y (mm/s)',
    'met_vel_med':    'Vel. media CoP (mm/s)',
    'met_vel_pico_x': 'Vel. pico X (mm/s)',
    'met_vel_pico_y': 'Vel. pico Y (mm/s)',
    'met_desl':       'Desplazamiento (mm)',
    'met_time':       'Tiempo (s)',
    'met_ea95':       'Área elipse 95% (mm²)',
    'met_leng_a':     'Semieje a (mm)',
    'met_leng_b':     'Semieje b (mm)',
    'met_ratio_ml_ap':'Rel. amp. ML/AP',
    'met_ratio_vel':  'Rel. vel. ML/AP',
    'met_stiff_x':    'Rigidez X (1/s)',
    'met_stiff_y':    'Rigidez Y (1/s)',
    'met_cov_xx':     'Var CoP X (mm²)',
    'met_cov_yy':     'Var CoP Y (mm²)',
    'met_cov_xy':     'Cov CoP XY (mm²)',
    'met_rms_x':      'RMS ML (mm)',
    'met_rms_y':      'RMS AP (mm)',
    'met_rms_r':      'RMS Radio (mm)',

    'aba_dados':   'DATOS',
    'aba_grupo':   'GRUPO',
    'aba_spss':    'SPSS',
    'aba_elipse':  'elipse',
    'aba_estab':   'estabilograma',

    'word_titulo':          'Informe de Estabilidad Postural',
    'word_subtitulo':       'Análisis Biomecánico - {protocolo}',
    'word_velocidades':     'Velocidades',
    'word_elipse':          'Elipse de Confianza 95%',
    'word_estabilograma':   'Estabilograma',

    'licenca_titulo':   'Acuerdo de Licencia de Usuario Final (EULA)',
    'licenca_subtitulo':'Lea detenidamente antes de continuar.',

    'pass_titulo':      'Contraseña de acceso',
    'pass_hint':        'Disponible en github.com/andremassuca',
    'pass_erro':        'Contraseña incorrecta. Inténtelo de nuevo.',

    'poder_titulo':     'Calculadora de Tamaño Muestral',
    'poder_subtitulo':  'Basada en Cohen (1988)  |  t-Student pareada / Wilcoxon',
    'poder_calcular':   'Calcular',
    'poder_fechar':     'Cerrar',
    'poder_presets':    'Presets rápidos:',

    'tip_embedded':     'Usa el intervalo de tiempo incrustado en el archivo .xls.',
    'tip_elipse':       'Crea la pestaña "elipse_..." con el diagrama de dispersión del COP y la elipse del 95%.',
    'tip_estab':        'Crea la pestaña "estabilograma_..." con el gráfico del desplazamiento del COP en el tiempo.',
    'tip_individuais':  'Crea un archivo Excel individual por sujeto.',
    'tip_n_ensaios':    'Reemplaza el número de ensayos definido por el protocolo.',

    'btn_stop':         '■  Detener',
    'btn_limpar':       '⊘ Limpiar',
    'btn_fechar':       'Cerrar',
    'btn_calcular':     'Calcular',
    'btn_download':     'Descargar →',
    'btn_exportar_xl':  '💾 Exportar Excel',
    'tema_claro':       'Cambiar a Modo Claro',
    'tema_escuro':      'Cambiar a Modo Oscuro',
    'lbl_idioma':       'Idioma',
    'shortcuts_titulo': '⌨  Atajos de Teclado',
    'perfis_titulo':    '📌  Perfiles de Configuración',
    'perfis_descr':     'Guarda y carga conjuntos de configuración con nombre.',
    'perfis_nome_lbl':  'Nombre:',
    'perfis_guardar':   '💾 Guardar actual',
    'perfis_carregar':  '↑ Cargar',
    'perfis_apagar':    '✕ Eliminar',
    'perfil_guardado':  '✓ Perfil guardado: "{nm}"',
    'perfil_carregado': '✓ Perfil cargado: "{nm}"',
    'quick_titulo':     '⚡  Análisis Rápido - Archivo Único',
    'quick_exportar':   '💾 Exportar Excel',
    'upd_banner_txt':   '  ⬆  Versión v{nova} disponible',
    'hist_filtrar':     '  filtrar por protocolo / fecha',
    'poder_titulo_win': 'Calculadora de Potencia  |  BSP',
    'lbl_sep':          'Sep:',
    'lbl_dec':          'Dec:',
    'lbl_dist_m':       'Distancia (m):',
    'iscpsi_colab': '',
    'ck_estats':        'Pruebas estadísticas automáticas (pestaña ESTATS)',
    'ck_estats_gr':     'Descriptivos + SW + IC 95%',
    'ck_estats_de':     'Der. vs Izq. + d de Cohen',
    'ck_estats_pd':     'Posición vs Disparo + IP',
    'ck_estats_at':     'Variab. intra-sujeto (CV)',
    'ck_estats_fr':     'Friedman entre intervalos',
    'ck_estats_ph':     '  + post-hoc Bonferroni',
    'ck_estats_co':     'Correl. EA95 vs Score',
    'ck_html':          'Generar informe HTML interactivo (Chart.js)',
    'tip_posthoc':      'Post-hoc de Wilcoxon con corrección de Bonferroni\nsi Friedman p < 0.05.',
    'tip_csv':          'Genera archivos .csv para R / SPSS / Excel:\n  *_grupo.csv   - medias y DE por sujeto/lado\n  *_trials.csv  - un ensayo por fila\nSeparador y decimal configurables.',
    'tip_csv_sep':      'Separador de columnas: ; para Excel ES, , para Excel EN, Tab para R',
    'tip_csv_dec':      'Separador decimal: , para PT/ES, . para EN',
    'tip_docx':         'Genera informe Word con tablas estadísticas (ESTATS).\nRequiere: pip install python-docx',
    'tip_html':         'Genera un archivo .html standalone con gráficos interactivos.\nSe abre en cualquier navegador sin instalación.',
    'tip_dist_extra':   'Distancia adicional (ej: 10, 25).\nLos archivos deben seguir el formato trial[dist]_[ensayo].',
    'tip_n_ensaios_tiro':'Reemplaza el número de ensayos del protocolo.\nÚtil para sesiones con más o menos ensayos.',
    'tip_match':        'Define cómo se asocia cada carpeta de sujeto con el archivo de tiempos.\n\nID de carpeta (ID_Nombre): usa el número al inicio del nombre.\n  Ej: "66_García" -> ID=66.\n\nPosición en lista: usa el orden alfabético de las carpetas.',
    'tip_rb_id':        'Extrae el número al inicio del nombre de la carpeta (ej: "66_García" -> ID 66).\nRecomendado.',
    'tip_rb_idx':       'Asocia por orden: 1.ª carpeta -> fila 1, 2.ª carpeta -> fila 2.\nUsar solo si las carpetas no tienen ID numérico.',
    'tip_bipodal':      'Analiza los archivos der_N y izq_N de cada sujeto\nusando los tiempos de la hoja "inicio_fim (Hurdle Step)".',
    'tip_n_ens_hs':     'Número de ensayos por pie (der. e izq.) para Hurdle Step.\nDefault: 5.',
    'log_tempos':       'Tiempos: {nome}',
    'log_dist_det':     '  Distancias detectadas: {dists}',
    'log_hs_n':         '  Hurdle Step: {n} sujetos',
    'log_n_ind':        '  {n} sujetos.',
    'log_tempo_nao_enc':'Archivo de tiempos no encontrado.',
    'log_scores':       'Puntuaciones: {n} sujetos.',
    'log_sem_subs':     'No se encontraron subcarpetas.',
    'log_n_total':      '{n} sujeto(s)  [{proto}]',
    'log_ens_override': '  N.º ensayos override: {n}',
    'log_gerar_excel':  '\nGenerando Excel (DATOS + GRUPO + SPSS {estats})...',
    'log_excel_ok':     'Excel: {f}',
    'log_csv':          'CSV: {f}',
    'log_aviso_csv':    'Aviso CSV: {e}',
    'log_word_ok':      'Word: {f}',
    'log_word_nao':     'Word no generado: {msg}',
    'log_aviso_docx':   'Aviso DOCX: {e}',
    'log_word_sem_estats': 'Word: activa las pruebas estadísticas para generar el informe.',
    'log_html_ok':      'HTML: {f}',
    'log_aviso_html':   'Aviso HTML: {e}',
    'log_fich_ind':     '\nArchivos individuales: {pasta}',
    'log_tiro_dist':    '  {nome} (tiro, {n} dist.)',
    'log_hs_ok':        '  {fn} (hurdle step)',
    'log_sel_ok':       '  {nome} – {lado} (sel archivos)',
    'log_aviso_sel':    '  aviso sel {lado}: {e}',
    'log_gerar_pdf':    '\nGenerando PDF...',
    'log_concluido_n':  '\nCompletado  -  {n} sujeto(s).',
    'log_cancelado':    'Cancelado.',
    'log_cfg_guardada': '✓ Configuración guardada  ({mod}+S)',
    'log_cfg_erro':     'Error al guardar: {e}',
    'log_sem_dados':    'Sin datos.',
    'log_erro_tb':      '\nError:\n{tb}',
    'val_titulo':       'Validación de entradas',
    'val_intro':        'Corrija los siguientes problemas antes de ejecutar:\n\n',
    'val_pasta_vazia':  'Carpeta de sujetos no definida.',
    'val_pasta_nao_enc':'Carpeta no encontrada:\n{p}',
    'val_pasta_sem_sub':'No hay subcarpetas en la carpeta de sujetos.',
    'val_saida_vazia':  'Archivo de salida Excel no definido.',
    'val_ifd_nao_enc':  'Archivo de tiempos no encontrado:\n{f}',
    'inst_pasta':       'Carpeta de instalación:',
    'inst_opcoes':      'Opciones:',
    'inst_btn_inst':    '  Instalar  ',
    'inst_btn_cancel':  'Cancelar',
    'inst_btn_instalar':'Instalando...',
    'inst_sucesso':     '¡{app} instalado correctamente!',
    'inst_btn_abrir':   '  Abrir {app}  ',
    'inst_btn_fechar':  '  Cerrar  ',
    'desinst_confirm':  '¿Seguro que deseas eliminar\n{app}?',
    'desinst_btn':      '  Desinstalar  ',
    'desinst_removendo':'Eliminando...',
    'desinst_sucesso':  '{app} se ha eliminado correctamente.',
    'desinst_ficheiros':'Los archivos se eliminarán al cerrar.',
},

# ─────────────────────────────────────────────────────────────────────────
'DE': {
    'app_title':           'BSP  |  Biomechanical Stability Program  v{v}',
    'prog_nome':           'Biomechanical Stability Program',
    'autor_label':         'Autor',
    'individuo_label':     'Proband',
    'data_label':          'Datum',
    'versao_label':        'v{v}',
    'lang_tooltip':        'Sprache der Benutzeroberfläche und Berichte ändern',

    'sec_ficheiros_entrada':  'EINGABEDATEIEN',
    'sec_ficheiros_saida':    'AUSGABEDATEIEN',
    'sec_opcoes_gerais':      'ALLGEMEINE OPTIONEN',
    'sec_protocolo_tiro':     'SCHIESSPROT0KOLL',
    'sec_opcoes_estacao':     'STATIONSOPTIONEN',
    'sec_analise_bipodal':    'BIPEDALE ANALYSE',
    'sec_analise_unipodal':   'EINBEINIGE ANALYSE',
    'sec_resultados':         'ERGEBNISSE',
    'sec_log':                'AUSFÜHRUNGSPROTOKOLL',

    'pasta_individuos':    'Probandenordner *',
    'pasta_individuos_s':  'Probanden *',
    'pasta_individuos_tip':'Hauptordner mit einem Unterordner pro Proband',
    'fich_tempos':         'Start/Ende-Datei',
    'fich_tempos_s':       'Start/Ende',
    'fich_tempos_tip':     'Excel-Datei mit Start/Ende-Intervallen je Versuch',
    'fich_tempos_tiro':    'Zeitdatei (Schießen) *',
    'fich_tempos_tiro_s':  'Zeitdatei *',
    'fich_tempos_tiro_tip':'Excel-Datei mit Berührungs-, Ziel- und Schusszeitstempeln',
    'fich_atletas_ref':    'Demographische Referenz (142 Athleten)',
    'fich_atletas_ref_s':  'Demogr. Referenz',
    'fich_atletas_ref_tip':'Excel-Datei mit Gewicht, Größe, Geschlecht, Stil und Scores der Referenzathleten',
    'fich_scores':         'Punktedatei (optional)',
    'fich_scores_s':       'Punkte',
    'fich_scores_tip':     'Excel-Datei mit Genauigkeitswerten je Proband (optional)',
    'saida_excel':         'Excel-Zusammenfassung *',
    'saida_excel_s':       'Excel Ausgabe *',
    'saida_excel_tip':     'Excel-Ausgabedatei mit der Ergebniszusammenfassung',
    'pasta_individuais':   'Ordner für Einzeldateien',
    'pasta_individuais_s': 'Einzeldateien',
    'pasta_individuais_tip':'Ordner für die einzelnen Excel-Dateien',
    'relatorio_pdf':       'PDF-Bericht (leer = überspringen)',
    'relatorio_pdf_s':     'PDF',
    'relatorio_pdf_tip':   'PDF-Berichtspfad (leer lassen zum Überspringen)',
    'relatorio_word':      'Word-Bericht (optional)',
    'n_ensaios_label':     'Anzahl Versuche (leer = Protokollstandard)',
    'associar_individuo':  'Probanden zuordnen nach:',
    'id_pasta':            'Ordner-ID (ID_Name)',
    'posicao_lista':       'Listenposition',
    'intervalos_calcular': 'Zu berechnende Intervalle:',

    'ck_embedded':   'Eingebettetes Zeitintervall aus .xls-Datei verwenden',
    'ck_elipse':     '95%-Ellipsen-Tab mit Diagramm generieren',
    'ck_estab':      'Stabilogramm-Tab generieren (COP X und Y vs. Zeit)',
    'ck_individuais':'Einzelne Excel-Dateien pro Proband generieren',
    'ck_pdf':        'PDF-Bericht generieren',
    'ck_word':       'Word-Bericht generieren',
    'ck_bipodal':    'Bipedale Analyse einbeziehen (Hurdle Step)',

    'btn_executar':     '▶  Analyse starten',
    'btn_cancelar':     '✖  Abbrechen',
    'btn_guardar':      '💾 Speichern',
    'btn_historico':    '📋 Verlauf',
    'btn_poder':        '🔬 Teststärke',
    'btn_protocolo':    '⇄ Protokoll',
    'btn_tema':         '🌙 Design',
    'btn_abrir_pasta':  '📂 Ergebnisordner öffnen',
    'btn_demografia':       'Demographische Analyse',
    'demo_titulo':          'Demographische Analyse (Bogenschießen)',
    'demo_intro':           'Metrik und Faktor/Variable wählen, dann unten Schaltfläche klicken.',
    'demo_metrica':         'Metrik',
    'demo_factor':          'Faktor',
    'demo_var_dem':         'Demographische Variable',
    'demo_btn_comparar':    'Gruppen vergleichen',
    'demo_btn_corr_dem':    'Demographische Korrelation',
    'demo_btn_corr_score':  'Score vs CoP',
    'btn_aceitar':      '  ✓  Gelesen und Akzeptiert  ',
    'btn_recusar':      'Ablehnen und Beenden',
    'btn_entrar':       '  Anmelden  ',
    'btn_abrir_github': '🔗  GitHub öffnen',
    'btn_aceitar_lic':  '  ✓  Akzeptieren  ',
    'btn_recusar_lic':  '  ✗  Ablehnen  ',

    # ── Status bar / progresso ────────────────────────────────────────────
    'status_pronto':    '● Bereit',
    'status_prog':      '▶ Verarbeitung…',
    'acesso_pass':      'Zugangskennwort',
    'acesso_pass_err':  'Falsches Kennwort.',
    'acesso_github':    'Kennwort verfügbar auf',
    'acesso_entrar':    '  Anmelden  ',
    'log_bsp_header':   'BSP  -  {prog} v{versao}  -  {proto}',
    'log_atalhos':      'Tastenkürzel: {mod}+Enter ausführen  |  {mod}+S speichern  |  {mod}+H Verlauf',

    # ── Protokollauswahl-Bildschirm ───────────────────────────────────────
    'proto_titulo':     'Analyseprotokoll auswählen:',
    'proto_confirmar':  'Bestätigen und weiter',
    'proto_fechar':     'Schließen',
    'proto_func_titulo':'Funktionsaufgabe',
    'proto_func_sub':   'Spezifisches Protokoll auswählen',
    'proto_func_label': 'Verfügbare Funktionsaufgaben-Protokolle:',
    'proto_confirmar2': 'Bestätigen',
    'proto_voltar':     'Zurück',
    'proto_fms_nome':   'FMS Bipodal',
    'proto_fms_descr':  '5 Versuche pro Fuß, bipodal\n95%-Ellipse, Dir/Links Asymmetrie, alle Metriken',
    'proto_uni_nome':   'Einbeiniger Stand',
    'proto_uni_descr':  '5 Versuche pro Fuß, unipodal\n95%-Ellipse, laterale Schwankungsmetriken',
    'proto_func_nome':  'Funktionsaufgabe',
    'proto_func_descr': 'Spezifische Funktionsaufgaben-Analyse\nUntermenü: Schießen und zukünftige Aufgaben',
    'proto_tiro_nome':  'Schießen',
    'proto_tiro_descr': '5 bipodal Versuche, 2 Fenster pro Versuch\nPosition vs. Schuss Vergleich, Genauigkeitskorrelation',
    'proto_arco_nome':  'Bogenschießen',
    'proto_arco_descr': 'Bis zu 30 Versuche, einzelnes Fenster (Bestätigung 1 > Bestätigung 2)\nDemografische Analyse: Geschlecht, Kategorie, Stil, Präzision',
    'curso_label':         '',
    'hist_titulo':          'Sitzungsverlauf',
    'hist_subtitulo':      'Letzte {n} Analysen auf diesem Computer.',
    'hist_col_data':       'Datum',
    'hist_col_proto':      'Protokoll',
    'hist_col_atletas':    'Athleten',
    'hist_col_excel':      'Excel',
    'hist_col_pdf':        'PDF',
    'hist_sem':            'Kein Verlauf',
    'hist_btn_abrir':      '📂  Ordner öffnen',
    'hist_btn_reutilizar': '↺  Konfiguration wiederverwenden',
    'hist_btn_fechar':     'Schließen',
    'hist_cfg_recarregada':'Konfiguration aus Sitzung {data} geladen',
    'hist_pasta_nao_enc':  'Ordner nicht gefunden.',
    'sec_ficheiros_ent':   'EINGABEDATEIEN',
    'sec_ficheiros_sai':   'AUSGABEDATEIEN',
    'sec_opcoes_gerais':   'ALLGEMEINE OPTIONEN',
    'sec_exportacao':      'EXPORT',
    'sec_proto_tiro':      'SCHIESSPROTOKOLL',
    'sec_dist_teste':      'TESTDISTANZEN',
    'sec_anal_estat':      'STATISTISCHE ANALYSE',
    'tiro_stat_pos_disp':  'Position vs Schuss + IP',
    'export_csv_label':    'CSV exportieren (Gruppe + Versuche)',
    'export_docx_label':   'Word-Bericht exportieren (.docx)',
    'tiro_assoc_por':      'Person zuordnen nach:',
    'tiro_id_pasta':       'Ordner-ID (ID_Name)',
    'tiro_pos_lista':      'Position in der Liste',
    'tiro_itvs_calc':      'Zu berechnende Intervalle:',
    'tiro_bipodal':        'Bipodal-Analyse einbeziehen (Hurdle Step)',
    'tiro_n_ensaios':      '  N.Versuche:',
    'tiro_p_pe':           'p/Fuß',
    'tiro_dist_info':      'Distanzen in der Zeitdatei (automatisch erkannt).\nZusätzliche Distanzen unten hinzufügen:',
    'tiro_add_dist':       '+ Distanz hinzufügen',
    'tiro_itv_pont':       'Berührung bis Ziel',
    'tiro_itv_disp':       'Berührung bis Schuss',
    'tiro_itv_pont_disp':  'Ziel bis Schuss',
    'tiro_itv_disp_fim':   'Schuss bis Ende',
    'tiro_itv_total':      'Gesamtversuch',
    'pdf_capa_data':       'Datum: {data}',
    'pdf_capa_n_indiv':    'Anzahl Personen: {n}',
    'pdf_capa_analisados': 'ANALYSIERTE PERSONEN',
    'pdf_capa_analisados2':'ANALYSIERTE PERSONEN (Forts.)',
    'pdf_capa_distancias': 'Distanzen: {dists}',
    'pdf_capa_intervalos': 'Intervalle: {itvs}',
    'pdf_capa_colab': '',
    'pdf_capa_colab2': '',
    'pdf_capa_mais':       '... und {n} weitere Personen (siehe nächste Seite)',
    'meses': ['Januar','Februar','März','April','Mai','Juni','Juli','August','September','Oktober','November','Dezember'],

    'log_inicio':       '▶  Analyse wird gestartet…',
    'log_concluido':    '✔  Analyse abgeschlossen',
    'log_erro':         '✖  Fehler: {msg}',
    'log_aviso':        '⚠  {msg}',
    'log_update':       '⬆  Neue Version verfügbar: v{nova}  (aktuell: v{atual})',
    'log_individuo':    '→  Verarbeite: {nome}',
    'log_sem_dados':    '⚠  Keine gültigen Daten: {nome}',

    'pdf_metrica':       'Metrik',
    'pdf_pe_direito':    'RECHTER FUß',
    'pdf_pe_esquerdo':   'LINKER FUß',
    'pdf_posicao':       'POSITION',
    'pdf_disparo':       'SCHUSS',
    'pdf_ai':            'AI (%)',
    'pdf_max':           'max',
    'pdf_med':           'MW',
    'pdf_dp':            'SD',
    'pdf_resumo_grupo':  'GRUPPENÜBERSICHT',
    'pdf_estatisticas':  'DESKRIPTIVE STATISTIK',
    'pdf_normalidade':   'NORMALITÄT',
    'pdf_comparacao':    'VERGLEICH',
    'pdf_correlacoes':   'KORRELATIONEN',
    'pdf_relatorio':     'BERICHT',
    'pdf_bilateral':     'Bilaterale Analyse',
    'pdf_unilateral':    'Unilaterale Analyse',
    'pdf_estabilidade':  'Posturale Stabilität',
    'pdf_individuo':     'Proband',
    'pdf_grupo':         'Gruppe',
    'pdf_n':             'n',
    'pdf_media':         'Mittelwert',
    'pdf_desvio_padrao': 'SD',
    'pdf_cv':            'VK (%)',
    'pdf_min':           'Min.',
    'pdf_ic95':          '95%-KI',
    'pdf_p_valor':       'p',
    'pdf_shapiro':       'Shapiro-Wilk',
    'pdf_normal':        'Normal',
    'pdf_nao_normal':    'Nicht normal',
    'pdf_teste':         'Test',
    'pdf_cohen_d':       "Cohen's d",
    'pdf_efeito':        'Effekt',
    'pdf_grande':        'Groß',
    'pdf_medio':         'Mittel',
    'pdf_pequeno':       'Klein',
    'pdf_sem_efeito':    'Vernachlässigbar',
    'pdf_assimetria':    'Asymmetrieindex (%)',
    'pdf_direito':       'Re.',
    'pdf_esquerdo':      'Li.',
    'pdf_diferenca':     'Differenz',
    'pdf_nota_clinica':  'Hinweis: Werte über 10% Asymmetrie können auf ein klinisch relevantes Ungleichgewicht hinweisen.',
    'pdf_gerado_por':    'Erstellt von {prog} v{v} | {autor} | {univ}',
    'pdf_data_relatorio':'Bericht erstellt am {data}',
    'pdf_pag':           'S. {n}',

    'met_amp_x':      'Amplitude X / ML (mm)',
    'met_amp_y':      'Amplitude Y / AP (mm)',
    'met_vel_x':      'Mittl. Geschwindigkeit X (mm/s)',
    'met_vel_y':      'Mittl. Geschwindigkeit Y (mm/s)',
    'met_vel_med':    'Mittl. CoP-Geschwindigkeit (mm/s)',
    'met_vel_pico_x': 'Spitzengeschwindigkeit X (mm/s)',
    'met_vel_pico_y': 'Spitzengeschwindigkeit Y (mm/s)',
    'met_desl':       'Verschiebung (mm)',
    'met_time':       'Zeit (s)',
    'met_ea95':       '95%-Ellipsenfläche (mm²)',
    'met_leng_a':     'Halbachse a (mm)',
    'met_leng_b':     'Halbachse b (mm)',
    'met_ratio_ml_ap':'ML/AP-Amplitudenverhältnis',
    'met_ratio_vel':  'ML/AP-Geschwindigkeitsverhältnis',
    'met_stiff_x':    'Steifigkeit X (1/s)',
    'met_stiff_y':    'Steifigkeit Y (1/s)',
    'met_cov_xx':     'CoP X Varianz (mm²)',
    'met_cov_yy':     'CoP Y Varianz (mm²)',
    'met_cov_xy':     'CoP XY Kovarianz (mm²)',
    'met_rms_x':      'RMS ML (mm)',
    'met_rms_y':      'RMS AP (mm)',
    'met_rms_r':      'RMS Radius (mm)',

    'aba_dados':   'DATEN',
    'aba_grupo':   'GRUPPE',
    'aba_spss':    'SPSS',
    'aba_elipse':  'ellipse',
    'aba_estab':   'stabilogramm',

    'word_titulo':          'Bericht zur posturalen Stabilität',
    'word_subtitulo':       'Biomechanische Analyse - {protocolo}',
    'word_velocidades':     'Geschwindigkeiten',
    'word_elipse':          '95%-Konfidenzellipse',
    'word_estabilograma':   'Stabilogramm',

    'licenca_titulo':   'Endbenutzer-Lizenzvertrag (EULA)',
    'licenca_subtitulo':'Bitte lesen Sie sorgfältig, bevor Sie fortfahren.',

    'pass_titulo':      'Zugangspasswort',
    'pass_hint':        'Verfügbar unter github.com/andremassuca',
    'pass_erro':        'Falsches Passwort. Bitte erneut versuchen.',

    'poder_titulo':     'Stichprobengrößenrechner',
    'poder_subtitulo':  'Basierend auf Cohen (1988)  |  Gepaarter t-Test / Wilcoxon',
    'poder_calcular':   'Berechnen',
    'poder_fechar':     'Schließen',
    'poder_presets':    'Schnellvorlagen:',

    'tip_embedded':     'Verwendet das in der .xls-Datei gespeicherte Zeitintervall.',
    'tip_elipse':       'Erstellt den Tab „ellipse_..." mit dem CoP-Streudiagramm und der 95%-Ellipse.',
    'tip_estab':        'Erstellt den Tab „stabilogramm_..." mit dem CoP-Verschiebungsdiagramm.',
    'tip_individuais':  'Erstellt eine Excel-Datei pro Proband.',
    'tip_n_ensaios':    'Überschreibt die im Protokoll definierte Versuchsanzahl.',

    'btn_stop':         '■  Stopp',
    'btn_limpar':       '⊘ Löschen',
    'btn_fechar':       'Schließen',
    'btn_calcular':     'Berechnen',
    'btn_download':     'Herunterladen →',
    'btn_exportar_xl':  '💾 Excel exportieren',
    'tema_claro':       'Hell-Modus aktivieren',
    'tema_escuro':      'Dunkel-Modus aktivieren',
    'lbl_idioma':       'Sprache',
    'shortcuts_titulo': '⌨  Tastenkürzel',
    'perfis_titulo':    '📌  Konfigurationsprofile',
    'perfis_descr':     'Konfigurationen mit Namen speichern und laden.',
    'perfis_nome_lbl':  'Name:',
    'perfis_guardar':   '💾 Aktuell speichern',
    'perfis_carregar':  '↑ Laden',
    'perfis_apagar':    '✕ Löschen',
    'perfil_guardado':  '✓ Profil gespeichert: „{nm}"',
    'perfil_carregado': '✓ Profil geladen: „{nm}"',
    'quick_titulo':     '⚡  Schnellanalyse - Einzeldatei',
    'quick_exportar':   '💾 Excel exportieren',
    'upd_banner_txt':   '  ⬆  Version v{nova} verfügbar',
    'hist_filtrar':     '  nach Protokoll / Datum filtern',
    'poder_titulo_win': 'Teststärkerechner  |  BSP',
    'lbl_sep':          'Sep:',
    'lbl_dec':          'Dez:',
    'lbl_dist_m':       'Distanz (m):',
    'iscpsi_colab': '',
    'ck_estats':        'Automatische statistische Tests (ESTATS-Blatt)',
    'ck_estats_gr':     'Deskriptiva + SW + 95%-KI',
    'ck_estats_de':     'Rechts vs. Links + Cohen\'s d',
    'ck_estats_pd':     'Position vs. Schuss + IP',
    'ck_estats_at':     'Intra-Proband-Variabilität (VK)',
    'ck_estats_fr':     'Friedman zwischen Intervallen',
    'ck_estats_ph':     '  + Post-hoc Bonferroni',
    'ck_estats_co':     'Korrel. EA95 vs. Score',
    'ck_html':          'Interaktiven HTML-Bericht erstellen (Chart.js)',
    'tip_posthoc':      'Wilcoxon-Post-hoc mit Bonferroni-Korrektur\nwenn Friedman p < 0,05.',
    'tip_csv':          'Erstellt .csv-Dateien für R / SPSS / Excel:\n  *_gruppe.csv   - Mittelwerte und SD\n  *_versuche.csv - ein Versuch pro Zeile\nTrennzeichen und Dezimalzeichen konfigurierbar.',
    'tip_csv_sep':      'Spaltentrennzeichen: ; für Excel DE, , für Excel EN, Tab für R',
    'tip_csv_dec':      'Dezimalzeichen: , für PT/ES/DE, . für EN',
    'tip_docx':         'Erstellt einen Word-Bericht mit statistischen Tabellen (ESTATS).\nErfordert: pip install python-docx',
    'tip_html':         'Erstellt eine eigenständige .html-Datei mit interaktiven Diagrammen.\nÖffnet in jedem Browser ohne Installation.',
    'tip_dist_extra':   'Zusätzliche Distanz (z. B. 10, 25).\nDateien müssen dem Format trial[dist]_[versuch] folgen.',
    'tip_n_ensaios_tiro':'Überschreibt die Versuchsanzahl des Protokolls.\nNützlich für Sitzungen mit mehr oder weniger Versuchen.',
    'tip_match':        'Legt fest, wie jeder Probandenordner der Zeitdatei zugeordnet wird.\n\nOrdner-ID (ID_Name): verwendet die Zahl am Anfang des Ordnernamens.\n  Bsp.: „66_Müller" -> ID=66.\n\nListenposition: verwendet alphabetische Reihenfolge.',
    'tip_rb_id':        'Liest die Zahl am Anfang des Ordnernamens (z. B. „66_Müller" -> ID 66).\nEmpfohlen.',
    'tip_rb_idx':       'Ordnet nach Reihenfolge zu: 1. Ordner -> Zeile 1.\nNur verwenden, wenn Ordner keine numerische ID haben.',
    'tip_bipodal':      'Analysiert die Dateien rechts_N und links_N jedes Probanden\nanhand der Zeiten aus dem Blatt „inicio_fim (Hurdle Step)".',
    'tip_n_ens_hs':     'Anzahl Versuche pro Fuß (rechts und links) für Hurdle Step.\nStandard: 5.',
    'log_tempos':       'Zeitdatei: {nome}',
    'log_dist_det':     '  Erkannte Distanzen: {dists}',
    'log_hs_n':         '  Hurdle Step: {n} Probanden',
    'log_n_ind':        '  {n} Probanden.',
    'log_tempo_nao_enc':'Zeitdatei nicht gefunden.',
    'log_scores':       'Scores: {n} Probanden.',
    'log_sem_subs':     'Keine Unterordner gefunden.',
    'log_n_total':      '{n} Proband(en)  [{proto}]',
    'log_ens_override': '  Versuchsanzahl Override: {n}',
    'log_gerar_excel':  '\nExcel wird erstellt (DATEN + GRUPPE + SPSS {estats})...',
    'log_excel_ok':     'Excel: {f}',
    'log_csv':          'CSV: {f}',
    'log_aviso_csv':    'CSV-Warnung: {e}',
    'log_word_ok':      'Word: {f}',
    'log_word_nao':     'Word nicht erstellt: {msg}',
    'log_aviso_docx':   'DOCX-Warnung: {e}',
    'log_word_sem_estats': 'Word: Statistische Tests aktivieren, um den Bericht zu erstellen.',
    'log_html_ok':      'HTML: {f}',
    'log_aviso_html':   'HTML-Warnung: {e}',
    'log_fich_ind':     '\nEinzeldateien: {pasta}',
    'log_tiro_dist':    '  {nome} (Schießen, {n} Dist.)',
    'log_hs_ok':        '  {fn} (Hurdle Step)',
    'log_sel_ok':       '  {nome} – {lado} (Sel-Dateien)',
    'log_aviso_sel':    '  Sel-Warnung {lado}: {e}',
    'log_gerar_pdf':    '\nPDF wird erstellt...',
    'log_concluido_n':  '\nAbgeschlossen  -  {n} Proband(en).',
    'log_cancelado':    'Abgebrochen.',
    'log_cfg_guardada': '✓ Konfiguration gespeichert  ({mod}+S)',
    'log_cfg_erro':     'Fehler beim Speichern: {e}',
    'log_sem_dados':    'Keine Daten.',
    'log_erro_tb':      '\nFehler:\n{tb}',
    'val_titulo':       'Eingabevalidierung',
    'val_intro':        'Bitte beheben Sie folgende Probleme:\n\n',
    'val_pasta_vazia':  'Probandenordner nicht definiert.',
    'val_pasta_nao_enc':'Ordner nicht gefunden:\n{p}',
    'val_pasta_sem_sub':'Keine Unterordner im Probandenordner.',
    'val_saida_vazia':  'Excel-Ausgabedatei nicht definiert.',
    'val_ifd_nao_enc':  'Zeitdatei nicht gefunden:\n{f}',
    'inst_pasta':       'Installationsordner:',
    'inst_opcoes':      'Optionen:',
    'inst_btn_inst':    '  Installieren  ',
    'inst_btn_cancel':  'Abbrechen',
    'inst_btn_instalar':'Wird installiert...',
    'inst_sucesso':     '{app} erfolgreich installiert!',
    'inst_btn_abrir':   '  {app} öffnen  ',
    'inst_btn_fechar':  '  Schließen  ',
    'desinst_confirm':  'Möchten Sie wirklich\n{app} entfernen?',
    'desinst_btn':      '  Deinstallieren  ',
    'desinst_removendo':'Wird entfernt...',
    'desinst_sucesso':  '{app} wurde erfolgreich entfernt.',
    'desinst_ficheiros':'Dateien werden beim Schließen gelöscht.',
},

}  # fim _STRINGS


# ═══════════════════════════════════════════════════════════════════════════
#  TEXTOS DE LICENÇA EULA (completos, por língua)
# ═══════════════════════════════════════════════════════════════════════════

LICENCA = {

'PT': """\
BSP - Biomechanical Stability Program  v{versao}
André Massuça, P. Aleixo & Luís M. Massuça

ACORDO DE LICENÇA DE UTILIZADOR FINAL (EULA)
© 2025-2026 André Massuça, P. Aleixo & Luís M. Massuça. Todos os direitos reservados.

LEIA ATENTAMENTE ESTE ACORDO ANTES DE INSTALAR OU UTILIZAR O SOFTWARE.
A instalação, cópia ou utilização do BSP constitui aceitação plena e
irrevogável dos presentes termos.

─────────────────────────────────────────────────────────────────────

1. CONCESSÃO DE LICENÇA
O autor concede ao utilizador uma licença pessoal, não exclusiva,
intransmissível e gratuita para instalar e utilizar o BSP
exclusivamente para fins académicos, de investigação científica
e avaliação clínica em contexto profissional qualificado.

2. RESTRIÇÕES
É expressamente proibido, sem autorização prévia por escrito do autor:
  a) reproduzir, distribuir ou sublicenciar o software ou qualquer parte;
  b) modificar, descompilar, desmontar ou realizar engenharia inversa;
  c) utilizar o software para fins comerciais ou de lucro;
  d) remover ou alterar avisos de direitos de autor ou de marca.

3. PROPRIEDADE INTELECTUAL
O BSP, incluindo código-fonte, algoritmos, documentação, ícones e
materiais associados, é propriedade de André Massuça, Pedro Aleixo e Luís M. Massuça.
Todos os direitos reservados. A utilização não implica transferência
de qualquer direito de propriedade intelectual.

4. ISENÇÃO DE GARANTIAS
O software é fornecido "NO ESTADO EM QUE SE ENCONTRA" (AS IS),
sem garantias de qualquer natureza, expressas ou implícitas, incluindo,
sem limitação, garantias de comerciabilidade, adequação a um fim
específico ou não infração. O autor não garante que o software seja
isento de erros ou funcione sem interrupção.

5. LIMITAÇÃO DE RESPONSABILIDADE
Em nenhuma circunstância o autor será responsável por danos diretos,
indiretos, incidentais, especiais ou consequentes resultantes da
utilização ou impossibilidade de utilização do software.

6. USO CLÍNICO E SEGURANÇA
O BSP é uma ferramenta de suporte à análise biomecânica.
NÃO constitui dispositivo médico, nem substitui avaliação clínica,
diagnóstico médico ou decisão terapêutica. A responsabilidade
pela interpretação dos resultados é exclusiva do profissional.

7. PRIVACIDADE E DADOS
Os dados clínicos e de análise (ficheiros de plataforma de força,
referências demográficas, scores) NÃO saem da máquina do utilizador -
todo o processamento dos dados ocorre exclusivamente em local.

Telemetria mínima: ao aceitar estes termos, é enviado uma única vez
um registo anónimo aos autores, contendo: data/hora da aceitação,
versão do BSP, língua da interface, sistema operativo e um identificador
anónimo da máquina (hash não-reversível). Não inclui nome, email,
caminhos de ficheiros ou dados clínicos. Para desactivar:
definir a variável de ambiente BSP_TELEMETRY_URL com valor vazio.

8. CITAÇÃO ACADÉMICA
  Massuça, A., Aleixo, P., & Massuça, L. M. (2026). BSP - Biomechanical Stability Program (v{versao}).
  https://github.com/andremassuca/BSP

9. RESCISÃO
Esta licença cessa automaticamente em caso de violação de qualquer
das suas cláusulas. O utilizador obriga-se a eliminar todas as cópias.

10. LEI APLICÁVEL
Este acordo é regulado pela lei portuguesa. Quaisquer litígios
serão submetidos à jurisdição dos tribunais de Lisboa, Portugal.

─────────────────────────────────────────────────────────────────────

Ao clicar em "Aceitar", o utilizador declara ter lido, compreendido
e aceite integralmente os termos deste Acordo de Licença.

Autores: {autor}  |  {univ}
""",

'EN': """\
BSP – Biomechanical Stability Program  v{versao}
André Massuça, P. Aleixo & Luís M. Massuça

END USER LICENCE AGREEMENT (EULA)
© 2025-2026 André Massuça, P. Aleixo & Luís M. Massuça. All rights reserved.

PLEASE READ THIS AGREEMENT CAREFULLY BEFORE INSTALLING OR USING THE SOFTWARE.
Installing, copying or using BSP constitutes full and irrevocable acceptance
of these terms.

─────────────────────────────────────────────────────────────────────

1. LICENCE GRANT
The author grants the user a personal, non-exclusive, non-transferable,
royalty-free licence to install and use BSP solely for academic,
scientific research and clinical assessment purposes in a qualified
professional context.

2. RESTRICTIONS
Without prior written authorisation from the author, it is strictly
prohibited to:
  a) reproduce, distribute or sublicense the software or any part thereof;
  b) modify, decompile, disassemble or reverse-engineer the software;
  c) use the software for commercial or for-profit purposes;
  d) remove or alter any copyright or trademark notices.

3. INTELLECTUAL PROPERTY
BSP, including source code, algorithms, documentation, icons and
associated materials, is the property of André Massuça and Luís M. Massuça.
All rights reserved. Use does not imply transfer of any intellectual
property rights.

4. DISCLAIMER OF WARRANTIES
The software is provided "AS IS", without warranties of any kind,
express or implied, including, without limitation, warranties of
merchantability, fitness for a particular purpose or non-infringement.
The author does not warrant that the software is error-free or will
operate without interruption.

5. LIMITATION OF LIABILITY
Under no circumstances shall the author be liable for any direct,
indirect, incidental, special, exemplary or consequential damages
resulting from the use of or inability to use the software.

6. CLINICAL USE AND SAFETY
BSP is a biomechanical analysis support tool. It does NOT constitute
a medical device and does not replace clinical assessment, medical
diagnosis or therapeutic decisions. Responsibility for the
interpretation of results rests solely with the professional user.

7. PRIVACY AND DATA
Clinical and analysis data (force-plate files, demographic references,
scores) NEVER leave the user's machine - all data processing is performed
exclusively locally.

Minimal telemetry: upon accepting these terms, a single anonymous
record is sent once to the authors, containing: timestamp of
acceptance, BSP version, interface language, operating system, and
an anonymous machine identifier (one-way hash). It does NOT include
name, email, file paths, or clinical data. To disable: set the
environment variable BSP_TELEMETRY_URL to an empty value.

8. ACADEMIC CITATION
  Massuça, A., Aleixo, P., & Massuça, L. M. (2026). BSP - Biomechanical Stability Program (v{versao}).
  https://github.com/andremassuca/BSP

9. TERMINATION
This licence terminates automatically upon breach of any of its terms.
Upon termination, the user must delete all copies of the software.

10. GOVERNING LAW
This agreement is governed by Portuguese law. Any disputes shall be
submitted to the exclusive jurisdiction of the courts of Lisbon, Portugal.

─────────────────────────────────────────────────────────────────────

By clicking "Accept", the user declares to have read, understood and
fully accepted the terms of this Licence Agreement.

Authors: {autor}  |  {univ}
""",

'ES': """\
BSP – Biomechanical Stability Program  v{versao}
André Massuça, P. Aleixo & Luís M. Massuça

ACUERDO DE LICENCIA DE USUARIO FINAL (EULA)
© 2025-2026 André Massuça, P. Aleixo & Luís M. Massuça. Todos los derechos reservados.

LEA DETENIDAMENTE ESTE ACUERDO ANTES DE INSTALAR O UTILIZAR EL SOFTWARE.
La instalación, copia o uso del BSP constituye la aceptación plena e
irrevocable de los presentes términos.

─────────────────────────────────────────────────────────────────────

1. CONCESIÓN DE LICENCIA
El autor concede al usuario una licencia personal, no exclusiva,
intransferible y gratuita para instalar y utilizar BSP exclusivamente
con fines académicos, de investigación científica y evaluación clínica
en un contexto profesional cualificado.

2. RESTRICCIONES
Sin autorización previa por escrito del autor, queda expresamente
prohibido:
  a) reproducir, distribuir o sublicenciar el software o cualquier parte;
  b) modificar, descompilar, desensamblar o realizar ingeniería inversa;
  c) utilizar el software con fines comerciales o lucrativos;
  d) eliminar o alterar avisos de derechos de autor o marca.

3. PROPIEDAD INTELECTUAL
BSP, incluido el código fuente, algoritmos, documentación, iconos y
materiales asociados, es propiedad de André Massuça y Luís M. Massuça.
Todos los derechos reservados.

4. EXCLUSIÓN DE GARANTÍAS
El software se proporciona "TAL CUAL" (AS IS), sin garantías de ningún
tipo, expresas o implícitas, incluyendo, sin limitación, garantías de
comerciabilidad, idoneidad para un fin particular o no infracción.

5. LIMITACIÓN DE RESPONSABILIDAD
En ningún caso el autor será responsable de daños directos, indirectos,
incidentales, especiales o consecuentes derivados del uso o la
imposibilidad de uso del software.

6. USO CLÍNICO Y SEGURIDAD
BSP es una herramienta de apoyo al análisis biomecánico. NO constituye
un dispositivo médico ni sustituye la evaluación clínica, el diagnóstico
médico ni las decisiones terapéuticas.

7. PRIVACIDAD Y DATOS
Los datos clínicos y de análisis (archivos de plataforma de fuerza,
referencias demográficas, puntuaciones) NUNCA salen del equipo del
usuario - todo el procesamiento se realiza exclusivamente de forma
local.

Telemetría mínima: al aceptar estos términos, se envía una única vez
un registro anónimo a los autores con: fecha/hora de la aceptación,
versión del BSP, idioma de la interfaz, sistema operativo y un
identificador anónimo de la máquina (hash no reversible). No incluye
nombre, email, rutas de archivos ni datos clínicos. Para desactivar:
definir la variable de entorno BSP_TELEMETRY_URL como vacía.

8. CITA ACADÉMICA
  Massuça, A., Aleixo, P., & Massuça, L. M. (2026). BSP - Biomechanical Stability Program (v{versao}).
  https://github.com/andremassuca/BSP

9. RESCISIÓN
Esta licencia cesa automáticamente ante cualquier incumplimiento de sus
cláusulas. El usuario se compromete a eliminar todas las copias.

10. LEY APLICABLE
Este acuerdo se rige por la legislación portuguesa. Cualquier litigio
será sometido a la jurisdicción de los tribunales de Lisboa, Portugal.

─────────────────────────────────────────────────────────────────────

Al hacer clic en "Aceptar", el usuario declara haber leído, comprendido
y aceptado íntegramente los términos de este Acuerdo de Licencia.

Autores: {autor}  |  {univ}
""",

'DE': """\
BSP – Biomechanical Stability Program  v{versao}
André Massuça, P. Aleixo & Luís M. Massuça

ENDBENUTZER-LIZENZVERTRAG (EULA)
© 2025-2026 André Massuça, P. Aleixo & Luís M. Massuça. Alle Rechte vorbehalten.

BITTE LESEN SIE DIESEN VERTRAG SORGFÄLTIG, BEVOR SIE DIE SOFTWARE
INSTALLIEREN ODER VERWENDEN.
Die Installation, Vervielfältigung oder Nutzung von BSP stellt die
vollständige und unwiderrufliche Annahme dieser Bedingungen dar.

─────────────────────────────────────────────────────────────────────

1. LIZENZERTEILUNG
Der Autor gewährt dem Benutzer eine persönliche, nicht ausschließliche,
nicht übertragbare, gebührenfreie Lizenz zur Installation und Nutzung
von BSP ausschließlich für akademische, wissenschaftliche und klinische
Bewertungszwecke in einem qualifizierten professionellen Umfeld.

2. EINSCHRÄNKUNGEN
Ohne vorherige schriftliche Genehmigung des Autors ist Folgendes
ausdrücklich untersagt:
  a) Vervielfältigung, Vertrieb oder Weiterlizenzierung der Software;
  b) Modifizierung, Dekompilierung, Disassemblierung oder Reverse Engineering;
  c) kommerzielle oder gewinnbringende Nutzung der Software;
  d) Entfernung oder Änderung von Urheberrechts- oder Markenhinweisen.

3. GEISTIGES EIGENTUM
BSP, einschließlich Quellcode, Algorithmen, Dokumentation, Icons und
zugehörigen Materialien, ist ausschließliches Eigentum von André Massuca.

4. GEWÄHRLEISTUNGSAUSSCHLUSS
Die Software wird „WIE SIE IST" (AS IS) ohne Gewährleistungen jeglicher
Art bereitgestellt, weder ausdrücklich noch stillschweigend, einschließlich
der Gewährleistung der Marktgängigkeit oder der Eignung für einen
bestimmten Zweck.

5. HAFTUNGSBESCHRÄNKUNG
Der Autor haftet unter keinen Umständen für direkte, indirekte,
zufällige, besondere oder Folgeschäden aus der Nutzung oder der
Unmöglichkeit der Nutzung der Software.

6. KLINISCHE NUTZUNG UND SICHERHEIT
BSP ist ein Unterstützungswerkzeug für die biomechanische Analyse.
Es stellt KEIN Medizinprodukt dar und ersetzt keine klinische Untersuchung,
medizinische Diagnose oder therapeutische Entscheidung.

7. DATENSCHUTZ
Klinische und Analyse-Daten (Kraftplatten-Dateien, demografische
Referenzen, Scores) verlassen NIEMALS den Computer des Benutzers -
die gesamte Datenverarbeitung erfolgt ausschließlich lokal.

Minimale Telemetrie: bei Annahme dieser Bedingungen wird einmalig ein
anonymer Datensatz an die Autoren gesendet, der Folgendes enthält:
Zeitstempel der Annahme, BSP-Version, Sprache der Oberfläche,
Betriebssystem und eine anonyme Geräte-Kennung (Einweg-Hash). Er
enthält KEINEN Namen, keine E-Mail, keine Dateipfade oder klinischen
Daten. Zum Deaktivieren: die Umgebungsvariable BSP_TELEMETRY_URL
auf einen leeren Wert setzen.

8. WISSENSCHAFTLICHE ZITATION
  Massuça, A., Aleixo, P., & Massuça, L. M. (2026). BSP - Biomechanical Stability Program (v{versao}).
  https://github.com/andremassuca/BSP

9. KÜNDIGUNG
Diese Lizenz erlischt automatisch bei Verstoß gegen ihre Bedingungen.
Der Benutzer ist verpflichtet, alle Kopien zu löschen.

10. ANWENDBARES RECHT
Dieser Vertrag unterliegt portugiesischem Recht. Streitigkeiten werden
den Gerichten in Lissabon, Portugal, vorgelegt.

─────────────────────────────────────────────────────────────────────

Durch Klicken auf „Akzeptieren" erklärt der Benutzer, die Bedingungen
dieses Lizenzvertrags gelesen, verstanden und vollständig akzeptiert
zu haben.

Autoren: {autor}  |  {univ}
""",

}  # fim LICENCA


def licenca_texto(lang=None, versao='23', autor='André O. Massuça',
                  univ='Pedro Aleixo  |  Luís M. Massuça') -> str:
    """Devolve o texto da licença na língua indicada (default: língua actual)."""
    lang = lang or _LANG_ATUAL[0]
    tpl = LICENCA.get(lang, LICENCA['PT'])
    return tpl.format(versao=versao, autor=autor, univ=univ)


# Preencher referência na tabela PT (evita duplicação)
_STRINGS['PT']['licenca_texto'] = lambda **kw: licenca_texto('PT', **kw)
_STRINGS['EN']['licenca_texto'] = lambda **kw: licenca_texto('EN', **kw)
_STRINGS['ES']['licenca_texto'] = lambda **kw: licenca_texto('ES', **kw)
_STRINGS['DE']['licenca_texto'] = lambda **kw: licenca_texto('DE', **kw)


# ═══════════════════════════════════════════════════════════════════════════
#  METS_PDF e METS_XL localizadas
# ═══════════════════════════════════════════════════════════════════════════

_MET_KEYS_PDF = [
    'met_amp_x', 'met_amp_y', 'met_vel_x', 'met_vel_y', 'met_vel_med',
    'met_vel_pico_x', 'met_vel_pico_y', 'met_desl', 'met_time',
    'met_ea95', 'met_leng_a', 'met_leng_b',
    'met_rms_x', 'met_rms_y', 'met_rms_r',
    'met_ratio_ml_ap', 'met_ratio_vel',
    'met_stiff_x', 'met_stiff_y',
    'met_cov_xx', 'met_cov_yy', 'met_cov_xy',
]

_MET_INTERNAL_KEYS = [
    'amp_x', 'amp_y', 'vel_x', 'vel_y', 'vel_med',
    'vel_pico_x', 'vel_pico_y', 'desl', 'time',
    'ea95', 'leng_a', 'leng_b',
    'rms_x', 'rms_y', 'rms_r',
    'ratio_ml_ap', 'ratio_vel',
    'stiff_x', 'stiff_y',
    'cov_xx', 'cov_yy', 'cov_xy',
]

def mets_pdf_localizadas():
    """Devolve lista [(chave_interna, label_traduzido)] para o PDF."""
    return [(ik, T(lk)) for ik, lk in zip(_MET_INTERNAL_KEYS, _MET_KEYS_PDF)]

def mets_xl_localizadas():
    """Devolve lista [(chave_interna, label_traduzido)] para o Excel (inclui SPSS)."""
    return [(ik, T(lk)) for ik, lk in zip(_MET_INTERNAL_KEYS, _MET_KEYS_PDF)]

def lados_pdf_localizados():
    """Devolve mapeamento chave→label para os lados no PDF."""
    return {
        'dir':  T('pdf_pe_direito'),
        'esq':  T('pdf_pe_esquerdo'),
        'pos':  T('pdf_posicao'),
        'disp': T('pdf_disparo'),
    }
