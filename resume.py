import json, os, re, time, urllib.error, urllib.request
from openpyxl.styles import Font, PatternFill

# Resume com um modelo local (Ollama) as descricoes que passam de LIMITE caracteres.
# O modelo escreve texto livre; quem garante "so letras, numeros e espacos" e o limpar() do main.py.

URL_BASE = os.environ.get('OLLAMA_URL', 'http://127.0.0.1:11434')
MODELO = 'qwen2.5:7b'
LIMITE = 250          # descricoes acima disso sao enviadas ao modelo
META = 200            # tamanho pedido ao modelo na 1a tentativa (cai 40 a cada nova tentativa)
TENTATIVAS = 3
ENTRADA_MAX = 6000    # descricao maior que isso (em caracteres) nao e enviada: fica para revisao manual
TEMPERATURA = 0.2
NUM_CTX = 4096
NUM_PREDICT = 150     # teto de tokens gerados por resposta
TIMEOUT = 600         # segundos por chamada

# Exemplos genericos (nada do cliente). Cada par: (descricao longa, resumo bom). Para ensinar o formato
# com casos reais, troque ou acrescente pares aqui, de tipos de perfil diferentes e com resumos de ate ~200 caracteres.
EXEMPLOS = [
    ('Perfil destinado aos colaboradores do almoxarifado responsaveis por receber mercadorias conferir notas fiscais registrar entradas e saidas de materiais controlar o estoque e solicitar reposicao junto ao setor de compras',
     'Perfil do almoxarifado para receber e conferir mercadorias, registrar entradas e saidas, controlar o estoque e pedir reposicao ao setor de compras.'),
    ('Objetivo analisar e decidir sobre solicitacoes de viagem corporativa visualizar e aprovar pedidos da equipe consultar o historico de gastos e delegar aprovacoes temporariamente com acesso restrito a unidade de Curitiba sem criar editar ou excluir solicitacoes',
     'Aprova ou rejeita solicitacoes de viagem da equipe, consulta o historico de gastos e delega aprovacoes temporariamente, com acesso restrito a unidade de Curitiba e sem criar ou editar solicitacoes.'),
    ('Perfil administrativo do sistema financeiro com acesso total a configuracao de usuarios licencas e integracoes e a relatorios de todas as areas sujeito a revisao semestral e a monitoramento por ser de alto risco',
     'Administra usuarios, licencas e integracoes do sistema financeiro e acessa relatorios de todas as areas. Perfil de alto risco, sujeito a revisao semestral e monitoramento.'),
]

PROMPT = """Voce e um redator tecnico que reescreve descricoes de perfis de acesso de sistemas corporativos para um portal de solicitacao de acesso. Quem le o texto e um solicitante que precisa decidir se pede o perfil.

Tarefa: reescreva a descricao abaixo em ate {META} caracteres, em uma ou duas frases completas em portugues, claras e coerentes.

Regras:
1. Mantenha o mesmo sentido. Nao invente nenhuma informacao que nao esteja no texto.
2. Preserve o que mais importa: para que serve o perfil, o escopo (pais, area ou empresa) e as principais restricoes, ou seja, o que o perfil nao permite.
3. Comece direto pelo conteudo. Nao copie rotulos do texto original como Objetivo, Escopo ou Acao.
4. Nao repita palavras em sequencia, nao copie trechos soltos e nao faca lista de palavras.
5. Mantenha nomes de sistemas, paises e siglas como aparecem no texto.
6. Responda somente com o texto reescrito, sem comentarios, aspas ou titulo.

{EXEMPLOS}

Texto: {TEXTO}
Resumo:"""


def _montar_prompt(texto, meta):
    exemplos = '\n\n'.join(f'Texto: {t}\nResumo: {r}' for t, r in EXEMPLOS)
    # {TEXTO} por ultimo, para que chaves dentro do texto nao sejam trocadas
    return PROMPT.replace('{META}', str(meta)).replace('{EXEMPLOS}', exemplos).replace('{TEXTO}', texto)


def _verificar_servidor():
    """Devolve None se o Ollama e o modelo estao prontos, ou a mensagem do problema."""
    try:
        with urllib.request.urlopen(URL_BASE + '/api/tags', timeout=5) as r:
            nomes = {m.get('name', '') for m in json.load(r).get('models', [])}
    except (urllib.error.URLError, OSError, ValueError):
        return f'Ollama nao respondeu em {URL_BASE} (rode: docker start ollama)'
    if MODELO not in nomes and MODELO + ':latest' not in nomes:
        return f'modelo {MODELO} nao esta baixado (rode: docker exec ollama ollama pull {MODELO})'
    return None


def _chamar(prompt):
    corpo = json.dumps({
        'model': MODELO, 'prompt': prompt, 'stream': False, 'keep_alive': '5m',
        'options': {'temperature': TEMPERATURA, 'num_ctx': NUM_CTX, 'num_predict': NUM_PREDICT},
    }).encode()
    req = urllib.request.Request(URL_BASE + '/api/generate', data=corpo,
                                 headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
        return json.load(r)


def _resumir_um(texto, limpar):
    """Devolve (resumo ja limpo, tentativas usadas, motivo da falha). Resumo vazio = falhou."""
    motivo = ''
    for t in range(1, TENTATIVAS + 1):
        meta = max(META - 40 * (t - 1), 80)
        prompt = _montar_prompt(texto, meta)
        try:
            r = _chamar(prompt)
        except (urllib.error.URLError, OSError, ValueError) as e:
            motivo = f'erro ao chamar o Ollama: {e}'
            continue
        if r.get('done_reason') != 'stop':
            motivo = f'resposta cortada ({r.get("done_reason")})'
            continue
        bruto = re.sub(r'^\s*resumo\s*:?\s*', '', r.get('response', ''), flags=re.IGNORECASE)
        resumo = limpar(bruto)
        if not resumo:
            motivo = 'resposta vazia'
            continue
        if len(resumo) > LIMITE:
            motivo = f'resumo ainda com {len(resumo)} caracteres'
            continue
        return resumo, t, ''
    return '', TENTATIVAS, motivo


def resumir(resultado, limpar):
    """resultado: lista de (nome, descricao ja limpa). Devolve (novo resultado, linhas da aba revisao,
    conjunto de nomes resumidos). Descricao que nao deu para resumir fica como estava, marcada 'revisar'."""
    finais = [d for _, d in resultado]
    status = ['sem necessidade'] * len(resultado)
    pendentes = [i for i, (_, d) in enumerate(resultado) if len(d) > LIMITE]
    for i in pendentes:
        status[i] = ''

    if pendentes:
        erro = _verificar_servidor()
        if erro:
            print(f'AVISO: {len(pendentes)} descricao(oes) acima de {LIMITE} caracteres, mas {erro}. Nada foi resumido.')
            for i in pendentes:
                status[i] = 'revisar: ' + erro
            pendentes = []
        else:
            print(f'{len(pendentes)} descricao(oes) acima de {LIMITE} caracteres, resumindo com {MODELO} (Ctrl+C interrompe e salva o que ja foi feito)')

    for k, i in enumerate(pendentes, 1):
        nome, desc = resultado[i]
        if len(desc) > ENTRADA_MAX:
            status[i] = f'revisar: descricao muito grande ({len(desc)} caracteres)'
            print(f'[{k}/{len(pendentes)}] {nome[:60]} -> pulada, muito grande')
            continue
        inicio = time.time()
        try:
            resumo, tentativas, motivo = _resumir_um(desc, limpar)
        except KeyboardInterrupt:
            print('\nInterrompido: o restante fica marcado para revisao.')
            break
        if resumo:
            finais[i] = resumo
            status[i] = f'resumida ({tentativas} tentativa(s))'
            print(f'[{k}/{len(pendentes)}] {nome[:60]} {len(desc)} -> {len(resumo)} ({time.time() - inicio:.0f}s)')
        else:
            status[i] = 'revisar: ' + motivo
            print(f'[{k}/{len(pendentes)}] {nome[:60]} -> falhou: {motivo}')

    for i in pendentes:
        if status[i] == '':
            status[i] = 'revisar: interrompido'

    novo = [(nome, finais[i]) for i, (nome, _) in enumerate(resultado)]
    revisao = [(nome, desc, finais[i], len(desc), len(finais[i]), status[i])
               for i, (nome, desc) in enumerate(resultado)]
    resumidos = {r[0] for r in revisao if r[5].startswith('resumida')}
    n_resumidas = sum(r[5].startswith('resumida') for r in revisao)
    a_revisar = sum(r[5].startswith('revisar') for r in revisao)
    if n_resumidas or a_revisar:
        print(f'Resumidas: {n_resumidas} | para revisar manualmente: {a_revisar}')
    return novo, revisao, resumidos


def adicionar_aba_revisao(wb, revisao):
    ws = wb.create_sheet('revisao')
    ws.append(['nome do perfil', 'descricao original (limpa)', 'descricao final', 'tamanho antes', 'tamanho depois', 'status'])
    for c in ws[1]:
        c.font = Font(bold=True)
    vermelho = PatternFill('solid', fgColor='F8CBAD')
    for linha in revisao:
        ws.append(list(linha))
        if linha[5].startswith('revisar'):
            for c in ws[ws.max_row]:
                c.fill = vermelho
    for col, largura in zip('ABCDEF', (50, 70, 70, 14, 14, 45)):
        ws.column_dimensions[col].width = largura
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions