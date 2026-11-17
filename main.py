import sys, os, re, csv, io, unicodedata
from openpyxl import Workbook, load_workbook

COLUNA_DESC = 'descricao do perfil'   # o cabecalho so precisa COMECAR com isso
COLUNA_PERFIL = 'perfil'              # cabecalho exato da coluna com o nome completo do perfil
TITULO_NOME = 'nome do perfil'
TITULO_DESC = 'descricao do perfil'
PREFIXO_SAIDA = 'descricaoAjustada'


def limpar(texto):
    t = unicodedata.normalize('NFD', texto)
    t = ''.join(c for c in t if unicodedata.category(c) != 'Mn')
    t = re.sub(r'[^A-Za-z0-9 ]', ' ', t)
    return re.sub(r'\s+', ' ', t).strip()


def norm(v):
    return limpar(v).lower()


def ler_linhas(caminho):
    ext = os.path.splitext(caminho)[1].lower()
    if ext == '.xlsx':
        wb = load_workbook(caminho, read_only=True, data_only=True)
        linhas = [['' if c is None else str(c) for c in row] for row in wb.worksheets[0].iter_rows(values_only=True)]
        wb.close()
        return linhas
    with open(caminho, encoding='utf-8-sig', newline='') as f:
        conteudo = f.read()
    primeira = conteudo.split('\n', 1)[0]
    delim = ';' if primeira.count(';') > primeira.count(',') else ','
    return list(csv.reader(io.StringIO(conteudo), delimiter=delim))


def proximo_arquivo(pasta):
    padrao = re.compile(rf'^{PREFIXO_SAIDA}_(\d+)\.xlsx$', re.IGNORECASE)
    usados = [int(m.group(1)) for n in os.listdir(pasta) if (m := padrao.match(n))]
    versao = max(usados, default=0) + 1
    return os.path.join(pasta, f'{PREFIXO_SAIDA}_{versao:02d}.xlsx')


def achar_entrada(pasta):
    saida = re.compile(rf'^{PREFIXO_SAIDA}_\d+\.xlsx$', re.IGNORECASE)
    cand = sorted(n for n in os.listdir(pasta)
                  if n.lower().endswith(('.csv', '.xlsx')) and not n.startswith('~$') and not saida.match(n))
    if len(cand) == 1:
        return os.path.join(pasta, cand[0])
    if not cand:
        sys.exit(f'Nao achei nenhum .csv ou .xlsx em {pasta}')
    sys.exit('Achei mais de uma planilha na pasta, deixe so a base:\n  ' + '\n  '.join(cand))


pasta_script = os.path.dirname(os.path.abspath(__file__))
entrada = sys.argv[1] if len(sys.argv) > 1 else achar_entrada(pasta_script)
print('Lendo:', entrada)
linhas = ler_linhas(entrada)

# acha a linha do cabecalho (nas 6 primeiras) e as duas colunas
cab = i_desc = i_perfil = None
for n, linha in enumerate(linhas[:6]):
    nomes = [norm(v) for v in linha]
    i_d = next((i for i, nome in enumerate(nomes) if nome.startswith(COLUNA_DESC)), None)
    if i_d is not None and COLUNA_PERFIL in nomes:
        cab, i_desc, i_perfil = n, i_d, nomes.index(COLUNA_PERFIL)
        break
if cab is None:
    sys.exit('Nao achei as colunas "Perfil" e "Descricao do perfil..." nas 6 primeiras linhas.')

resultado, originais, sem_desc, sem_nome, vistos, repetidos = [], [], [], 0, set(), set()
for l in linhas[cab + 1:]:
    l = l + [''] * (max(i_desc, i_perfil) + 1 - len(l))
    nome = l[i_perfil].strip()
    desc = limpar(l[i_desc])
    if not nome:
        if desc:
            sem_nome += 1
        continue
    if nome in vistos:
        repetidos.add(nome)
    vistos.add(nome)
    if not desc:
        sem_desc.append(nome)
    resultado.append((nome, desc))
    originais.append((nome, l[i_desc]))

wb = Workbook()
ws = wb.active
ws.title = 'descricoes'
ws.append([TITULO_NOME, TITULO_DESC])
for nome, desc in resultado:
    ws.append([nome, desc])
ws.column_dimensions['A'].width = 50
ws.column_dimensions['B'].width = 110
ws.freeze_panes = 'A2'

saida = proximo_arquivo(os.path.dirname(os.path.abspath(entrada)))
wb.save(saida)

print(f'{len(resultado)} perfis gravados, {len(sem_desc)} sem descricao.')
for n in sem_desc:
    print('  sem descricao:', n)
if sem_nome:
    print(f'AVISO: {sem_nome} linha(s) com descricao mas sem nome do perfil foram ignoradas.')
for n in sorted(repetidos):
    print('AVISO: nome repetido na planilha:', n)
print('Gerado:', saida)

from valida import validar
validar(originais, saida)