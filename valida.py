import os, re, unicodedata
from openpyxl import load_workbook
from fpdf import FPDF

# Teste independente do main.py: reabre o XLSX salvo e confere, perfil a perfil (pelo nome),
# se a descricao gravada e exatamente a descricao original depois da limpeza.


def _limpar(texto):
    t = unicodedata.normalize('NFD', texto)
    t = ''.join(c for c in t if unicodedata.category(c) != 'Mn')
    t = re.sub(r'[^A-Za-z0-9 ]', ' ', t)
    return re.sub(r'\s+', ' ', t).strip()


def _chave(nome):
    return re.sub(r'\s+', ' ', nome).strip().lower()


def _ler_saida(caminho):
    wb = load_workbook(caminho, read_only=True, data_only=True)
    linhas = list(wb.worksheets[0].iter_rows(min_row=2, values_only=True))
    wb.close()
    gerado = {}
    for l in linhas:
        nome = '' if l[0] is None else str(l[0]).strip()
        desc = '' if len(l) < 2 or l[1] is None else str(l[1])
        if nome:
            gerado.setdefault(nome, []).append(desc)
    return gerado


def _pdf_texto(t, limite=110):
    t = t if len(t) <= limite else t[:limite - 3] + '...'
    return t.encode('latin-1', 'replace').decode('latin-1')


def _gerar_pdf(caminho, arquivo, resumo, problemas):
    pdf = FPDF(format='A4')
    pdf.set_auto_page_break(False)          # garante uma pagina so
    pdf.add_page()
    pdf.set_margins(12, 12, 12)
    pdf.set_xy(12, 12)
    pdf.set_font('Helvetica', 'B', 14)
    pdf.cell(0, 8, 'Validacao das descricoes', new_x='LMARGIN', new_y='NEXT')
    pdf.set_font('Helvetica', '', 9)
    pdf.cell(0, 5, _pdf_texto('Arquivo conferido: ' + arquivo, 120), new_x='LMARGIN', new_y='NEXT')
    pdf.ln(2)
    for rotulo, valor in resumo:
        pdf.set_font('Helvetica', 'B', 10)
        pdf.cell(95, 6, rotulo)
        pdf.set_font('Helvetica', '', 10)
        pdf.cell(0, 6, str(valor), new_x='LMARGIN', new_y='NEXT')
    pdf.ln(3)
    pdf.set_font('Helvetica', 'B', 11)
    if not problemas:
        pdf.cell(0, 7, 'Nenhum problema encontrado: tudo o que esta na base esta na planilha gerada.')
    else:
        pdf.cell(0, 7, f'O que faltou ou ficou diferente ({len(problemas)}) e por que', new_x='LMARGIN', new_y='NEXT')
        pdf.set_font('Helvetica', '', 8)
        mostrados = 0
        for tipo, nome, motivo, detalhe in problemas:
            linhas = 2 + (1 if detalhe else 0)
            if pdf.get_y() + linhas * 4 > 280:
                break
            pdf.set_font('Helvetica', 'B', 8)
            pdf.cell(0, 4, _pdf_texto(f'[{tipo}] {nome}', 120), new_x='LMARGIN', new_y='NEXT')
            pdf.set_font('Helvetica', '', 8)
            pdf.cell(0, 4, _pdf_texto('   Motivo: ' + motivo, 130), new_x='LMARGIN', new_y='NEXT')
            if detalhe:
                pdf.cell(0, 4, _pdf_texto('   ' + detalhe, 130), new_x='LMARGIN', new_y='NEXT')
            mostrados += 1
        if mostrados < len(problemas):
            pdf.ln(1)
            pdf.set_font('Helvetica', 'I', 8)
            pdf.cell(0, 4, f'... e mais {len(problemas) - mostrados} item(ns) que nao cabem nesta pagina. Veja a lista completa no terminal.')
    pdf.output(caminho)


def validar(originais, caminho_saida):
    """originais: lista de (nome do perfil, descricao ORIGINAL, sem limpar).
    caminho_saida: XLSX gerado pelo main.py."""
    gerado = _ler_saida(caminho_saida)
    por_chave = {}
    for n in gerado:
        por_chave.setdefault(_chave(n), []).append(n)

    iguais = sem_desc = 0
    problemas = []
    usados = set()
    for nome, original in originais:
        esperado = _limpar(original)
        if nome in gerado and gerado[nome]:
            obtido = gerado[nome].pop(0)
            usados.add(nome)
            if obtido == esperado:
                if esperado:
                    iguais += 1
                else:
                    sem_desc += 1
            else:
                problemas.append(('DIFERENTE', nome, 'a descricao gravada nao e a original depois da limpeza',
                                  f'esperado: {esperado or "(vazio)"} | gravado: {obtido or "(vazio)"}'))
            continue
        quase = [n for n in por_chave.get(_chave(nome), []) if n != nome]
        if nome in usados:
            motivo = 'nome repetido na base: so existe uma linha na planilha gerada'
        elif quase:
            motivo = f'o nome aparece diferente na planilha gerada ("{quase[0]}"), so em maiusculas ou espacos'
        else:
            motivo = 'o nome do perfil nao aparece na planilha gerada (linha nao gravada ou nome alterado)'
        problemas.append(('FALTOU', nome, motivo, ''))

    sobrando = [n for n, restantes in gerado.items() for _ in restantes]
    for n in sobrando:
        problemas.append(('SOBROU', n, 'o perfil esta na planilha gerada mas nao esta na base', ''))

    resumo = [
        ('Perfis na base', len(originais)),
        ('Iguais (descricao conferida)', iguais),
        ('Iguais, sem descricao na base', sem_desc),
        ('Diferentes', sum(p[0] == 'DIFERENTE' for p in problemas)),
        ('Faltando na planilha gerada', sum(p[0] == 'FALTOU' for p in problemas)),
        ('Sobrando na planilha gerada', len(sobrando)),
    ]
    print('\n== Validacao ==')
    for r, v in resumo:
        print(f'{r}: {v}')
    for tipo, nome, motivo, detalhe in problemas:
        print(f'  [{tipo}] {nome} -> {motivo}')
        if detalhe:
            print('      ', detalhe)
    pdf = os.path.splitext(caminho_saida)[0] + '_validacao.pdf'
    _gerar_pdf(pdf, os.path.basename(caminho_saida), resumo, problemas)
    print('Relatorio:', pdf)
    return not problemas