"""Prepara o ambiente: instala o Ollama (Linux e Windows), sobe o servidor e baixa o modelo.

Rode uma vez, num terminal:   python instalar.py      (-y responde "sim" a todas as perguntas)
Usa so a biblioteca padrao do Python, entao roda antes de instalar qualquer dependencia."""
import json, os, re, shutil, subprocess, sys, tempfile, time, urllib.error, urllib.request

URL_BASE = os.environ.get('OLLAMA_URL', 'http://127.0.0.1:11434')
URL_LINUX = 'https://ollama.com/install.sh'
URL_WINDOWS = 'https://ollama.com/download/OllamaSetup.exe'
ESPERA_APOS_INSTALAR = 40   # segundos esperando o servico subir sozinho depois de instalar
ESPERA_APOS_INICIAR = 30    # segundos esperando depois de iniciar o servidor na mao
TIMEOUT_DOWNLOAD = 600


def _modelo():
    """Le o MODELO do resume.py, para os dois arquivos nunca divergirem."""
    try:
        pasta = os.path.dirname(os.path.abspath(__file__))
        with open(os.path.join(pasta, 'resume.py'), encoding='utf-8') as f:
            m = re.search(r"^MODELO\s*=\s*['\"]([^'\"]+)['\"]", f.read(), re.MULTILINE)
        if m:
            return m.group(1)
    except OSError:
        pass
    return 'qwen2.5:7b'


def _modelos_instalados():
    """Conjunto de modelos instalados, ou None se o Ollama nao responde."""
    try:
        with urllib.request.urlopen(URL_BASE + '/api/tags', timeout=3) as r:
            return {m.get('name', '') for m in json.load(r).get('models', [])}
    except (urllib.error.URLError, OSError, ValueError):
        return None


def _perguntar(texto, sim):
    if sim:
        print(texto + 's')
        return True
    try:
        return input(texto).strip().lower() == 's'
    except EOFError:
        return False


def _esperar_servidor(segundos):
    print('Aguardando o Ollama responder', end='', flush=True)
    fim = time.time() + segundos
    while time.time() < fim:
        if _modelos_instalados() is not None:
            print(' ok')
            return True
        print('.', end='', flush=True)
        time.sleep(2)
    print(' nao respondeu')
    return False


def _achar_exe():
    exe = shutil.which('ollama')
    if exe:
        return exe
    if sys.platform == 'win32':
        candidato = os.path.expandvars(r'%LOCALAPPDATA%\Programs\Ollama\ollama.exe')
        if os.path.exists(candidato):
            return candidato
    return None


def _subir_servidor(exe):
    """Inicia 'ollama serve' em segundo plano (para quando o Ollama esta instalado mas parado)."""
    extra = {}
    if sys.platform == 'win32':
        extra['creationflags'] = (getattr(subprocess, 'CREATE_NEW_PROCESS_GROUP', 0)
                                  | getattr(subprocess, 'DETACHED_PROCESS', 0))
    else:
        extra['start_new_session'] = True
    subprocess.Popen([exe, 'serve'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, **extra)


def _baixar_arquivo(url, destino):
    req = urllib.request.Request(url, headers={'User-Agent': 'instalar.py'})
    with urllib.request.urlopen(req, timeout=TIMEOUT_DOWNLOAD) as r, open(destino, 'wb') as f:
        total = int(r.headers.get('Content-Length') or 0)
        baixado, ultimo = 0, -1
        while True:
            bloco = r.read(1024 * 256)
            if not bloco:
                break
            f.write(bloco)
            baixado += len(bloco)
            if total:
                pct = baixado * 100 // total
                if pct != ultimo:
                    print(f'\rBaixando instalador... {pct}%'.ljust(40), end='', flush=True)
                    ultimo = pct
    print()


def _instalar_linux(sim):
    if not shutil.which('curl'):
        print('Falta o curl. Instale com: sudo apt install curl   (e rode este script de novo)')
        return False
    if not sys.stdin.isatty():
        print('Rode este script num terminal de verdade (nao no console Run do PyCharm):')
        print('o instalador pede a senha do sudo.')
        return False
    print(f'Vai rodar o instalador oficial do Ollama:  curl -fsSL {URL_LINUX} | sh')
    print('Ele pede a senha do sudo e instala o Ollama como servico do sistema.')
    if not _perguntar('Instalar o Ollama agora? [s/n] ', sim):
        return False
    return subprocess.run(f'curl -fsSL {URL_LINUX} | sh', shell=True).returncode == 0


def _instalar_windows(sim):
    print('Vai baixar o instalador oficial do Ollama (OllamaSetup.exe, arquivo grande) e abri-lo.')
    print('Conclua a instalacao na janela que aparecer.')
    if not _perguntar('Baixar e instalar o Ollama agora? [s/n] ', sim):
        return False
    destino = os.path.join(tempfile.gettempdir(), 'OllamaSetup.exe')
    try:
        _baixar_arquivo(URL_WINDOWS, destino)
    except (urllib.error.URLError, OSError) as e:
        print(f'Falha ao baixar o instalador: {e}')
        return False
    print('Abrindo o instalador... (este script continua quando ele terminar)')
    return subprocess.run([destino]).returncode == 0


def _garantir_servidor(sim):
    """Deixa o Ollama instalado e respondendo. Devolve True se deu certo."""
    if _modelos_instalados() is not None:
        print('Ollama ja esta rodando.')
        return True
    if not re.match(r'https?://(127\.0\.0\.1|localhost)(:|/|$)', URL_BASE):
        print(f'O Ollama configurado em {URL_BASE} nao responde, e ele nao fica nesta maquina.')
        print('Este script so instala e inicia o Ollama local.')
        return False

    exe = _achar_exe()
    if exe is None:
        if sys.platform.startswith('linux'):
            ok = _instalar_linux(sim)
        elif sys.platform == 'win32':
            ok = _instalar_windows(sim)
        else:
            print('Neste sistema a instalacao e manual: https://ollama.com/download')
            print('Depois de instalar, rode este script de novo.')
            return False
        if not ok:
            print('Instalacao nao concluida.')
            return False
        if _esperar_servidor(ESPERA_APOS_INSTALAR):
            return True
        exe = _achar_exe()
        if exe is None:
            print('Nao encontrei o comando ollama depois da instalacao.')
            print('Abra um terminal novo (para atualizar o PATH) e rode este script de novo.')
            return False

    print('Ollama instalado, mas o servidor nao esta rodando. Iniciando...')
    _subir_servidor(exe)
    if _esperar_servidor(ESPERA_APOS_INICIAR):
        return True
    print('O servidor nao subiu. Tente manualmente em outro terminal: ollama serve')
    return False


def _baixar_modelo(modelo):
    print(f'Baixando {modelo} (pode levar varios minutos)...')
    req = urllib.request.Request(
        URL_BASE + '/api/pull',
        data=json.dumps({'model': modelo, 'stream': True}).encode(),
        headers={'Content-Type': 'application/json'})
    ultimo = ''
    with urllib.request.urlopen(req, timeout=TIMEOUT_DOWNLOAD) as r:
        for linha in r:
            d = json.loads(linha)
            if d.get('error'):
                raise OSError(d['error'])
            if d.get('total') and d.get('completed'):
                txt = f"{d.get('status', '')} {int(d['completed'] * 100 / d['total'])}%"
            else:
                txt = d.get('status', '')
            if txt != ultimo:
                print('\r' + txt.ljust(60), end='', flush=True)
                ultimo = txt
    print()


def _preparar(silencioso=False):
    sim = any(a in ('-y', '--sim') for a in sys.argv[1:])
    modelo = _modelo()
    print(f'Preparando o ambiente (modelo: {modelo})')

    if not _garantir_servidor(sim):
        return 1

    nomes = _modelos_instalados() or set()
    if modelo in nomes or modelo + ':latest' in nomes:
        print(f'Modelo {modelo} ja esta instalado.')
    else:
        if not _perguntar(f'Modelo {modelo} nao esta instalado. Baixar agora? [s/n] ', sim):
            print('Sem o modelo o script principal nao consegue resumir as descricoes.')
            return 1
        try:
            _baixar_modelo(modelo)
        except (urllib.error.URLError, OSError, ValueError) as e:
            print(f'Falha ao baixar {modelo}: {e}')
            return 1
        nomes = _modelos_instalados() or set()
        if modelo not in nomes and modelo + ':latest' not in nomes:
            print(f'O download terminou, mas o {modelo} nao apareceu na lista do Ollama.')
            return 1

    if silencioso:
        print('Ambiente pronto.\n')
    else:
        print('\nPronto. Agora rode:  python main.py planilha.xlsx')
    return 0


def main(silencioso=False):
    try:
        return _preparar(silencioso)
    except KeyboardInterrupt:
        print('\nInterrompido pelo usuario. Rode de novo para continuar de onde parou.')
        return 1


if __name__ == '__main__':
    sys.exit(main())