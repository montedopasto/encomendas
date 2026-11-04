#!/usr/bin/env python3
"""
Gera public/index.html a partir de src/encomendas-zolve.html (a app tal como está publicada no Claude).

Uso:  python3 tools/build.py

O que muda em relação à versão do Claude:
  - documento HTML completo (doctype, head, body);
  - carrega config.js, a biblioteca de entrada Microsoft (MSAL) e platform.js antes da app;
  - ExcelJS e JSZip vêm da pasta vendor/ (sem depender de CDN);
  - os textos que mandam "ligar o conector no claude.ai" passam a falar da conta Microsoft.
A lógica da app não é alterada.
"""
import pathlib, re, sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
src = (ROOT / 'src' / 'encomendas-zolve.html').read_text(encoding='utf-8')

# Cabeça (title, meta, fontes, estilos) e corpo (header, main, scripts).
cut = src.index('</style>') + len('</style>')
head, body = src[:cut], src[cut:]

def must_replace(text, old, new, count=1):
    n = text.count(old)
    if n != count:
        sys.exit(f'build.py: esperava {count}x {old[:60]!r}, encontrei {n}x — a app mudou; rever o build.')
    return text.replace(old, new)

body = must_replace(body, '<script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"></script>', '<script src="vendor/exceljs.min.js"></script>')
body = must_replace(body, '<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>', '<script src="vendor/jszip.min.js"></script>')

# Mensagens de erro: na versão do GitHub não há "conectores do claude.ai".
TEXTS = [
    (r"Ligue-o em claude\.ai → Definições → Conectores, com a conta @montedopasto\.pt\.", 'Entre de novo com a conta @montedopasto.pt (botão com as suas iniciais, no topo).'),
    (r"Ligue-o em claude\.ai → Definições → Conectores, ou use \"Descarregar Excel\"\.", 'Entre de novo com a conta Microsoft da empresa, ou use "Descarregar Excel".'),
    (r"Ligue-o em claude\.ai → Definições → Conectores\.", 'Esta ligação não existe na versão do GitHub.'),
    (r"Volte a ligá-la em claude\.ai → Definições → Conectores\.", 'Recarregue a página e entre de novo com a conta Microsoft.'),
    (r"Volte a ligá-lo em claude\.ai → Definições → Conectores\.", 'Esta ligação não existe na versão do GitHub.'),
    (r"O Microsoft 365 não está ligado à sua conta claude\.ai\.", 'A sessão Microsoft não está ativa.'),
    (r"quando o claude\.ai perguntar", 'quando o Microsoft perguntar'),
]
for pat, rep in TEXTS:
    body = re.sub(pat, rep, body)

# Na versão do GitHub o rascunho é criado diretamente pelo Microsoft Graph (sem fluxo Power Automate).
body = must_replace(body, 'através do fluxo Power Automate da pasta SalesVision.', 'diretamente pelo Microsoft 365.')
body = must_replace(body, 'Pedido enviado. O rascunho com o Excel anexado vai aparecer na pasta <b>Rascunhos</b> do Outlook dentro de cerca de 1 minuto. A app vai verificar.', 'Rascunho criado na pasta <b>Rascunhos</b> do Outlook, com o Excel anexado. A app vai confirmar.')

html = f'''<!doctype html>
<html lang="pt">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#0B0B0B">
<link rel="icon" href="icon.png">
<link rel="apple-touch-icon" href="icon.png">
<link rel="manifest" href="manifest.webmanifest">
{head}
<script src="config.js"></script>
<script src="vendor/msal-browser.min.js"></script>
<script src="platform.js"></script>
</head>
<body>
{body}
</body>
</html>
'''
(ROOT / 'public' / 'index.html').write_text(html, encoding='utf-8')
print('public/index.html gerado (' + str(len(html)) + ' bytes)')
