# Encomendas

Aplicação web estática. Autenticação Microsoft Entra ID e integração com Microsoft Graph.

## Publicação

O fluxo GitHub Actions gera e publica a pasta `public/` no GitHub Pages.

## Configuração

`public/config.js` contém apenas configuração pública de uma SPA. Nunca adicionar dados, tokens, credenciais ou client secrets ao repositório.

## Gerar a página

`python3 tools/build.py`
