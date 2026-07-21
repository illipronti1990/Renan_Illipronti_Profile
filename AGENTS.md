# AGENTS.md

## Instruções específicas do Cursor Cloud

Este repositório é um **README de perfil do GitHub** (página pessoal "Sobre mim").
Na branch `main` ele contém um único arquivo versionado, `README.md`. Não há:

- Código de aplicação, backend ou frontend
- Gerenciador de pacotes, lockfile ou manifesto de dependências
- Build, lint, teste ou serviço para executar

Por causa disso, **não há nada para instalar** e o script de atualização é um no-op.

### Pré-visualizando o README (a única atividade de "desenvolvimento" relevante)

`README.md` é Markdown que o GitHub renderiza automaticamente na página de perfil.
Ele usa serviços externos de imagem (shields.io, capsule-render, github-readme-stats,
streak-stats, komarev) referenciados apenas como URLs de imagem.

Para pré-visualizar localmente da forma como o GitHub renderiza:

1. Renderize para HTML no formato do GitHub: `gh api --method POST /markdown/raw -H "Content-Type: text/plain" --input README.md > body.html`
2. Envolva `body.html` em uma página HTML mínima e sirva-a, por exemplo `python3 -m http.server 8899`
3. Abra `http://localhost:8899/` em um navegador.

Observações:
- Os widgets `github-readme-stats` e `streak-stats` consultam a API do GitHub sem
  autenticação e são limitados por taxa (rate limit) a partir de IPs arbitrários,
  então podem exibir um placeholder "Failed to retrieve contributions" localmente.
  Isso é uma limitação do serviço externo, não um problema do repositório — eles
  renderizam normalmente no perfil real.
- Não adicione ferramentas de build ou dependências a este repositório apenas para
  pré-visualizá-lo.
