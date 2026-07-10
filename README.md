# Livro-Razão · Controle Financeiro

App de controle financeiro para 2 empresas (CNPJs), com DRE mensal/anual, Balanço
Patrimonial, Dashboard e Fluxo de Caixa diário/mensal/anual. Login com Supabase
(suporta quantos usuários você quiser, sincroniza entre celular/computador, e
funciona como app instalável no celular — PWA).

Este guia assume **zero conhecimento técnico prévio**. São 4 etapas:

1. Criar o banco de dados (Supabase) — grátis
2. Colocar o código no ar (Vercel) — grátis
3. Ligar seu domínio próprio
4. Instalar no celular

Tempo estimado: 30–40 minutos na primeira vez.

---

## Etapa 1 — Criar o banco de dados no Supabase

O Supabase guarda os dados (lançamentos, balanços, empresas) e cuida do login dos
4 usuários. É gratuito para este porte de uso.

1. Acesse **https://supabase.com** e crie uma conta (dá para usar login do Google).
2. Clique em **New project**. Escolha um nome (ex: `controle-financeiro`), crie uma
   senha do banco (guarde-a num lugar seguro) e escolha a região mais próxima
   (ex: São Paulo/`sa-east-1`). Aguarde ~2 minutos até o projeto ficar pronto.
3. No menu lateral, vá em **SQL Editor** → **New query**. Abra o arquivo
   `supabase.sql` (está junto com este projeto), copie todo o conteúdo, cole no
   editor e clique em **Run**. Isso cria a tabela onde os dados ficam guardados.
4. Vá em **Authentication → Users** → **Add user** → **Create new user**.
   Crie os 4 usuários, um de cada vez. Para cada um:
   - **Email**: o e-mail que a pessoa vai usar para logar (ex: `financeiro@suaempresa.com.br`)
   - **Password**: a senha dela
   - Marque **Auto Confirm User** (assim ela já pode logar direto, sem confirmar e-mail)
   - Em **User Metadata**, cole algo assim (ajuste nome e cargo):
     ```json
     { "nome": "Ana Ribeiro", "cargo": "Financeiro" }
     ```
5. Vá em **Project Settings → API**. Copie dois valores que você vai usar na Etapa 2:
   - **Project URL** (algo como `https://xxxxx.supabase.co`)
   - **anon public key** (uma chave longa)

> Essa chave "anon" é pública por natureza (fica visível no código do site) — a
> segurança real vem das políticas de acesso (RLS) que o `supabase.sql` já configura,
> exigindo login para ler/gravar dados.

---

## Etapa 2 — Colocar o app no ar (Vercel)

1. Crie uma conta em **https://github.com** (se ainda não tiver).
2. Crie um repositório novo (ex: `controle-financeiro`) e suba os arquivos deste
   projeto para ele. Formas mais fáceis:
   - Pelo site do GitHub: **Add file → Upload files**, arraste todos os arquivos
     e pastas deste projeto (menos `node_modules`, que nem existe ainda) e clique
     em **Commit changes**.
   - Ou, se preferir linha de comando: `git init`, `git add .`, `git commit -m "primeira versão"`,
     depois siga as instruções do GitHub para conectar e enviar (`git push`).
3. Acesse **https://vercel.com**, crie conta com login do GitHub, clique em
   **Add New → Project** e selecione o repositório que você acabou de criar.
4. Na tela de configuração do projeto, abra **Environment Variables** e adicione:
   - `VITE_SUPABASE_URL` = a Project URL que você copiou na Etapa 1
   - `VITE_SUPABASE_ANON_KEY` = a anon public key que você copiou na Etapa 1
5. Antes de clicar em Deploy, confira o campo **Project Name** — deixe escrito
   `controle-financeiro` (o projeto já vem configurado com esse nome no
   `package.json`, então a Vercel deve sugerir isso automaticamente). Esse nome
   vira o endereço do site.
6. Clique em **Deploy**. Em 1–2 minutos o Vercel te entrega o link
   **`controle-financeiro.vercel.app`** — o site já está no ar, com HTTPS,
   pronto para acessar de qualquer lugar.

   > Se aparecer aviso de que o nome já está em uso (é um espaço global, então
   > pode acontecer), a Vercel sugere uma variação automática, tipo
   > `controle-financeiro-abc.vercel.app`. Tudo funciona igual, só muda o link.

Teste o login com um dos 4 usuários criados na Etapa 1 antes de seguir.

---

## Etapa 3 — Ligar seu domínio próprio

Se você ainda não comprou um domínio, no Brasil o mais comum é comprar um
`.com.br` em **registro.br** (~R$40/ano) ou um `.com` em serviços como Namecheap
ou no próprio painel da Vercel.

Depois de ter o domínio:

1. No painel da Vercel, abra seu projeto → **Settings → Domains** → digite seu
   domínio (ex: `financeiro.suaempresa.com.br`) → **Add**.
2. A Vercel mostra exatamente quais registros DNS criar (geralmente um registro
   tipo `CNAME` ou `A`). Copie esses valores.
3. Entre no painel do site onde você comprou o domínio (registro.br, GoDaddy etc.),
   procure por **DNS** ou **Zona DNS**, e cadastre os registros que a Vercel indicou.
4. Aguarde a propagação (de alguns minutos até algumas horas). A Vercel emite o
   certificado de segurança (HTTPS) automaticamente assim que detectar o DNS
   correto — você não precisa fazer nada além de esperar.

Pronto: seu app estará acessível pelo seu domínio.

---

## Etapa 4 — Instalar no celular

Como o projeto já vem com PWA configurado, não é preciso loja de aplicativos:

- **Android (Chrome)**: abra o site pelo seu domínio → menu (⋮) → **Adicionar à
  tela inicial** / **Instalar app**.
- **iPhone (Safari)**: abra o site → toque no ícone de compartilhar (□↑) →
  **Adicionar à Tela de Início**.

O ícone aparece na tela do celular como um app normal, abrindo em tela cheia.

---

## Rodar localmente (opcional, para quem quiser mexer no código)

```bash
npm install
cp .env.example .env       # depois edite o .env com suas chaves do Supabase
npm run dev                # abre em http://localhost:5173
```

Para gerar a versão de produção manualmente: `npm run build` (gera a pasta `dist`).

---

## Estrutura do projeto

```
├─ src/
│  ├─ App.jsx            → toda a lógica e telas do app
│  ├─ main.jsx            → ponto de entrada React
│  └─ supabaseClient.js   → conexão com o Supabase
├─ public/icons/          → ícones do app (PWA)
├─ supabase.sql           → script para criar a tabela/permissões no Supabase
├─ vite.config.js         → configuração do Vite + PWA
├─ .env.example           → modelo das variáveis de ambiente
└─ package.json
```

## Gerenciando usuários depois

Para adicionar, remover ou trocar senha de usuários: Supabase → **Authentication
→ Users**. Não é preciso mexer no código nem republicar o site.

## Dúvidas comuns

- **"Erro ao carregar / gravar dados"**: confira se `VITE_SUPABASE_URL` e
  `VITE_SUPABASE_ANON_KEY` estão certos no painel da Vercel (Settings →
  Environment Variables) e se rodou o `supabase.sql`.
- **Login não funciona**: confirme que marcou "Auto Confirm User" ao criar o
  usuário no Supabase.
- **Quero adicionar uma 3ª empresa**: o modelo atual foi feito para 2 CNPJs (A e
  B) + consolidado. Para adicionar mais, é preciso mexer no código (`App.jsx`) —
  posso te ajudar com isso se precisar.
