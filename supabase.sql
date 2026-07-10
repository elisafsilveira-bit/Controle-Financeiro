-- =========================================================
-- Execute este script no Supabase: SQL Editor > New query > Run
-- =========================================================

-- Tabela única de armazenamento (guarda lançamentos, balanços e empresas)
create table if not exists app_storage (
  chave text primary key,
  valor jsonb not null,
  atualizado_em timestamptz default now()
);

-- Ativa Row Level Security
alter table app_storage enable row level security;

-- Qualquer usuário autenticado (logado) pode ler e gravar
create policy "Usuários autenticados podem ler"
  on app_storage for select
  to authenticated
  using (true);

create policy "Usuários autenticados podem gravar"
  on app_storage for insert
  to authenticated
  with check (true);

create policy "Usuários autenticados podem atualizar"
  on app_storage for update
  to authenticated
  using (true);

-- Registros iniciais (o app cria sozinho se não existirem, mas não custa nada já deixar prontos)
insert into app_storage (chave, valor) values
  ('lancamentos', '[]'::jsonb),
  ('balancos', '{"a":{},"b":{}}'::jsonb),
  ('empresas', '{"a":{"nome":"Empresa A","cnpj":""},"b":{"nome":"Empresa B","cnpj":""}}'::jsonb)
on conflict (chave) do nothing;
