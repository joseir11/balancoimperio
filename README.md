# BALANÇO IMPÉRIO

Aplicativo mobile-first de conferência de inventário, hospedado no GitHub Pages e integrado ao Supabase.

## Estrutura
- `index.html`: login, interface das equipes e painel administrativo.
- `ESTOQUE.xlsx`: planilha inicial de materiais.
- `supabase/migrations/`: histórico da estrutura do banco.
- `supabase/functions/manage-team/`: gestão segura de contas de equipes.

## Supabase
Project ref: `znivpoqcvxkgscqoszfj`

## Primeiro administrador
1. No Supabase, abra Authentication > Users e crie o usuário administrador (e-mail e senha).
2. O trigger cria um perfil padrão.
3. No SQL Editor, promova o usuário:
```sql
update public.profiles p set role='admin', team_id=null
from auth.users u where p.id=u.id and u.email='SEU-EMAIL';
``
4. Entre no aplicativo e cadastre as equipes no painel.

## Equipes
O sistema cria contas individuais pelo Supabase Auth. O login gerado será `nome-da-equipe@teams.balancoimperio.app`. A senha é definida pelo administrador.

## Importação
Planilha com cabeçalhos `CODIGO`, `DESCRICAO` e `UN`. A importação inclui e atualiza materiais sem apagar contagens existentes.

## Segurança
RLS protege todas as tabelas públicas. Equipes acessam apenas as próprias quantidades; a tabela de status compartilha apenas AGUARDANDO, OK ou DIVERGENTE, sem expor quantidades de outras equipes.
