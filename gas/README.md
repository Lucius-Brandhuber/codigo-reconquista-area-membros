# Backend de contas (Rota A) — Área de Membros

Login/registro com e-mail e senha usando **Google Apps Script + Google Sheets**.
A senha é guardada como **sal + hash SHA-256** (nunca em texto puro). Os dados que
a aluna insere (diário etc.) ficam salvos no servidor, por e-mail — **não no cache**.

## Como publicar (uma vez)

1. Acesse https://script.google.com → **Novo projeto**.
2. Apague o conteúdo e cole o `Codigo.gs` deste repositório.
3. (Opcional) Em `COMPRADORES`, cole os e-mails de quem comprou para só esses
   poderem criar conta. Deixe `[]` para liberar qualquer e-mail.
4. Rode a função **`autorizar`** uma vez e aceite as permissões (Drive/Sheets).
5. **Implantar → Nova implantação → App da Web**
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**
6. Copie a **URL `/exec`** e cole na constante `GAS_AUTH` do `index.html`
   (no topo do bloco `AUTH`). Faça commit/push → o Vercel publica sozinho.

A planilha **"Reconquista — Membros DB"** é criada automaticamente no seu Drive
na 1ª execução (aba `usuarios`).

## Endpoints (POST com corpo JSON, `Content-Type: text/plain`)

| action     | envia                                  | retorna                          |
|------------|----------------------------------------|----------------------------------|
| `register` | `nome, email, senha`                   | `{ok, nome, token, dados}`       |
| `login`    | `email, senha`                         | `{ok, nome, token, dados}`       |
| `save`     | `email, token, dados`                  | `{ok}`                           |
| `load`     | `email, token`                         | `{ok, nome, dados}`              |

O `token` é gerado a cada login e exigido em `save`/`load` (sessão simples).

## Observações de segurança

- SHA-256 com sal é simples e ok para um entregável. Para algo mais robusto
  (proteção contra força bruta, recuperação de senha por e-mail), use Supabase/
  Firebase (Rota B).
- Não há "esqueci minha senha" nesta versão — dá para adicionar com `MailApp`.
- O e-mail sozinho não comprova a compra; para travar de verdade, preencha
  `COMPRADORES` (ou integre com a lista da Payt).
