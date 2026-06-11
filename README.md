# Controle de Produtividade Banho — Pintura Eletrostática

Sistema de monitoramento em tempo real do tratamento de peças, do **pré-banho**
(preparação dos cestos) ao **banho**, com grade visual de cestos, leitura de
código de barras da OP, dashboards, painel público para a gerência e
exportação de relatórios em Excel. Feito para rodar 24h na nuvem (Railway) com
banco PostgreSQL, e otimizado para tablets (Samsung Galaxy Tab A9).

---

## Fluxo (sem etapa de aprovação do líder)

1. **Operador de preparação** abre a tela e vê a **grade de cestos 1 a 19**.
   Cestos livres aparecem coloridos; cestos em uso aparecem com **cadeado**.
2. Toca num cesto livre → **Iniciar preparação** (cronômetro começa).
3. Ao terminar de encher o cesto, toca no cesto de novo, **bipa o código de
   barras da OP** (preenche código/material, texto breve e quantidade a partir
   da lista mestra — tudo editável), escolhe processo, tipo e observação, e
   conclui. O cesto **vai direto para a fila do banho**.
4. **Operador de banho** vê a fila, inicia o banho (novo cronômetro) e finaliza
   quando o cesto sai. Aí o cesto **volta a ficar livre** na grade.
5. Qualquer card pode ser **editado** depois de criado (tocar no cesto ocupado),
   para corrigir informação errada.

Todas as telas atualizam sozinhas a cada poucos segundos.

---

## Telas e acessos

| Caminho | Quem | O que faz |
|---------|------|-----------|
| `/preparacao` | operadores de prep (e banho) | grade de cestos, abrir/preencher/editar |
| `/banho` | operador de banho | fila e cronômetro de banho |
| `/dashboard` | admin | KPIs, gráficos, histórico, **export Excel** |
| `/painel` | **público (sem senha)** | só leitura, para a gerência, com filtro por data |
| `/admin/mestre` | admin | importar a lista mestra do SAP |
| `/admin/usuarios` | admin | criar/remover usuários, trocar senha |

### Usuários criados automaticamente (troque as senhas depois)

| Login | Senha | Perfil |
|-------|-------|--------|
| `admin` | `admin123` | admin |
| `banho` | `banho123` | operador de banho |
| `op1`…`op6` | `op1234` | operador de preparação |

São 3 perfis: **prep** (pré-banho, pode ter vários), **banho** (normalmente 1,
mas pode criar mais) e **admin**. O admin cria quantos usuários quiser na tela
de Usuários.

---

## Lista mestra (relatório do SAP)

Na tela "Lista Mestra", o admin importa o relatório do SAP em **.xlsx** ou
**.csv/.txt** (separado por TAB). O sistema usa estas colunas:

```
Ordem | Nº do item | Material | Texto breve material | Quantidade total | ...
```

- **Ordem** = a OP (o que é bipado)
- **Material** = código
- **Texto breve material** = descrição
- **Quantidade total** = quantidade

O histórico guarda exatamente: **OP, Código, Texto breve do material e
Quantidade total**. Veja `exemplo_lista_mestra_sap.txt`.

---

## Relatórios e backup

No dashboard do admin há dois botões de exportação Excel, para você salvar
backups quando quiser:

- **Excel pré-banho** — dados e tempos da preparação.
- **Excel banho** — dados e tempos do banho.

Cada arquivo vem com data/hora no nome. Os dados ficam no PostgreSQL, então
**nunca se perdem** em reinício/deploy.

---

## Deploy — GitHub + Railway

### 1) GitHub
```bash
cd v2
git init
git add .
git commit -m "Controle de Produtividade Banho"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/SEU_REPO.git
git push -u origin main
```

### 2) Railway
1. railway.app → **New Project → Deploy from GitHub repo** → escolha o repo.
2. **New → Database → Add PostgreSQL** (cria `DATABASE_URL` automaticamente).
3. No serviço web, em **Variables**, adicione `SECRET_KEY` = uma frase longa e
   aleatória.
4. O Railway sobe com gunicorn (via `Procfile`). Em **Settings → Networking →
   Generate Domain** para obter a URL pública.

> Sem o PostgreSQL conectado, o app usa um SQLite local que **se perde** a cada
> reinício. Para uso 24/7, adicionar o PostgreSQL (passo 2) é obrigatório.

O painel da gerência fica em `SUA_URL/painel` (link aberto, só leitura).

---

## Rodar localmente

```bash
pip install -r requirements.txt
python app.py    # http://localhost:5000
```

Sem `DATABASE_URL`, usa SQLite local (`dados_local.db`) só para teste.

---

## Estrutura

```
v2/
├── app.py
├── requirements.txt   Procfile   railway.json   runtime.txt   .gitignore
├── exemplo_lista_mestra_sap.txt
├── static/
│   ├── style.css
│   └── dash.js
└── templates/
    ├── base.html  login.html  prep.html  banho.html
    ├── dashboard.html  painel.html  usuarios.html  mestre.html
```
