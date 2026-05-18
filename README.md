# CEMFS — Cronograma A3

Geração automática do cronograma A3 do Centro Espírita Manoel Felipe Santiago.

Toda **segunda às 6h da manhã**, um script lê os dados das atividades cadastradas no Notion e atualiza o arquivo `CRONOGRAMA_A3_CEMFS.xlsx` deste repositório. Basta abrir o link do arquivo e imprimir.

## Onde fica o cronograma pronto

O arquivo sempre atualizado fica em:

```
https://github.com/SEU_USUARIO/cemfs-cronograma/raw/main/CRONOGRAMA_A3_CEMFS.xlsx
```

(Troque `SEU_USUARIO` pelo seu usuário do GitHub. Esse link é fixo — pode imprimir e colar na parede da secretaria pra qualquer pessoa baixar.)

## Como atualizar fora de hora

Se alguém mudou algo no Notion e quer ver o cronograma novo imediatamente sem esperar até segunda:

1. Vá em **Actions** (aba no topo do GitHub)
2. Clique em **Gerar Cronograma CEMFS** (menu da esquerda)
3. Clique em **Run workflow** (botão à direita) → **Run workflow**
4. Em ~30 segundos o arquivo está atualizado no link acima

---

## Setup inicial (uma vez só)

### 1. Criar o repositório no GitHub

- Acesse [github.com/new](https://github.com/new)
- **Repository name**: `cemfs-cronograma` (ou outro nome)
- **Public** ou **Private**: tanto faz. *Public* deixa o link do xlsx acessível sem login (mais fácil pra voluntários). *Private* exige login pra baixar mas é mais discreto.
- **NÃO** marque "Add a README" — vamos subir o nosso
- Clique em **Create repository**

### 2. Subir os arquivos deste pacote

Existe duas formas. A mais simples:

**Pelo navegador**:
- Na página do repo recém-criado, clique no link **uploading an existing file**
- Arraste TODOS os arquivos desta pasta (incluindo a pasta oculta `.github`)
- Escreva uma mensagem de commit qualquer (ex: "setup inicial")
- Clique em **Commit changes**

> ⚠️ Pelo navegador o GitHub não aceita pastas vazias. A pasta `.github/workflows` precisa do arquivo `cronograma.yml` dentro pra ser criada. Se der trabalho arrastar a pasta inteira, faça assim:
> 1. Crie um arquivo novo chamado `.github/workflows/cronograma.yml` (use o botão "Add file > Create new file" e digite o caminho com as barras)
> 2. Cole o conteúdo do `cronograma.yml` deste pacote

### 3. Adicionar o token do Notion como secret

1. No repo, vá em **Settings** (aba do topo)
2. Menu da esquerda: **Secrets and variables** → **Actions**
3. Clique em **New repository secret**
4. **Name**: `NOTION_TOKEN`
5. **Secret**: cole o token do Notion (aquele que começa com `ntn_...`)
6. **Add secret**

### 4. Garantir que a integration do Notion tem acesso à página

Esse passo é o que mais quebra. Sem ele, o script roda mas não consegue ler os dados do Notion.

1. Abra a página **CEMFS — Cronograma** no Notion
2. No canto superior direito, clique nos **três pontinhos (•••)**
3. Role até **Connections** (Conexões)
4. Clique em **Add connections**
5. Busque o nome da connection que você criou (ex: "CEMFS Cronograma")
6. Selecione → confirme

Isso libera acesso à página inteira, incluindo todos os bancos dentro (Atividades, Horários, Salas).

### 5. Testar

Na aba **Actions** do repo, clique em **Gerar Cronograma CEMFS** → **Run workflow** → **Run workflow**. Em ~30 segundos:

- Se der ✅ (verde) — sucesso, o arquivo `CRONOGRAMA_A3_CEMFS.xlsx` foi atualizado no repo
- Se der ❌ (vermelho) — clique no nome do run e olha os logs do step "Rodar gerador" pra ver o erro

Depois disso, dura anos sem precisar mexer.

---

## Customizações comuns

### Mudar o horário em que roda automaticamente

Abra `.github/workflows/cronograma.yml` e edite a linha:

```yaml
    - cron: '0 9 * * 1'
```

Formato: `minuto hora * * dia-da-semana` (em UTC; Brasília = UTC menos 3 horas).

- `'0 9 * * 1'` = toda segunda 9h UTC = 6h Brasília
- `'0 12 * * *'` = todo dia ao meio-dia UTC = 9h Brasília
- `'30 14 * * 5'` = toda sexta 14:30 UTC = 11:30 Brasília

### Mudar o nome curto exibido no xlsx

Algum programa aparece com nome ruim no cronograma impresso? Abra `gerador.py`, ache o dicionário `NOMES_CURTOS` (linhas ~80), edite ou adicione:

```python
NOMES_CURTOS = {
    'Consultas Odontológicas': 'CONSULTAS ODONT.',
    'Nome novo do programa no Notion': 'NOME CURTO QUE QUERO NO XLSX',
    # ...
}
```

Faça commit e o próximo run usa o novo nome.

### Adicionar uma sala nova

1. Crie a sala no banco Salas do Notion
2. Pegue o **ID da página da sala** (na URL, é a parte longa sem hífens)
3. Em `gerador.py`, no dicionário `SALA_TO_COL`, adicione a entrada com o número da coluna nova:

```python
'id_da_sala_nova_aqui': 18,  # → coluna R
```

4. Pode ser necessário também:
   - Ajustar o `merge_cells('K2:Q6')` pra cobrir mais colunas
   - Ajustar header (linha 9) com o nome da sala

---

## Quando algo dá errado

### O run vermelho diz "ERRO API Notion: 401"

Token errado ou expirado. Confere o secret `NOTION_TOKEN` em **Settings → Secrets**.

### O run vermelho diz "ERRO API Notion: 404"

A integration não tem acesso à página. Volta no passo 4 do Setup (compartilhar página com a connection).

### O xlsx gerado está vazio

Provavelmente os IDs dos bancos mudaram (raríssimo, mas se você recriou as databases pode acontecer). Olha o log do run pra ver quantas atividades/horários foram lidos — se for 0, conferir `DB_ATIVIDADES` e `DB_HORARIOS` em `gerador.py`.

### Apareceu uma atividade que não existe mais

Confere o status dela no Notion. Status diferente de "Ativo" são pulados automaticamente.

---

## Estrutura do repositório

```
cemfs-cronograma/
├── README.md                       ← este arquivo
├── gerador.py                      ← script Python que faz a mágica
├── requirements.txt                ← lista de dependências Python
├── QUADRO_GERAL_TEMPLATE.xlsx      ← template original que serve de base
├── CRONOGRAMA_A3_CEMFS.xlsx        ← arquivo final, atualizado automaticamente
└── .github/
    └── workflows/
        └── cronograma.yml          ← config do agendamento + botão manual
```

Não precisa mexer em nenhum arquivo no dia-a-dia. As mudanças de horários, salas, responsáveis, etc, são feitas direto no Notion — o script puxa dali toda vez que roda.
