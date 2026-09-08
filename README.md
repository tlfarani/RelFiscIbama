# ⚖️ FiscFlow — IBAMA: Gestão de Esteira, Atribuição Automática e Automação de Relatórios

Aplicativo web desenvolvido em **Streamlit** para acompanhamento macro de processos, distribuição automática balanceada de carga de trabalho, gestão de atribuições e geração em lote de **Relatórios de Fiscalização** em formato Word (`.docx`), integrado de forma bidirecional ao **SharePoint via Power Automate**.

---

## 🔑 1. Configuração dos Segredos (`secrets.toml`)

O aplicativo utiliza chaves de configuração no `secrets.toml` para autenticação com a API do Power Automate, navegação para a planilha e validação da chave de acesso do sistema.

### 🛡️ Estrutura nos Secrets do Streamlit Cloud (*Settings > Secrets*):

```toml
# Chave de segurança (PIN) para validação de logins de coordenadores/fiscais
senha_acesso = "SUA_SENHA_DE_ACESSO_AQUI"

[sharepoint]
# URL de API do Power Automate (Gatilho HTTP POST unificado)
url_planilha = "https://SUA_URL_DO_POWER_AUTOMATE_AQUI"

# URL Web direta da planilha no SharePoint (para o botão de atalho na barra lateral)
url_visualizacao = "https://ibamagov.sharepoint.com/:x:/r/teams/..."

```

> ⚠️ **Segurança:** Nunca comite o arquivo `.streamlit/secrets.toml` no GitHub. Certifique-se de mantê-lo listado no seu arquivo `.gitignore`.

---

## 🔄 2. Arquitetura do Sistema e Fluxo do Power Automate

O aplicativo comunica-se com o SharePoint através de **um único fluxo unificado HTTP POST no Power Automate** (`FT_Plataformar_Ler_Atualizar_Dados`), chaveado pelo parâmetro `"acao"` no corpo da requisição JSON.

```
                    ┌──► acao == 'LER' ───────► Lista tabelas [Processos_FT] e [Equipe]
                    │                           └─► Devolve JSON { "processos": [...], "equipe": [...] }
[ Streamlit App ] ──┼
                    │                           ┌─► Atualiza linhas selecionadas no Excel
                    └──► acao == 'ATUALIZAR' ──► Loop "Aplicar a cada" no Power Automate
                                                └─► Grava Servidor Laudo / Fiscal / Situação

```

### ⚡ Como funciona o Fluxo do Power Automate:

1. **Gatilho Inicial (*manual*):** Recebe requisições HTTP POST do Streamlit contendo o JSON da ação.
2. **Nó de Condição (`Condição`):**
* **Caminho FALSO (`acao: "LER"`):**
* Executa a ação *Listar linhas presentes em uma tabela* para a tabela **`Processos_FT`**.
* Executa a ação *Listar linhas presentes em uma tabela* para a tabela **`Equipe`**.
* Retorna a resposta HTTP **200 OK** com o corpo montado dinamicamente:
```json
{
  "processos": [body/value de Processos_FT],
  "equipe": [body/value de Equipe]
}

```




* **Caminho VERDADEIRO (`acao: "ATUALIZAR"`):**
* Recebe as listas `processos_sei`, `ids` e os novos valores (`novo_servidor_laudo`, `novo_fiscal`, `nova_situacao`).
* Executa um loop **Aplicar a cada** para atualizar os registros correspondentes na planilha do SharePoint.
* Retorna a resposta HTTP **200 OK** com o corpo `{"status": "sucesso"}`.





---

## 🔒 3. Autenticação e Controle de Acesso por Perfil (RBAC)

O sistema possui uma camada híbrida e estrita de segurança e controle de permissões:

* **Captura SSO (Streamlit Cloud):** Detecta nativamente o e-mail verificado do usuário logado via `st.user.email`.
* **Validação por Chave de Acesso (PIN/Senha):** Quando o e-mail não é fornecido automaticamente pelo SSO, exige a digitação do e-mail e da chave de acesso cadastrada em `senha_acesso`.
* **Múltiplos E-mails por Servidor:** A coluna `E_Mail` na aba `Equipe` suporta múltiplos e-mails por pessoa (ex: `tiago.farani@ibama.gov.br; tlfarani@hotmail.com`).
* **Níveis de Acesso por Perfil (`Perfil`):**
* `Coordenação` / `Admin`: Acesso completo a todos os módulos, incluindo a tela restrita **👑 Coordenação** (Dashboard macro, Planilha Geral e Atribuição Automática).
* `Análise Técnica`: Acesso aos módulos de Análise Técnica e Fiscalização.
* `Fiscalização`: Acesso ao módulo de Fiscalização e geração de minutas.
* `Não Cadastrado`: Bloqueio preventivo dos módulos restritos.



---

## 📊 4. Mapeamento das Bases de Dados (SharePoint)

### 📋 Tabela 1: `Processos_FT` (Planilha Principal)

| Coluna Original | Nome Interno | Descrição / Uso |
| --- | --- | --- |
| `ID` / `Num_Doc` | `num_doc` | Identificador único do processo |
| `PROCESSO` | `processo_sei` | Número do processo no SEI |
| `SIEMA` | `siema` | Código SIEMA |
| `SITUACAO` | `situacao` | Situação no fluxo (`Fazer Laudo`, `Autuar`, `Auto Lavrado`, etc.) |
| `LAUDO_SEI` | `laudo_sei` | Número do Laudo Técnico SEI |
| `SERVIDOR_LAUDO` | `servidor_laudo` | Analista técnico responsável pela instrução do laudo |
| `Fiscal` | `fiscal` | Fiscal responsável pela autuação |
| `DATA_ACIDENTE` | `data_acid` | Data da ocorrência do acidente |
| `PRODUTO` | `produto` | Sustância ou produto envolvido |
| `CLASS_OL` | `class_ol` | Classificação (`Oleoso` ou `Não Oleoso`) |
| `CLASS_RISCO` | `class_risco` | Classe de risco da empresa (`A`, `B`, `C`, `D`, `E`) |
| `VOL` | `vol_char` | Volume derramado em m³ (formatação decimal brasileira) |
| `Lat` / `Lon` | `lat` / `lon` | Coordenadas do acidente |
| `MULTA_PREVISTA` | `multa_prevista` | Valor da multa estimada |
| `Grandeza` | `grandeza` | Potencial, Reduzida, Fraca, Moderada, Grave |
| `Nivel` / `Nivel_Pontos` | `nivel` / `nivel_pontos` | Nível de gravidade para enquadramento da multa |

### 👥 Tabela 2: `Equipe` (Gestão de Integrantes)

| Coluna Original | Descrição / Uso |
| --- | --- |
| `Nome` | Nome completo do servidor |
| `E_Mail` | E-mail funcional ou e-mail de acesso (aceita múltiplos e-mails separados por `;`) |
| `Perfil` | Perfil de atuação (`Coordenação`, `Análise Técnica`, `Fiscalização`) |

---

## 📱 5. Módulos e Funcionalidades da Interface

### 👑 Módulo 1: Coordenação (Exclusivo para Coordenadores)

* **📊 Dashboard de Indicadores:**
* Cards de métricas gerais (Total de processos, Pendentes de Laudo, Pendentes de Auto, Proc. AI Gerados).
* Gráfico interativo de **Carga por Servidor de Laudo** (destacando a carga ativa pendente em amarelo).
* Gráfico de **Carga por Fiscal Responsável**.
* Gráfico de **Distribuição por Bacia Sedimentar**.


* **📋 Planilha Geral & Atribuições em Lote:**
* Painel de filtros superiores por Bacia, Servidor, Fiscal e Situação.
* Painel de alteração em lote com gravação direta no SharePoint via Power Automate.
* Seleção unificada por checkboxes.


* **🎲 Atribuição Automática Balanceada:**
* Algoritmo inteligente que calcula a **carga ativa real** de trabalho de cada integrante da equipe:
* *Para Laudos:* Considera apenas processos em `"Fazer Laudo"` e sem número de `laudo_sei`.
* *Para Autos:* Considera apenas processos em `"Autuar"` e sem número de `auto`.


* Filtro de distribuição por Bacia Sedimentar ou para todas as bacias.
* Distribuição imparcial pela menor carga ativa com sorteio aleatório em caso de empate (*tie-break*).
* Exibição de prévia da distribuição antes da gravação no SharePoint.



### 🔬 Módulo 2: Análise Técnica

* Acompanhamento dos processos na esteira de instrução técnica.
* Filtros por situação, servidor de laudo e pendências.
* Métricas e tabela consolidada de processos para elaboração de laudos.

### ⚖️ Módulo 3: Fiscalização

* Gestão da fila de autuação de infrações.
* Filtros por situação, fiscal responsável e exibição de minutas.
* **Geração de Minutas em Word (`.docx`):** Preenchimento automático de modelos Word com tags personalizadas e destaque em amarelo para campos que exigem edição manual.
* **Download Unificado (`.zip`):** Empacotamento em lote de múltiplos relatórios preenchidos em um único arquivo compactado.
* **Prévia do Texto de Infração:** Visualização em caixa formatada da descrição da infração antes de baixar o documento.

---

## 📦 6. Dependências e Tema Visual

### 📦 Dependências (`requirements.txt`)

```text
streamlit
pandas
python-docx
requests
plotly

```

### 🎨 Tema Visual Institucional Verde Musgo (`.streamlit/config.toml`)

```toml
[theme]
primaryColor = "#4E5D30"
backgroundColor = "#F8F9F9"
secondaryBackgroundColor = "#E9EDDE"
textColor = "#000000"
base = "light"

```

---

❓ Perguntas Frequentes (FAQ) & Guia de Solução de Problemas
🔐 1. Autenticação, Perfis de Acesso e Governança
Como um novo servidor solicita acesso ao FiscFlow?
Na tela inicial do aplicativo, o servidor acessa a aba 📝 Criar Conta / Solicitar Acesso.

Informa seu nome completo, e-mail funcional (@ibama.gov.br), define uma senha de acesso e seleciona as atribuições desejadas (Análise Técnica, Fiscalização e/ou Coordenação).

Ao submeter o formulário, o aplicativo envia uma requisição CADASTRAR_USUARIO para o Power Automate, que grava o registro com status PENDENTE e dispara um cartão de aprovação interativo diretamente para a caixa de entrada da Coordenação.

Como a Coordenação autoriza uma nova solicitação de acesso?
A Coordenação recebe uma notificação interativa no Outlook (Actionable Message / Cartão do Power Automate Approvals).

No próprio corpo do e-mail, é possível visualizar os dados do solicitante e clicar em [Aprovar] ou [Rejeitar].

O Power Automate processa a decisão e atualiza o campo Status para APROVADO ou REJEITADO na planilha restrita FiscFlow_Equipe.xlsx.

As senhas dos usuários ficam visíveis para quem acessa o SharePoint?
Não. Antes do envio, a senha é processada pelo algoritmo PBKDF2-HMAC-SHA256 com derivação de 100.000 iterações e geração de um Salt criptográfico individual aleatório de 16 bytes. O valor gravado no Excel segue o formato salt$hash (exemplo: a3f89...$e7b21...), sendo irreversível.

Por que a tabela de equipe foi isolada em outro arquivo (FiscFlow_Equipe.xlsx)?
A funcionalidade nativa do Excel de proteger planilhas por senha bloqueia os conectores do Power Automate. Separar a equipe em uma pasta restrita (/Gestao_Acesso/) garante proteção por permissões de rede corporativa sem interferir na esteira de automação.

⚙️ 2. Arquitetura do Fluxo Power Automate
Para que serve o menu "🛠️ Diagnóstico da Conexão" na barra lateral?
É um recurso exclusivo para usuários com perfil de Coordenação. Ele exibe a quantidade exata de linhas brutas recebidas de cada planilha e lista todas as colunas detectadas. Caso um cabeçalho seja renomeado acidentalmente no Excel (por exemplo, de PROCESSO para Processo SEI), a divergência é identificada na hora.

🛠️ 3. Resolução de Erros Comuns (Troubleshooting)
❌ Erro: 400 Client Error: Bad Request for url: ...
Causa: O gatilho HTTP do Power Automate possui Validação de Esquema (Schema Validation) ativada ou campos marcados como "required": [...] no JSON. Quando o Streamlit envia {"acao": "LER"}, a requisição é barrada no gatilho.

Solução: No primeiro bloco do fluxo, clique em ... ➔ Configurações, desative a Validação de Esquema e apague qualquer trecho "required": [...] do esquema JSON.

❌ Erro: O editor do Power Automate fica carregando em loop infinito (círculos girando)
Causa: Instabilidade do Novo Designer (Modern Designer) da Microsoft ao processar condições aninhadas em cascata.

Solução: Clique no menu ... no canto superior direito do Power Automate e selecione "Alternar para o designer clássico".

❌ Erro: Expecting value: line 1 column 1 (char 0) no Streamlit
Causa: O Power Automate retornou HTTP 200, mas o corpo da resposta (Body) estava completamente vazio.

Solução: No bloco Resposta de leitura, defina o cabeçalho Content-Type: application/json e insira o JSON mapeado com os blocos dinâmicos value:

JSON
{
  "processos": @{body('Listar_linhas_presentes_em_uma_tabela_(Processos_FT)')?['value']},
  "equipe": @{body('Listar_linhas_presentes_em_uma_tabela_(Equipe)')?['value']}
}
❌ Erro: StreamlitAPIException no st.data_editor
Causa: A tabela de processos ficou vazia e o componente tentou instanciar uma coluna booleana com [marcar_todos] * 0, gerando uma lista vazia de tipo object.

Solução: Forçar o tipo da série como booleano:

Python
"Selecionar": pd.Series([marcar_todos] * len(df_f), dtype=bool)
e encapsular a renderização em if df_f.empty: st.warning(...) else: st.data_editor(...).

❌ A tabela de processos exibe nomes e e-mails em vez de processos SEI
Causa: No bloco Resposta do Power Automate, o campo "processos":  recebeu acidentalmente o value da ação da tabela Equipe.

Solução: No campo "processos": , selecione estritamente o value da ação Listar linhas presentes em uma tabela (Processos FT).

❌ Ao clicar em "Aprovar" no e-mail, o usuário é marcado como REJEITADO
Causa: Na Condição 3, o campo de comparação foi preenchido com aspas literais: 'Approve'.

Solução: Digite a palavra limpa, sem aspas: Approve.

❌ Fiz alterações no SharePoint, mas o aplicativo continua mostrando dados antigos
Causa: Cache local do Streamlit (@st.cache_data(ttl=300)).

Solução: Pressione a tecla C no navegador para limpar o cache e recarregue com Ctrl + F5.
