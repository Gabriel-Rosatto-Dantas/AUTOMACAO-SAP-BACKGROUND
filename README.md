# Automação Extrator On-Time SAP

Este projeto consiste em um script de automação (RPA) desenvolvido em Python para extrair dados de performance de entregas (On-Time) diretamente do ambiente SAP S/4HANA. A automação simula a interação humana com a interface do SAP GUI para executar transações, extrair relatórios e consolidar informações de forma eficiente e livre de erros.

## 📋 Funcionalidades Principais

  * **Login Automático:** Acessa o SAP de forma segura utilizando credenciais armazenadas em variáveis de ambiente.
  * **Extração de Múltiplas Fontes:** Coleta dados de transações customizadas (`ZPMMT_287`) e de tabelas padrão do SAP via `SE16N` (`EBAN`, `EKET`, `LIPS`, `VBFA`, `J_1BNFLIN`, `J_1BNFDOC`, `MARA`).
  * **Processamento de Dados:** Utiliza a biblioteca `pandas` para manipular, limpar e consolidar os dados extraídos.
  * **Fluxo de Trabalho Encadeado:** Orquestra um processo complexo onde a saída de uma extração é utilizada como filtro para a etapa seguinte.
  * **Geração de Arquivos:** Exporta os dados brutos e processados para arquivos nos formatos `.xlsx` e `.txt`, organizados em um diretório local.
  * **Gerenciamento de Janelas:** Garante que instâncias anteriores do SAP ou de planilhas Excel sejam fechadas antes de iniciar o processo para evitar conflitos.

## 🛠️ Pré-requisitos

Para executar este projeto, os seguintes componentes são necessários:

  * **Software:**
      * Python 3.8 ou superior.
      * SAP GUI for Windows (versão 7.60 ou superior recomendada).
  * **Bibliotecas Python:**
      * `pywin32`
      * `pandas`
      * `pywinauto`
      * `python-dotenv`
      * (Todas as dependências estão listadas no arquivo `requirements.txt`)

## ⚙️ Configuração e Instalação

Siga estes passos para configurar o ambiente de desenvolvimento.

**1. Clone o Repositório**

```bash
git clone <https://github.com/Gabriel-Rosatto-Dantas/AUTOMACAO-SAP-BACKGROUND>
cd <AUTOMACAO-SAP-BACKGROUND>
```

**2. Crie um Ambiente Virtual (Recomendado)**

```bash
python -m venv venv
# No Windows
venv\Scripts\activate
# No macOS/Linux
source venv/bin/activate
```

**3. Instale as Dependências**

```bash
pip install -r requirements.txt
```

**4. Habilite o SAP GUI Scripting**
Para que a automação funcione, o scripting precisa estar habilitado no seu cliente SAP:

  * Abra o **SAP Logon**.
  * Clique no menu no canto superior esquerdo e vá em **Opções**.
  * Navegue até **Acessibilidade e Scripts \> Script**.
  * Marque a opção **Ativar scripting**.
  * Desmarque as opções **Notificar quando um script for anexado** e **Notificar quando um script abrir uma conexão**.

**5. Configure as Variáveis de Ambiente**
Crie um arquivo chamado `.env` na raiz do projeto e preencha com suas credenciais SAP. Este arquivo é ignorado pelo Git para proteger suas informações.

```env
SAP_USER="SEU_USUARIO_SAP"
SAP_PASSWORD="SUA_SENHA_SAP"
```

## ▶️ Executando o Script

Antes de executar, verifique o script e ajuste os caminhos de diretório conforme necessário.

**1. Ajuste dos Caminhos (Passo Crítico)**
No código, os caminhos para salvar os arquivos estão fixos (hardcoded), por exemplo: `C:\Users\3976339\Desktop\ONTIME`. **Você deve substituir estes caminhos** pelo diretório que deseja usar em sua máquina.

**2. Execute o Script**
Com o ambiente virtual ativado e as configurações prontas, execute o script principal:

```bash
python extrator_ontime.py
```

O progresso da execução será exibido no terminal.

## 🔄 Estrutura do Processo

O script executa um fluxo de trabalho lógico para coletar e relacionar os dados:

1.  **Preparação:** Fecha qualquer sessão ativa do SAP ou do Excel.
2.  **Login:** Acessa o SAP S/4HANA.
3.  **Transação ZPMMT\_287:** Extrai a base inicial de requisições de compras e materiais.
4.  **Tabelas EBAN e EKET:** Usa os dados da extração anterior para buscar detalhes dos pedidos.
5.  **Tabela LIPS:** Consolida os pedidos para encontrar as remessas correspondentes.
6.  **Tabela VBFA:** Rastreia o fluxo de documentos a partir das remessas para identificar os movimentos de mercadoria.
7.  **Tabelas J\_1BNFLIN e J\_1BNFDOC:** Busca dados das notas fiscais associadas aos movimentos de mercadoria.
8.  **Tabela MARA:** Extrai informações mestras dos materiais envolvidos.
9.  **Finalização:** Salva o último conjunto de dados e encerra a automação.

## 📂 Estrutura de Pastas e Arquivos Gerados

O script cria e utiliza uma série de arquivos intermediários e finais. A estrutura de saída esperada no diretório `ONTIME` (ou o nome que você definir) é a seguinte:

```
ONTIME/
├── EBAN.xlsx
├── EKET.XLSX
├── J_1BNFDOC.XLSX
├── J_1BNFLIN.xlsx
├── JLIN.txt
├── LIPS.XLSX
├── MARA.txt
├── MARA.XLSX
├── PEDIDOS_CONSOLIDADO.txt
├── REMESSA.txt
├── VBFA.XLSX
├── VBFA_CONSOLIDADO.txt
├── ZPMMT.xlsx
└── ZPMMT_REQ.txt
```

## ⚠️ Observações Importantes

  * **Dependência de Interface:** A automação depende da estrutura da interface do SAP GUI. Mudanças na interface, como IDs de elementos ou layouts de tela, podem quebrar o script.
  * **Caminhos Fixos:** Como mencionado, todos os caminhos de salvamento de arquivos são fixos. Para portabilidade, considere refatorar o código para usar caminhos relativos ou configuráveis.
  * **Performance:** Extrações de grandes volumes de dados podem ser lentas. O script foi otimizado com filtros, mas a performance depende da resposta do sistema SAP.
