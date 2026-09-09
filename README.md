# Financial Transaction Manager

Sistema para cadastro e gerenciamento de transações de cartão de crédito, desenvolvido em **Visual Basic 6.0** com **SQL Server**.

---

## 📋 Funcionalidades

| Funcionalidade | Descrição |
|----------------|-----------|
| **CRUD Completo** | Criar, editar, excluir e consultar transações |
| **Filtros Avançados** | Por cartão, período (data), faixa de valor, status |
| **Paginação Server-Side** | 50 registros por página (configurável), navegação completa |
| **Regras de Negócio** | Edição bloqueada para status "Approved", exclusão com confirmação |
| **Exportação Excel** | Relatório do último mês completo com formatação profissional |
| **Logs de Erro** | Arquivo diário com detalhes técnicos (`logs/erros_YYYYMMDD.txt`) |
| **Mensagens Amigáveis** | Erros técnicos logados, usuário vê mensagens claras |

---

## 🗄️ Banco de Dados

### Objetos Criados

| Tipo | Nome | Descrição |
|------|------|-----------|
| Tabela | `Transactions` | Dados principais com índices otimizados |
| Stored Procedure | `sp_GetTransactionTotalsByPeriod` | Totais por cartão/status no período |
| Função Escalar | `fn_GetValueCategory` | Categoriza valor (Premium/High/Medium/Low) |
| TVF | `tvf_GetCategorizedTransactionsByPeriod` | Transações com categoria no período |
| View | `vw_ConsolidatedTransactions` | Consolidação com campos derivados |

### Categorização de Valores

| Faixa (USD) | Categoria |
|-------------|-----------|
| > $2.000,00 | Premium |
| $1.000,00 - $2.000,00 | High |
| $500,00 - $1.000,00 | Medium |
| < $500,00 | Low |

---

## 🚀 Como Executar

### Opção 1: Docker Compose (Recomendada — Inclui Auto-Seed)

**Pré-requisitos:**
- Docker Desktop instalado e rodando
- Windows 10/11 ou Linux/macOS

**Subir o banco:**
```bash
docker-compose up -d
```

**O que acontece automaticamente:**
1. SQL Server 2022 Developer inicia na porta **1433**
2. Banco `TransactionManagement` é criado
3. Todos os objetos de schema são criados (tabela, índices, SP, funções, view)
4. **200 transações de exemplo** são carregadas do `scripts-database/06_dados_exemplo_load_test.sql`:
   - 10 cartões de teste (Visa, Mastercard, Amex, Discover, JCB, UnionPay)
   - 6 meses de dados (Abr–Set 2026)
   - Distribuição: ~130 Low, ~14 Medium, ~29 High, ~27 Premium
   - Status: ~145 Approved, ~33 Pending, ~22 Canceled

**Detalhes de conexão:**
```
Server: localhost,1433
Database: TransactionManagement
Username: sa
Password: YourStrong@Password123
```

**Parar e remover:**
```bash
docker-compose down              # Para (preserva dados no volume)
docker-compose down -v           # Para e apaga todos os dados
```

**Ver logs:**
```bash
docker logs sqlserver_db -f
```

---

### Opção 2: SQL Server Local (Windows)

**Pré-requisitos:**
- Windows 10/11 (32-bit ou 64-bit com WoW64)
- SQL Server (LocalDB, Express ou Full) — versão 2012+ (para OFFSET/FETCH)
- Visual Basic 6.0 IDE (para desenvolvimento)
- Microsoft Excel (para exportação)
- OCXs registrados: `MSFLXGRD.OCX`, `MSADO28.tlb`

**1. Banco de Dados (escolha uma):**

**A. Script consolidado (recomendado):**
```sql
-- Execute no SSMS ou sqlcmd:
sqlcmd -S (local) -i database\script_completo.sql
```

**B. Scripts individuais (ordem obrigatória):**
```sql
1. database\01_create_database.sql
2. database\02_sp_total_periodo.sql
3. database\03_fn_categoria_valor.sql
4. database\04_tvf_transacoes_categorizadas.sql
5. database\05_view_consolidada.sql
6. database\06_dados_exemplo_load_test.sql   -- 200 registros para teste de carga
```

**C. Linux/macOS com sqlcmd:**
```bash
chmod +x scripts-scripts-database/init-database.sh
./scripts-database/init-database.sh localhost TransactionManagement sa YourStrong@Password123
```

**2. Configurar Conexão**

Edite `FinancialTtransactionManager.frm` no `Form_Load`:
```vb
Call InicializarConexao( _
    Servidor:="(local)", _         ' Seu SQL Server
    Banco:="TransactionManagement", _
    AutenticacaoWindows:=True)      ' False se usar SQL Auth
```

Para Docker:
```vb
Call InicializarConexao( _
    Servidor:="localhost,1433", _  ' Mapeamento de porta do Docker
    Banco:="TransactionManagement", _
    Usuario:="sa", _
    Senha:="YourStrong@Password123", _
    AutenticacaoWindows:=False)
```

**3. Abrir no VB6**
1. Abra `FinancialTransactionManager.vbp` no VB6 IDE
2. Verifique referências (Projeto > Referências):
   - ✅ Microsoft ActiveX Data Objects 2.8 Library
   - ✅ Microsoft FlexGrid Control 6.0
3. Pressione **F5** para executar

**4. Compilar (Opcional)**
```
Arquivo > Make FinancialTransactionManager.exe
```
O executável será gerado na pasta do projeto.

---

## 📁 Estrutura do Projeto

```
FinancialTransactionManager/
├── FinancialTransactionManager.vbp       # Projeto VB6
├── FinancialTtransactionManager.frm      # Formulário principal (lista/grid)
├── frmTransacao.frm                      # Formulário cadastro/edição (modal)
├── FinancialTransactionManager.vbw       # Workspace
├── docker-compose.yml                    # Orquestração Docker
├── src/
│   ├── modConexao.bas                    # Conexão ADO + helpers
│   ├── modLogErros.bas                   # Log em arquivo texto
│   ├── modUtilitarios.bas                # Validações, formatação
│   └── modExportacaoExcel.bas            # Exportação Excel via COM
├── scripts-database/
│   ├── 01_create_database.sql            # DB + Tabela Transactions + índices
│   ├── 02_sp_total_periodo.sql           # Stored Procedure
│   ├── 03_fn_categoria_valor.sql         # Função escalar (categoria)
│   ├── 04_tvf_transacoes_categorizadas.sql # Função table-valued
│   ├── 05_view_consolidada.sql           # View consolidada
│   ├── 06_dados_exemplo.sql              # ~50 registros originais
│   ├── 06_dados_exemplo_load_test.sql    # 200 registros teste carga (usado pelo Docker)
│   ├── script_completo.sql               # Executa todos acima (seed original)
│   ├── init-database.sql                 # Script mestre init (roda 01-06)
│   └── init-database.sh                  # Script init Linux/macOS
├── logs/                                 # Criado em runtime
│   └── erros_YYYYMMDD.txt
├── Relatorios/                           # Criado em runtime
│   └── Relatorio_Transacoes_YYYYMM.xlsx  # Gerado pela exportação
└── docs/
    ├── proposta.md                       # Requisitos originais
    ├── PLANO_EXECUCAO.md                 # Plano detalhado por fases
    ├── DOCUMENTACAO_TECNICA.md           # Documentação técnica completa
    ├── EVOLUCAO_FUTURA.md                # Roadmap e ideias futuras
    └── README.md                         # Este arquivo
```

---

## 🐳 Detalhes do Docker

### `docker-compose.yml`
```yaml
version: '3.8'

services:
  mssql:
    image: mcr.microsoft.com/mssql/server:2022-latest
    container_name: sqlserver_db
    restart: unless-stopped
    environment:
      - ACCEPT_EULA=Y
      - MSSQL_SA_PASSWORD=YourStrong@Password123
      - MSSQL_PID=Developer
    ports:
      - "1433:1433"
    volumes:
      - mssql_data:/var/opt/mssql
      - ./database:/docker-entrypoint-initdb.d
    command:
      - /bin/bash
      - -c
      - |
        /opt/mssql/bin/sqlservr &
        sleep 30
        echo "Running database initialization..."
        /opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "YourStrong@Password123" -C -i "/docker-entrypoint-initdb.d/01_create_database.sql"
        /opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "YourStrong@Password123" -C -i "/docker-entrypoint-initdb.d/02_sp_total_periodo.sql"
        /opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "YourStrong@Password123" -C -i "/docker-entrypoint-initdb.d/03_fn_categoria_valor.sql"
        /opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "YourStrong@Password123" -C -i "/docker-entrypoint-initdb.d/04_tvf_transacoes_categorizadas.sql"
        /opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "YourStrong@Password123" -C -i "/docker-entrypoint-initdb.d/05_view_consolidada.sql"
        /opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "YourStrong@Password123" -C -i "/docker-entrypoint-initdb.d/06_dados_exemplo_load_test.sql"
        echo "Database initialization completed."
        wait

volumes:
  mssql_data:
```

### Pontos Importantes
- **Auto-seed a cada inicialização**: O script de teste de carga roda `DELETE FROM dbo.Transactions` + `DBCC CHECKIDENT` para reseed limpo
- **Persistência via volume**: Dados sobrevivem a `docker-compose down` (use `-v` para apagar)
- **Health check**: Verifique com `docker logs sqlserver_db` — procure por "Database initialization completed."

---

## 🎯 Regras de Negócio Implementadas

| Regra | Implementação |
|-------|---------------|
| Cartão: 16 dígitos | `ValidarCartao()` em `modUtilitarios.bas` |
| Valor: Decimal > 0 | `ValidarValor()` + CHECK constraint no DB |
| Descrição: ≤ 255 chars | `ValidarDescricao()` + VARCHAR(255) |
| Status: Approved/Pending/Canceled | CHECK constraint + ComboBox (Style=2) |
| Editar Approved = Bloqueado | Verificação no `cmdEditar_Click` + `frmTransacao` |
| Excluir = Confirmação | `MsgBox vbYesNo` antes do DELETE |
| Paginação = Server-side | `OFFSET/FETCH` no SQL (50 registros/página) |
| Exportação = Último mês | Query com `DATEADD/DATEDIFF` no `modExportacaoExcel` |

---

## 📊 Exportação Excel

**Arquivo gerado**: `Relatorios/Relatorio_Transacoes_YYYYMM.xlsx`

**Formatação incluída:**
- Cabeçalho: Azul escuro, fonte branca, negrito
- Valores: Formato moeda USD ($#,##0.00)
- Datas: MM/DD/YYYY HH:NN
- Cartões: Mascarados (XXXX XXXX XXXX XXXX)
- Zebra striping: Linhas alternadas em azul claro
- Linha TOTAL: Fórmula SUM, negrito, borda superior grossa
- Freeze panes: Linha 1 fixa ao rolar
- Bordas completas na tabela de dados

---

## 🐛 Logs e Debug

**Arquivo de log**: `logs/erros_YYYYMMDD.txt`

**Formato:**
```
[DD/MM/YYYY HH:NN:SS] Rotina | Nível | Mensagem | Detalhes
```

**Níveis**: ERROR, INFO, WARN

**Exemplo:**
```
[04/09/2026 14:30:15] Form1.Form_Load | INFO | Aplicação iniciada
[04/09/2026 14:35:22] cmdExcluir_Click | INFO | Transação excluída - ID: 15
[04/09/2026 14:40:01] ExportarUltimoMesParaExcel | ERROR | ActiveX component can't create object | SQL: SELECT...
```

---

## 🔧 Solução de Problemas

| Problema | Solução |
|----------|---------|
| "ActiveX component can't create object" (Excel) | Instale Excel ou use versão EPPlus (.NET) |
| Erro de conexão SQL Server | Verifique connection string, serviço SQL rodando, firewall |
| "Provider cannot be found" | Instale MDAC 2.8+ / Microsoft Access Database Engine |
| Grid não aparece / Erro MSFlexGrid | `regsvr32 C:\Windows\SysWOW64\MSFLXGRD.OCX` (admin) |
| Timeout na query | Verifique índices, reduza `m_recordsPerPage` |
| Acentos quebrados | Salve .frm/.bas como ANSI (padrão VB6) |
| Docker: porta 1433 em uso | Pare SQL Server local ou altere `ports:` no compose |

---

## 📚 Documentação

- [`docs/DOCUMENTACAO_TECNICA.md`](docs/DOCUMENTACAO_TECNICA.md) - Documentação técnica completa
- [`docs/EVOLUCAO_FUTURA.md`](docs/EVOLUCAO_FUTURA.md) - Roadmap, migração .NET, features futuras
- [`docs/PLANO_EXECUCAO.md`](docs/PLANO_EXECUCAO.md) - Plano de desenvolvimento por fases
- [`docs/proposta.md`](docs/proposta.md) - Requisitos originais

---

## 📝 Licença

Projeto desenvolvido para fins educacionais/demonstração - Administradora de Cartões XYZ.

---

**Versão**: 1.0  
**Data**: 08/09/2026  
**Tecnologia**: VB6 + SQL Server + Excel COM Interop + Docker