# Automatização Conferencia entregas de notas vencidas

Este repositório contém o script Python `nfs_vencidas.py`, responsável por automatizar o fluxo de verificação e atualização de notas fiscais vencidas não entregues no SAP, combinando extração via SAP GUI Scripting, processamento de dados em Excel, consulta ao banco de dados interno e atualização automática no SAP.

## Visão Geral do Fluxo

1. **Exportação SAP (FBL5N)**
   - Inicia o SAP GUI via `saplogon.exe` e conecta à instância configurada.
   - Executa transação **FBL5N** com a variante `MERC_N_ENTREGU` para filtrar notas não entregues.
   - Exporta a saída ALV para Excel (`merc_nao_entregue.XLSX`).

2. **Processamento em Python**
   - Remove arquivos antigos de exportação e renomeia o arquivo gerado.
   - Lê o Excel e extrai a coluna `Referência`, formatando números e séries.
   - Monta lista de notas e executa query em `MV_PBI_DATASLOGISTICAS` no SQL Server.
   - Formata e mescla dados de Excel e resultados SQL, gerando:
     - `Notas Fiscais.xlsx` (todas as notas analisadas);
     - `Notas Não Localizadas.xlsx` (caso existam notas sem log no banco);
     - `Nfs entregues.xlsx` (caso existam notas já entregues).

3. **Atualização no SAP**
   - Para cada nota identificada como já entregue, abre a tela de ajuste no SAP e remove a marcação de "mercadorias não entregues".

4. **Limpeza Final**
   - Encerra processos do Excel e do SAP GUI.
   - Imprime `concluido` no console como sinal para controle externo.

## Pré-requisitos

- **Windows** com SAP GUI Scripting habilitado.
- **Python 3.8+** instalado.
- Dependências Python (via `pip install pandas pywin32 psutil python-dateutil pyautogui pyodbc`).
- Arquivo de credenciais SAP (`SAP_Credentials.xlsx`) contendo colunas `usuario` e `senha`.
- Conexão ODBC/SQL Server configurada (Driver SQL Server, `Server`, `Database`, Trusted Connection). 

## Estrutura do Script (`nfs_vencidas.py`)

- **Imports e configuração de datas**: define `today` e `ha_30_dias` em formato `DD.MM.YYYY`.
- **Cleanup inicial**: remove arquivo antigo e finaliza processos SAP existentes.
- **Login SAP GUI**: abre SAP Logon, conecta na instância (`pe1`), preenche cliente, usuário, senha e idioma.
- **Exportação FBL5N**: carrega variante, executa relatório e exporta ALV para Excel.
- **Processamento de Excel**:
  - Renomeia `export.XLSX` para `merc_nao_entregue.XLSX`;
  - Lê planilha, formata referências e gera lista para consulta SQL.
- **Consulta SQL e merge**:
  - Query em `MV_PBI_DATASLOGISTICAS` para obter `PREV LOGIST` e `DT ENT`;
  - Mescla com dados de Excel e gera relatórios em Excel.
- **Ajuste no SAP**:
  - Para cada nota entregue, navega via scripting e desmarca o flag de não entrega.
- **Encerramento**: finaliza processos Excel e SAP e imprime `concluido`.

## Uso

Execute diretamente:

```powershell
python nfs_vencidas.py
```

Monitore o console para mensagens:

- `Login realizado com sucesso!`
- Relatórios gerados ou mensagens de inexistência de notas
- `concluido` (ponto de sincronização com fluxo de automação)

## Troubleshooting

- **Erro de conexão ODBC**: verifique driver e DSN ou parâmetros na string de conexão.
- **SAP GUI Scripting não habilitado**: ative em `SAP GUI Options` > `Accessibility & Scripting`.
- **Permissões de arquivo**: execute script com permissão para ler/excluir/criar arquivos no diretório.
- **Delays insuficientes**: ajuste valores de `time.sleep` para adequar ao desempenho do ambiente.


