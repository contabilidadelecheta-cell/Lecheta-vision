# Conversor Platinox

**Status: AINDA NÃO IMPLEMENTADO no Lecheta Vision.**

## Destino planejado
- Sistema principal: Lecheta Control
- Módulo: Lecheta Vision
- Ferramenta: Conversor Platinox
- Integração ao menu/interface: pendente. **Não ativar nem divulgar como disponível.**

## Finalidade
Receber um número variável de relatórios de clientes e fornecedores (matriz e filial, inclusive arquivos .xls antigos), limpar os dados e produzir **dois arquivos Excel separados**:
- `Clientes.xlsx`
- `Fornecedores.xlsx`

Formato esperado: colunas **Conta** e **CNPJ** (incluindo CPF quando presente), sem células mescladas, com filtro; preservar zeros à esquerda e manter duplicados por padrão. Identificar clientes/fornecedores a partir do título, aba ou nome do arquivo.

## Histórico do projeto independente
Já houve desenvolvimento e testes de um script Python autônomo, **fora do Vision**. O último teste documentado com quatro relatórios atualizados gerou 43.440 registros de clientes e 9.784 de fornecedores. **Isso não significa que o conversor esteja implementado neste repositório.**

Arquivos de origem citados no histórico, ainda não importados nem validados neste repositório:
- `juntar_relatorios.py`
- `EXECUTAR.bat`
- `INSTALAR.bat`
- `requirements.txt`
- `LEIA-ME.txt`
- `Juntar_Relatorios_Python.zip`

Localização histórica informada: `C:\\Users\\TI\\Documents\\Codex\\2026-10-07\\quero\\outputs\\`. Verificar a existência e a versão atual antes de importar.

## Pendências para implementação
1. Recuperar e revisar os arquivos originais do conversor.
2. Definir interface e fluxo de upload/resultado no Lecheta Vision.
3. Integrar a rotina Python sem alterar ferramentas existentes.
4. Testar entradas múltiplas, classificação, preservação de documentos e resultados.
5. Somente após testes, mudar o status para implementado e adicionar ao menu.

> Este diretório contém apenas documentação da futura ferramenta; não há código executável do conversor aqui.
