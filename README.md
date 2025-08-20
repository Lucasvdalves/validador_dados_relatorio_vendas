# Analisador e Comparador de Planilhas de Contratos

Este projeto é uma ferramenta de auditoria e reconciliação de dados, projetada para comparar e identificar contratos em comum entre duas planilhas Excel. É ideal para equipes de finanças, cobrança ou operações que precisam cruzar informações de diferentes fontes e gerar um relatório unificado de forma rápida e precisa.

-----

## Funcionalidades

  * **Comparação de Planilhas**: Faz a leitura e a junção (merge) de dados de duas planilhas diferentes, usando o número do contrato como chave de comparação.
  * **Seleção de Colunas**: Lê apenas as colunas essenciais, otimizando o processamento. Ele busca o número do contrato e o valor renegociado em uma planilha, e o número do contrato e o valor pago em outra.
  * **Limpeza e Padronização**: Remove linhas com valores vazios e garante que os dados dos contratos estejam no formato correto para a comparação.
  * **Relatório Unificado**: Gera uma nova planilha Excel com os contratos que foram encontrados em ambas as bases, exibindo o número do contrato, o valor renegociado e o valor pago, facilitando a análise.

-----

## Tecnologias Utilizadas

  * **Python**: A base do projeto.
  * **Pandas**: A biblioteca mais poderosa para manipulação e análise de dados em Python.
  * **openpyxl**: Necessária para que o Pandas consiga ler e salvar arquivos no formato `.xlsx`.

-----

## Como Usar

### Pré-requisitos

Certifique-se de que o **Python** e as bibliotecas **Pandas** e **openpyxl** estão instaladas. Se você não as tem, pode instalá-las facilmente com o seguinte comando:

```sh
pip install pandas openpyxl
```

### Configuração

1.  **Ajuste os caminhos**: No final do script, altere os caminhos nas variáveis `caminho_abr`, `caminho_suelen` e `caminho_saida` para os locais dos seus arquivos de entrada e onde você deseja salvar o arquivo de saída.

<!-- end list -->

```python
comparar_planilhas(
    r'caminho/para/seu/Dados_Abril.xlsx',
    r'caminho/para/sua/Planilha_Suelen.xlsx',
    r'caminho/para/o/diretorio_de_saida'
)
```

2.  **Verifique os nomes das abas**: O script busca as abas `'Abril'` e `'1_BASE'`. Se o nome das suas abas for diferente, ajuste o código na linha `sheet_name`.

### Execução

Basta rodar o script no seu terminal:

```sh
python nome_do_seu_script.py
```

Ao final do processamento, você verá a mensagem de sucesso e o arquivo `Contratos_Identificados.xlsx` será gerado no diretório de saída que você especificou.

-----

## ✍️ Licença

Este projeto está sob a licença **MIT**.
