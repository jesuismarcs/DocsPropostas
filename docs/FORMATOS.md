# Contrato de dados: Excel e Word

## Excel de entrada

Folha obrigatória: **Doc Proposta**, lida sem cabeçalho de tabela.

| Campo | Célula/intervalo |
| --- | --- |
| ID_INTERNO | K3 |
| NOME_DA_EMPREITADA | K4 |
| LOCAL_DA_OBRA | L6 |
| DURACAO_DA_EMPREITADA | N14 |
| DOCUMENTOS_DA_PROPOSTA | J25:J45, células não vazias, separadas por nova linha |
| Habilitações importadas | M25:M36 |

O leitor usa posições fixas. Garanta que o livro se estende até à linha 45 e coluna N; o gerador de exemplo faz isso explicitamente.

A duração é tratada como texto. A ferramenta não verifica o prazo nem calcula o valor da proposta.

## Variáveis dos modelos Word

Além dos campos anteriores, são disponibilizados:

- `DATA_HOJE`: data de geração, formatada segundo a localização disponível no sistema.
- `ALVARA_SELECOES`: lista de dicionários para apresentar habilitações.

Exemplo de ciclo num modelo DOCX:

```jinja
{% for item in ALVARA_SELECOES %}
{{ item.cat_num_part }} / {{ item.sub_prefix_display }}
{{ item.sub_desc }} — {{ item.class_val }} — {{ item.valor_executar }}
{% endfor %}
```

Os campos `full_cat_name` e `full_sub_name` permitem reabrir seleções no editor.

Os marcadores devem ser inseridos nos modelos Word; esta documentação não é um modelo DOCX. O comando `python examples/create_demo.py` cria um modelo real de demonstração.

## Importação das habilitações

O parser reconhece expressões como `1, 2 sub, 1 cat`. A expressão `valor global` é reconhecida; quando não existe, o parser usa o valor demonstrativo `1`.

**É obrigatório rever esta informação:** o parser não interpreta uma especificação de habilitações em linguagem livre nem determina classes pelo valor da obra. As classes distribuídas estão vazias e precisam de preenchimento. O catálogo preserva as descrições existentes no código e não é apresentado como lista normativa atualizada.

## Saída

Cada modelo válido da pasta produz:

```text
ID_INTERNO_NomeDoModelo_LOCAL_DA_OBRA.docx
```

São ignorados ficheiros temporários Word começados por `~$`. Caracteres inválidos usuais são substituídos nos nomes. Valores ausentes de referência/local recebem `SEM-ID`/`SEM-LOCAL`.

O processamento para no primeiro erro. Não substitui ficheiros de saída existentes; documentos anteriores no mesmo lote podem já estar guardados. Não há transação global nem reversão automática.

