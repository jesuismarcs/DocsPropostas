# Análise técnica da edição autónoma

## Responsabilidades

| Módulo | Responsabilidade |
| --- | --- |
| main.py | Localização do sistema, verificação opcional de acesso e arranque |
| gui.py | Interface principal, importação Excel e recolha do contexto |
| alvara_gui.py | Edição e pesquisa de habilitações |
| constants.py | Catálogo de seleção, com classes empresariais removidas |
| data_logic.py | Extração de listas e renderização DOCX |
| config_manager.py | Preferências JSON locais |
| utils.py | Recursos, nomes de ficheiro e acesso opcional |

## Achados e alterações

1. **Modelos num caminho fixo:** substituído por seleção na interface, preferência persistida e variável de ambiente.
2. **Dependências não declaradas:** criados requirements de execução e desenvolvimento.
3. **Imports dependentes da pasta corrente:** organizados em pacote, com arranque pela raiz.
4. **Classes específicas da organização:** removidas da distribuição; o campo de classe passou a editável.
5. **Seleções perdidas ao reabrir/fechar o editor:** preservados os nomes completos e a lista anterior no cancelamento.
6. **Tooltips a consultar um índice de edição em botões:** posicionamento corrigido.
7. **Variáveis de modelo ausentes:** renderização estrita, com erro visível.
8. **Caracteres XML em campos:** ativado escape no motor DOCX.
9. **Sobrescrita de saída:** bloqueada com mensagem de erro.
10. **Ícones privados:** tornados opcionais, mantendo os botões textuais.

Os originais no repositório de trabalho foram preservados.

## Testes reproduzíveis

```powershell
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe -m pytest -q
```

Cobertura funcional: contrato do exemplo Excel, geração Word com caracteres especiais, variável inexistente, ficheiro já existente e ausência das classes empresariais.

## Limitações conhecidas

- Leitura de posições fixas e validação limitada de campos vazios.
- Datas dependentes da localização disponível.
- Catálogo de habilitações não validado como fonte normativa.
- Trabalho de geração na thread da interface.
- Ausência de PDF, assinatura digital, histórico centralizado e comparação visual Word.
- O formato de nomes não cobre todos os limites possíveis de comprimento/caminhos do Windows.
- Não foi realizado teste manual completo da interface ou empacotamento em executável.

A apresentação do projeto descreve funcionalidades observadas no código e testadas quando indicado. Não são atribuídas métricas de produtividade sem medição.

