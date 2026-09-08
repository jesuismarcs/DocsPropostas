# Utilização e configuração

## Instalação

Execute os comandos do README a partir da raiz do repositório. Use o Python do ambiente virtual tanto para instalar como para arrancar. Tkinter deve estar disponível na instalação Python.

A aplicação lê `.xlsx` e `.xlsm` através de openpyxl/pandas; o suporte a `.xls` usa xlrd. Não executa macros do livro. Se houver fórmulas, guarde primeiro os resultados calculados no Excel.

## Percurso na interface

1. Selecione o Excel ou arraste um ficheiro para a janela.
2. Reveja os quatro campos principais e a lista de documentos.
3. Se necessário, abra **Inserir/Editar Habilitações**. Selecione categoria/subcategoria, introduza classe e valor, e confirme.
4. Selecione a pasta de modelos Word. Também pode defini-la em `DOCSPROPOSTAS_TEMPLATES`.
5. Escolha uma pasta de saída e gere os documentos.

A pasta de saída pode reutilizar uma preferência anterior; confira-a em cada concurso. A seleção de modelos é persistida ao fechar a aplicação. Uma variável de ambiente de modelos preenchida tem prioridade sobre a escolha guardada.

## Preferências e ambiente

Copie `.env.example` para `.env` apenas se precisar de configuração externa.

| Variável | Finalidade |
| --- | --- |
| DOCSPROPOSTAS_TEMPLATES | Caminho da pasta de modelos DOCX |
| DOCSPROPOSTAS_CONFIG | Caminho alternativo do JSON de preferências |
| APP_ENV | `development` por defeito; `production` exige configuração de licença |
| DOCSPROPOSTAS_LICENSE_KEY | Valor esperado pela verificação opcional |
| DOCSPROPOSTAS_LICENSE_FILE | Ficheiro local com o valor de licença |
| DOCSPROPOSTAS_LICENSE_URL | Fonte HTTP(S) alternativa, consultada apenas se configurada |

O ficheiro de licença tem prioridade sobre o URL. Sem chave esperada, o arranque é permitido fora de `production`. Este mecanismo foi preservado como compatibilidade opcional; não distribui chaves nem constitui proteção forte de software.

As preferências são guardadas em `local/config.json`. Caminhos absolutos evitam ambiguidades. Variáveis já definidas no ambiente têm prioridade sobre `.env`.

## Recursos visuais

A interface funciona com os rótulos dos botões, sem depender dos ícones privados. Ícones próprios podem ser colocados em `resources/assets/`: excel, folder, clear, alvara, generate, add e confirm, em PNG. Essa pasta não é versionada por defeito.

## Diagnóstico

| Situação | Verificação |
| --- | --- |
| Folha não encontrada | Nome exato `Doc Proposta` |
| Erro de índices na importação | Modelo com extensão suficiente até à linha 45 e coluna N |
| Variável desconhecida no Word | Usar os nomes documentados; manter o marcador íntegro no Word |
| Documento já existe | Escolher outra pasta ou gerir o ficheiro existente |
| Dados desatualizados | Recalcular e guardar o Excel antes de importar |
| Classe por preencher | Abrir o editor de habilitações e introduzir a classe aplicável |
| Acesso negado | Rever a configuração opcional de licença e APP_ENV |

Não existe execução em background da geração: lotes grandes podem ocupar a interface até terminarem.

