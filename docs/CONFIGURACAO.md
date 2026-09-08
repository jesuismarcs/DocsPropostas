# Definições e perfis

Abra **Definições** no topo da aplicação. O perfil **Padrão Marcos** vem selecionado e mantém as folhas, células e variáveis do modelo original.

## Adaptar outro Excel

1. Duplique o perfil e atribua-lhe um nome.
2. Em Dados de entrada, abra um Excel de exemplo e escolha a folha.
3. Escolha o campo de proposta. Clique numa célula da pré-visualização para associar a posição, ou indique um intervalo, nome definido ou tabela.
4. Defina o tipo: texto, número, data, lista ou tabela. Campos simples exigem uma célula; listas usam intervalos; tabelas usam a primeira linha como cabeçalhos.
5. Clique em **Aplicar associação** e em **Testar mapeamento**. Corrija os erros apresentados.
6. Configure modelos e destinos em Saídas e clique em **Guardar e aplicar**.

O teste mostra os valores interpretados. Erros indicam campo e origem. A pré-visualização mostra até 100 linhas e 30 colunas; pode introduzir qualquer endereço válido no editor mesmo quando não está visível.

## Campos adicionais e habilitações

Adicionar campo cria uma nova variável de modelo. Valores manuais/constantes são editáveis sem Excel. Em Editar valor longo, introduza uma linha por elemento de lista; tabelas manuais usam cabeçalhos e colunas separados por tabulação.

Habilitações podem ser desativadas. O catálogo é editável em Catálogo de habilitações. A importação antiga usa o texto `sub, cat`; uma tabela estruturada deve fornecer os campos documentados em Formatos para continuar a funcionar com os modelos existentes.

## Regras de leitura

- Números: separadores decimal e de milhares definidos no perfil.
- Datas: formato de entrada indicado por `%d`, `%m` e `%Y`.
- Linhas ocultas: incluir ou ignorar.
- Células unidas: usar a âncora, repetir o valor da âncora ou rejeitar.
- Linhas vazias de tabela: ignorar ou terminar a leitura.
- `.xlsx`/`.xlsm`: não executa macros nem recalcula fórmulas. Fórmulas sem resultado disponível exigem recalcular e guardar no Excel.
- `.xls`: leitura por xlrd; para nomes definidos e tabelas estruturadas, guardar como `.xlsx` ou associar intervalos. A identificação de fórmulas sem cache é limitada pelo formato legado.

## Guardar e partilhar

As definições são guardadas em `%LOCALAPPDATA%/MarcosTools/DocsPropostas/profiles.json`. A configuração antiga é importada uma única vez. Preferências novas são guardadas explicitamente na janela de definições.

Exportar gera um JSON versão 1 sem ficheiros de trabalho, caminhos pessoais ou credenciais. Valores constantes/manuais e configurações identificadoras de email são removidos. Reassocie os ficheiros no destino. Cancelar não aplica o rascunho ao perfil ativo.

## Geração

Selecione modelos pelo nome, separados por `;`, ou deixe vazio para processar todos os DOCX da pasta. O nome de saída admite as variáveis do contexto e `{modelo}`. Todos os modelos são renderizados numa pasta temporária antes da publicação; colisões de nomes são rejeitadas e ficheiros existentes são preservados.
