# Escopo de RPA para Web Scraping da Leveros Integra

## Objetivo
Automatizar a coleta de dados de produtos de ar-condicionado de diferentes categorias no site Leveros Integra, incluindo nome, voltagem, preço à vista, preço parcelado, quantidade de parcelas e URL da imagem.

## Fluxo de Processo

### 1. Login no Sistema

- Acessar a URL: https://leverosintegra.dev.br/login
- Preencher o campo de usuário com: 22429301000178@22429301000178
- Preencher o campo de senha com: 22429301000178
- Clicar no botão "Entrar"
- Tratar o popup de boas-vindas clicando no botão de fechar (ícone "close")

### 2. Navegação por Categorias

Criar uma lista de categorias a serem acessadas:
- Inverter
- Convencional
- Multi-Split
- Ar Janela
- Cassete
- Piso Teto
- VRF
- Ar Portátil
- Climatizador
- Ventilador

Para cada categoria na lista:
- Clicar no elemento da categoria correspondente
- Proceder à extração dos dados dos produtos

### 3. Extração de Dados dos Produtos

Para cada página na categoria atual:
- Identificar todos os cards de produtos na página
- Para cada card de produto:
  - Extrair o nome do produto (div com classe "menuItems text-caption q-pt-sm ellipsis-2-lines")
  - Extrair a voltagem (dentro do elemento com classe "q-chip--outline")
  - Extrair o preço principal (elemento com classe "text-h6 text-weight-bold text-teal-9")
  - Extrair informações de parcelamento (elementos com classe "text-caption text-weight-bold")
  - Extrair o preço à vista (elemento contendo "à vista")
  - Capturar a URL da imagem do produto (src do elemento img dentro da classe "q-img__image")
  - Armazenar os dados extraídos em uma tabela ou planilha

- Verificar se existe um botão de próxima página
  - Se existir, clicar no botão para avançar para a próxima página
  - Se não existir, finalizar a extração para a categoria atual e passar para a próxima categoria

### 4. Armazenamento de Dados

Salvar todos os dados extraídos em uma planilha Excel com as seguintes colunas:
- Categoria
- Nome do Produto
- Voltagem
- Preço Principal
- Preço à Vista
- Quantidade de Parcelas
- Valor da Parcela
- URL da Imagem

## Estrutura do Projeto em UiPath/WinAutomation

### Variáveis Principais
- `categorias`: Array contendo todas as categorias a serem analisadas
- `produtos`: Tabela para armazenar temporariamente os dados extraídos
- `paginaAtual`: Controle da página atual na navegação
- `temProximaPagina`: Flag para verificar se existe próxima página

### Principais Seletores
- Campo de usuário: `input[id^='f_'][aria-label='Informe seu usuário']`
- Campo de senha: `input[id^='f_'][aria-label='Informe sua senha']`
- Botão Entrar: `button span.block:contains('Entrar')`
- Botão fechar popup: `button i.material-icons:contains('close')`
- Cards de produtos: `div.q-card.my-card`
- Nome do produto: `div.menuItems.text-caption.q-pt-sm.ellipsis-2-lines`
- Voltagem: `div.q-chip--outline`
- Preço principal: `div.text-h6.text-weight-bold.text-teal-9`
- Informações de parcelamento: `div.text-caption.text-weight-bold`
- Botão próxima página: `button i.material-icons:contains('fast_forward')`
- Categorias: `div.text-teal-10.q-pa-md.text-center:contains('NOME_CATEGORIA')`

### Tratamento de Exceções
- Implementar timeouts adequados para carregamento de páginas
- Tratamento de erro para login falho
- Verificação de elementos visíveis antes de interagir
- Tratamento para casos em que a categoria não possui produtos

## Exemplo de Pseudocódigo

```
Início
  // Login no sistema
  Navegar para "https://leverosintegra.dev.br/login"
  Digitar "22429301000178@22429301000178" no campo de usuário
  Digitar "22429301000178" no campo de senha
  Clicar no botão Entrar
  Esperar página carregar
  Se popup estiver visível então
    Clicar no botão fechar
  Fim Se

  // Criar planilha para armazenar dados
  Criar nova planilha Excel "ProdutosLeveros.xlsx"
  Adicionar cabeçalhos na planilha

  // Lista de categorias
  categorias = ["Inverter", "Convencional", "Multi-Split", "Ar Janela", "Cassete", "Piso Teto", "VRF", "Ar Portátil", "Climatizador", "Ventilador"]

  // Percorrer cada categoria
  Para cada categoria em categorias faça
    Clicar no elemento da categoria
    Esperar página carregar
    paginaAtual = 1
    temProximaPagina = Verdadeiro

    // Percorrer todas as páginas da categoria
    Enquanto temProximaPagina faça
      // Extrair dados de todos os produtos da página atual
      cards = Encontrar todos os cards de produtos
      
      Para cada card em cards faça
        nome = Extrair texto do elemento nome do produto
        voltagem = Extrair texto do elemento voltagem
        preçoPrincipal = Extrair texto do elemento preço principal
        infoParcelamento = Extrair texto do elemento informações de parcelamento
        preçoÀVista = Extrair texto do elemento preço à vista
        urlImagem = Extrair atributo src da imagem do produto
        
        // Adicionar dados à planilha
        Adicionar linha na planilha com [categoria, nome, voltagem, preçoPrincipal, preçoÀVista, infoParcelamento, urlImagem]
      Fim Para

      // Verificar se existe próxima página
      Se botão próxima página estiver habilitado então
        Clicar no botão próxima página
        paginaAtual = paginaAtual + 1
        Esperar página carregar
      Senão
        temProximaPagina = Falso
      Fim Se
    Fim Enquanto
  Fim Para

  // Salvar planilha
  Salvar e fechar planilha Excel
  Mostrar mensagem "Extração de dados concluída com sucesso!"
Fim
```

## Armazenamento e Organização dos Dados em Excel

### Estrutura do Arquivo Excel

O arquivo Excel será organizado da seguinte forma:

1. **Planilha Principal**: "Produtos Leveros"
   - Contém todos os produtos extraídos de todas as categorias
   - Inclui filtros e formatação condicional para facilitar a análise

2. **Planilhas por Categoria**:
   - Uma planilha separada para cada categoria
   - Criadas dinamicamente durante a execução do RPA
   - Contêm apenas os produtos da categoria específica

3. **Planilha de Resumo**:
   - Tabela dinâmica mostrando totais e médias por categoria
   - Gráficos para visualização rápida de dados

### Formatação e Organização

1. **Cabeçalhos**:
   - Formatados com cores corporativas (tons de verde e azul)
   - Fonte em negrito e tamanho maior
   - Filtros automáticos habilitados

2. **Dados**:
   - Formatação condicional para preços (verde para valores abaixo da média, vermelho para acima)
   - Formatação de moeda para valores monetários
   - Hiperlinks para as URLs das imagens

3. **Fórmulas e Automação**:
   - Fórmulas PROCV para relacionar dados entre planilhas
   - Cálculo automático de valores médios e estatísticas por categoria
   - Validação de dados para garantir consistência

### Processo de Gravação

1. **Inicialização**:
   - Criar o arquivo com todas as planilhas necessárias
   - Aplicar formatação e configurar cabeçalhos
   - Preparar fórmulas e validações

2. **Durante a Extração**:
   - Adicionar cada produto na planilha principal
   - Simultaneamente adicionar na planilha da categoria correspondente
   - Verificar e tratar possíveis erros de formatação

3. **Finalização**:
   - Atualizar todas as fórmulas e tabelas dinâmicas
   - Ajustar largura das colunas automaticamente
   - Aplicar proteção para evitar alterações acidentais
   - Salvar em formato .xlsx com macros habilitadas para funcionalidades avançadas

### Exemplo de Estrutura de Dados

| Categoria   | Nome do Produto           | Voltagem | Preço Principal | Preço à Vista | Qtd. Parcelas | Valor Parcela | URL da Imagem                |
|-------------|---------------------------|----------|-----------------|---------------|---------------|---------------|------------------------------|
| Inverter    | Split Hi Wall Carrier     | 220V     | R$ 2.499,00     | R$ 2.249,10   | 10x           | R$ 249,90     | https://exemplo.com/img1.jpg |
| Convencional| Janela Springer Midea     | 110V     | R$ 1.299,00     | R$ 1.169,10   | 10x           | R$ 129,90     | https://exemplo.com/img2.jpg |
| Piso Teto   | Split Piso Teto LG        | 220V     | R$ 5.999,00     | R$ 5.399,10   | 10x           | R$ 599,90     | https://exemplo.com/img3.jpg |

## Considerações Adicionais

- **Robustez**: Implementar esperas dinâmicas para garantir que os elementos sejam carregados antes de interagir com eles.
- **Tratamento de erros**: Adicionar blocos try-catch para lidar com exceções durante a execução.
- **Log de atividades**: Registrar cada ação realizada pelo robô para facilitar a depuração.
- **Checkpoints**: Implementar pontos de verificação para garantir que os dados foram extraídos corretamente.
- **Reuso de código**: Criar funções reutilizáveis para tarefas comuns, como extrair dados de um card ou navegar entre páginas.
- **Validação de dados**: Incluir validações para garantir que os dados extraídos estão no formato esperado.
- **Backup automático**: Implementar salvamento periódico dos dados para evitar perda em caso de falha.
- **Relatório de execução**: Gerar um relatório ao final do processo informando estatísticas da extração.
