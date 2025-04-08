# RPA para Web Scraping da Leveros Integra

Este projeto implementa um RPA (Robotic Process Automation) para extrair dados de produtos de ar-condicionado do site Leveros Integra.

## Funcionalidades

- Login automático no sistema Leveros Integra
- Navegação por diferentes categorias de produtos
- Extração de dados dos produtos (nome, voltagem, preços, etc.)
- Armazenamento dos dados em arquivo Excel formatado
- Organização por categorias e resumo estatístico

## Requisitos

- Python 3.8 ou superior
- Chrome ou Chromium
- Pacotes Python listados no arquivo `requirements.txt`

## Instalação

1. Clone ou baixe este repositório
2. Instale as dependências:

```bash
pip install -r requirements.txt
```

## Como Usar

Execute o script principal:

```bash
python leveros_rpa.py
```

O RPA irá:
1. Abrir o navegador Chrome
2. Fazer login no sistema Leveros Integra
3. Navegar pelas categorias de produtos
4. Extrair os dados dos produtos
5. Salvar os dados em um arquivo Excel formatado

## Estrutura do Projeto

- `leveros_rpa.py`: Script principal de automação
- `requirements.txt`: Lista de dependências
- `Escopo_RPA_Leveros_Integra.md`: Documentação detalhada do escopo
- `Template_Produtos_Leveros.xlsx`: Template da estrutura de dados esperada

## Customização

Você pode personalizar o RPA modificando as seguintes variáveis no início da classe `LeverosRPA`:

- `url_login`: URL da página de login
- `usuario` e `senha`: Credenciais de acesso
- `categorias`: Lista de categorias a serem processadas

## Logs

O RPA cria logs detalhados da execução no arquivo `leveros_rpa.log`.

## Observações Importantes

- O RPA foi configurado para funcionar com o layout atual do site Leveros Integra. Alterações no site podem exigir ajustes nos seletores CSS.
- Para executar em modo headless (sem interface gráfica), descomente a linha `chrome_options.add_argument("--headless")` no método `iniciar_navegador()`.
- Os arquivos Excel são salvos com timestamp no nome para evitar sobrescrever dados anteriores.

## Tratamento de Erros

O RPA inclui tratamento robusto de erros para lidar com:
- Falhas de login
- Elementos que não carregam
- Timeouts de página
- Categorias ou produtos não encontrados
