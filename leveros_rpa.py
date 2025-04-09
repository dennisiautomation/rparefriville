"""
RPA para Web Scraping da Leveros Integra
Script principal de automação para extrair dados de produtos de ar-condicionado
"""

import os
import time
import logging
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import platform
import subprocess
import traceback
import requests
import zipfile
import shutil
import tempfile
from fpdf import FPDF
from urllib.parse import urlparse
from io import BytesIO
from PIL import Image

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("leveros_rpa.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class LeverosRPA:
    """Classe principal do RPA para extração de dados da Leveros Integra"""
    
    def __init__(self, headless=False):
        """Inicializa o RPA com as configurações básicas"""
        self.url_login = "https://leverosintegra.dev.br/login"
        self.usuario = "22429301000178@22429301000178"
        self.senha = "22429301000178@22429301000178"
        self.categorias = [
            "Inverter", "Convencional", "Multi-Split", "Ar Janela", 
            "Cassete", "Piso Teto", "VRF", "Ar Portátil", 
            "Climatizador", "Ventilador"
        ]
        self.dados_produtos = []
        self.driver = None
        self.timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        self.arquivo_excel = f"ProdutosLeveros_{self.timestamp}.xlsx"
        self.arquivo_pdf = f"ProdutosLeveros_{self.timestamp}.pdf"
        self.headless = headless
        
        # Seletores CSS para os elementos de interesse
        self.seletores = {
            "campo_usuario": "input[id^='f_'][aria-label='Informe seu usuário']",
            "campo_senha": "input[id^='f_'][aria-label='Informe sua senha']",
            "botao_entrar": "button span.block:contains('Entrar')",
            "botao_fechar_popup": "button i.material-icons:contains('close')",
            "cards_produtos": "div.q-card.my-card",
            "nome_produto": "div.menuItems.text-caption.q-pt-sm.ellipsis-2-lines",
            "voltagem": "div.q-chip--outline",
            "preco_principal": "div.text-h6.text-weight-bold.text-teal-9",
            "info_parcelamento": "div.text-caption.text-weight-bold",
            "preco_a_vista": "div.text-caption:contains('à vista')",
            "url_imagem": "div.q-img > img.q-img__image",
            "botao_proxima_pagina": "button i.material-icons:contains('fast_forward')",
        }
    
    def inicializar_navegador(self):
        """Inicializa o navegador Chrome com as configurações necessárias"""
        try:
            logger.info("Inicializando o navegador Chrome...")
            
            # Detectar plataforma e arquitetura
            plataforma = platform.system()
            arquitetura = platform.machine()
            logger.info(f"Plataforma: {plataforma}, Arquitetura: {arquitetura}")
            
            # Configurações do Chrome
            opcoes = webdriver.ChromeOptions()
            opcoes.add_argument("--start-maximized")
            opcoes.add_argument("--disable-extensions")
            opcoes.add_argument("--disable-notifications")
            opcoes.add_argument("--disable-popup-blocking")
            
            # Para testes, deixamos o navegador visível, mas em produção pode ser headless
            if self.headless:
                opcoes.add_argument("--headless")
            
            # Configuração específica para Mac com chips M1/M2
            if plataforma == "Darwin" and arquitetura == "arm64":
                logger.info("Detectado Mac com chip Apple Silicon (M1/M2)")
                opcoes.binary_location = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
                
                # Download direto do ChromeDriver para Mac ARM
                temp_dir = tempfile.mkdtemp()
                chromedriver_url = "https://storage.googleapis.com/chrome-for-testing-public/135.0.7049.42/mac-arm64/chromedriver-mac-arm64.zip"
                zip_path = os.path.join(temp_dir, "chromedriver.zip")
                
                try:
                    # Baixar o arquivo ZIP
                    logger.info(f"Baixando ChromeDriver de {chromedriver_url}")
                    response = requests.get(chromedriver_url)
                    with open(zip_path, 'wb') as f:
                        f.write(response.content)
                    
                    # Extrair o arquivo ZIP
                    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                        zip_ref.extractall(temp_dir)
                    
                    # Procurar o executável chromedriver
                    chromedriver_path = None
                    for root, dirs, files in os.walk(temp_dir):
                        if "chromedriver" in files:
                            chromedriver_path = os.path.join(root, "chromedriver")
                            break
                    
                    if not chromedriver_path:
                        raise Exception("Não foi possível encontrar o executável chromedriver no pacote baixado")
                    
                    # Garantir que o executável tem permissões de execução
                    os.chmod(chromedriver_path, 0o755)
                    logger.info(f"ChromeDriver executável em: {chromedriver_path}")
                    
                    # Criar o driver com o executável
                    service = Service(executable_path=chromedriver_path)
                    self.driver = webdriver.Chrome(service=service, options=opcoes)
                    
                except Exception as e:
                    logger.error(f"Erro ao configurar ChromeDriver para M1/M2: {str(e)}")
                    logger.error(traceback.format_exc())
                    
                    # Fallback para o método padrão
                    logger.info("Tentando método alternativo com ChromeDriverManager")
                    self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opcoes)
            else:
                # Configuração padrão para outras plataformas
                self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opcoes)
            
            # Configurar tempos de espera
            self.driver.implicitly_wait(10)  # espera implícita de 10 segundos
            self.wait = WebDriverWait(self.driver, 15)  # espera explícita de até 15 segundos
            
            logger.info("Navegador Chrome inicializado com sucesso.")
            return True
        except Exception as e:
            logger.error(f"Erro ao inicializar navegador: {str(e)}")
            return False
    
    def fazer_login(self):
        """Realiza o login no sistema Leveros Integra"""
        try:
            logger.info("Acessando a página de login...")
            self.driver.get(self.url_login)
            time.sleep(3)  # Aguarda o carregamento da página
            
            # Preenche o campo de usuário
            logger.info("Preenchendo campo de usuário...")
            campo_usuario = self.driver.find_element(By.CSS_SELECTOR, 
                                                   "input[aria-label='Informe seu usuário']")
            campo_usuario.clear()
            campo_usuario.send_keys(self.usuario)
            
            # Preenche o campo de senha
            logger.info("Preenchendo campo de senha...")
            campo_senha = self.driver.find_element(By.CSS_SELECTOR, 
                                                 "input[aria-label='Informe sua senha']")
            campo_senha.clear()
            campo_senha.send_keys(self.senha)
            
            # Clica no botão entrar
            logger.info("Clicando no botão Entrar...")
            botao_entrar = self.driver.find_element(By.XPATH, 
                                                  "//button//span[contains(text(), 'Entrar')]")
            botao_entrar.click()
            
            # Aguarda o carregamento da página após o login
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.q-layout"))
            )
            logger.info("Login realizado com sucesso!")
            
            # Trata o popup de boas-vindas com uma abordagem mais robusta
            try:
                logger.info("Verificando se há popup de boas-vindas...")
                # Aguarda um pouco para o popup aparecer completamente
                time.sleep(2)
                
                # Verifica se existe o backdrop do diálogo
                backdrop = self.driver.find_elements(By.CSS_SELECTOR, "div.q-dialog__backdrop")
                if backdrop:
                    logger.info("Popup detectado. Tentando fechar...")
                    
                    # Tenta localizar o botão de fechar de várias maneiras
                    try:
                        # Método 1: Botão com ícone close
                        botao_fechar = self.driver.find_element(By.CSS_SELECTOR, 
                                                            "button i.material-icons")
                        logger.info("Fechando popup de boas-vindas com método 1...")
                        self.driver.execute_script("arguments[0].click();", botao_fechar)
                    except:
                        try:
                            # Método 2: Qualquer botão no diálogo
                            botao_fechar = self.driver.find_element(By.CSS_SELECTOR, 
                                                                "div.q-dialog button")
                            logger.info("Fechando popup de boas-vindas com método 2...")
                            self.driver.execute_script("arguments[0].click();", botao_fechar)
                        except:
                            # Método 3: Clicar no backdrop do diálogo
                            logger.info("Fechando popup de boas-vindas com método 3...")
                            self.driver.execute_script("arguments[0].click();", backdrop[0])
                    
                    # Aguarda o popup desaparecer
                    time.sleep(1)
                    logger.info("Popup de boas-vindas fechado com sucesso.")
                else:
                    logger.info("Não foi detectado popup de boas-vindas.")
            except Exception as e:
                logger.warning(f"Erro ao tentar fechar popup (não crítico): {str(e)}")
                # Continuamos mesmo se não conseguir fechar o popup
            
            return True
        except Exception as e:
            logger.error(f"Erro durante o login: {str(e)}")
            return False
    
    def navegar_para_categoria(self, categoria):
        """Navega para a página da categoria especificada"""
        try:
            logger.info(f"Navegando para a categoria: {categoria}")
            
            # Aguardar um pouco para garantir que a página esteja totalmente carregada
            time.sleep(3)
            
            # Verificar se existe algum popup ou overlay e tentar fechar
            try:
                backdrops = self.driver.find_elements(By.CSS_SELECTOR, "div.q-dialog__backdrop")
                if backdrops:
                    logger.info("Detectado overlay/popup. Tentando fechar...")
                    self.driver.execute_script("arguments[0].click();", backdrops[0])
                    time.sleep(1)
            except Exception as e:
                logger.warning(f"Erro ao tentar fechar overlay (não crítico): {str(e)}")
            
            # Tentar localizar o elemento da categoria de várias maneiras
            elemento_categoria = None
            
            # Método 1: Usando XPath com texto exato
            try:
                xpath = f"//div[contains(@class, 'text-teal-10') and contains(text(), '{categoria}')]"
                elementos = self.driver.find_elements(By.XPATH, xpath)
                if elementos:
                    elemento_categoria = elementos[0]
                    logger.info(f"Elemento da categoria {categoria} encontrado com método 1.")
            except Exception:
                pass
            
            # Método 2: Usando querySelector com JavaScript
            if not elemento_categoria:
                try:
                    script = f"""
                    return Array.from(document.querySelectorAll('div.text-teal-10')).find(el => 
                        el.textContent.includes('{categoria}')
                    );
                    """
                    elemento_categoria = self.driver.execute_script(script)
                    if elemento_categoria:
                        logger.info(f"Elemento da categoria {categoria} encontrado com método 2.")
                except Exception:
                    pass
            
            # Se ainda não encontrou, tenta um método mais genérico
            if not elemento_categoria:
                try:
                    elementos = self.driver.find_elements(By.CSS_SELECTOR, "div.text-teal-10.q-pa-md.text-center")
                    for el in elementos:
                        if categoria.lower() in el.text.lower():
                            elemento_categoria = el
                            logger.info(f"Elemento da categoria {categoria} encontrado com método 3.")
                            break
                except Exception:
                    pass
            
            if not elemento_categoria:
                raise Exception(f"Não foi possível encontrar o elemento da categoria {categoria}")
            
            # Usar JavaScript para clicar no elemento (mais confiável para elementos sobrepostos)
            logger.info(f"Clicando na categoria {categoria} usando JavaScript...")
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento_categoria)
            time.sleep(1)  # Aguardar um pouco após rolagem
            self.driver.execute_script("arguments[0].click();", elemento_categoria)
            
            # Aguardar carregamento dos produtos com retry
            max_tentativas = 3
            for tentativa in range(max_tentativas):
                try:
                    logger.info(f"Aguardando carregamento dos produtos (tentativa {tentativa+1}/{max_tentativas})...")
                    WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "div.q-card.my-card"))
                    )
                    logger.info(f"Navegação para categoria {categoria} realizada com sucesso.")
                    return True
                except TimeoutException:
                    if tentativa < max_tentativas - 1:
                        logger.warning(f"Timeout ao aguardar produtos. Tentando novamente...")
                        time.sleep(2)
                    else:
                        raise
            
            return True
        except Exception as e:
            logger.error(f"Erro ao navegar para a categoria {categoria}: {str(e)}")
            return False
    
    def processar_categorias(self, categorias):
        """Processa todas as categorias para extração de dados"""
        todos_produtos = []
        ultima_categoria_processada = None
        
        for categoria in categorias:
            produtos_da_categoria = []
            logger.info(f"Iniciando processamento da categoria: {categoria}")
            
            try:
                self.navegar_para_categoria(categoria)
                
                pagina = 1
                while True:
                    logger.info(f"Processando página {pagina} da categoria {categoria}...")
                    produtos_da_pagina = self.extrair_produtos_da_pagina(categoria)
                    
                    if produtos_da_pagina:
                        produtos_da_categoria.extend(produtos_da_pagina)
                        logger.info(f"Extraídos {len(produtos_da_pagina)} produtos da página {pagina}.")
                    else:
                        logger.warning(f"Nenhum produto encontrado na página {pagina} da categoria {categoria}.")
                    
                    # Verificar se existe próxima página
                    proxima_pagina_existe = self.ir_para_proxima_pagina()
                    if not proxima_pagina_existe:
                        logger.info(f"Não há mais páginas para a categoria {categoria}.")
                        break
                    
                    pagina += 1
                
                # Atualizar a última categoria processada com sucesso
                ultima_categoria_processada = categoria
                todos_produtos.extend(produtos_da_categoria)
                
            except Exception as e:
                logger.error(f"Erro ao processar categoria {categoria}: {str(e)}")
                logger.error(traceback.format_exc())
                
                # Se o erro for relacionado ao navegador fechado, tentar reiniciar
                if "no such window" in str(e).lower() or "window not found" in str(e).lower():
                    logger.warning("Navegador fechado ou travado. Tentando reiniciar...")
                    try:
                        # Fechar o navegador se ainda estiver aberto
                        try:
                            self.driver.quit()
                        except:
                            pass
                        
                        # Reiniciar o navegador
                        self.inicializar_navegador()
                        self.fazer_login()
                        
                        # Se estava no meio de uma categoria, continuar dela
                        if ultima_categoria_processada != categoria:
                            logger.info(f"Retomando processamento da categoria: {categoria}")
                            # Tentar novamente esta categoria
                            continue
                    except Exception as reinit_error:
                        logger.error(f"Erro ao reiniciar navegador: {str(reinit_error)}")
        
        return todos_produtos
    
    def extrair_produtos_da_pagina(self, categoria):
        """Extrai todos os produtos de uma página"""
        logger.info(f"Extraindo produtos da página atual para a categoria {categoria}...")
        produtos = []
        
        max_tentativas = 3
        tentativa = 1
        
        while tentativa <= max_tentativas:
            try:
                # Primeiro, vamos aguardar que a página carregue completamente
                time.sleep(3)
                
                # Tentar diferentes seletores para encontrar os produtos
                seletores = [
                    'div.q-card.q-hoverable',
                    'div.q-card',
                    'div.my-card',
                    'div.q-card.my-card',
                    'div[class*="card"]'
                ]
                
                cards = []
                for seletor in seletores:
                    logger.info(f"Tentando encontrar produtos com seletor: {seletor}")
                    # Usar JavaScript para obter todos os cards de produtos
                    js_script = f"""
                    return document.querySelectorAll('{seletor}');
                    """
                    cards = self.driver.execute_script(js_script)
                    if len(cards) > 0:
                        logger.info(f"Encontrados {len(cards)} produtos com seletor: {seletor}")
                        break
                
                # Se ainda não encontrou, tentar encontrar pelo conteúdo característico
                if len(cards) == 0:
                    logger.info("Tentando encontrar produtos pela estrutura interna...")
                    js_script = """
                    // Buscar elementos que provavelmente são cards de produtos
                    let potentialCards = [];
                    
                    // Cards geralmente contêm preços
                    document.querySelectorAll('div.text-h6.text-weight-bold.text-teal-9').forEach(priceEl => {
                        let card = priceEl.closest('div.q-card') || priceEl.closest('div[class*="card"]');
                        if (card && !potentialCards.includes(card)) {
                            potentialCards.push(card);
                        }
                    });
                    
                    // Cards também podem conter informações de parcelamento
                    document.querySelectorAll('div.text-caption.text-weight-bold').forEach(infoEl => {
                        let card = infoEl.closest('div.q-card') || infoEl.closest('div[class*="card"]');
                        if (card && !potentialCards.includes(card)) {
                            potentialCards.push(card);
                        }
                    });
                    
                    return potentialCards;
                    """
                    cards = self.driver.execute_script(js_script)
                
                logger.info(f"Encontrados {len(cards)} produtos na página atual.")
                
                # Se não encontrou nenhum card, tentar rolar a página
                if len(cards) == 0 and tentativa < max_tentativas:
                    logger.info(f"Nenhum produto encontrado. Tentando rolar a página (tentativa {tentativa}/{max_tentativas})...")
                    # Rolar para baixo
                    self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight/2);")
                    time.sleep(2)
                    # Rolar mais para baixo
                    self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(2)
                    tentativa += 1
                    continue
                
                # Capturar screenshot para debug se necessário
                if len(cards) == 0:
                    logger.info("Capturando screenshot para análise...")
                    try:
                        screenshot_path = f"screenshot_categoria_{categoria}_pagina.png"
                        self.driver.save_screenshot(screenshot_path)
                        logger.info(f"Screenshot salvo em {screenshot_path}")
                    except Exception as e:
                        logger.error(f"Erro ao capturar screenshot: {str(e)}")
                
                # Extrair dados de cada card
                for i, card in enumerate(cards, 1):
                    logger.info(f"Processando produto {i}/{len(cards)}...")
                    produto = self.extrair_dados_produto(card, categoria)
                    if produto:
                        # Verificar se o nome do produto contém "instalação" (case insensitive)
                        nome_produto = produto.get("Nome do Produto", "").lower()
                        if "instalacao" in nome_produto or "instalação" in nome_produto:
                            logger.info(f"Produto ignorado por conter 'instalação' no nome: {produto.get('Nome do Produto')}")
                        else:
                            produtos.append(produto)
                
                # Se tivemos sucesso, sair do loop
                break
                
            except Exception as e:
                # Se for um erro de "no such window", propagar a exceção para ser tratada no nível superior
                if "no such window" in str(e).lower() or "window not found" in str(e).lower():
                    raise
                    
                logger.error(f"Erro ao extrair produtos da página (tentativa {tentativa}/{max_tentativas}): {str(e)}")
                logger.error(traceback.format_exc())
                
                if tentativa < max_tentativas:
                    logger.info(f"Tentando novamente em 3 segundos...")
                    time.sleep(3)
                    tentativa += 1
                else:
                    logger.error("Número máximo de tentativas atingido. Continuando com próxima etapa.")
                    break
    
        logger.info(f"Extraídos {len(produtos)} produtos da página atual.")
        return produtos
    
    def extrair_dados_produto(self, card, categoria):
        """Extrai os dados de um card de produto"""
        max_tentativas = 3
        tentativa = 1
        
        while tentativa <= max_tentativas:
            try:
                # Aguardar um pouco para garantir que o card esteja totalmente carregado
                time.sleep(0.5)
                
                # Usar JavaScript para extrair os dados mais confiáveis
                script = """
                function getTextOrDefault(card, selector, defaultValue = "N/A") {
                    const el = card.querySelector(selector);
                    return el ? el.textContent.trim() : defaultValue;
                }
                
                function getAttributeOrDefault(card, selector, attribute, defaultValue = "N/A") {
                    const el = card.querySelector(selector);
                    return el ? (el.getAttribute(attribute) || defaultValue) : defaultValue;
                }
                
                function findElementWithText(card, selector, text) {
                    const elements = card.querySelectorAll(selector);
                    for(let el of elements) {
                        if(el.textContent && el.textContent.includes(text)) {
                            return el.textContent.trim();
                        }
                    }
                    return "";
                }
                
                function getPublicImageUrl(privateUrl) {
                    // Extrair o ID da imagem ou caminho da URL privada
                    if (!privateUrl || privateUrl === 'N/A') return 'N/A';
                    
                    try {
                        // Tentar extrair o nome do arquivo da URL
                        const urlObj = new URL(privateUrl);
                        const pathname = urlObj.pathname;
                        const filename = pathname.split('/').pop();
                        
                        // Construir URL pública baseada no domínio de vendas da Leveros
                        return `https://www.vendas.leveros.com.br/upload/produto/imagem/${filename}`;
                    } catch (e) {
                        // Se falhar, retornar a URL original
                        return privateUrl;
                    }
                }
                
                const result = {};
                
                // Nome do produto
                result.nome = getTextOrDefault(arguments[0], 'div.menuItems.text-caption.q-pt-sm.ellipsis-2-lines');
                
                // Voltagem
                result.voltagem = getTextOrDefault(arguments[0], 'div.q-chip--outline');
                
                // Preço principal
                result.precoPrincipal = getTextOrDefault(arguments[0], 'div.text-h6.text-weight-bold.text-teal-9');
                
                // Info de parcelamento
                result.infoParcelamento = getTextOrDefault(arguments[0], 'div.text-caption.text-weight-bold');
                
                // Preço à vista (usando função personalizada para encontrar o elemento com o texto "à vista")
                result.precoVista = findElementWithText(arguments[0], 'div.text-caption', 'à vista');
                
                // Se não encontrou "à vista", pega qualquer text-caption como fallback
                if (!result.precoVista) {
                    result.precoVista = getTextOrDefault(arguments[0], 'div.text-caption');
                }
                
                // URL da imagem privada (área logada)
                result.urlImagem = getAttributeOrDefault(arguments[0], 'div.q-img img', 'src');
                
                // URL pública da imagem
                result.urlImagemPublica = getPublicImageUrl(result.urlImagem);
                
                return result;
                """
                
                # Executar o script JavaScript
                resultado = self.driver.execute_script(script, card)
                
                # Processar informações de parcelamento
                info_parcelamento = resultado.get('infoParcelamento', 'N/A')
                
                # Extrair quantidade de parcelas e valor da parcela
                qtd_parcelas = "N/A"
                valor_parcela = "N/A"
                
                if 'de' in info_parcelamento:
                    partes = info_parcelamento.split('de')
                    qtd_parcelas = partes[0].strip()
                    valor_parcela = partes[1].strip()
                
                # Criar dicionário com os dados do produto
                produto = {
                    "Categoria": categoria,
                    "Nome do Produto": resultado.get('nome', 'N/A'),
                    "Voltagem": resultado.get('voltagem', 'N/A'),
                    "Preço Principal": resultado.get('precoPrincipal', 'N/A'),
                    "Preço à Vista": resultado.get('precoVista', 'N/A'),
                    "Qtd. Parcelas": qtd_parcelas,
                    "Valor Parcela": valor_parcela,
                    "URL da Imagem": resultado.get('urlImagem', 'N/A'),
                    "URL Pública da Imagem": resultado.get('urlImagemPublica', 'N/A')
                }
                
                logger.info(f"Produto extraído: {resultado.get('nome', 'N/A')}")
                return produto
                
            except Exception as e:
                # Se for um erro de "no such window", propagar a exceção para ser tratada no nível superior
                if "no such window" in str(e).lower() or "window not found" in str(e).lower():
                    raise
                    
                logger.error(f"Erro ao extrair dados do produto (tentativa {tentativa}/{max_tentativas}): {str(e)}")
                logger.error(traceback.format_exc())
                
                if tentativa < max_tentativas:
                    logger.info(f"Tentando novamente em 2 segundos...")
                    time.sleep(2)
                    tentativa += 1
                else:
                    logger.error("Número máximo de tentativas atingido. Retornando None.")
                    return None
    
    def ir_para_proxima_pagina(self):
        """Verifica se existe um botão de próxima página e clica nele se estiver disponível"""
        try:
            logger.info("Verificando se existe próxima página...")
            
            # Usar JavaScript para verificar e clicar no botão de próxima página
            script_check = """
            const buttons = Array.from(document.querySelectorAll('button'));
            // Procura botões com ícone 'fast_forward' ou com texto contendo 'próxima'
            const nextButton = buttons.find(btn => {
                const icon = btn.querySelector('i.material-icons');
                return (icon && icon.textContent.includes('fast_forward')) || 
                       btn.textContent.toLowerCase().includes('próxima');
            });
            
            if (nextButton && !nextButton.disabled) {
                return nextButton;
            }
            return null;
            """
            
            botao_proxima = self.driver.execute_script(script_check)
            
            if botao_proxima:
                logger.info("Botão de próxima página encontrado. Clicando...")
                # Scrollar para o botão e clicar
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_proxima)
                time.sleep(1)
                self.driver.execute_script("arguments[0].click();", botao_proxima)
                time.sleep(3)  # Aguardar carregamento da próxima página
                return True
            else:
                logger.info("Não há mais páginas disponíveis.")
                return False
                
        except Exception as e:
            logger.error(f"Erro ao verificar próxima página: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def salvar_dados_excel(self):
        """Salva os dados extraídos em um arquivo Excel formatado"""
        try:
            if not self.dados_produtos:
                logger.warning("Não há dados para salvar!")
                return False
            
            logger.info(f"Salvando {len(self.dados_produtos)} produtos no Excel...")
            
            # Cria um DataFrame com os dados coletados
            df = pd.DataFrame(self.dados_produtos)
            
            # Cria um ExcelWriter para formatar o arquivo
            writer = pd.ExcelWriter(self.arquivo_excel, engine='xlsxwriter')
            
            # Escreve a planilha principal com todos os produtos
            df.to_excel(writer, sheet_name='Produtos Leveros', index=False)
            
            # Cria uma planilha para cada categoria
            for categoria in self.categorias:
                df_categoria = df[df['Categoria'] == categoria]
                if not df_categoria.empty:
                    # Limita o nome da planilha a 31 caracteres (limite do Excel)
                    df_categoria.to_excel(writer, sheet_name=categoria[:31], index=False)
            
            # Cria uma planilha de resumo
            resumo = df.groupby('Categoria').agg({
                'Nome do Produto': 'count',
            }).reset_index()
            resumo.columns = ['Categoria', 'Quantidade de Produtos']
            resumo.to_excel(writer, sheet_name='Resumo', index=False)
            
            # Formato para a planilha principal
            workbook = writer.book
            worksheet = writer.sheets['Produtos Leveros']
            
            # Formato para cabeçalhos
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#4F6228',
                'font_color': 'white',
                'border': 1
            })
            
            # Aplica o formato nos cabeçalhos
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Ajusta a largura das colunas
            worksheet.set_column('A:A', 15)  # Categoria
            worksheet.set_column('B:B', 40)  # Nome do Produto
            worksheet.set_column('C:C', 10)  # Voltagem
            worksheet.set_column('D:E', 15)  # Preços
            worksheet.set_column('F:G', 15)  # Parcelas
            worksheet.set_column('H:H', 40)  # URL da Imagem
            worksheet.set_column('I:I', 40)  # URL Pública da Imagem
            
            # Adiciona filtros automáticos
            worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
            
            # Salva o arquivo
            writer.close()
            
            logger.info(f"Dados salvos com sucesso no arquivo: {self.arquivo_excel}")
            return True
        except Exception as e:
            logger.error(f"Erro ao salvar os dados no Excel: {str(e)}")
            return False
    
    def salvar_dados_pdf(self):
        """
        Salva os dados extraídos em um arquivo PDF.
        """
        logging.info("Salvando dados em PDF...")
        
        agora = datetime.now()
        timestamp = agora.strftime("%Y%m%d_%H%M%S")
        filename = f"ProdutosLeveros_{timestamp}.pdf"
        
        try:
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            
            # Título
            pdf.set_font("Arial", "B", 16)
            pdf.cell(0, 10, txt="Relatório de Produtos Leveros", ln=True, align="C")
            pdf.ln(5)
            
            # Data de geração
            pdf.set_font("Arial", "", 12)
            pdf.cell(0, 8, txt=f"Data de geração: {agora.strftime('%d/%m/%Y %H:%M')}", ln=True)
            
            # Total de produtos
            total_produtos = len(self.dados_produtos)
            pdf.cell(0, 8, txt=f"Total de produtos: {total_produtos}", ln=True)
            
            pdf.ln(10)
            
            # Página de resumo - contagem por categoria
            pdf.set_font("Arial", "B", 14)
            pdf.cell(0, 10, txt="Resumo por Categoria", ln=True)
            pdf.ln(5)
            
            pdf.set_font("Arial", "", 12)
            for categoria in self.categorias:
                produtos_categoria = [p for p in self.dados_produtos if p['Categoria'] == categoria]
                pdf.cell(0, 8, txt=f"{categoria}: {len(produtos_categoria)} produtos", ln=True)
            
            # Detalhes dos produtos por categoria
            for categoria in self.categorias:
                produtos_categoria = [p for p in self.dados_produtos if p['Categoria'] == categoria]
                if produtos_categoria:
                    pdf.add_page()
                    pdf.set_font("Arial", "B", 14)
                    pdf.cell(0, 10, txt=f"Categoria: {categoria}", ln=True)
                    pdf.ln(5)
                    
                    for idx, produto in enumerate(produtos_categoria, 1):
                        pdf.set_font("Arial", "B", 12)
                        pdf.cell(0, 8, txt=f"Produto {idx}: {produto['Nome do Produto']}", ln=True)
                        
                        pdf.set_font("Arial", "", 10)
                        pdf.cell(0, 8, txt=f"Categoria: {produto['Categoria']}", ln=True)
                        if 'Modelo' in produto:
                            pdf.cell(0, 8, txt=f"Modelo: {produto['Modelo']}", ln=True)
                        
                        # Adicionar link da imagem com texto amigável
                        img_url = produto.get('URL Pública da Imagem', '')
                        if img_url and img_url != 'N/A':
                            # Texto explicativo acima do link
                            pdf.set_font("Arial", "", 10)
                            pdf.set_text_color(0, 0, 0)  # Preto
                            pdf.cell(0, 8, txt="Segue link da foto do produto:", ln=True)
                            
                            # Texto do link
                            link_text = "LINK PARA FOTO DO PRODUTO (Clique para visualizar)"
                            
                            # Configurar fonte e cor para o link
                            pdf.set_font("Arial", "BU", 10)  # Negrito e Sublinhado
                            pdf.set_text_color(0, 0, 255)   # Azul
                            
                            # Obter largura do texto para calcular posição X
                            link_width = pdf.get_string_width(link_text)
                            
                            # Adicionar o texto com link
                            pdf.cell(link_width, 8, txt=link_text, link=img_url)
                            pdf.ln()
                            
                            # Restaurar fonte e cor
                            pdf.set_text_color(0, 0, 0)  # Preto
                            pdf.set_font("Arial", "", 10)  # Normal
                        else:
                            pdf.cell(0, 8, txt="Foto do produto: Não disponível", ln=True)
                        
                        # Adicionar outras informações do produto
                        for chave, valor in produto.items():
                            if chave not in ['Nome do Produto', 'Categoria', 'Modelo', 'URL da Imagem', 'URL Pública da Imagem']:
                                pdf.cell(0, 8, txt=f"{chave}: {valor}", ln=True)
                        
                        pdf.ln(8)  # Espaço entre produtos
                        
                        # Adicionar uma linha separadora entre produtos
                        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
                        pdf.ln(8)
            
            # Salvar o PDF
            pdf_path = os.path.join(os.getcwd(), filename)
            pdf.output(pdf_path)
            logging.info(f"Arquivo PDF salvo com sucesso: {pdf_path}")
            
        except Exception as e:
            logging.error(f"Erro ao salvar PDF: {str(e)}")
    
    def executar(self):
        """Executa o fluxo completo do RPA"""
        try:
            logger.info("Iniciando execução do RPA Leveros Integra...")
            
            # Inicializa o navegador
            if not self.inicializar_navegador():
                logger.error("Não foi possível inicializar o navegador. Abortando execução.")
                return False
            
            # Faz login no sistema
            if not self.fazer_login():
                logger.error("Não foi possível realizar o login. Abortando execução.")
                self.finalizar()
                return False
            
            # Processa cada categoria
            self.dados_produtos = self.processar_categorias(self.categorias)
            
            # Salva os dados no Excel
            self.salvar_dados_excel()
            
            # Salva os dados no PDF
            self.salvar_dados_pdf()
            
            # Finaliza a execução
            self.finalizar()
            
            logger.info("Execução do RPA concluída com sucesso!")
            return True
        except Exception as e:
            logger.error(f"Erro durante a execução do RPA: {str(e)}")
            self.finalizar()
            return False
    
    def finalizar(self):
        """Finaliza o navegador e libera recursos"""
        try:
            if self.driver:
                logger.info("Finalizando navegador...")
                self.driver.quit()
                logger.info("Navegador finalizado.")
        except Exception as e:
            logger.error(f"Erro ao finalizar navegador: {str(e)}")


if __name__ == "__main__":
    # Verifica argumentos de linha de comando
    import sys
    headless_mode = "--headless" in sys.argv
    
    # Executa o RPA
    logger.info(f"Iniciando RPA em modo {'headless' if headless_mode else 'normal'}")
    rpa = LeverosRPA(headless=headless_mode)
    rpa.executar()
