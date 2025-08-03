#IMPORTAÇÕES
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import pandas as pd

servico = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=servico)

#importar base de clientes
tabela = pd.read_excel(r'C:\Users\smart\Desktop\emissaoNF\Arquivos\NotasEmitir.xlsx')
print(tabela)

#Entrar na pagina de Login
driver.get(r'C:\Users\smart\Desktop\emissaoNF\Arquivos\login.html')

#Preencher o LOGIN e SENHA
driver.find_element(By.XPATH, '/html/body/div/form/input[1]').send_keys('sergio.bfj')
driver.find_element(By.XPATH, '/html/body/div/form/input[2]').send_keys('sergio123')

#Clicar no botão de fazer LOGIN
driver.find_element(By.XPATH, '/html/body/div/form/button').click()


#Preencher os dados da NF
for linha in tabela.index:
    # nome/razao social
    driver.find_element(By.NAME, 'nome').send_keys(tabela.loc[linha, 'Cliente'])
    # endereco
    driver.find_element(By.NAME, 'endereco').send_keys(tabela.loc[linha, 'Endereço'])
    # bairro
    driver.find_element(By.NAME, 'bairro').send_keys(tabela.loc[linha, 'Bairro'])
    # municipio
    driver.find_element(By.NAME, 'municipio').send_keys(tabela.loc[linha, 'Municipio'])
    # cep
    driver.find_element(By.NAME, 'cep').send_keys(str(tabela.loc[linha, 'CEP']))
    # uf
    driver.find_element(By.NAME, 'uf').send_keys(tabela.loc[linha, 'UF'])
    # cpf/cnpj
    driver.find_element(By.NAME, 'cnpj').send_keys(str(tabela.loc[linha, 'CPF/CNPJ']))
    # inscriçao estadual
    driver.find_element(By.NAME, 'inscricao').send_keys(str(tabela.loc[linha, 'Inscricao Estadual']))
    # descriçao
    driver.find_element(By.NAME, 'descricao').send_keys(tabela.loc[linha, 'Descrição'])
    # quantidade
    driver.find_element(By.NAME, 'quantidade').send_keys(str(tabela.loc[linha, 'Quantidade']))
    # valor uni
    driver.find_element(By.NAME, 'valor_unitario').send_keys(str(tabela.loc[linha, 'Valor Unitario']))
    # valor total
    driver.find_element(By.NAME, 'total').send_keys(str(tabela.loc[linha, 'Valor Total']))

    #Recarregar pagina
    driver.refresh()

    #Clicar em EMITIR
    driver.find_element(By.CLASS_NAME, 'registerbtn').click()

