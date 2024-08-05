from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import pyodbc

driver = webdriver.Firefox()
driver.get("http://vitibrasil.cnpuv.embrapa.br/index.php?opcao=opt_02")
wait = WebDriverWait(driver, 60)
table = wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'table.tb_base:nth-child(2)')))

# Configurações de conexão
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\User\PycharmProjects\WebScrapy\base.accdb;' # substitua pelo caminho do seu banco de dados
    )


cnxn = pyodbc.connect(conn_str)
cursor = cnxn.cursor()

rows = table.find_elements(By.TAG_NAME, "tr")

# Crie uma lista para armazenar todos os dados da tabela
table_data = []

# Adicione cada linha da tabela à lista de dados da tabela
for row in rows:
    cols = row.find_elements(By.TAG_NAME, "td")  # note que isso muda para "th" para cabeçalhos de tabela
    row_data = [col.text for col in cols]
    table_data.append(row_data)

# Agora, 'table_data' é uma lista de listas que representa a tabela inteira
table_data.pop(0)
print(table_data)
for row in table_data:
    # Inserir dados na tabela
    cursor.execute('''
                INSERT INTO Producao (Produto,Quantidade) 
                VALUES (?, ?)
                ''', row)

# Commit nas alterações e fechar a conexão
cnxn.commit()


# Feche o navegador

button = driver.find_element(By.CSS_SELECTOR, 'button.btn_opt:nth-child(4)')  # substitua 'buttonId' pelo ID correto do botão
# Clique no botão
button.click()
table = wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'table.tb_base:nth-child(2)')))

rows = table.find_elements(By.TAG_NAME, "tr")

# Crie uma lista para armazenar todos os dados da tabela
table_data = []

# Adicione cada linha da tabela à lista de dados da tabela
for row in rows:
    cols = row.find_elements(By.TAG_NAME, "td")  # note que isso muda para "th" para cabeçalhos de tabela
    row_data = [col.text for col in cols]
    table_data.append(row_data)

# Agora, 'table_data' é uma lista de listas que representa a tabela inteira
table_data.pop(0)
print(table_data)
for row in table_data:
    # Inserir dados na tabela
    cursor.execute('''
                INSERT INTO Comercializacao (Produto,Quantidade) 
                VALUES (?, ?)
                ''', row)

# Commit nas alterações e fechar a conexão
cnxn.commit()




button = driver.find_element(By.CSS_SELECTOR, 'button.btn_opt:nth-child(3)')  # substitua 'buttonId' pelo ID correto do botão
# Clique no botão
button.click()
table = wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'table.tb_base:nth-child(2)')))

rows = table.find_elements(By.TAG_NAME, "tr")

# Crie uma lista para armazenar todos os dados da tabela
table_data = []

# Adicione cada linha da tabela à lista de dados da tabela
for row in rows:
    cols = row.find_elements(By.TAG_NAME, "td")  # note que isso muda para "th" para cabeçalhos de tabela
    row_data = [col.text for col in cols]
    table_data.append(row_data)

# Agora, 'table_data' é uma lista de listas que representa a tabela inteira
table_data.pop(0)
print(table_data)
for row in table_data:
    # Inserir dados na tabela
    cursor.execute('''
                INSERT INTO ProcViniferas (Cultivar,Quantidade) 
                VALUES (?, ?)
                ''', row)

button = driver.find_element(By.CSS_SELECTOR, 'button.btn_sopt:nth-child(3)')  # substitua 'buttonId' pelo ID correto do botão
# Clique no botão
button.click()
table = wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'table.tb_base:nth-child(2)')))

rows = table.find_elements(By.TAG_NAME, "tr")

# Crie uma lista para armazenar todos os dados da tabela
table_data = []

# Adicione cada linha da tabela à lista de dados da tabela
for row in rows:
    cols = row.find_elements(By.TAG_NAME, "td")  # note que isso muda para "th" para cabeçalhos de tabela
    row_data = [col.text for col in cols]
    table_data.append(row_data)

# Agora, 'table_data' é uma lista de listas que representa a tabela inteira
table_data.pop(0)
print(table_data)
for row in table_data:
    # Inserir dados na tabela
    cursor.execute('''
                INSERT INTO ProcAmericanasHibridas (Cultivar,Quantidade) 
                VALUES (?, ?)
                ''', row)
# Commit nas alterações e fechar a conexão

button = driver.find_element(By.CSS_SELECTOR, 'button.btn_sopt:nth-child(5)')  # substitua 'buttonId' pelo ID correto do botão
# Clique no botão
button.click()
table = wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'table.tb_base:nth-child(2)')))

rows = table.find_elements(By.TAG_NAME, "tr")

# Crie uma lista para armazenar todos os dados da tabela
table_data = []

# Adicione cada linha da tabela à lista de dados da tabela
for row in rows:
    cols = row.find_elements(By.TAG_NAME, "td")  # note que isso muda para "th" para cabeçalhos de tabela
    row_data = [col.text for col in cols]
    table_data.append(row_data)

# Agora, 'table_data' é uma lista de listas que representa a tabela inteira
table_data.pop(0)
print(table_data)
for row in table_data:
    # Inserir dados na tabela
    cursor.execute('''
                INSERT INTO ProcUvasDeMesa (Cultivar,Quantidade) 
                VALUES (?, ?)
                ''', row)
# Commit nas alterações e fechar a conexão

button = driver.find_element(By.CSS_SELECTOR, 'button.btn_opt:nth-child(5)')  # substitua 'buttonId' pelo ID correto do botão
# Clique no botão
button.click()
table = wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'table.tb_base:nth-child(2)')))

rows = table.find_elements(By.TAG_NAME, "tr")

# Crie uma lista para armazenar todos os dados da tabela
table_data = []

# Adicione cada linha da tabela à lista de dados da tabela
for row in rows:
    cols = row.find_elements(By.TAG_NAME, "td")  # note que isso muda para "th" para cabeçalhos de tabela
    row_data = [col.text for col in cols]
    table_data.append(row_data)

# Agora, 'table_data' é uma lista de listas que representa a tabela inteira
table_data.pop(0)
print(table_data)
for row in table_data:
    # Inserir dados na tabela
    cursor.execute('''
                INSERT INTO ImportVinhosDeMesa (Paises,Quantidade,Valor) 
                VALUES (?, ?, ?)
                ''', row)
# Commit nas alterações e fechar a conexão

button = driver.find_element(By.CSS_SELECTOR, 'button.btn_sopt:nth-child(3)')  # substitua 'buttonId' pelo ID correto do botão
# Clique no botão
button.click()
table = wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'table.tb_base:nth-child(2)')))

rows = table.find_elements(By.TAG_NAME, "tr")

# Crie uma lista para armazenar todos os dados da tabela
table_data = []

# Adicione cada linha da tabela à lista de dados da tabela
for row in rows:
    cols = row.find_elements(By.TAG_NAME, "td")  # note que isso muda para "th" para cabeçalhos de tabela
    row_data = [col.text for col in cols]
    table_data.append(row_data)

# Agora, 'table_data' é uma lista de listas que representa a tabela inteira
table_data.pop(0)
print(table_data)
for row in table_data:
    # Inserir dados na tabela
    cursor.execute('''
                INSERT INTO ImportEspumantes (Paises,Quantidade,Valor) 
                VALUES (?, ?, ?)
                ''', row)
# Commit nas alterações e fechar a conexão

button = driver.find_element(By.CSS_SELECTOR, 'button.btn_sopt:nth-child(5)')  # substitua 'buttonId' pelo ID correto do botão
# Clique no botão
button.click()
table = wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'table.tb_base:nth-child(2)')))

rows = table.find_elements(By.TAG_NAME, "tr")

# Crie uma lista para armazenar todos os dados da tabela
table_data = []

# Adicione cada linha da tabela à lista de dados da tabela
for row in rows:
    cols = row.find_elements(By.TAG_NAME, "td")  # note que isso muda para "th" para cabeçalhos de tabela
    row_data = [col.text for col in cols]
    table_data.append(row_data)

# Agora, 'table_data' é uma lista de listas que representa a tabela inteira
table_data.pop(0)
print(table_data)
for row in table_data:
    # Inserir dados na tabela
    cursor.execute('''
                INSERT INTO ImportUvasFrescas (Paises,Quantidade,Valor) 
                VALUES (?, ?, ?)
                ''', row)
# Commit nas alterações e fechar a conexão
button = driver.find_element(By.CSS_SELECTOR, 'button.btn_sopt:nth-child(7)')  # substitua 'buttonId' pelo ID correto do botão
# Clique no botão
button.click()
table = wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'table.tb_base:nth-child(2)')))

rows = table.find_elements(By.TAG_NAME, "tr")

# Crie uma lista para armazenar todos os dados da tabela
table_data = []

# Adicione cada linha da tabela à lista de dados da tabela
for row in rows:
    cols = row.find_elements(By.TAG_NAME, "td")  # note que isso muda para "th" para cabeçalhos de tabela
    row_data = [col.text for col in cols]
    table_data.append(row_data)

# Agora, 'table_data' é uma lista de listas que representa a tabela inteira
table_data.pop(0)
print(table_data)
for row in table_data:
    # Inserir dados na tabela
    cursor.execute('''
                INSERT INTO ImportUvasPassas (Paises,Quantidade,Valor) 
                VALUES (?, ?, ?)
                ''', row)
# Commit nas alterações e fechar a conexão

button = driver.find_element(By.CSS_SELECTOR, 'button.btn_sopt:nth-child(9)')  # substitua 'buttonId' pelo ID correto do botão
# Clique no botão
button.click()
table = wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'table.tb_base:nth-child(2)')))

rows = table.find_elements(By.TAG_NAME, "tr")

# Crie uma lista para armazenar todos os dados da tabela
table_data = []

# Adicione cada linha da tabela à lista de dados da tabela
for row in rows:
    cols = row.find_elements(By.TAG_NAME, "td")  # note que isso muda para "th" para cabeçalhos de tabela
    row_data = [col.text for col in cols]
    table_data.append(row_data)

# Agora, 'table_data' é uma lista de listas que representa a tabela inteira
table_data.pop(0)
print(table_data)
for row in table_data:
    # Inserir dados na tabela
    cursor.execute('''
                INSERT INTO ImportSucoDeUva (Paises,Quantidade,Valor) 
                VALUES (?, ?, ?)
                ''', row)


button = driver.find_element(By.CSS_SELECTOR, 'button.btn_opt:nth-child(6)')  # substitua 'buttonId' pelo ID correto do botão
# Clique no botão
button.click()
table = wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'table.tb_base:nth-child(2)')))

rows = table.find_elements(By.TAG_NAME, "tr")

# Crie uma lista para armazenar todos os dados da tabela
table_data = []

# Adicione cada linha da tabela à lista de dados da tabela
for row in rows:
    cols = row.find_elements(By.TAG_NAME, "td")  # note que isso muda para "th" para cabeçalhos de tabela
    row_data = [col.text for col in cols]
    table_data.append(row_data)

# Agora, 'table_data' é uma lista de listas que representa a tabela inteira
table_data.pop(0)
print(table_data)
for row in table_data:
    # Inserir dados na tabela
    cursor.execute('''
                INSERT INTO ExportVinhosDeMesa (Paises,Quantidade,Valor) 
                VALUES (?, ?, ?)
                ''', row)

button = driver.find_element(By.CSS_SELECTOR, 'button.btn_sopt:nth-child(3)')  # substitua 'buttonId' pelo ID correto do botão
# Clique no botão
button.click()
table = wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'table.tb_base:nth-child(2)')))

rows = table.find_elements(By.TAG_NAME, "tr")

# Crie uma lista para armazenar todos os dados da tabela
table_data = []

# Adicione cada linha da tabela à lista de dados da tabela
for row in rows:
    cols = row.find_elements(By.TAG_NAME, "td")  # note que isso muda para "th" para cabeçalhos de tabela
    row_data = [col.text for col in cols]
    table_data.append(row_data)

# Agora, 'table_data' é uma lista de listas que representa a tabela inteira
table_data.pop(0)
print(table_data)
for row in table_data:
    # Inserir dados na tabela
    cursor.execute('''
                INSERT INTO ExportEspumantes (Paises,Quantidade,Valor) 
                VALUES (?, ?, ?)
                ''', row)

button = driver.find_element(By.CSS_SELECTOR, 'button.btn_sopt:nth-child(5)')  # substitua 'buttonId' pelo ID correto do botão
# Clique no botão
button.click()
table = wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'table.tb_base:nth-child(2)')))

rows = table.find_elements(By.TAG_NAME, "tr")

# Crie uma lista para armazenar todos os dados da tabela
table_data = []

# Adicione cada linha da tabela à lista de dados da tabela
for row in rows:
    cols = row.find_elements(By.TAG_NAME, "td")  # note que isso muda para "th" para cabeçalhos de tabela
    row_data = [col.text for col in cols]
    table_data.append(row_data)

# Agora, 'table_data' é uma lista de listas que representa a tabela inteira
table_data.pop(0)
print(table_data)
for row in table_data:
    # Inserir dados na tabela
    cursor.execute('''
                INSERT INTO ExportUvasFrescas (Paises,Quantidade,Valor) 
                VALUES (?, ?, ?)
                ''', row)

button = driver.find_element(By.CSS_SELECTOR, 'button.btn_sopt:nth-child(7)')  # substitua 'buttonId' pelo ID correto do botão
# Clique no botão
button.click()
table = wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'table.tb_base:nth-child(2)')))

rows = table.find_elements(By.TAG_NAME, "tr")

# Crie uma lista para armazenar todos os dados da tabela
table_data = []

# Adicione cada linha da tabela à lista de dados da tabela
for row in rows:
    cols = row.find_elements(By.TAG_NAME, "td")  # note que isso muda para "th" para cabeçalhos de tabela
    row_data = [col.text for col in cols]
    table_data.append(row_data)

# Agora, 'table_data' é uma lista de listas que representa a tabela inteira
table_data.pop(0)
print(table_data)
for row in table_data:
    # Inserir dados na tabela
    cursor.execute('''
                INSERT INTO ExportSucoDeUva (Paises,Quantidade,Valor) 
                VALUES (?, ?, ?)
                ''', row)
cnxn.commit()
cnxn.close()
driver.quit()
