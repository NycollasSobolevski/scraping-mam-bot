import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from datetime import date
import xlsxwriter
 
from environment import UserInfo, TicketData

login = input("Insira o nome de usuário:\n> ")
senha = input("Insira a senha:\n> ")
startDate = input("Digite a data de inicio (dd.mm.aaaa): ")
startDate = startDate.split('.')
startDate = date(int(startDate[2]), int(startDate[1]), int(startDate[0]))
endDate   = input("Digite a data de final  (dd.mm.aaaa): ")
endDate   = endDate.split('.')
endDate   = date(int(endDate[2]), int(endDate[1]), int(endDate[0]))


edge_options = webdriver.EdgeOptions()
edge_options.add_argument('--headless')
driver = webdriver.Edge(edge_options)
# url = "https://dev.maminfo.com.br:62443/mamapp/login"
url = "https://mamerp.maminfo.com.br:62443/mamapp/login"

driver.get(url)

inputLoginPath = "/html/body/div[1]/div[2]/main/div/div/section/div/form/div[2]/div/input"
inputPassPath  = "/html/body/div[1]/div[2]/main/div/div/section/div/form/div[3]/div/input"

element = driver.find_element(by=By.XPATH, value=inputLoginPath)
element.send_keys(login)

element = driver.find_element(by=By.XPATH, value=inputPassPath)
element.send_keys(senha)

element.submit()

# Indo para ver todos os Chamados
seeAllPath = "/html/body/div[1]/div[2]/div/main/div/section[1]/div[1]/div[1]/div[2]/div/div/div[5]/div/a[2]"
element = driver.find_element(by=By.XPATH, value=seeAllPath).click()

allTicketsPath = "/html/body/div[1]/div[2]/div/main/div/section[1]/div[2]/div[2]"
allTickets = driver.find_element(by=By.XPATH, value=allTicketsPath).find_elements(by=By.TAG_NAME, value='a')
selectedTickets = []

index = 1
for ticket in allTickets:
    ticketDateSplited = ticket.text.split("\n")[0].split(" ")[1].split('-')
    ticketDate = date(int(ticketDateSplited[0]), int(ticketDateSplited[1]), int(ticketDateSplited[2]))
    if(ticketDate.month >= startDate.month and ticketDate.day >= startDate.day and ticketDate.year >= startDate.year and ticketDate.month <= endDate.month and ticketDate.day <= endDate.day and ticketDate.year <= endDate.year):
        selectedTickets.append(f"{allTicketsPath}/a[{index}]")
    index += 1

# criando planilha
workbook = xlsxwriter.Workbook('chamados.xlsx')
worksheet = workbook.add_worksheet("Chamados")

sheetColumns = {
    "Numero": "A",
    "Senha": "C",
    "Cliente": "D",
    "Inicio": "J",
    "Inicio Hora": "L",
    "Endereco": "M",
}
#get adress
worksheet.write(f"{sheetColumns['Numero']}1", "Numero")
worksheet.write(f"{sheetColumns['Senha']}1", "Senha")
worksheet.write(f"{sheetColumns['Cliente']}1", "Cliente")
worksheet.write(f"{sheetColumns['Inicio']}1", "Inicio")
worksheet.write(f"{sheetColumns['Endereco']}1", "Endereço")

excellIndex = 2

for i in range(len(selectedTickets)):
    print(selectedTickets[i])
    try:
        ticket = driver.find_element(by=By.XPATH, value=selectedTickets[i]).click()
        time.sleep(1)

        ticketNumber        = driver.find_element(by=By.XPATH, value="/html/body/div[1]/div[2]/div/main/div/section[1]/div[2]/div[2]/div/div[3]/div/p[1]/strong").text
        ticketPass          = driver.find_element(by=By.XPATH, value="/html/body/div[1]/div[2]/div/main/div/section[1]/div[2]/div[2]/div/div[3]/div/p[15]/strong").text
        ticketStartDateTime = driver.find_element(by=By.XPATH, value="/html/body/div[1]/div[2]/div/main/div/section[1]/div[2]/div[2]/div/div[3]/div/p[11]/strong").text
        ticketClient        = driver.find_element(by=By.XPATH, value="/html/body/div[1]/div[2]/div/main/div/section[1]/div[2]/div[2]/div/div[3]/div/p[4]/strong").text
        ticketAdress        = driver.find_element(by=By.XPATH, value="/html/body/div[1]/div[2]/div/main/div/section[1]/div[2]/div[2]/div/div[3]/div/p[5]/strong").text
        ticketData = TicketData(ticketNumber, ticketPass, ticketClient, ticketStartDateTime, ticketAdress)
        print(ticketData.Number)
        worksheet.write(f"{sheetColumns['Numero']}{excellIndex}", ticketData.Number)
        worksheet.write(f"{sheetColumns['Senha']}{excellIndex}", ticketData.Pass)
        worksheet.write(f"{sheetColumns['Cliente']}{excellIndex}", ticketData.Client)
        worksheet.write(f"{sheetColumns['Inicio Hora']}{excellIndex}", ticketData.getStringTime())
        worksheet.write(f"{sheetColumns['Inicio']}{excellIndex}", ticketData.getStringDate())
        worksheet.write(f"{sheetColumns['Endereco']}{excellIndex}", ticketData.Adress)
        
        print("\n ------------------------------------------------------")
        print(f"Numero: {ticketData.Number}")
        print(f"Senha: {ticketData.Pass}")
        print(f"Cliente: {ticketData.Client}")
        print(f"Inicio: {ticketData.getStringDate()}")
        print(f"Endereco: {ticketData.Adress}")
        print("------------------------------------------------------\n")
        

        driver.back()
    except Exception as e:
        print(e)
        driver.back()
    excellIndex += 1
    time.sleep(1)
    
    
workbook.close()