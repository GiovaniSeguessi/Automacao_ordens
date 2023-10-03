import pandas as pd
import pyautogui
import time
import datetime
from pywinauto import application
from PIL import Image, ImageGrab
import threading

encerrar_programa = False

# Função que localiza a imagem e retorna sua coordenada central
def locate_image(nome_imagem, confianca):
    try:
        centro = pyautogui.locateCenterOnScreen(nome_imagem,confidence=confianca)
        if centro is not None:
            return centro
        else:
            print(f'não encontrou a imagem {nome_imagem} e parou na linha {contador_linhas}')
            exit()
    except Exception as e:
        print(f'Erro do tipo : {str(e)}')

# Função que localiza a imagem exclamacao ou interrogacao, captura tela e muda a váriavel encerrar_programa
def verificar_imagem():
    global encerrar_programa
    while True:
        if pyautogui.locateOnScreen('interroga1.PNG', confidence=0.7):
            print("Imagem interrogacao encontrada!")
            encerrar_programa = True
            screenshot = ImageGrab.grab()
            screenshot.save("screenshot.png", "PNG")
            exit()

# Começa a rodar duas threads, fica verificando a imagem enquanto o programa roda
thread = threading.Thread(target=verificar_imagem)
thread.daemon = True
thread.start()

# Inicia as váriaveis de tempo
hoje = datetime.date.today()
data_sete_dias = hoje + datetime.timedelta(days=7)    
    
hoje_f = hoje.strftime('%d/%m/%Y')         
data_sete_dias_f = data_sete_dias.strftime('%d/%m/%Y')

# Inicia o contador de linhas
contador_linhas = 0

# Iniciar o PIMS
executable_path = r'I:\XXXX\XXX\XXXXX.exe'
app = application.Application()

try:
    app.start(executable_path)
except Exception as e:
    print('Erro ao iniciar o Programa')
    exit()

# Abre o planjamento de atividade e geração de os
time.sleep(10) 
pyautogui.press('down')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.write('giovanis')
time.sleep(3)
pyautogui.press('tab')
pyautogui.write('134749')
pyautogui.press('tab')
pyautogui.press('down')
pyautogui.press('tab')
pyautogui.press('enter') 
time.sleep(10)
planejamento_atividades_recursos = locate_image('planejamento_atividades_recursos.PNG',confianca=0.7) 
pyautogui.moveTo(planejamento_atividades_recursos)
pyautogui.doubleClick()
time.sleep(2)
planejamento_atvidades = locate_image('planejamento_atividades.PNG',confianca=0.7)
pyautogui.moveTo(planejamento_atvidades)
pyautogui.doubleClick()
time.sleep(3)
pyautogui.press('left')
pyautogui.press('enter')
time.sleep(45)
processamento = locate_image('processamento.PNG',confianca=0.7)
pyautogui.moveTo(processamento)
pyautogui.click()
time.sleep(1)
pyautogui.press('down')
pyautogui.press('down')
pyautogui.press('down')
pyautogui.press('right')
pyautogui.press('enter')
time.sleep(10)

# Confirma se está na página de geração de OS Avulsa
confirma_os = pyautogui.locateOnScreen('geracao_os.PNG', confidence=0.7)
if confirma_os is not None:
    pass
else:
    print('Não chegou na geração de OS')
    exit()

# Caminho para o arquivo Excel
excel_file_path = excel_file_path = r'C:\Users\XXXX.xlsx'
sheet_name = "Planilha1"

# Carregar o arquivo Excel em um DataFrame
df = pd.read_excel(excel_file_path,sheet_name=sheet_name)

# Insumos que são registrados no receituário agronômico
insumos_especiais = [
    '15300098', '15300101', '15300108', '15300109', '15300080', '15300083', '15300097',
    '15300004', '15300002', '15300005', '15300007', '15300009', '15300011', '15300012',
    '15300014', '15300017', '15300018', '15300022', '15300024', '15300026', '15300027',
    '15300029', '15300030', '15300036', '15300037', '15300041', '15300050', '15300053',
    '15300054', '15300058', '15300062', '15300013'
]

# Loop para processar cada linha na planilha
for index, row in df.iterrows():
    contador_linhas +=1
    # Extrair informações do DataFrame
    plano = str(row['Plano'])
    CCusto = str(row['C.Custo'])
    GrupoOpera = str(row['Grupo.Oper'])
    Operacao = str(row['Operacao'])
    Bloco = str(row['Bloco'])
    Prestador = str(row['Prestador de serviço'])
    Responsavel = str(row['Responsável'])
    Regiao = str(row['Região'])
    Quadra = str(row['Quadra'])
    Tipo = str(row['Tipo'])
    Quantos_insumos = int(row['Quantos insumos'])
    
    # Loop para criar uma lista de tuplas com valores do numero do insumo e quantidade
    insumos = []
    for i in range(1,Quantos_insumos+1):
        numero_insumo_col = f'Numero do insumo {i}'
        qtd_insumo_col = f'Qtd Insumo {i}'
        numero_insumo = str(int(row[numero_insumo_col]))
        qtd_insumo = str(row[qtd_insumo_col])
        if numero_insumo and qtd_insumo:
            insumos.append((numero_insumo, qtd_insumo))
    
    # Cria uma lista com os insumos especiais         
    valores_encontrados = []     
    for item in insumos:
        if item[0] in insumos_especiais:
            valores_encontrados.append(item[0])
       
    
    if len(valores_encontrados) == 0:
        # Começa a preencher a ordem
        time.sleep(1)
        pyautogui.write('02/10/2023')
        pyautogui.press('tab')
        pyautogui.write('08/10/2023')
        pyautogui.press('tab') 
        pyautogui.press('tab')
        pyautogui.write(plano)
        pyautogui.press('tab')
        pyautogui.write(CCusto)
        pyautogui.press('tab')
        pyautogui.write(GrupoOpera)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.write(Operacao)
        pyautogui.press('tab')
        time.sleep(3)
        pyautogui.press('tab')
        pyautogui.write(Prestador)
        pyautogui.press('tab')
        pyautogui.write(Responsavel)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.write(Regiao)
        pyautogui.press('tab')
        pyautogui.write(Quadra)
        pyautogui.press('tab')
        pyautogui.write(Bloco)
        time.sleep(1)
        recurso = locate_image('recursos.PNG',confianca=0.7)
        pyautogui.moveTo(recurso)
        pyautogui.click()
        time.sleep(1)
        
        # Loop para adicionar os insumos na ordem
        for numero, qtd in insumos:
            
            estrela = locate_image('estrela.PNG',confianca=0.7)
            pyautogui.moveTo(estrela)
            pyautogui.click()
            pyautogui.write(Tipo)
            pyautogui.press('tab')
            pyautogui.write(numero)
            pyautogui.press('tab')
            pyautogui.write(qtd)
            pyautogui.press('tab')
            pyautogui.write('1')
        
        time.sleep(1)
        local = locate_image('local.PNG',confianca=0.7)
        pyautogui.moveTo(local)
        pyautogui.click()
        time.sleep(1)
        gerar_os = locate_image('gerar_os.PNG',confianca=0.7)
        pyautogui.moveTo(gerar_os)
        pyautogui.click()
        time.sleep(15)
        bloco_notas = locate_image('bloco_notas.PNG',confianca=0.7)
        pyautogui.moveTo(bloco_notas)
        time.sleep(5)
        pyautogui.click()
        time.sleep(15)
        fechar_bloco_notas = locate_image('fechar_bloco_notas.PNG',confianca=0.7)
        pyautogui.moveTo(fechar_bloco_notas)
        pyautogui.click()
        time.sleep(15)
        ok = locate_image('ok.PNG',confianca=0.7)
        pyautogui.moveTo(ok)
        pyautogui.click()
        time.sleep(2)
        sair = locate_image('sair.PNG',confianca=0.7)
        pyautogui.moveTo(sair)
        pyautogui.click()
        time.sleep(2)
        processamento = locate_image('processamento.PNG',confianca=0.7)
        pyautogui.moveTo(processamento)
        pyautogui.click()
        time.sleep(1)
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('right')
        pyautogui.press('enter')
        
        if encerrar_programa == True:
            print(f'saindo do programa pelo erro de Interrogação')
            exit()
        else:
            pass
    
    else:
        time.sleep(1)
        pyautogui.write('02/10/2023')
        pyautogui.press('tab')
        pyautogui.write('08/10/2023')
        pyautogui.press('tab') 
        pyautogui.press('tab')
        pyautogui.write(plano)
        pyautogui.press('tab')
        pyautogui.write(CCusto)
        pyautogui.press('tab')
        pyautogui.write(GrupoOpera)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.write(Operacao)
        pyautogui.press('tab')
        time.sleep(3)
        pyautogui.press('tab')
        pyautogui.write(Prestador)
        pyautogui.press('tab')
        pyautogui.write(Responsavel)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.write(Regiao)
        pyautogui.press('tab')
        pyautogui.write(Quadra)
        pyautogui.press('tab')
        pyautogui.write(Bloco)
        time.sleep(1)
        recurso = locate_image('recursos.PNG',confianca=0.7)
        pyautogui.moveTo(recurso)
        pyautogui.click()
        time.sleep(1)
        
        # Loop para adicionar os insumos na ordem
        for numero, qtd in insumos:
            
            estrela = locate_image('estrela.PNG',confianca=0.7)
            pyautogui.moveTo(estrela)
            pyautogui.click()
            pyautogui.write(Tipo)
            pyautogui.press('tab')
            pyautogui.write(numero)
            pyautogui.press('tab')
            pyautogui.write(qtd)
            pyautogui.press('tab')
            pyautogui.write('1')
        
        time.sleep(1)
        manejo = locate_image('manejo.PNG',confianca=0.7)
        pyautogui.moveTo(manejo)
        pyautogui.click()
        time.sleep(2)
        numero_aplicacoes = locate_image('numero_aplicacoes.PNG',confianca=0.7)
        pyautogui.moveTo(numero_aplicacoes)
        pyautogui.click()
        time.sleep(1)
        pyautogui.move(0, 33)
        pyautogui.click()
        time.sleep(1)
        pyautogui.write('1')
        time.sleep(1)
        pyautogui.move(0, 25)
        pyautogui.click()
        pyautogui.write('1')
        pyautogui.move(0, 25)
        pyautogui.click()
        pyautogui.write('1')
        pyautogui.move(0, 25)
        pyautogui.click()
        pyautogui.write('1')
        pyautogui.move(0, 25)
        pyautogui.click()
        pyautogui.write('1')
        pyautogui.move(0, 24)
        pyautogui.click()
        pyautogui.move(0, 25)
        pyautogui.click()
        pyautogui.write('1')
        pyautogui.move(0, 24)
        pyautogui.click()
        pyautogui.write('1')
        
        time.sleep(2)
        local = locate_image('local.PNG',confianca=0.7)
        pyautogui.moveTo(local)
        pyautogui.click()
        time.sleep(2)
        gerar_os = locate_image('gerar_os.PNG',confianca=0.7)
        pyautogui.moveTo(gerar_os)
        pyautogui.click()
        time.sleep(15)
        bloco_notas = locate_image('bloco_notas.PNG',confianca=0.7)
        pyautogui.moveTo(bloco_notas)
        pyautogui.click()
        time.sleep(15)
        fechar_bloco_notas = locate_image('fechar_bloco_notas.PNG',confianca=0.7)
        pyautogui.moveTo(fechar_bloco_notas)
        pyautogui.click()
        time.sleep(15)
        ok = locate_image('ok.PNG',confianca=0.7)
        pyautogui.moveTo(ok)
        pyautogui.click()
        time.sleep(3)
        sair = locate_image('sair.PNG',confianca=0.7)
        pyautogui.moveTo(sair)
        pyautogui.click()
        time.sleep(1)
        processamento = locate_image('processamento.PNG',confianca=0.7)
        pyautogui.moveTo(processamento)
        pyautogui.click()
        time.sleep(1)
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('right')
        pyautogui.press('enter')
        if encerrar_programa == True:
            print(f'saindo do programa pelo erro de exclamação')
            exit()
        else:
            pass
        
print(f'o programa parou na linha {index+1}')
thread.join()
exit()