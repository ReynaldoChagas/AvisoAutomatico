#Importação das bibliotecas
import pandas as pd
from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
import urllib
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException

#Conversor para corrigir as especialidades que saem cortadas no relatório
conversor = {'Alergia e Imu': 'Alergia e Imunologia',
 'Assistente So': 'Assistente Social - Intervenção Precoce',
 'Cirurgia de C': 'Cirurgia de Cabeça e Pescoço',
 'Cirurgia Gera': 'Cirurgia Geral',
 'Cirurgia Plás': 'Cirurgia Plástica',
 'Cirurgia Vasc': 'Cirurgia Vascular',
 'Endocrinologi': 'Endocrinologia',
 'Fisioterapia ': 'Fisioterapia',
 'Fonoaudiologi': 'Fonoaudiologia',
 'Gastroenterol': 'Gastroenterologia',
 'Ginecologia -': 'Ginecologia',
 'Laqueadura - ': 'Laqueadura - Avaliação Cirúrgica',
 'Laqueadura Ge': 'Laqueadura Gestante- Avaliação Cirúrgica',
 'Medico Planej': 'Medico Planejamento Familiar',
 'Nefrologia - ': 'Nefrologia - Avaliação',
 'Neurologia - ': 'Neurologia - Pacientes Especiais',
 'Neurologia Pe': 'Neurologia Pediátrica',
 'Odonto - Diag': 'Odonto - Diagnostico Lesoes (Estomatolog',
 'Odonto-Avalia': 'Odonto-Avaliação Paciente Renal Crônico',
 'Odonto.Cir.Tr': 'Odonto.Cir.Traum. Buco-Maxilo-Facial ',
 'Odontologia P': 'Odontologia PNE',
 'Oftalmologia ': 'Oftalmologia - Refração/Outros',
 'Oftalmologia-': 'Oftalmologia- Glaucoma/Retinopatias',
 'Oncologia Clí': 'Oncologia Clínica',
 'Otorrinolarin': 'Otorrinolaringologia',
 'Traumatologia': 'Traumatologia - Atendimento Pós PS',
 'Urologia - Pe': 'Urologia - Pediátrica',
 'Vasectomia - ': 'Vasectomia - Avaliação Cirúrgica'}

#Escolhe o tipo de agendamento para avisar
while True:
    escolha = input('Escolha o tipo de aviso:\n<1>Exames\n<2>Consultas\n') 
    match escolha:
        case '1': #Mensagens para Exames
            df_aviso = pd.read_excel('exames.xlsx')
            df_aviso = df_aviso.fillna('').astype(str)
            df_aviso = df_aviso.drop_duplicates(subset=['Nome','Procedimento','Data/Hora'],keep='first').reset_index(drop=True)
            break
        case '2': #Mensagens para Consultas
            df_aviso = pd.read_excel('consultas.xlsx')
            df_aviso['Hora'] = df_aviso['Hora'].apply(lambda x: x[:5] if type(x) == str else x.strftime('%H:%M')) #Ajusta a hora
            df_aviso['Data'] = df_aviso['Data'].apply(lambda x: x[:10] if type(x) == str else x.strftime('%d/%m/%Y')) #Ajusta a data
            df_aviso = df_aviso.fillna('').astype(str)
            df_aviso['Procedimento'] = df_aviso['Procedimento'].replace(conversor)
            df_aviso = df_aviso.drop_duplicates(subset=['Nome','Profissional','Data','Hora'],keep='first').reset_index(drop=True)
            break          
        case _:
            print('Escolha errada, tente novamente.')                
assinatura = str(input('Digite a assinatura das mensagens (Ex: Usafa xxxxx):\n'))


def login_whats(driver): #função para realizar o login no whatsapp
    print('Verificando login.\n')
    while True:
        try:
            campo_busca = driver.find_element(By.XPATH,'//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div[1]')
            print("Whatsapp logado.")
            break #whatsapp já está logado
        except NoSuchElementException:
            sleep(5) #tempo para aguardar o login
    sleep(3) #tempo para carregar a página após o login

def cria_mensagem(escolha,df,i):
    match escolha:
        case '1':
            mensagem = f'''Olá, Sr(a) *{df.loc[i,'Nome']}*
Estamos entrando em contato para comunicar o agendamento de: 
*{df.loc[i,'Procedimento']}*
*{df.loc[i,'Data/Hora']}*

Digite *1* Para *CONFIRMAR*
Digite *2* Para *REMARCAR* (Justifique o motivo)
Digite *3* Para *EXCLUIR* (Justifique o motivo)

at.te
{assinatura}
'''
        case '2':
            if df.loc[i,'Tipo Consulta'] == 'N':
                tipo_consulta = 'Consulta'
            else:
                tipo_consulta = 'Retorno'
            mensagem = f'''Olá, Sr(a) *{df.loc[i,'Nome']}*
Estamos entrando em contato para comunicar o agendamento de:
*{tipo_consulta}* 
*{df.loc[i,'Procedimento']}* 
*{df.loc[i,'Profissional']}*
*{df.loc[i,'Data']}*
*{df.loc[i,'Hora']}*

Digite *1* Para *CONFIRMAR*
Digite *2* Para *REMARCAR* (Justifique o motivo)
Digite *3* Para *EXCLUIR* (Justifique o motivo)

at.te
{assinatura}
'''
    return mensagem

#Abre navegador
driver = webdriver.Chrome()
driver.get('https://web.whatsapp.com/')
wait = WebDriverWait(driver,15)
#Loop para logar o captcha
login_whats(driver) 

qtd_avisos = df_aviso.loc[df_aviso['Controle']!='Enviado','Controle'].count() #contabiliza a quantidade de envios que serão realizados
try:
        #Envia as mensagens para cada paciente na listagem
        for y,i in enumerate(df_aviso.loc[df_aviso['Controle']!='Enviado','Controle'].index): #percorre apenas os indices que serão utilizados para o contador
                print(f'Enviando mensagens, aguarde! Progresso: {y+1}/{qtd_avisos}')
                if df_aviso.loc[i,'Controle'] == 'Enviado': 
                        continue #pula a repetição caso já tenha enviado a mensagem
                else:
                        mensagem = cria_mensagem(escolha,df_aviso,i)
                        contato = df_aviso.loc[i,'Telefone']
                        if contato == '': #evita a página ficar em loop quando o contato é vazio
                                df_aviso.loc[i,'Controle'] = 'Sem número para contato'
                                continue
                        texto = urllib.parse.quote(mensagem)
                        try:
                                driver.get(f'https://web.whatsapp.com/send/?phone=55${contato}&text={texto}')
                                campo_texto = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p')))
                                sleep(5)
                                while True:
                                    campo_texto.send_keys(Keys.ENTER)
                                    sleep(2)
                                    if assinatura in driver.find_elements(By.XPATH,"//span[@dir='ltr']")[-1].text:
                                        df_aviso.loc[i,'Controle'] = 'Enviado'
                                        break
                                sleep(5)
                        except NoSuchElementException: #Exceção para quando não existe whatsapp cadastrado no número
                                df_aviso.loc[i,'Controle'] = 'Falha no envio. Verificar contato.'
                                continue
                        except TimeoutException: #Exceção para erros de timeout
                                df_aviso.loc[i,'Controle'] = 'Falha no envio. Verificar contato.'
                                continue
                        except StaleElementReferenceException: #Exceção para quando não existe whatsapp cadastrado no número
                                df_aviso.loc[i,'Controle'] = 'Falha no envio. Tentar novamente.'
                                continue
                        except Exception as e:
                                print(f'Erro {type(e)} ao enviar para {contato}.\nAplicação encerrada. Entre em contato com o administrador') #detecta outros possíveis erros para posterior tratamento
                                break
                        finally:
                            match escolha: #atualiza a planilha de acordo com o aviso realizado
                                case '1':
                                        df_aviso.to_excel('exames.xlsx',index=False)
                                case '2':
                                        df_aviso.to_excel('consultas.xlsx',index=False)     
        
                  
finally:
    print('Fim da aplicação.\nProgresso salvo na planilha.')

