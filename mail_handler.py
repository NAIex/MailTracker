import sys
import subprocess

def install(package):
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

import imaplib
import datetime
import os
try:
    import xlwings
except:
    install("xlwings")
    subprocess.check_call([sys.executable, "xlwings", "addin", "install"])
    import xlwings
import email
from email.mime.text import MIMEText
from getpass import getpass

def month_days(month):
    month_days = [-1,31,28,31,30,31,30,31,31,30,31,30,31]
    return month_days[month]
def get_week(date):
    """Returns the week like an interval of 2 dates, format: yyyy-mm-dd - yyyy-mm-dd """
    first_space = date.find("-")
    second_space = date.rfind("-")
    
    day = int(date[second_space+1:])
    month = int(date[first_space+1:second_space])
    year = int(date[:first_space])
    
    real_date = datetime.date(year,month,day)
    index_of_week = real_date.weekday()
    if day - index_of_week < 1:
        first_day = datetime.date(year,month-1,month_days(month-1) + (day - index_of_week))
    else:
        first_day = datetime.date(year,month,day - index_of_week)
        
    if day - index_of_week + 6 > month_days(month):
        last_day = datetime.date(year,month+1,day - index_of_week + 6 - month_days(month))
    else:
        last_day = datetime.date(year,month,day - index_of_week + 6)


    return str(first_day) + " - " + str(last_day)

def change_date_format(original_date:str):
    """Changes date format from dd mmm yyyy to the retarded yyyy-mm-dd """
    first_space = original_date.find(" ")
    second_space = original_date.rfind(" ")
    
    day = original_date[:first_space]
    month = original_date[first_space+1:second_space]
    year = original_date[second_space:]

    
    match month:
        case "Jan":
            month = "01"
        case "Feb":
            month = "02"
        case "Mar":
            month = "03"
        case "Apr":
            month = "04"
        case "May":
            month = "05"
        case "Jun":
            month = "06"
        case "Jul":
            month = "07"
        case "Aug":
            month = "08"
        case "Sep":
            month = "09"
        case "Oct":
            month = "10"
        case "Nov":
            month = "11"
        case "Dec":
            month = "12"
    return year + "-" +month + "-" + day

def get_name_from_mail(mail):
    mail:str
    punct_pos = mail.find(".")
    nume = mail[:punct_pos].capitalize() + " " + mail[punct_pos+1:].capitalize()

    return nume
def get_project_name_from_list(name_list):
    full_name =''
    for name in name_list:
        full_name += name + ' '

    return full_name

def read_email(sender, password):
    M = imaplib.IMAP4_SSL("imap.gmail.com")
    M.login(sender,password)

    M.select('inbox', readonly= True)
    adrese_re =["silviu.pavel"]
    for adresa in adrese_re:
        current_date = datetime.date.today().strftime("%d-%b-%Y")
        #resp_code, mails = M.search(None, f'FROM {adresa}')
        resp_code, mails = M.search(None, f'SINCE {current_date}')
        mail_ids = mails[0].decode().split()
        #print(" Mail IDs : {}\n".format(len(mail_ids)))
        print(f'Mails received on {sender} since {current_date}: {len(mail_ids)}')


    M.close()
    M.logout()

#read_email(sender,password)
def open_emails_txt(mode = "r+"):
    ## === GENERATING EMAIL FILE PATH ===
    emails_path = os.path.abspath(__file__)
    emails_path = emails_path[0:emails_path.rfind('\\')]
    emails_path += '\\adreseRE.txt'

    try:
        email_file = open(emails_path, mode)
    except:
        email_file = open(emails_path,"x+")
        email_file.writelines(datetime.date.today().strftime("%d-%b-%Y"))
        pass
    
    if mode == "w":
        return email_file

    ## === PRINT LAST SEARCH DATE ===
    lastDate = email_file.readline()
    lastDate = lastDate[:-1]
    if lastDate == str(0):
        lastDate = datetime.date.today().strftime("%d-%b-%Y")
        #print(f'Cauta mailuri incepand cu data de: {lastDate}')

    ## === PRINT EMAILS SEARCHED
    emails=[]
    for line in email_file:
        if line.find('\n') != -1:
            line = line[0:line.find('\n')]
        emails.append(line)
    #print(f'Mailurile membrilor sunt: {emails}')
    
    email_file.seek(0,0)
    return email_file
def close_emails_txt(email_file):
    email_file.seek(0,0)    
    email_file.writelines(datetime.date.today().strftime("%d-%b-%Y") + '\n')

    email_file.close()

def check_emails(sender, password, project_name:list,email_file):
    M = imaplib.IMAP4_SSL("imap.gmail.com")
    M.login(sender,password)
    emails = []
  
    lastDate = email_file.readline()
    lastDate = lastDate[:-1]
    if lastDate == str(0):
        lastDate = datetime.date.today().strftime("%d-%b-%Y")
  
    for line in email_file:
        if line.find('\n') != -1:
            line = line[0:line.find('\n')]
        emails.append(line)
    
    M.select("INBOX",readonly= True)

    print(f'\n\n\nContactari pentru {get_project_name_from_list(project_name)}:\n')
    lista_mailuri_trimise = []
    for om in emails:
       lista_dati_om = []
       subjs = []
       for names in project_name:
           subjs.append("SUBJECT " + names)

       typ, raw_data = M.search(None,f'FROM {om}', f'SENTSINCE {lastDate}', *subjs)
       mail_info = raw_data[0].decode().split()
       print(f'{get_name_from_mail(om)} a trimis {len(mail_info)} mailuri')
       for mail_id in mail_info:
            typ, mail_data = M.fetch(mail_id,'(BODY.PEEK[HEADER.FIELDS (DATE)])')
            message = email.message_from_bytes(mail_data[0][1])

            date_as_str = str(message.get("Date"))[5:-15]
            week_as_str = get_week(change_date_format(date_as_str))
            lista_dati_om.append(week_as_str)

       om_nr_mail_trimise = (om, lista_dati_om)
       lista_mailuri_trimise.append(om_nr_mail_trimise)


    email_file.seek(0,0)    
    M.close()
    M.logout()

    return lista_mailuri_trimise

def open_excel():
       ## === GENERATING DATA FILE PATH ===
    emails_path = os.path.abspath(__file__)
    emails_path = emails_path[0:emails_path.rfind('\\')]
    emails_path += '\\emailuriTrimise.xlsx'
    #print(emails_path)
    try:
        evidenta_excel = xlwings.Book(emails_path)
    except:
        evidenta_excel = xlwings.Book()

    return evidenta_excel
def close_excel(evidenta_excel):
        ## === GENERATING DATA FILE PATH ===
    emails_path = os.path.abspath(__file__)
    emails_path = emails_path[0:emails_path.rfind('\\')]
    emails_path += '\\emailuriTrimise.xlsx'

    evidenta_excel.save(emails_path)
    evidenta_excel.close()

def update_email_data(evidenta_excel,new_data, project_name):
    full_name = get_project_name_from_list(project_name)

    workSheet: xlwings.Sheet
    try:
        workSheet = evidenta_excel.sheets[full_name]    
    except:
        evidenta_excel.sheets.add(full_name)
        workSheet = evidenta_excel.sheets[full_name]

    total_weeks = 0
    for column in range(1,100):
        if workSheet[0,column].value != None:
            total_weeks += 1
        else:
            break
    total_names = 0
    for row in range(1,100):
        if workSheet[row,0].value != None:
            total_names += 1
        else:
            break

    for nume_adresa, dati_mailuri in new_data:
        if len(dati_mailuri) == 0:
            continue

        nume = get_name_from_mail(nume_adresa)
        index_nume = 1+total_names

        try:
            names_list = workSheet[1:1+total_names,0].value
            names_list:list
            
            index_nume = names_list.index(nume)+1
        except:
            workSheet[1+total_names,0].value = nume
            total_names += 1      

        for date in dati_mailuri:
            index_date = total_weeks + 1
            try:
                dates_list = workSheet[0,1:index_date].value
                
                index_date = dates_list.index(date)+1

                workSheet[index_nume,index_date].value +=1
            except:
                workSheet[0,index_date].value = date
                workSheet[index_nume,index_date].value = 1

                total_weeks += 1
    workSheet[0:100,0:100].columns.autofit()

#check_emails(sender,password)

nume_file_adrese = 'adreseRE.txt'
nume_file_date = 'emailuriTrimise.xlsx'
def error_handle(error_code):
    """0 - Greseala la apelarea modulului"""
    print("Nu e corect cum ai apelat modulul, apeleaza cu 'help' pt o descriere a sa", end=" ")
    exit()

def module_login_and_mail_scraping():
  #=== LOGGING SYSTEM ===
            adresa_email:str
            password:str
            failed_login = True

            #=== TRY LOGGING IN UNTIL IT WORKS ===
            while failed_login:
                adresa_email =  input("adresa_email: ")
                password =  getpass(prompt= "parola(chiar daca nu apare, merge sa scrii): ")

                try:
                    M = imaplib.IMAP4_SSL("imap.gmail.com")
                    M.login(adresa_email,password)
                    M.logout()
                    failed_login = False
                    pass
                except:
                    print(f'\nDate de logare incorecte, try again bro')
                
            #=== OPEN TXT FILE === 
            emails_list = open_emails_txt()

            # === CHECK MAILS FOR PROJECTS IN LIST ===
            results_list = []
            projects_list = [['FII','IT-ist'],['FIICode'],['FII','Practic']]
            for i, project in enumerate(projects_list):
                result = check_emails(adresa_email,password,projects_list[i],emails_list)
                results_list.append(result)
            
            #=== STORE FOUND DATA IN EXCEL ===
            excel = open_excel()
            for i, results in enumerate(results_list):
                update_email_data(excel,results_list[i],projects_list[i])
            close_excel(excel)

            close_emails_txt(emails_list)
def module_reset():
    result = input("Doresti sa dai reset? y/n: ")
    if result == "n":
        exit()
    elif result == "y":
        result = input("SIGUR SIGUR, doresti sa dai reset? y/n: ")

    if result == "n":
        exit()
    elif result == "y":
        print("Incepe procesul de resetare")
        excel = open_excel()
        for sheet in excel.sheets:
            sheet:xlwings.Sheet
            sheet.clear()
        close_excel(excel)

        mail_file = open_emails_txt()
        data_noua = input("Data noua de cautare(cu format 01-Jan-1999): ")
        mail_file.write(data_noua)
        mail_file.close()
def module_modify_list():
    while True:
        emails = open_emails_txt()
        data:str =""
        for line in emails:
            data += line
        print(data)

        result = input("Asteapta comanda( del prenume.nume, add prenume.nume, exit ): ")
        space_index = result.find(" ")
        if space_index == -1 and result == "exit":
            break
        command = result[:space_index] 
        name = result[space_index+1:]
        match command:
            case "del":
                name_start_index = data.find(name) - 1
                name_last_index = name_start_index + len(name) + 1
                if data[name_start_index] == '\n' and (name_last_index == len(data) or data[name_last_index] == '\n'):
                    data = data[:name_start_index] + data[name_last_index:]
                else:
                    print(f'Numele nu exista!')
                
            case "add":
                if name.find(".") != -1:
                    data += '\n'+name 
                else:
                    print(f'Numele ar trebui sa fie in formatul prenume.nume!')
                pass

        emails.close()
        emails = open_emails_txt("w")
        emails.write(data)
        emails.close()
            
        


if __name__ == "__main__":
    if sys.argv.__len__() != 2:
        error_handle(0)
    elif sys.argv.__len__() == 2:
        if sys.argv[1] == 'help':
            print(f'\nModulul in principal interactioneaza cu doua fisiere: {nume_file_adrese} si {nume_file_date} \n')
            print(f'{nume_file_adrese}\n \
            Prima linie - Ultima data cand a fost apelat modulul, si implicit data de cand sunt cautate mailurile \n \
            Celelalte linii - Contin adresele de email ale oamenilor din departament ce contacteaza, sub forma prenume.nume(fara @asii.ro) \n')
            print(f'{nume_file_date}\n \
            Prima linie - Pe fiecare coloana (mai putin prima) este prezenta cate o saptamana \n \
            Celelalte linii - Pe prima coloana este prezent numele membrului, iar pe restu cate mailuri a trimis in saptamanile respective \n')
            print(f'Apeleaza cu "login" pentru a rula modulul\n\n')
            print(f'Resetarea Modulului\n\
                  Cand se apeleaza comanda de reset, se goleste fila de excel, si se seteaza data de cand se fie apelat modulul.\n\
                  Formatul datii este de forma: "01-Jan-1999"\n')
            print(f'Apeleaza cu "reset" pentru a reseta modului\n\n')
            print(f'Modificare Listei de mailuri\n\
                  Cand se apeleaza comanda de modify, apare lista cu mailurile puse.\n\
                  Pentru a da delete la un nume, se scrie "del prenume.nume", iar pentru a adauga un nume se scrie "add prenume.nume"')
            print(f'Apeleaza cu "modify_list" pentru a modifica list')
            
        elif sys.argv[1] == 'login':
            module_login_and_mail_scraping()
        elif sys.argv[1] == 'reset':
            module_reset()
        elif sys.argv[1] == 'modify_list':
            module_modify_list()
        else:
            error_handle(0)



