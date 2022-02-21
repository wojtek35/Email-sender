from cProfile import run
import time
import os
import subprocess
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import getpass
import random
import string
import sys

# Ustawienie bieżącego katalogu, tam gdzie znajduje się program
if getattr(sys, 'frozen', False):	# Jeśli program jest w postaci .exe
    application_path = os.path.dirname(sys.executable)
elif __file__:						# Jeśli program jest w postaci .py
    application_path = os.path.dirname(__file__)

def get_data_from_excel_and_send_email(sender_email, sender_password, name_and_surname):
	"""Funkcja pozwalająca zebrać dane z excela a następnie wysłać email"""
	current_year = time.strftime("%Y") # Zwraca obecny rok
	previous_year = str(int(current_year) -1) # Zwraca poprzedni rok. Jeżeli email będzie wysłany w 2022 roku 
											  # w wiadomości pojawi się "Pit za rok 2021"

	def send_mail(to, subject, content, attatchment_name, last_name, first_name, sender_email, sender_password, name_and_surname):
		"""Funkcja pozwalająca wysłać maila ze skrzynki gmail"""

		# Utworzenia klienta SMTP na który zaloguje się na nasz gmail
		with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp: # Nawiązanie bezpiecznego połączenia SSL
			# Tworzenie wiadomości
			msg = MIMEMultipart() 
			msg['Subject'] = subject
			msg['From'] = sender_email
			msg['To'] = to
			body_part = MIMEText(content, 'plain') # Tworzenie treści maila
			msg.attach(body_part)

			smtp.login(sender_email, sender_password) # Logowanie do serwera

			# Jeśli mamy załącznik 
			if attatchment_name != '':
				# Załączamy załącznik 
				attatchment = MIMEApplication(open(attatchment_name, 'rb').read()) # Odczyt pliku
				attatchment.add_header("Content-Disposition", 'attachment; filename="%s"' % f'{last_name.title()} {first_name[0].upper()}.rar.txt')
				msg.attach(attatchment) 
				smtp.send_message(msg) # Wysłanie maila

			# Jeśli nie ma załącznika
			else:
				smtp.send_message(msg) # Wysłanie maila
	
	file = r'mailing-list.xlsx' 	# excel z danymi
	try:
		df = pd.read_excel(os.getcwd()+"\\" +file) # Załaduj plik z excela do pandas dataframe
	except FileNotFoundError:
		# Prawdopodobnie zła nazwa pliku lub jest w innym folderze
		print("Brakuje pliku z danymi lub program został przeniesiony do innego folderu.")
		run_pit_script()
	except PermissionError:
		print('Zamknij excella')

	df_ready = pd.DataFrame()  # Tworzenie pustej ramki: Empty DataFrame Columns: [] Index: []
	df_ready["pracownik nazwisko"] = df["pracownik nazwisko"]  # przypisanie kolumny "pracownik nazwisko"
	df_ready["pracownik imie"] = df["pracownik imie"]  # przypisanie kolumny "pracownik imie"
	df_ready["adres mailowy"] = df["adres mailowy"]  # przypisanie kolumny "adres mailowy"
	df_ready["haslo"] = df["haslo"]  # przypisanie kolumny "hasło"
	subject = f"Rozliczenie roczne {previous_year} PIT-11" # Temat maila, który będzie widać w wiadomości
	os.system('cls')
	for index, row in df_ready.iterrows(): # pętla przechodzi przez każdy rząd i wysyła maila do opowiedniej osoby na podstawie danych
		to = row['adres mailowy']
		last_name = row['pracownik nazwisko']
		first_name = row['pracownik imie']
		attatchment_name = f"./archiwa/{last_name.title()} {first_name[0].title()}.rar.txt" # w nazwie dodano .txt aby google nie wykrył jako plik.rar zabezpieczony hasłem
		no_attatchment = ''
		content_password = "Haslo do pliku: " + str(row['haslo']) # Treść maila z hasłem
		# Treść maila z archiwum w załączniku
		content = f"""Cześć {first_name}.\nW załączeniu przesyłam rozliczenie roczne za {previous_year} rok. W odrębnym mailu otrzymasz haslo do pliku.
		Pamiętaj, że rozliczenie dostajesz tylko w formie elektronicznej dlatego zachowaj zarówno plik jak i hasło do pliku.
		Gdybyś miał dodatkowe pytania co do rozliczenia lub byłyby problemy z otwarciem pliku śmiało pisz.
		Aby wypakować archiwum proszę zmienić nazwę poprzez usunięcie końcówki .txt jak w przykładzie: Nazwisko I.rar.txt na Nazwisko I.rar\n\n{name_and_surname}
		"""
		content = (''.join([c for c in content if c not in ['\t']])) # Żeby entery w kodzie nie pojawił się w wiadomości
		
		# PIT
		print('########### PIT #############\n' + str(row["pracownik imie"]) + " " + str(row["pracownik nazwisko"]) + "\n")
		try:
			# Wysyła maila z załącznikiem
			send_mail(to, subject, content, attatchment_name, last_name, first_name, sender_email, sender_password, name_and_surname)
		except smtplib.SMTPAuthenticationError:
			print("Zły email lub hasło")
			break
		# Hasla
		print('########### HASŁO #############\n' + str(row["pracownik imie"]) + " " + str(row["pracownik nazwisko"]) + "\n")
		try:
			# wysyła maila z hasłem
			send_mail(to, subject, content_password, no_attatchment, last_name, first_name, sender_email, sender_password, name_and_surname)
		except smtplib.SMTPAuthenticationError:
			print("Zły email lub hasło")
			break
		print('******************************** Nastęna osoba **************************************')

def generate_passwords():
	''' Generuje losowe hasła i wpisuje je do kolumny hasła w excellu'''
	file = r'mailing-list.xlsx'
	try:
		df = pd.read_excel(os.getcwd()+"\\" +file) # Załaduj plik z excela do pandas dataframe
	except FileNotFoundError:
		# Prawdopodobnie zła nazwa pliku lub jest w innym folderze
		print("Brakuje pliku z danymi lub program został przeniesiony do innego folderu.")
		run_pit_script()
	except PermissionError:
		print('Zamknij excella')

	for index, row in df.iterrows():
		# Generuje losowe hasło i df.at wypisuje w kolumnie hasło w każdym rzędzie 
		characters = string.ascii_letters + string.digits + string.punctuation 
		password = ''.join(random.choice(characters) for i in range(8))
		df.at[index, 'haslo'] = password
	try:
		df.to_excel(os.getcwd()+'\\mailing-list.xlsx')
	except FileNotFoundError:
		# Prawdopodobnie zła nazwa pliku lub jest w innym folderze
		print("Brakuje pliku z danymi lub program został przeniesiony do innego folderu.")
		run_pit_script()
	except PermissionError:
		print('Zamknij excella')
    
def pack_archives():
	""" Tworzy archiwa a folderów i nadaje im hasło. Rozszerzenie jest .rar.txt dlatego że gmail blokuje pliki .rar zabezpieczone hasłem"""
	file = r'mailing-list.xlsx'
	try:
		df = pd.read_excel(os.getcwd()+"\\" +file) # Załaduj plik z excela do pandas dataframe
	except FileNotFoundError:
		# Prawdopodobnie zła nazwa pliku lub jest w innym folderze
		print("Brakuje pliku z danymi lub program został przeniesiony do innego folderu.")
		run_pit_script()
	except PermissionError:
		print('Zamknij excella')

	for index, row in df.iterrows():
		password = df.at[index, 'haslo']
		last_name = row['pracownik nazwisko']
		first_name = row['pracownik imie']
		folder_name = f'.\\archiwa\\foldery\\"{last_name.title().rstrip()} {first_name[0].title().rstrip()}"'
		zip_name = f'.\\archiwa\\"{last_name.title().rstrip()} {first_name[0].title().rstrip()}.rar.txt"'
		zip_command = f'7zG a -t7z {zip_name} {folder_name} -p"{password}"' # Używa 7 zipa żeby zapakować pliki
		print(zip_command)
		
		if os.system(zip_command):
			print("\nPlik 7zg.exe msui znajdować się w tym samym folderze\n")

			

def menu():
	"""Funkcja wyświetlająca opcje menu"""
	menu_string = ('Wpisz "1" jeżeli chcesz dodać dane.\n'
				'Wpisz "2" jeżeli chcesz sprawdzić dane.\n'
				'Wpisz "3" jeżeli chcesz wysłać maile.\n'
				'Wpisz "4" żeby wyjść z programu.\n'
				'Wpisz "5" żeby uzyskać pomoc.\n'
				'Wpisz "6" żeby spakować archiwa.\n'
                'Wpisz "7" żeby wygenerować hasła.\n')
	your_choice = input(menu_string)
	if your_choice.isnumeric() == False:
			message = 'Tutaj należało wpisać numer'
			print(message)
			run_pit_script()
	return int(your_choice)


def run_pit_script():
	"""Funkcja odpalająca cały skrypt"""
	email = ''
	password = ''
	name_and_surname = ''
	while True:
		your_choice = menu()
		if your_choice == 1:
			email = input(f'Proszę wprowadź gmail z którego chcesz rozsyłać pity'
							f'{email}, zatwierdź przyciskiem <ENTER>:\n')
			password = getpass.getpass(f'Proszę wprowadź swoje haslo do gmaila{password}, zatwierdź przyciskiem <ENTER>"\n')
			name_and_surname = input(f'Proszę wprowadź swoje imię i nazwisko np. Kasia Sibilska'
										f'{name_and_surname}, zatwiedź przyciskiem <ENTER>:\n')
		elif your_choice == 2:
			if email == '' or password == '' or name_and_surname == '':
				print("Nie uzupełniłeś wszystkich danych!\n")
			else:
				print(f'Podałeś email: {email}, oraz imię i nazwisko: {name_and_surname}')
				print("Jeżeli dane są poprawne, wpisz 3")
		elif your_choice == 3:
			if email == '' or password == '' or name_and_surname == '':
				print("Nie uzupełniłeś wszystkich danych!\n")
				continue
			else:
				get_data_from_excel_and_send_email(email, password, name_and_surname)
		elif your_choice == 4:
			break
		elif your_choice == 5:
			subprocess.call(['notepad.exe', 'README.md'])
		elif your_choice == 6:
			pack_archives()
		elif your_choice == 7:
			generate_passwords()
		
run_pit_script()
