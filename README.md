Aplikacja wysyłająca pity
==============================

Co to jest?
-------------

Prosta aplikacja działająca w Pythonie wysyłająca maile z archiwami wg list w excelu.

Co zrobić przed pierwszym użyciem?
---------------
Aby umożliwić aplikacji wykorzystanie twojego e-maila:
1. Przejdź do https://mail.google.com/
2. Zaloguj się na skrzynkę pocztową.
3. Na stronie startowej twojej skrzynki pocztowej, w prawym rogu ekranu widoczna jest ikona twojego profilu "Konto Google".
4. Przyciśnij ikonę "Konto Google" a następnie kliknij przycisk "Zarządzaj kontem Google".
5. Przejdż do podstrony "Bezpieczeństwo" (Sekcja ta znajduje się po lewej stronie ekranu).
6. W sekcji "Bezpieczeństwo", znajdz sekcję zatytułowaną "Dostęp mniej bezpiecznych aplikacji"
   (Domyślnie ustawienie jest "WYŁĄCZONE")
7. Zmień ustawienia z "Wyłączone" na "Włączone".
8. Prawdopodobnie dostaniesz informujący e-mail na temat zagrożenia bezpieczeństwa, co jest oczekiwanym zachowaniem.
9. Możesz zamknać przeglądarkę.


Jak korzystać z programu?
-------
1. W głównym folderze umieść plik z excela o nazwie 'mailing-list.xlsxl.
2. W głównym folderze znajduje się folder 'archiwa' i w nim folder o nazwie 'foldery;. Umieść foldery w tym folderze.
3. Archiwa zostaną wygenerowane w folderze 'archiwa'
4. Excell powinien mieć budowę jak w przykładzie znajdującym się w głównym folderze: kolumny: 'rok', 'pracownik nazwisko', 'pracownik imie', 'adres mailowy', 'haslo'
   - bez polskich znaków. Jeśli nie zostanie podane hasło w kolumnie "haslo" - paczka nie zostanie wysłana.
5. Kliknij dwukrotnie na plik send_emails.exe
6. Pojawi się menu. Wpisz nr 1 żeby podać swojego maila i hasło. Podając hasło nie będą się pojawiać litery.
7. Po podaniu danych można sprawdzić je wpisując nr 2.
8. Wpisując nr 3 wyślemy maile.
9. Wpisując nr 4 wyjdziemy z programu a wpisując nr 5 otworzy się ten plik który właśnie czytasz.
10. Wpisując 6 spakujemy pliki do archiwów.
11. Wpisując 7 wypełnimy w excelu rząd 'hasła' losowymi hasłami 
12. Jeśli pojawią się błędy, powodem może być zła nazwa pliku excel lub archiw. Pliki programu nie powinny być przenoszone do innych folderów.

Nie jest wymaganie instalowanie pythona ani żadnych bibliotek aby móc korzystać z programu.

Dlaczego archiwa pojawiają się w formacie Duda A.rar.txt ?
------------
Gmail nie pozwala wysyłać zaszyfrowanych archiwów. Jeśli zostanie dopisane .txt, gmail rozpozna go jako plik tekstowy. W odebraniu maila wystarczy
usunąć w nazwie .txt i plik będzie można bez problemów rozpakować.

Jak wprowadzić zmianę w programie?
------------

1. Otwórz plik main.py (potrzebujesz jakiegoś rodzaju edytor)
2. Wprowadź zmianę do programu
3. Przejdź do folderu update
4. Kliknij update.bat
5. W głównym folderze pojawi się nowy plik 'send_emails_new.exe'
7. Gdyby brakowało któreś z bibliotek - uruchom plik 'install_dependencies.bat'

Do dokonywania zmian wymagany jest zainstalowany python.

Jeżeli podczas wykonywania skryptu wyświetli się błąd:

Jeśli Gmail wyświetla błąd „Ta wiadomość została zablokowana, ponieważ jej treść stanowi potencjalne zagrożenie”, 
przyczyn może być kilka. Gmail blokuje wiadomości, które mogą rozprzestrzeniać wirusy – 
w tym wiadomości zawierające pliki wykonywalne i niektóre linki.

Więcej: https://support.google.com/mail/answer/6590?hl=pl#zippy=%2Cwiadomo%C5%9Bci-z-za%C5%82%C4%85cznikami

Kontakt
----------------
wojciech.kowcz@testarmy.com