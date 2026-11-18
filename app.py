from flask import Flask, render_template_string, request, redirect, url_for, send_file, session
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime, timedelta
import os
import json
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)
app.secret_key = 'tajny_klucz_sesji_produkcyjnej'

# Bezpieczna konfiguracja bazy danych (odczyt z Render.com lub fallback do lokalnego URI z hasłem)
SUPABASE_URI = os.getenv("DATABASE_URL", "postgresql://postgres.dbgjiavdnlgnvqnrkykx:Prod_Glo_2026@aws-0-eu-west-1.pooler.supabase.com:6543/postgres")
app.config['SQLALCHEMY_DATABASE_URI'] = SUPABASE_URI
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# Definicje modeli bazodanowych w PostgreSQL
class ArtykulModel(db.Model):
    __tablename__ = 'produkty_baza'
    kod = db.Column(db.String(50), primary_key=True)
    nazwa = db.Column(db.String(200), nullable=False)
    odbiorca = db.Column(db.String(100), nullable=False)

class UzytkownikModel(db.Model):
    __tablename__ = 'uzytkownicy_baza'
    username = db.Column(db.String(50), primary_key=True)
    haslo = db.Column(db.String(100), nullable=False)
    rola = db.Column(db.String(50), nullable=False)
    pelna_nazwa = db.Column(db.String(100), nullable=False)
    dostep = db.Column(db.Text, nullable=False) # przechowywane jako JSON string

class PlanModel(db.Model):
    __tablename__ = 'plan_produkcyjny'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    data = db.Column(db.String(20), nullable=False)
    kod = db.Column(db.String(50), nullable=False)
    ilosc = db.Column(db.Integer, nullable=False)

class WpisProdukcyjnyModel(db.Model):
    __tablename__ = 'historia_produkcji'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    data = db.Column(db.String(20), nullable=False)
    linia = db.Column(db.String(10), nullable=False)
    zmiana = db.Column(db.String(10), nullable=False)
    kod = db.Column(db.String(50), nullable=False)
    nazwa = db.Column(db.String(200), nullable=False)
    odbiorca = db.Column(db.String(100), nullable=False)
    ilosc = db.Column(db.Integer, nullable=False)
    data_przyd = db.Column(db.String(20), nullable=False)
    partia_prod = db.Column(db.String(50), nullable=False)
    etykieta_gora = db.Column(db.String(100), nullable=False)
    etykieta_dol = db.Column(db.String(100), nullable=False)
    czas_start = db.Column(db.String(10), nullable=False)
    czas_koniec = db.Column(db.String(10), nullable=False)
    potwierdzenie_kontroli = db.Column(db.String(10), nullable=True)
    potwierdzenie_oczyszczenia = db.Column(db.String(10), nullable=True)
    operator = db.Column(db.String(100), nullable=False)

# Inicjalizacja domyślnych danych w bazie, jeśli jest pusta
with app.app_context():
    db.create_all()
    if UzytkownikModel.query.count() == 0:
        domyslni_uzyt = {
            "admin": {"haslo": "admin123", "rola": "admin", "pelna_nazwa": "Administrator Systemu", "dostep": ["1", "2", "3", "4", "plan", "progress", "wydajnosc", "db_art", "admin"]},
            "tetiana": {"haslo": "operator123", "rola": "operator", "pelna_nazwa": "Tetiana Kantymyr", "dostep": ["1", "2", "3", "4", "progress", "wydajnosc"]},
            "iryna": {"haslo": "operator123", "rola": "operator", "pelna_nazwa": "Poberezhna Iryna", "dostep": ["1", "2", "3", "4", "progress", "wydajnosc"]}
        }
        for usr, info in domyslni_uzyt.items():
            u = UzytkownikModel(username=usr, haslo=info["haslo"], rola=info["rola"], pelna_nazwa=info["pelna_nazwa"], dostep=json.dumps(info["dostep"]))
            db.session.add(u)
        
        domyslnie_produkty = {
            "1354": {"nazwa": "Hähnchen Teriyaki 190g", "odbiorca": "Aldi SUD"},
            "1378": {"nazwa": "Hähnchen Sweet Chili 190g", "odbiorca": "Aldi SUD"},
            "1422": {"nazwa": "Falafel-Hummus 190g", "odbiorca": "Aldi SUD"},
            "1553": {"nazwa": "Kanapka z jajkiem i pastą jajeczną", "odbiorca": "Perla PL"},
            "1861": {"nazwa": "Wrap Azja tortilla pomidorowa 190g HU", "odbiorca": "Lidl Węgry"},
            "1878": {"nazwa": "Wrap Włoski tortilla pomidorowa 190g HU", "odbiorca": "Lidl Węgry"},
            "1885": {"nazwa": "Wrap Grecki tortilla pomidorowa 190g HU", "odbiorca": "Lidl Węgry"},
            "2690": {"nazwa": "Thunfisch 190g", "odbiorca": "Aldi SUD"},
            "3101": {"nazwa": "G&G Paprykowy 190g", "odbiorca": "EDEKA"},
            "3071": {"nazwa": "G&G Wrap Cezar 190g", "odbiorca": "EDEKA"},
            "3132": {"nazwa": "G&G Wrap Gyros 190g", "odbiorca": "EDEKA"},
            "3163": {"nazwa": "G&G Wrap Tajski 190g", "odbiorca": "EDEKA"},
            "3194": {"nazwa": "G&G Wrap tuńczyk MSC 190g", "odbiorca": "EDEKA"}
        }
        for kod, info in domyslnie_produkty.items():
            p = ArtykulModel(kod=kod, nazwa=info["nazwa"], odbiorca=info["odbiorca"])
            db.session.add(p)
            
        dzis_str = datetime.now().strftime('%Y-%m-%d')
        domyslny_plan = {"1354": 2500, "2690": 1500, "3101": 2000}
        for kod, szt in domyslny_plan.items():
            pl = PlanModel(data=dzis_str, kod=kod, ilosc=szt)
            db.session.add(pl)
            
        db.session.commit()

T = {
    "pl": {
        "tytul": "Panel Realizacji Produkcji i Raportowania",
        "wybierz_zm": "Wybierz zmianę roboczą:",
        "rozpocznij": "Rozpocznij pracę",
        "zmiana": "Aktualna zmiana",
        "zmien_zm": "Zmień zmianę",
        "zalog_op": "Zalogowany operator",
        "linia_1": "Linia 1", "linia_2": "Linia 2", "linia_3": "Linia 3", "linia_4": "Linia 4",
        "plan_tab": "Plan Produkcyjny", "prog_tab": "Postęp / Realizacja", "wydajnosc_tab": "Wydajność (Szt/h)", "db_art_tab": "Baza Artykułów", "admin_tab": "Panel Administratora",
        "zarejestruj": "Zarejestruj wyprodukowaną partię dla",
        "data_prod": "Data produkcji:", "kod_art": "Kod artykułu:", "wybierz_kod": "-- Wybierz kod artykułu --",
        "ilosc_szt": "Ilość sztuk:", "wpisz_ilosc": "wpisz ilość...", "nazwa_ind": "Nazwa indeksu:",
        "odbiorca": "Odbiorca:", "data_przyd": "Data przydatności (min. +4 dni):", "nr_partii": "Nr partii produktu:",
        "ety_gora": "Etykieta GÓRNA:", "ety_dol": "Etykieta DOLNA:", "operator": "Operator:",
        "dodaj_btn": "Zarejestruj wyprodukowaną paletę / partię", "realizacja_biez": "Realizacja na bieżąco",
        "brak_wpisow": "Brak wpisów dla tej linii i daty.", "lacznie": "Łącznie wyprodukowano na linii:",
        "zapisz_excel": "Zapisz raport", "wklej_plan": "Zarządzanie Planem Produkcyjnym",
        "instrukcja_plan": "Wybierz datę planu, wklej dane z Excela (Kod [Tab] Ilość) lub dodaj pojedynczą pozycję:",
        "zapisz_plan_btn": "Zapisz plan dla wybranej daty", "zarzadzanie_art": "Zarządzanie Bazą Artykułów",
        "dodaj_art": "Dodaj / Edytuj Artykuł", "nazwa": "Nazwa", "dodaj_art_btn": "Zapisz artykuł",
        "istniejace_art": "Istniejące Artykuły", "usun": "[Usuń]",
        "logowanie_tytul": "Logowanie do Systemu Produkcyjnego", "login": "Login:", "haslo": "Hasło:", "zaloguj": "Zaloguj się",
        "brak_dostepu": "Brak uprawnień do tej sekcji!", "wyloguj": "Wyloguj", "zarzadzanie_uzyt": "Zarządzanie Użytkownikami i Uprawnieniami",
        "bhp_tytul": "Wymagane potwierdzenia BHP i Kontroli:",
        "bhp_1": "Potwierdzam przeprowadzenie kontroli pracowników i czystości urządzeń przed przystąpieniem do pracy.",
        "bhp_2": "Potwierdzam prawidłowe oczyszczenie linii z pozostałych etykiet, banderol i kartonów przed rozpoczęciem produkcji kolejnego produktu.",
        "czas_pracy_tytul": "Czas pracy (WYMAGANE):",
        "czas_opis": "Kliknij przyciski, aby pobrać aktualny czas startu i końca lub wpisz ręcznie.",
        "start_palety": "Start Palety", "koniec_palety": "Koniec Palety",
        "czas_start_lbl": "Czas Startu (Wymagane):", "czas_koniec_lbl": "Czas Końca (Wymagane):",
        "postep_tytul": "Łączny Postęp Produkcji",
        "podsumowanie_dnia": "Podsumowanie Dnia (Wszystkie Artykuły)",
        "do_zrobienia": "Do zrobienia łącznie:", "wykonano": "Wykonano:", "pozostalo": "Pozostało:",
        "postep_art": "Postęp według poszczególnych artykułów:",
        "brak_planu_dzis": "Brak planu produkcyjnego na ten dzień.",
        "raport_wyd_tytul": "Raport Wydajności (Sztuk na Godzinę)",
        "raport_wyd_opis": "Poniższy raport przedstawia całkowity czas pracy oraz średnią produktywność w przeliczeniu na 1 godzinę dla każdej linii i zmiany.",
        "th_linia": "Linia", "th_zmiana": "Zmiana", "th_wyprodukowano": "Wyprodukowano sztuk", "th_czas_pracy": "Czas pracy (godziny)", "th_wydajnosc": "Wydajność (szt/h)",
        "brak_czasow": "Brak zarejestrowanych czasów pracy i palet dla wybranej daty.",
        "wybierz_dzien_plan": "Wybierz dzień planu:", "podglad_planu": "Podgląd zaplanowanej produkcji:", "brak_planow": "Brak zdefiniowanych planów produkcyjnych.",
        "dodaj_nowego_uzyt": "Dodaj nowego użytkownika / operatora", "imie_nazwisko": "Imię i Nazwisko operatora:", "rola": "Rola:",
        "uprawnienia_lbl": "Uprawnienia (wybierz zakładki):", "zapisz_uzyt_btn": "Zapisz użytkownika", "zarejestrowani_uzyt": "Zarejestrowani użytkownicy",
        "systemowy": "(systemowy)", "kol_kod": "Kod", "kol_nazwa": "Nazwa", "kol_odbiorca": "Odbiorca", "kol_akcja": "Akcja",
        "wybierz_dzien_hist": "Wybierz dzień do podglądu historii i generowania raportu (Admin):",
        "data_raportu_hist": "Data raportu / historii:", "tabela_nr": "#", "tabela_ilosc": "Ilość", "tabela_czas": "Czas pracy",
        "tabela_bhp": "Potwierdzenia BHP", "tabela_szczegoly": "Szczegóły",
        "ph_kod": "np. 9999", "ph_nazwa": "Nazwa produktu", "ph_odbiorca": "Odbiorca", "ph_ilosc": "wpisz ilość...",
        "ph_etygora": "Wpisz etykietę górną...", "ph_etydol": "Wpisz etykietę dolną..."
    },
    "en": {
        "tytul": "Production Management Panel",
        "wybierz_zm": "Select work shift:",
        "rozpocznij": "Start work",
        "zmiana": "Current shift",
        "zmien_zm": "Change shift",
        "zalog_op": "Logged operator",
        "linia_1": "Line 1", "linia_2": "Line 2", "linia_3": "Line 3", "linia_4": "Line 4",
        "plan_tab": "Production Plan", "prog_tab": "Progress / Realization", "wydajnosc_tab": "Productivity (Pcs/h)", "db_art_tab": "Articles DB", "admin_tab": "Admin Panel",
        "zarejestruj": "Register batch for",
        "data_prod": "Production date:", "kod_art": "Article code:", "wybierz_kod": "-- Select article code --",
        "ilosc_szt": "Quantity:", "wpisz_ilosc": "enter quantity...", "nazwa_ind": "Index name:",
        "odbiorca": "Recipient:", "data_przyd": "Expiry date (min. +4 days):", "nr_partii": "Batch number:",
        "ety_gora": "TOP label:", "ety_dol": "BOTTOM label:", "operator": "Operator:",
        "dodaj_btn": "Register produced palette / batch", "realizacja_biez": "Live progress",
        "brak_wpisow": "No entries for this line and date.", "lacznie": "Total produced on line:",
        "zapisz_excel": "Save report", "wklej_plan": "Production Plan Management",
        "instrukcja_plan": "Select plan date, paste from Excel (Code [Tab] Qty) or add single item:",
        "zapisz_plan_btn": "Save plan for selected date", "zarzadzanie_art": "Articles Management",
        "dodaj_art": "Add / Edit Article", "nazwa": "Name", "dodaj_art_btn": "Save Article",
        "istniejace_art": "Existing Articles", "usun": "[Delete]",
        "logowanie_tytul": "Production System Login",
        "brak_dostepu": "Access denied!", "wyloguj": "Log out", "zarzadzanie_uzyt": "Users and Permissions Management",
        "bhp_tytul": "Required OHS and Control confirmations:",
        "bhp_1": "I confirm the inspection of workers and equipment cleanliness before starting work.",
        "bhp_2": "I confirm the proper cleaning of the line from remaining labels, bands and cartons before starting the production of the next product.",
        "czas_pracy_tytul": "Working time (REQUIRED):",
        "czas_opis": "Click buttons to fetch current start and end time or enter manually.",
        "start_palety": "Palette Start", "koniec_palety": "Palette End",
        "czas_start_lbl": "Start Time (Required):", "czas_koniec_lbl": "End Time (Required):",
        "postep_tytul": "Total Production Progress",
        "podsumowanie_dnia": "Day Summary (All Articles)",
        "do_zrobienia": "Total to do:", "wykonano": "Completed:", "pozostalo": "Remaining:",
        "postep_art": "Progress by individual articles:",
        "brak_planu_dzis": "No production plan for this day.",
        "raport_wyd_tytul": "Productivity Report (Pcs per Hour)",
        "raport_wyd_opis": "The report below shows total working time and average productivity per 1 hour for each line and shift.",
        "th_linia": "Line", "th_zmiana": "Shift", "th_wyprodukowano": "Produced pcs", "th_czas_pracy": "Working time (hours)", "th_wydajnosc": "Productivity (pcs/h)",
        "brak_czasow": "No recorded working times for selected date.",
        "wybierz_dzien_plan": "Select plan day:", "podglad_planu": "Preview of planned production:", "brak_planow": "No production plans defined.",
        "dodaj_nowego_uzyt": "Add new user / operator",
        "imie_nazwisko": "Operator Full Name:", "rola": "Role:",
        "uprawnienia_lbl": "Permissions (select tabs):", "zapisz_uzyt_btn": "Save user", "zarejestrowani_uzyt": "Registered users",
        "systemowy": "(system)", "kol_kod": "Code", "kol_nazwa": "Name", "kol_odbiorca": "Recipient", "kol_akcja": "Action",
        "wybierz_dzien_hist": "Select day to view history and generate report (Admin):",
        "data_raportu_hist": "Report / history date:", "tabela_nr": "#", "tabela_ilosc": "Quantity", "tabela_czas": "Working time",
        "tabela_bhp": "OHS Confirmations", "tabela_szczegoly": "Details",
        "ph_kod": "e.g. 9999", "ph_nazwa": "Product name", "ph_odbiorca": "Recipient", "ph_ilosc": "enter quantity..."
    },
    "es": {
        "tytul": "Panel de Gestión de Producción",
        "wybierz_zm": "Seleccione el turno de trabajo:",
        "rozpocznij": "Iniciar trabajo",
        "zmiana": "Turno actual",
        "zmien_zm": "Cambiar turno",
        "zalog_op": "Operador conectado",
        "linia_1": "Línea 1", "linia_2": "Línea 2", "linia_3": "Línea 3", "linia_4": "Línea 4",
        "plan_tab": "Plan de Producción", "prog_tab": "Progreso / Realización", "wydajnosc_tab": "Productividad (Pcs/h)", "db_art_tab": "Base de Artículos", "admin_tab": "Panel de Admin",
        "zarejestruj": "Registrar lote para",
        "data_prod": "Fecha de producción:", "kod_art": "Código de artículo:", "wybierz_kod": "-- Seleccione código --",
        "ilosc_szt": "Cantidad:", "wpisz_ilosc": "ingrese cantidad...", "nazwa_ind": "Nombre del índice:",
        "odbiorca": "Destinatario:", "data_przyd": "Fecha de caducidad (mín. +4 días):", "nr_partii": "Nº de lote:",
        "ety_gora": "Etiqueta SUPERIOR:", "ety_dol": "Etiqueta INFERIOR:", "operator": "Operador:",
        "dodaj_btn": "Registrar palet / lote producido",
        "realizacja_biez": "Realización en curso",
        "brak_wpisow": "No hay entradas para esta línea y fecha.", "lacznie": "Total producido en la línea:",
        "zapisz_excel": "Guardar informe", "wklej_plan": "Gestión del Plan de Producción",
        "instrukcja_plan": "Seleccione fecha, pegue desde Excel (Código [Tab] Cantidad) o añada:",
        "zapisz_plan_btn": "Guardar plan para la fecha", "zarzadzanie_art": "Gestión de Artículos",
        "dodaj_art": "Añadir / Editar Artículo", "nazwa": "Nombre", "dodaj_art_btn": "Guardar Artículo",
        "istniejace_art": "Artículos Existentes", "usun": "[Eliminar]",
        "logowanie_tytul": "Iniciar sesión en el sistema",
        "brak_dostepu": "¡Acceso denegado!", "wyloguj": "Cerrar sesión", "zarzadzanie_uzyt": "Gestión de Usuarios y Permisos",
        "bhp_tytul": "Confirmaciones requeridas de SST y Control:",
        "bhp_1": "Confirmo la inspección de trabajadores y limpieza de equipos antes de comenzar el trabajo.",
        "bhp_2": "Confirmo la limpieza adecuada de la línea de etiquetas, bandas y cartones restantes antes de iniciar la producción del siguiente producto.",
        "czas_pracy_tytul": "Tiempo de trabajo (OBLIGATORIO):",
        "czas_opis": "Haga clic en los botones para obtener la hora actual o ingrese manualmente.",
        "start_palety": "Inicio de Palet", "koniec_palety": "Fin de Palet",
        "czas_start_lbl": "Hora de inicio (Obligatorio):", "czas_koniec_lbl": "Hora de fin (Obligatorio):",
        "postep_tytul": "Progreso Total de Producción",
        "podsumowanie_dnia": "Resumen del Día (Todos los Artículos)",
        "do_zrobienia": "Total por hacer:", "wykonano": "Completado:", "pozostalo": "Restante:",
        "postep_art": "Progreso por artículos individuales:",
        "brak_planu_dzis": "No hay plan para este día.",
        "raport_wyd_tytul": "Informe de Productividad (Pcs por Hora)",
        "raport_wyd_opis": "El informe muestra el tiempo total de trabajo y la productividad promedio por hora para cada línea.",
        "th_linia": "Línea", "th_zmiana": "Turno", "th_wyprodukowano": "Pcs producidas", "th_czas_pracy": "Tiempo de trabajo (horas)", "th_wydajnosc": "Productividad (pcs/h)",
        "brak_czasow": "No hay tiempos registrados.",
        "wybierz_dzien_plan": "Seleccionar día del plan:", "podglad_planu": "Vista previa de producción planificada:", "brak_planow": "No hay planes definidos.",
        "dodaj_nowego_uzyt": "Añadir nuevo usuario / operador",
        "imie_nazwisko": "Nombre completo del operador:", "rola": "Rol:",
        "uprawnienia_lbl": "Permisos (seleccionar pestañas):", "zapisz_uzyt_btn": "Guardar usuario", "zarejestrowani_uzyt": "Usuarios registrados",
        "systemowy": "(sistema)", "kol_kod": "Código", "kol_nazwa": "Nombre", "kol_odbiorca": "Destinatario", "kol_akcja": "Acción",
        "wybierz_dzien_hist": "Seleccionar día para ver historial e informe (Admin):",
        "data_raportu_hist": "Fecha de informe / historial:", "tabela_nr": "#", "tabela_ilosc": "Cantidad", "tabela_czas": "Tiempo de trabajo",
        "tabela_bhp": "Confirmaciones SST", "tabela_szczegoly": "Detalles",
        "ph_kod": "ej. 9999", "ph_nazwa": "Nombre del producto", "ph_odbiorca": "Destinatario", "ph_ilosc": "ingrese cantidad..."
    },
    "uk": {
        "tytul": "Панель управління виробництвом",
        "wybierz_zm": "Виберіть робочу зміну:",
        "rozpocznij": "Почати роботу",
        "zmiana": "Поточна зміна",
        "zmien_zm": "Змінити зміну",
        "zalog_op": "Зареєстрований оператор",
        "linia_1": "Лінія 1", "linia_2": "Лінія 2", "linia_3": "Лінія 3", "linia_4": "Лінія 4",
        "plan_tab": "Виробничий план", "prog_tab": "Прогрес / Виконання", "wydajnosc_tab": "Продуктивність (Шт/год)", "db_art_tab": "База артикулів", "admin_tab": "Панель адміна",
        "zarejestruj": "Зареєструвати партію для",
        "data_prod": "Дата виробництва:", "kod_art": "Код артикулу:", "wybierz_kod": "-- Виберіть код артикулу --",
        "ilosc_szt": "Кількість штук:", "wpisz_ilosc": "введіть кількість...", "nazwa_ind": "Назва індексу:",
        "odbiorca": "Одержувач:", "data_przyd": "Термін придатності (мін. +4 дні):", "nr_partii": "Партія продукту:",
        "ety_gora": "Етикетка ВЕРХНЯ:", "ety_dol": "Етикетка НИЖНЯ:", "operator": "Оператор:",
        "dodaj_btn": "Зареєструвати вироблену палету / партію",
        "realizacja_biez": "Поточне виконання",
        "brak_wpisow": "Немає записів для цієї лінії та дати.", "lacznie": "Всього вироблено на лінії:",
        "zapisz_excel": "Зберегти звіт", "wklej_plan": "Управління виробничим планом",
        "instrukcja_plan": "Виберіть дату, вставте з Excel (Код [Tab] Кількість) або додайте окремо:",
        "zapisz_plan_btn": "Зберегти план на дату", "zarzadzanie_art": "Управління артикулами",
        "dodaj_art": "Додати / Редагувати артикул", "nazwa": "Назва", "dodaj_art_btn": "Зберегти артикул",
        "istniejace_art": "Існуючі артикули", "usun": "[Видалити]",
        "logowanie_tytul": "Вхід до системи",
        "brak_dostepu": "Доступ заборонено!", "wyloguj": "Вийти", "zarzadzanie_uzyt": "Управління користувачами",
        "bhp_tytul": "Необхідні підтвердження з охорони праці та контролю:",
        "bhp_1": "Підтверджую проведення перевірки працівників та чистоти обладнання перед початком роботи.",
        "bhp_2": "Підтверджую належне очищення лінії від залишкових етикеток, бандеролей та картонів перед початком виробництва наступного продукту.",
        "czas_pracy_tytul": "Робочий час (ОБОВ'ЯЗКОВО):",
        "czas_opis": "Натисніть кнопки, щоб отримати поточний час початку та кінця або введіть вручну.",
        "start_palety": "Початок палети", "koniec_palety": "Кінець палети",
        "czas_start_lbl": "Час початку (Обов'язково):", "czas_koniec_lbl": "Час кінця (Обов'язково):",
        "postep_tytul": "Загальний прогрес виробництва",
        "podsumowanie_dnia": "Підсумок дня (Всі артикули)",
        "do_zrobienia": "Загалом виконати:", "wykonano": "Виконано:", "pozostalo": "Залишилось:",
        "postep_art": "Прогрес за окремими артикулами:",
        "brak_planu_dzis": "Немає плану на цей день.",
        "raport_wyd_tytul": "Звіт про продуктивність (Шт за годину)",
        "raport_wyd_opis": "Звіт нижче показує загальний час роботи та середню продуктивність за 1 годину для кожної лінії та зміни.",
        "th_linia": "Лінія", "th_zmiana": "Зміна", "th_wyprodukowano": "Вироблено штук", "th_czas_pracy": "Час роботи (години)", "th_wydajnosc": "Продуктивність (шт/год)",
        "brak_czasow": "Немає зареєстрованого часу.",
        "wybierz_dzien_plan": "Виберіть день плану:", "podglad_planu": "Попередній перегляд спланованого виробництва:", "brak_planow": "Немає визначених планів.",
        "dodaj_nowego_uzyt": "Додати нового користувача / оператора",
        "imie_nazwisko": "ПІБ оператора:", "rola": "Роль:",
        "uprawnienia_lbl": "Права доступу (виберіть вкладки):", "zapisz_uzyt_btn": "Зберегти користувача", "zarejestrowani_uzyt": "Зареєстровані користувачі",
        "systemowy": "(системний)", "kol_kod": "Код", "kol_nazwa": "Назва", "kol_odbiorca": "Одержувач", "kol_akcja": "Дія",
        "wybierz_dzien_hist": "Виберіть день для перегляду історії та створення звіту (Адмін):",
        "data_raportu_hist": "Дата звіту / історії:", "tabela_nr": "#", "tabela_ilosc": "Кількість", "tabela_czas": "Час роботи",
        "tabela_bhp": "Підтвердження ОП", "tabela_szczegoly": "Деталі",
        "ph_kod": "напр. 9999", "ph_nazwa": "Назва продукту", "ph_odbiorca": "Одержувач", "ph_ilosc": "введіть кількість..."
    }
}

sesja_dane = {"jezyk": "pl"}

LOGIN_TEMPLATE = """
<!DOCTYPE html>
<html lang="{{ lang }}">
<head>
    <meta charset="UTF-8">
    <title>{{ t.logowanie_tytul }}</title>
    <style>
        body { font-family: Arial, sans-serif; background: #f0f2f5; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
        .login-box { background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); width: 350px; }
        h3 { text-align: center; color: #333; margin-top: 0; }
        label { font-weight: bold; display: block; margin-top: 12px; color: #444; font-size: 13px; }
        input { width: 100%; padding: 10px; margin-top: 5px; border: 1px solid #ccc; border-radius: 5px; box-sizing: border-box; font-size: 14px; }
        button { background: #007bff; color: white; border: none; padding: 12px; margin-top: 20px; cursor: pointer; width: 100%; font-size: 15px; border-radius: 5px; font-weight: bold; }
        button:hover { background: #0056b3; }
        .error { color: red; font-size: 13px; text-align: center; margin-top: 10px; }
    </style>
</head>
<body>
<div class="login-box">
    <h3>🔐 {{ t.logowanie_tytul }}</h3>
    {% if error %}
        <div class="error">{{ error }}</div>
    {% endif %}
    <form method="POST" action="/login">
        <label>{{ t.login }}</label>
        <input type="text" name="username" required autocomplete="off">
        <label>{{ t.haslo }}</label>
        <input type="password" name="password" required>
        <button type="submit">{{ t.zaloguj }}</button>
    </form>
</div>
</body>
</html>
"""

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="{{ lang }}">
<head>
    <meta charset="UTF-8">
    <title>{{ t.tytul }}</title>
    <style>
        body { font-family: Arial, sans-serif; background: #f0f2f5; margin: 0; padding: 20px; }
        .container { max-width: 1100px; background: white; padding: 25px; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); margin: auto; }
        h2, h3 { color: #333; text-align: center; }
        .header-bar { display: flex; justify-content: space-between; align-items: center; background: #e9ecef; padding: 10px 15px; border-radius: 6px; margin-bottom: 20px; }
        .lang-switcher { display: flex; gap: 5px; }
        .lang-switcher a { text-decoration: none; font-weight: bold; padding: 5px 10px; border-radius: 4px; background: #fff; border: 1px solid #ccc; color: #333; font-size: 13px; }
        .lang-switcher a.active { background: #007bff; color: white; border-color: #007bff; }
        .tabs { display: flex; gap: 5px; margin-bottom: 20px; border-bottom: 2px solid #ddd; padding-bottom: 10px; flex-wrap: wrap; }
        .tabs a { padding: 10px 18px; background: #e2e6ea; color: #333; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 14px; }
        .tabs a.active { background: #007bff; color: white; }
        .tabs a.plan-tab { background: #17a2b8; color: white; }
        .tabs a.plan-tab.active { background: #117a8b; color: white; }
        .tabs a.progress-tab { background: #28a745; color: white; }
        .tabs a.progress-tab.active { background: #1e7e34; color: white; }
        .tabs a.wydajnosc-tab { background: #6f42c1; color: white; }
        .tabs a.wydajnosc-tab.active { background: #5a32a3; color: white; }
        .tabs a.db-tab { background: #ffc107; color: #333; }
        .tabs a.db-tab.active { background: #d39e00; color: white; }
        .tabs a.admin-tab { background: #dc3545; color: white; }
        .tabs a.admin-tab.active { background: #bd2130; color: white; }
        label { font-weight: bold; display: block; margin-top: 10px; color: #444; font-size: 13px; }
        input, select, textarea { width: 100%; padding: 8px; margin-top: 4px; border: 1px solid #ccc; border-radius: 5px; box-sizing: border-box; font-size: 14px; }
        .row { display: flex; gap: 15px; }
        .row > div { flex: 1; }
        button { background: #28a745; color: white; border: none; padding: 12px; margin-top: 15px; cursor: pointer; width: 100%; font-size: 15px; border-radius: 5px; font-weight: bold; }
        button:hover { background: #218838; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 13px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f8f9fa; }
        .summary-box { background: #e9ecef; padding: 12px; border-radius: 5px; margin-top: 20px; font-weight: bold; text-align: center; font-size: 16px; }
        .plan-info-box { background: #e8f4fd; border: 1px solid #b8daff; padding: 10px 15px; border-radius: 6px; margin-top: 8px; font-size: 13px; color: #004085; display: none; }
        .plan-section { background: #e2f0d9; padding: 20px; border-radius: 8px; border: 1px solid #c8e1b9; }
        .progress-section { background: #f8f9fa; padding: 20px; border-radius: 8px; border: 1px solid #ddd; }
        .wydajnosc-section { background: #f3eef8; padding: 20px; border-radius: 8px; border: 1px solid #dcd0e8; }
        .db-section { background: #fff3cd; padding: 20px; border-radius: 8px; border: 1px solid #ffeeba; }
        .admin-section { background: #f8d7da; padding: 20px; border-radius: 8px; border: 1px solid #f5c6cb; }
        .error-banner { background: #f8d7da; color: #721c24; padding: 12px; border-radius: 6px; margin-bottom: 15px; border: 1px solid #f5c6cb; font-weight: bold; text-align: center; }
        
        /* Styl tabeli postępu i nagłówków sortowalnych */
        .progress-table { width: 100%; border-collapse: collapse; margin-top: 15px; background: white; font-size: 13px; }
        .progress-table th, .progress-table td { border: 1px solid #cbd5e1; padding: 8px 10px; text-align: left; }
        .progress-table th { background-color: #f1f5f9; cursor: pointer; user-select: none; }
        .progress-table th:hover { background-color: #e2e8f0; }
        .filter-input { width: 100%; padding: 5px 6px; margin-top: 5px; font-size: 12px; border: 1px solid #94a3b8; border-radius: 4px; box-sizing: border-box; font-weight: normal; cursor: text; }
    </style>
    <script>
        const baza = {{ produkty_json | safe }};
        const calyPlan = {{ plan_json | safe }};
        const zrealizowanoDzien = {{ zrealizowane_json | safe }};
        
        function aktualizujDane() {
            const kod = document.getElementById('kodSelect').value;
            const nazwaInput = document.getElementById('nazwaInput');
            const odbiorcaInput = document.getElementById('odbiorcaInput');
            const infoBox = document.getElementById('planInfoBox');
            
            const dataProdInput = document.getElementById('dataProdInput');
            const dataProd = dataProdInput ? dataProdInput.value : "{{ dzis }}";

            if (kod && baza[kod]) {
                nazwaInput.value = baza[kod].nazwa;
                odbiorcaInput.value = baza[kod].odbiorca;
                
                const planNaDzien = calyPlan[dataProd] || {};
                if (planNaDzien[kod] !== undefined) {
                    const planKwota = planNaDzien[kod];
                    const wykonaneKwota = (zrealizowanoDzien[dataProd] && zrealizowanoDzien[dataProd][kod]) ? zrealizowanoDzien[dataProd][kod] : (zrealizowanoDzien[kod] || 0);
                    const pozostalo = Math.max(planKwota - wykonaneKwota, 0);

                    infoBox.style.display = 'block';
                    infoBox.innerHTML = '📋 <b>Plan (' + dataProd + '):</b> ' + planKwota + ' | ✅ <b>Wykonano:</b> ' + wykonaneKwota + ' | ⚠️ <b>Pozostało:</b> <b>' + pozostalo + '</b>';
                } else {
                    infoBox.style.display = 'block';
                    infoBox.innerHTML = 'ℹ️ Brak zdefiniowanego planu na dzień <b>' + dataProd + '</b> dla tego kodu.';
                }
            } else {
                nazwaInput.value = "";
                odbiorcaInput.value = "";
                infoBox.style.display = 'none';
            }
        }

        // Filtrowanie tabeli postępu kolumna po kolumnie
        function filtrKolumnTabeli() {
            const fKod = document.getElementById('filterKod').value.toLowerCase();
            const fNazwa = document.getElementById('filterNazwa').value.toLowerCase();
            const fPlan = document.getElementById('filterPlan').value.toLowerCase();
            const fWykonano = document.getElementById('filterWykonano').value.toLowerCase();
            const fPozostalo = document.getElementById('filterPozostalo').value.toLowerCase();
            const fProcent = document.getElementById('filterProcent').value.toLowerCase();

            const rows = document.getElementById('progressTableBody').getElementsByTagName('tr');

            for (let i = 0; i < rows.length; i++) {
                const tdKod = rows[i].getAttribute('data-kod') || '';
                const tdNazwa = rows[i].getAttribute('data-nazwa') || '';
                const tdPlan = rows[i].getAttribute('data-plan') || '';
                const tdWykonano = rows[i].getAttribute('data-wykonano') || '';
                const tdPozostalo = rows[i].getAttribute('data-pozostalo') || '';
                const tdProcent = rows[i].getAttribute('data-procent') || '';

                const match = 
                    tdKod.toLowerCase().includes(fKod) &&
                    tdNazwa.toLowerCase().includes(fNazwa) &&
                    tdPlan.toLowerCase().includes(fPlan) &&
                    tdWykonano.toLowerCase().includes(fWykonano) &&
                    tdPozostalo.toLowerCase().includes(fPozostalo) &&
                    tdProcent.toLowerCase().includes(fProcent);

                rows[i].style.display = match ? "" : "none";
            }
        }

        // Funkcja sortowania tabeli rosnąco / malejąco po kliknięciu w nagłówek
        let sortDirections = {};
        function sortTable(columnIndex, isNumeric = false) {
            const table = document.getElementById("progressTable");
            const tbody = document.getElementById("progressTableBody");
            const rowsArray = Array.from(tbody.querySelectorAll("tr"));
            
            // Określenie kierunku sortowania (przełączanie między rosnąco a malejąco)
            let dir = sortDirections[columnIndex] || 'asc';
            dir = dir === 'asc' ? 'desc' : 'asc';
            sortDirections[columnIndex] = dir;

            rowsArray.sort((rowA, rowB) => {
                let cellA = rowA.children[columnIndex].innerText.trim();
                let cellB = rowB.children[columnIndex].innerText.trim();

                if (isNumeric) {
                    cellA = parseFloat(cellA.replace(/[^0-9.-]+/g,"")) || 0;
                    cellB = parseFloat(cellB.replace(/[^0-9.-]+/g,"")) || 0;
                    return dir === 'asc' ? cellA - cellB : cellB - cellA;
                } else {
                    return dir === 'asc' ? cellA.localeCompare(cellB) : cellB.localeCompare(cellA);
                }
            });

            // Ponowne wstawienie posortowanych wierszy do tabeli
            rowsArray.forEach(row => tbody.appendChild(row));
        }

        window.onload = function() {
            const dataInput = document.getElementById('dataProdInput');
            if (dataInput) {
                dataInput.addEventListener('change', aktualizujDane);
            }
            const kodSelect = document.getElementById('kodSelect');
            if (kodSelect && kodSelect.value) {
                aktualizujDane();
            }
        };

        function walidujFormularz(event) {
            const start = document.getElementById('czasStart').value;
            const koniec = document.getElementById('czasKoniec').value;
            if (start && koniec && koniec <= start) {
                alert("Błąd: Czas zakończenia palety nie może być wcześniejszy lub równy czasowi rozpoczęcia!");
                event.preventDefault();
                return false;
            }
            return true;
        }
    </script>
</head>
<body>
<div class="container">
    <div class="header-bar">
        <div>
            <b>{{ t.zalog_op }}:</b> {{ pelna_nazwa }} ({{ role }}) | 
            {% if wypr_zmiana %}
                <b>{{ t.zmiana }}:</b> {{ wypr_zmiana }} | 
                <a href="/reset-zmiana" style="color: #dc3545; font-weight: bold; text-decoration: none;">🔄 {{ t.zmien_zm }}</a>
            {% endif %}
        </div>
        <div style="display: flex; gap: 10px; align-items: center;">
            <div class="lang-switcher">
                <a href="/ustaw-jezyk/pl" class="{{ 'active' if lang == 'pl' else '' }}">PL</a>
                <a href="/ustaw-jezyk/en" class="{{ 'active' if lang == 'en' else '' }}">EN</a>
                <a href="/ustaw-jezyk/es" class="{{ 'active' if lang == 'es' else '' }}">ES</a>
                <a href="/ustaw-jezyk/uk" class="{{ 'active' if lang == 'uk' else '' }}">UA</a>
            </div>
            <a href="/logout" style="background: #dc3545; color: white; padding: 5px 10px; border-radius: 4px; text-decoration: none; font-size: 13px; font-weight: bold;">{{ t.wyloguj }}</a>
        </div>
    </div>

    {% if not wypr_zmiana %}
        <div style="text-align: center; padding: 40px;">
            <h2>{{ t.wybierz_zm }}</h2>
            <form method="POST" action="/wybierz-zmiane">
                <select name="zmiana" style="max-width: 300px; margin: 0 auto 15px auto;">
                    <option value="1">Zmiana 1</option>
                    <option value="2">Zmiana 2</option>
                </select>
                <button type="submit" style="max-width: 300px; margin: 0 auto;">{{ t.rozpocznij }}</button>
            </form>
        </div>
    {% else %}
        <h2>🏭 {{ t.tytul }}</h2>

        <div class="tabs">
            {% if '1' in dostep %}
                <a href="/?linia=1" class="{{ 'active' if linia == '1' else '' }}">{{ t.linia_1 }}</a>
            {% endif %}
            {% if '2' in dostep %}
                <a href="/?linia=2" class="{{ 'active' if linia == '2' else '' }}">{{ t.linia_2 }}</a>
            {% endif %}
            {% if '3' in dostep %}
                <a href="/?linia=3" class="{{ 'active' if linia == '3' else '' }}">{{ t.linia_3 }}</a>
            {% endif %}
            {% if '4' in dostep %}
                <a href="/?linia=4" class="{{ 'active' if linia == '4' else '' }}">{{ t.linia_4 }}</a>
            {% endif %}
            {% if 'plan' in dostep %}
                <a href="/?linia=plan" class="plan-tab {{ 'active' if linia == 'plan' else '' }}">{{ t.plan_tab }}</a>
            {% endif %}
            {% if 'progress' in dostep %}
                <a href="/?linia=progress" class="progress-tab {{ 'active' if linia == 'progress' else '' }}">{{ t.prog_tab }}</a>
            {% endif %}
            {% if 'wydajnosc' in dostep %}
                <a href="/?linia=wydajnosc" class="wydajnosc-tab {{ 'active' if linia == 'wydajnosc' else '' }}">{{ t.wydajnosc_tab }}</a>
            {% endif %}
            {% if 'db_art' in dostep %}
                <a href="/?linia=db_art" class="db-tab {{ 'active' if linia == 'db_art' else '' }}">{{ t.db_art_tab }}</a>
            {% endif %}
            {% if 'admin' in dostep %}
                <a href="/?linia=admin" class="admin-tab {{ 'active' if linia == 'admin' else '' }}">{{ t.admin_tab }}</a>
            {% endif %}
        </div>

        {% if error_msg %}
            <div class="error-banner">⚠️ {{ error_msg }}</div>
        {% endif %}

        {% if linia == 'plan' and 'plan' in dostep %}
            <div class="plan-section">
                <h3 style="margin-top: 0; color: #385723;">📋 {{ t.wklej_plan }}</h3>
                <form action="/zapisz-plan" method="POST" style="margin-bottom: 20px;">
                    <label><b>{{ t.wybierz_dzien_plan }}</b></label>
                    <input type="date" name="data_planu" value="{{ wybrana_data_planu or dzis }}" style="max-width: 250px; margin-bottom: 15px;" required>
                    
                    <p style="font-size: 13px; color: #444;">{{ t.instrukcja_plan }}</p>
                    <textarea name="surowy_plan" rows="8" placeholder="Wklej skopiowane kolumny z Excela (Kod [Tab] Ilość)..." style="font-family: monospace; font-size: 13px;"></textarea>
                    <button type="submit" style="background: #385723;">{{ t.zapisz_plan_btn }}</button>
                </form>

                <h4 style="border-top: 1px solid #c8e1b9; padding-top: 15px;">{{ t.podglad_planu }}</h4>
                
                <form method="GET" action="/" style="margin-bottom: 15px; display: flex; gap: 10px; align-items: flex-end;">
                    <input type="hidden" name="linia" value="plan">
                    <div style="flex: 1; max-width: 300px;">
                        <label><b>Wybierz dzień do podglądu:</b></label>
                        <select name="wybrana_data_planu" onchange="this.form.submit()">
                            {% for data_p in dostepne_daty_planu %}
                                <option value="{{ data_p }}" {{ 'selected' if data_p == wybrana_data_planu else '' }}>📅 {{ data_p }}</option>
                            {% endfor %}
                        </select>
                    </div>
                </form>

                {% if wybrana_data_planu and wybrana_data_planu in plan_produkcyjny %}
                    <div style="background: white; padding: 15px; border-radius: 6px; margin-bottom: 10px; border: 1px solid #c8e1b9;">
                        <b>📅 Dzień: {{ wybrana_data_planu }}</b>
                        <ul style="margin: 10px 0 0 20px; padding: 0; font-size: 13px;">
                            {% for kod, szt in plan_produkcyjny[wybrana_data_planu].items() %}
                                <li style="margin-bottom: 5px;"><b>{{ kod }}</b> - {{ produkty_baza.get(kod, {}).get('nazwa', 'Nieznany') }}: <b>{{ szt }} szt.</b> 
                                    (<a href="/usun-z-planu?data={{ wybrana_data_planu }}&kod={{ kod }}" style="color: red; text-decoration: none;">{{ t.usun }}</a>)
                                </li>
                            {% endfor %}
                        </ul>
                    </div>
                {% else %}
                    <p style="color: #666; font-size: 13px;">{{ t.brak_planow }}</p>
                {% endif %}
            </div>
        {% elif linia == 'progress' and 'progress' in dostep %}
            <div class="progress-section">
                <h3 style="margin-top: 0; color: #333;">📊 {{ t.postep_tytul }}</h3>
                
                <form method="GET" action="/" style="margin-bottom: 20px; display: flex; gap: 10px; align-items: flex-end;">
                    <input type="hidden" name="linia" value="progress">
                    <div style="flex: 1; max-width: 300px;">
                        <label><b>Wybierz dzień do podglądu postępu:</b></label>
                        <select name="wybrana_data_progress" onchange="this.form.submit()">
                            {% for data_d in dostepne_daty_wpisow_lub_planu %}
                                <option value="{{ data_d }}" {{ 'selected' if data_d == wybrana_data_progress else '' }}>📅 {{ data_d }}</option>
                            {% endfor %}
                        </select>
                    </div>
                </form>

                <div style="background: white; padding: 20px; border-radius: 8px; margin-bottom: 25px; border: 2px solid #28a745; box-shadow: 0 2px 6px rgba(0,0,0,0.05);">
                    <div style="font-size: 16px; font-weight: bold; margin-bottom: 8px; color: #155724; display: flex; justify-content: space-between;">
                        <span>🚀 {{ t.podsumowanie_dnia }} (Dzień: {{ wybrana_data_progress }})</span>
                        <span><b>{{ suma_wykonana_calkowita_prog }}</b> / {{ suma_planu_calkowita_prog }} szt. (<b>{{ procent_calkowity_prog }}%</b>)</span>
                    </div>
                    <div style="background: #e9ecef; border-radius: 8px; height: 26px; overflow: hidden; border: 1px solid #ced4da; margin-bottom: 10px;">
                        <div style="background: {{ '#28a745' if procent_calkowity_prog >= 100 else '#007bff' }}; width: {{ [procent_calkowity_prog, 100] | min }}%; height: 100%; display: flex; align-items: center; justify-content: center; color: white; font-weight: bold; font-size: 12px;">
                            {{ procent_calkowity_prog }}%
                        </div>
                    </div>
                    <div style="font-size: 14px; color: #333; display: flex; justify-content: space-between;">
                        <span>📋 {{ t.do_zrobienia }} <b>{{ suma_planu_calkowita_prog }} szt.</b></span>
                        <span>✅ {{ t.wykonano }} <b>{{ suma_wykonana_calkowita_prog }} szt.</b></span>
                        <span>⚠️ {{ t.pozostalo }} <b style="color: #dc3545;">{{ pozostalo_calkowite_prog }} szt.</b></span>
                    </div>
                </div>

                <h4 style="margin-top: 20px; color: #333;">{{ t.postep_art }} (Kliknij nagłówek kolumny, aby sortować ↕️):</h4>
                
                <div style="max-height: 500px; overflow-y: auto; background: white; border-radius: 6px; border: 1px solid #cbd5e1;">
                    <table class="progress-table" id="progressTable">
                        <thead>
                            <tr>
                                <th onclick="sortTable(0, false)">
                                    Kod artykułu ↕️
                                    <input type="text" id="filterKod" class="filter-input" onkeyup="filtrKolumnTabeli()" onclick="event.stopPropagation()" placeholder="Filtruj kod...">
                                </th>
                                <th onclick="sortTable(1, false)">
                                    Nazwa produktu ↕️
                                    <input type="text" id="filterNazwa" class="filter-input" onkeyup="filtrKolumnTabeli()" onclick="event.stopPropagation()" placeholder="Filtruj nazwę...">
                                </th>
                                <th style="width: 100px;" onclick="sortTable(2, true)">
                                    Plan ↕️
                                    <input type="text" id="filterPlan" class="filter-input" onkeyup="filtrKolumnTabeli()" onclick="event.stopPropagation()" placeholder="Plan...">
                                </th>
                                <th style="width: 100px;" onclick="sortTable(3, true)">
                                    Wykonano ↕️
                                    <input type="text" id="filterWykonano" class="filter-input" onkeyup="filtrKolumnTabeli()" onclick="event.stopPropagation()" placeholder="Wykonano...">
                                </th>
                                <th style="width: 100px;" onclick="sortTable(4, true)">
                                    Pozostało ↕️
                                    <input type="text" id="filterPozostalo" class="filter-input" onkeyup="filtrKolumnTabeli()" onclick="event.stopPropagation()" placeholder="Pozostało...">
                                </th>
                                <th style="width: 130px;" onclick="sortTable(5, true)">
                                    % Realizacji ↕️
                                    <input type="text" id="filterProcent" class="filter-input" onkeyup="filtrKolumnTabeli()" onclick="event.stopPropagation()" placeholder="%...">
                                </th>
                            </tr>
                        </thead>
                        <tbody id="progressTableBody">
                            {% for kod, plan_val in plan_wybrany_prog.items() %}
                                {% set wykonane = wykonane_dla_wybranej_daty_prog.get(kod, 0) %}
                                {% set left_amt = [plan_val - wykonane, 0] | max %}
                                {% set procent = (wykonane / plan_val * 100) | round | int if plan_val > 0 else 0 %}
                                {% set nazwa_art = produkty_baza.get(kod, {}).get('nazwa', 'Nieznany') %}
                                
                                <tr data-kod="{{ kod }}" data-nazwa="{{ nazwa_art }}" data-plan="{{ plan_val }}" data-wykonano="{{ wykonane }}" data-pozostalo="{{ left_amt }}" data-procent="{{ procent }}%">
                                    <td><b>{{ kod }}</b></td>
                                    <td>{{ nazwa_art }}</td>
                                    <td>{{ plan_val }} szt.</td>
                                    <td><b>{{ wykonane }}</b> szt.</td>
                                    <td><b style="color: #dc3545;">{{ left_amt }}</b> szt.</td>
                                    <td>
                                        <div style="display: flex; align-items: center; gap: 8px;">
                                            <span><b>{{ procent }}%</b></span>
                                            <div style="background: #e9ecef; border-radius: 4px; height: 10px; flex: 1; overflow: hidden; border: 1px solid #ced4da;">
                                                <div style="background: {{ '#28a745' if procent >= 100 else '#007bff' }}; width: {{ [procent, 100] | min }}%; height: 100%;"></div>
                                            </div>
                                        </div>
                                    </td>
                                </tr>
                            {% else %}
                                <tr><td colspan="6" style="text-align: center; color: #777; padding: 20px;">{{ t.brak_planu_dzis }}</td></tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>

                <form action="/generate?data={{ wybrana_data_progress }}&linia=1" method="POST" style="margin-top: 25px;">
                    <button type="submit" style="background: #28a745;">📥 {{ t.zapisz_excel }} (Postęp / Dzień: {{ wybrana_data_progress }})</button>
                </form>
            </div>
        {% elif linia == 'wydajnosc' and 'wydajnosc' in dostep %}
            <div class="wydajnosc-section">
                <h3 style="margin-top: 0; color: #5a32a3;">⏱️ {{ t.raport_wyd_tytul }}</h3>
                <p style="font-size: 13px; color: #555;">{{ t.raport_wyd_opis }}</p>
                
                <form method="GET" action="/" style="margin-bottom: 20px; display: flex; gap: 10px; align-items: flex-end;">
                    <input type="hidden" name="linia" value="wydajnosc">
                    <div style="flex: 1; max-width: 300px;">
                        <label><b>Wybierz dzień do podglądu wydajności:</b></label>
                        <select name="wybrana_data_wydajnosc" onchange="this.form.submit()">
                            {% for data_d in dostepne_daty_wpisow %}
                                <option value="{{ data_d }}" {{ 'selected' if data_d == wybrana_data_wydajnosc else '' }}>📅 {{ data_d }}</option>
                            {% endfor %}
                        </select>
                    </div>
                </form>

                <table>
                    <tr>
                        <th>{{ t.th_linia }}</th>
                        <th>{{ t.th_zmiana }}</th>
                        <th>{{ t.th_wyprodukowano }}</th>
                        <th>{{ t.th_czas_pracy }}</th>
                        <th>{{ t.th_wydajnosc }}</th>
                    </tr>
                    {% for klucz, stat in statystyki_wydajnosci_wybrana.items() %}
                    <tr>
                        <td><b>{{ t.linia_1.split()[0] }} {{ stat.linia }}</b></td>
                        <td>{{ t.th_zmiana }} {{ stat.zmiana }}</td>
                        <td><b>{{ stat.suma_sztuk }} szt.</b></td>
                        <td>{{ stat.czas_godz }} h</td>
                        <td><b style="color: #6f42c1; font-size: 15px;">{{ stat.sztuk_na_godzine }} szt/h</b></td>
                    </tr>
                    {% else %}
                    <tr><td colspan="5" style="text-align: center; color: #777;">{{ t.brak_czasow }}</td></tr>
                    {% endfor %}
                </table>

                <form action="/generate?data={{ wybrana_data_wydajnosc }}&linia=1" method="POST" style="margin-top: 25px;">
                    <button type="submit" style="background: #6f42c1;">📥 {{ t.zapisz_excel }} (Wydajność / Dzień: {{ wybrana_data_wydajnosc }})</button>
                </form>
            </div>
        {% elif linia == 'db_art' and 'db_art' in dostep %}
            <div class="db-section">
                <h3 style="margin-top: 0; color: #856404;">📦 {{ t.zarzadzanie_art }}</h3>
                <form action="/dodaj-artykul" method="POST" style="margin-bottom: 25px; background: white; padding: 15px; border-radius: 6px;">
                    <div class="row">
                        <div>
                            <label>{{ t.kol_kod }}:</label>
                            <input type="text" name="kod" placeholder="{{ t.ph_kod }}" required>
                        </div>
                        <div>
                            <label>{{ t.kol_nazwa }}:</label>
                            <input type="text" name="nazwa" placeholder="{{ t.ph_nazwa }}" required>
                        </div>
                        <div>
                            <label>{{ t.kol_odbiorca }}:</label>
                            <input type="text" name="odbiorca" placeholder="{{ t.ph_odbiorca }}" required>
                        </div>
                    </div>
                    <button type="submit" style="background: #ffc107; color: #000; margin-top: 10px;">{{ t.dodaj_art_btn }}</button>
                </form>
                <h4>{{ t.istniejace_art }} ({{ produkty_baza|length }})</h4>
                <div style="max-height: 400px; overflow-y: auto; background: white; border-radius: 6px;">
                    <table>
                        <tr><th>{{ t.kol_kod }}</th><th>{{ t.kol_nazwa }}</th><th>{{ t.kol_odbiorca }}</th><th>{{ t.kol_akcja }}</th></tr>
                        {% for kod, info in produkty_baza.items() %}
                        <tr>
                            <td><b>{{ kod }}</b></td>
                            <td>{{ info.nazwa }}</td>
                            <td>{{ info.odbiorca }}</td>
                            <td><a href="/usun-artykul/{{ kod }}" style="color: red; text-decoration: none;">{{ t.usun }}</a></td>
                        </tr>
                        {% endfor %}
                    </table>
                </div>
            </div>
        {% elif linia == 'admin' and 'admin' in dostep %}
            <div class="admin-section">
                <h3 style="margin-top: 0; color: #721c24;">🔐 {{ t.zarzadzanie_uzyt }}</h3>
                
                <form action="/admin/dodaj-uzytkownika" method="POST" style="background: white; padding: 15px; border-radius: 6px; margin-bottom: 25px;">
                    <h4 style="margin-top:0;">{{ t.dodaj_nowego_uzyt }}</h4>
                    <div class="row">
                        <div>
                            <label>{{ t.login }}</label>
                            <input type="text" name="username" placeholder="np. anna" required>
                        </div>
                        <div>
                            <label>{{ t.haslo }}</label>
                            <input type="password" name="password" placeholder="hasło" required>
                        </div>
                        <div>
                            <label>{{ t.imie_nazwisko }}</label>
                            <input type="text" name="pelna_nazwa" placeholder="np. Anna Nowak" required>
                        </div>
                        <div>
                            <label>{{ t.rola }}</label>
                            <select name="rola">
                                <option value="operator">operator</option>
                                <option value="admin">admin</option>
                            </select>
                        </div>
                    </div>
                    <label>{{ t.uprawnienia_lbl }}</label>
                    <div style="display: flex; gap: 15px; margin-top: 5px; flex-wrap: wrap; font-size: 13px;">
                        <label><input type="checkbox" name="dostep" value="1" checked> {{ t.linia_1 }}</label>
                        <label><input type="checkbox" name="dostep" value="2" checked> {{ t.linia_2 }}</label>
                        <label><input type="checkbox" name="dostep" value="3" checked> {{ t.linia_3 }}</label>
                        <label><input type="checkbox" name="dostep" value="4" checked> {{ t.linia_4 }}</label>
                        <label><input type="checkbox" name="dostep" value="plan"> {{ t.plan_tab }}</label>
                        <label><input type="checkbox" name="dostep" value="progress" checked> {{ t.prog_tab }}</label>
                        <label><input type="checkbox" name="dostep" value="wydajnosc" checked> {{ t.wydajnosc_tab }}</label>
                        <label><input type="checkbox" name="dostep" value="db_art"> {{ t.db_art_tab }}</label>
                        <label><input type="checkbox" name="dostep" value="admin"> {{ t.admin_tab }}</label>
                    </div>
                    <button type="submit" style="background: #dc3545; color: white; margin-top: 15px;">{{ t.zapisz_uzyt_btn }}</button>
                </form>

                <h4>{{ t.zarejestrowani_uzyt }}:</h4>
                <div style="background: white; padding: 15px; border-radius: 6px;">
                    <table>
                        <tr><th>{{ t.login }}</th><th>{{ t.imie_nazwisko }}</th><th>{{ t.rola }}</th><th>{{ t.uprawnienia_lbl.split()[0] }}</th><th>{{ t.kol_akcja }}</th></tr>
                        {% for usr, info in uzytkownicy_baza.items() %}
                        <tr>
                            <td><b>{{ usr }}</b></td>
                            <td>{{ info.pelna_nazwa }}</td>
                            <td>{{ info.rola }}</td>
                            <td>{{ info.dostep | join(', ') }}</td>
                            <td>
                                {% if usr != 'admin' %}
                                    <a href="/admin/usun-uzytkownika/{{ usr }}" style="color: red; text-decoration: none;">{{ t.usun }}</a>
                                {% else %}
                                    <i>{{ t.systemowy }}</i>
                                {% endif %}
                            </td>
                        </tr>
                        {% endfor %}
                    </table>
                </div>
            </div>
        {% elif linia in ['1', '2', '3', '4'] and linia in dostep %}
            <form method="POST" action="/dodaj?linia={{ linia }}" onsubmit="return walidujFormularz(event)">
                <h3>➕ {{ t.zarejestruj }} {{ t.linia_1.split()[0] }} {{ linia }}</h3>
                <div class="row">
                    <div>
                        <label>{{ t.data_prod }}</label>
                        {% if role == 'admin' %}
                            <input type="date" id="dataProdInput" name="data" value="{{ wybrana_data }}" required onchange="aktualizujDane()">
                        {% else %}
                            <input type="date" id="dataProdInput" name="data" value="{{ dzis }}" readonly style="background: #e9ecef;" title="Operator może rejestrować tylko na dzisiejszy dzień">
                        {% endif %}
                    </div>
                    <div>
                        <label>{{ t.kod_art }}</label>
                        <select id="kodSelect" name="kod" onchange="aktualizujDane()" required>
                            <option value="">{{ t.wybierz_kod }}</option>
                            {% for kod, info in produkty_baza.items() %}
                                <option value="{{ kod }}">{{ kod }} - {{ info.nazwa }} ({{ info.odbiorca }})</option>
                            {% endfor %}
                        </select>
                    </div>
                    <div>
                        <label>{{ t.ilosc_szt }}</label>
                        <input type="number" name="ilosc" min="1" required placeholder="{{ t.ph_ilosc }}">
                    </div>
                </div>

                <div id="planInfoBox" class="plan-info-box"></div>

                <div class="row">
                    <div style="flex: 2;">
                        <label>{{ t.nazwa_ind }}</label>
                        <input type="text" id="nazwaInput" name="nazwa" required readonly style="background: #e9ecef;">
                    </div>
                    <div style="flex: 2;">
                        <label>{{ t.odbiorca }}</label>
                        <input type="text" id="odbiorcaInput" name="odbiorca" required readonly style="background: #e9ecef;">
                    </div>
                    <div style="flex: 1;">
                        <label>{{ t.data_przyd }}</label>
                        <input type="date" name="data_przyd" min="{{ min_data }}" required>
                    </div>
                </div>

                <div class="row">
                    <div>
                        <label>{{ t.nr_partii }}</label>
                        <input type="text" name="partia_prod" value="{{ domyslna_partia }}" required>
                    </div>
                    <div>
                        <label>{{ t.ety_gora }}</label>
                        <input type="text" name="etykieta_gora" placeholder="{{ t.ph_etygora }}" required>
                    </div>
                    <div>
                        <label>{{ t.ety_dol }}</label>
                        <input type="text" name="etykieta_dol" placeholder="{{ t.ph_etydol }}" required>
                    </div>
                    <div>
                        <label>{{ t.operator }}</label>
                        <input type="text" name="operator" value="{{ pelna_nazwa }}" readonly style="background: #e9ecef; font-weight: bold;">
                    </div>
                </div>

                <div style="background: #e8f4fd; padding: 15px; border-radius: 6px; margin-top: 15px; border: 1px solid #b8daff;">
                    <label style="margin-top:0; color: #004085; font-size: 14px;">🛡️ {{ t.bhp_tytul }}</label>
                    <div style="margin-top: 10px; display: flex; flex-direction: column; gap: 8px; font-size: 13px;">
                        <label style="font-weight: normal; display: flex; align-items: center; gap: 8px; cursor: pointer;">
                            <input type="checkbox" name="potwierdzenie_kontroli" required style="width: auto; margin:0;">
                            <span>{{ t.bhp_1 }}</span>
                        </label>
                        <label style="font-weight: normal; display: flex; align-items: center; gap: 8px; cursor: pointer;">
                            <input type="checkbox" name="potwierdzenie_oczyszczenia" required style="width: auto; margin:0;">
                            <span>{{ t.bhp_2 }}</span>
                        </label>
                    </div>
                </div>

                <div style="background: #f8f9fa; padding: 12px; border-radius: 6px; margin-top: 15px; border: 1px dashed #007bff; display: flex; gap: 15px; align-items: center;">
                    <div style="flex: 1;">
                        <label style="margin-top:0; color: #004085;">⏱️ {{ t.czas_pracy_tytul }}</label>
                        <span style="font-size: 12px; color: #666;">{{ t.czas_opis }}</span>
                    </div>
                    <div style="display: flex; gap: 10px;">
                        <button type="button" onclick="ustawCzas('start')" style="background: #17a2b8; margin-top:0; padding: 8px 15px; font-size: 13px;">▶️ {{ t.start_palety }}</button>
                        <button type="button" onclick="ustawCzas('koniec')" style="background: #6c757d; margin-top:0; padding: 8px 15px; font-size: 13px;">⏹️ {{ t.koniec_palety }}</button>
                    </div>
                </div>
                <div class="row" style="margin-top: 10px;">
                    <div>
                        <label>{{ t.czas_start_lbl }}</label>
                        <input type="time" id="czasStart" name="czas_start" required>
                    </div>
                    <div>
                        <label>{{ t.czas_koniec_lbl }}</label>
                        <input type="time" id="czasKoniec" name="czas_koniec" required>
                    </div>
                </div>

                <script>
                    function ustawCzas(typ) {
                        const now = new Date();
                        const godziny = String(now.getHours()).padStart(2, '0');
                        const minuty = String(now.getMinutes()).padStart(2, '0');
                        const czasStr = godziny + ':' + minuty;
                        if (typ === 'start') {
                            document.getElementById('czasStart').value = czasStr;
                        } else {
                            document.getElementById('czasKoniec').value = czasStr;
                        }
                    }
                </script>

                <button type="submit">{{ t.dodaj_btn }}</button>
            </form>

            <hr style="margin: 25px 0;">

            {% if role == 'admin' %}
                <h3>📊 {{ t.wybierz_dzien_hist }}</h3>
                <form method="GET" action="/" style="display: flex; gap: 10px; align-items: flex-end; margin-bottom: 15px;">
                    <input type="hidden" name="linia" value="{{ linia }}">
                    <div style="flex: 1;">
                        <label>{{ t.data_raportu_hist }}</label>
                        <input type="date" name="wybrana_data" value="{{ wybrana_data }}" onchange="this.form.submit()">
                    </div>
                </form>
            {% endif %}

            <h3>📊 {{ t.realizacja_biez }} {{ t.linia_1.split()[0] }} {{ linia }} (dla dnia: {{ wybrana_data }})</h3>
            <table>
                <tr>
                    <th>{{ t.tabela_nr }}</th>
                    <th>{{ t.kol_kod }}</th>
                    <th>{{ t.kol_nazwa }}</th>
                    <th>{{ t.tabela_ilosc }}</th>
                    <th>{{ t.tabela_czas }}</th>
                    <th>{{ t.tabela_bhp }}</th>
                    <th>{{ t.tabela_szczegoly }}</th>
                    <th>{{ t.kol_akcja }}</th>
                </tr>
                {% for wpis in wpisy_dla_linii %}
                <tr>
                    <td>{{ loop.index }}</td>
                    <td>{{ wpis.kod }}</td>
                    <td>{{ wpis.nazwa }}</td>
                    <td><b>{{ wpis.ilosc }}</b> szt.</td>
                    <td>{{ wpis.czas_start }} - {{ wpis.czas_koniec }}</td>
                    <td>✅ Sprawdzono (BHP + Linia)</td>
                    <td>Data przyd: {{ wpis.data_przyd }} | Partia: {{ wpis.partia_prod }} | Op: {{ wpis.operator }}</td>
                    <td><a href="/usun/{{ wpis.id }}?data={{ wybrana_data }}" style="color: red; text-decoration: none;">{{ t.usun }}</a></td>
                </tr>
                {% else %}
                <tr><td colspan="8" style="text-align: center; color: #777;">{{ t.brak_wpisow }}</td></tr>
                {% endfor %}
            </table>

            <div class="summary-box">
                {{ t.lacznie }} {{ suma_sztuk }} szt.
            </div>

            <form action="/generate?data={{ wybrana_data }}&linia={{ linia }}" method="POST">
                <button type="submit" style="background: #007bff; margin-top: 20px;">{{ t.zapisz_excel }} (Dzień: {{ wybrana_data }})</button>
            </form>
        {% else %}
            <div style="text-align: center; padding: 40px; color: #dc3545;">
                <h3>⚠️ {{ t.brak_dostepu }}</h3>
            </div>
        {% endif %}
    {% endif %}
</div>
</body>
</html>
"""

@app.route('/login', methods=['GET', 'POST'])
def login():
    lang = sesja_dane.get("jezyk", "pl")
    t = T[lang]
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        user_db = UzytkownikModel.query.filter_by(username=username, haslo=password).first()
        if user_db:
            session['user'] = user_db.username
            session['rola'] = user_db.rola
            session['pelna_nazwa'] = user_db.pelna_nazwa
            session['dostep'] = json.loads(user_db.dostep)
            return redirect(url_for('index'))
        else:
            return render_template_string(LOGIN_TEMPLATE, t=t, lang=lang, error="Nieprawidłowy login lub hasło!")
    return render_template_string(LOGIN_TEMPLATE, t=t, lang=lang, error=None)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/ustaw-jezyk/<lang>')
def ustaw_jezyk(lang):
    if lang in T:
        sesja_dane["jezyk"] = lang
    return redirect(request.referrer or url_for('index'))

@app.route('/')
def index():
    if 'user' not in session:
        return redirect(url_for('login'))

    lang = sesja_dane.get("jezyk", "pl")
    t = T[lang]
    user = session.get('user')
    role = session.get('rola')
    pelna_nazwa = session.get('pelna_nazwa')
    dostep = session.get('dostep', [])

    wypr_zmiana = session.get('wybrana_zmiana')
    if not wypr_zmiana:
        return render_template_string(HTML_TEMPLATE, sesja_dane=sesja_dane, t=t, lang=lang, wypr_zmiana=None, user=user, role=role, pelna_nazwa=pelna_nazwa, dostep=dostep)

    linia = request.args.get('linia', session.get('aktywna_linia', '1'))
    session['aktywna_linia'] = linia

    dzis = datetime.now().strftime('%Y-%m-%d')
    if role == 'admin':
        wybrana_data = request.args.get('wybrana_data', dzis)
    else:
        wybrana_data = dzis

    min_data = (datetime.now() + timedelta(days=4)).strftime('%Y-%m-%d')
    dof = datetime.now().timetuple().tm_yday
    domyslna_partia = f"{dof}/{wypr_zmiana}/{linia}"

    error_msg = session.pop('error_msg', None)

    # Pobieranie danych z bazy PostgreSQL
    produkty_baza = {p.kod: {"nazwa": p.nazwa, "odbiorca": p.odbiorca} for p in ArtykulModel.query.all()}
    
    plan_baza_list = PlanModel.query.all()
    plan_produkcyjny = {}
    for pl in plan_baza_list:
        if pl.data not in plan_produkcyjny:
            plan_produkcyjny[pl.data] = {}
        plan_produkcyjny[pl.data][pl.kod] = pl.ilosc

    dostepne_daty_planu = sorted(plan_produkcyjny.keys(), reverse=True)
    domyslna_data_planu = dostepne_daty_planu[0] if dostepne_daty_planu else dzis
    wybrana_data_planu = request.args.get('wybrana_data_planu', domyslna_data_planu)
    if wybrana_data_planu not in plan_produkcyjny and dostepne_daty_planu:
        wybrana_data_planu = dostepne_daty_planu[0]

    # Pobieranie wszystkich wpisów historycznych dla widoków postępu i wydajności
    wpisy_baza_wszystkie_obj = WpisProdukcyjnyModel.query.all()
    wpisy_baza_wszystkie = [{
        "id": w.id, "data": w.data, "linia": w.linia, "zmiana": w.zmiana, "kod": w.kod,
        "nazwa": w.nazwa, "odbiorca": w.odbiorca, "ilosc": w.ilosc,
        "data_przyd": w.data_przyd, "partia_prod": w.partia_prod,
        "etykieta_gora": w.etykieta_gora, "etykieta_dol": w.etykieta_dol,
        "czas_start": w.czas_start, "czas_koniec": w.czas_koniec,
        "operator": w.operator
    } for w in wpisy_baza_wszystkie_obj]

    dostepne_daty_wpisow = sorted(list(set([w['data'] for w in wpisy_baza_wszystkie])), reverse=True)
    if dzis not in dostepne_daty_wpisow:
        dostepne_daty_wpisow.insert(0, dzis)

    dostepne_daty_wpisow_lub_planu = sorted(list(set(list(plan_produkcyjny.keys()) + [w['data'] for w in wpisy_baza_wszystkie] + [dzis])), reverse=True)

    wybrana_data_progress = request.args.get('wybrana_data_progress', dzis)
    if wybrana_data_progress not in dostepne_daty_wpisow_lub_planu and dostepne_daty_wpisow_lub_planu:
        wybrana_data_progress = dostepne_daty_wpisow_lub_planu[0]

    wybrana_data_wydajnosc = request.args.get('wybrana_data_wydajnosc', dzis)
    if wybrana_data_wydajnosc not in dostepne_daty_wpisow and dostepne_daty_wpisow:
        wybrana_data_wydajnosc = dostepne_daty_wpisow[0]

    # Słownik zrealizowanych sztuk per dzień dla formularza wprowadzania palet
    zrealizowane_z_datami = {}
    for w in wpisy_baza_wszystkie:
        d_w = w['data']
        k_w = w['kod']
        if d_w not in zrealizowane_z_datami:
            zrealizowane_z_datami[d_w] = {}
        zrealizowane_z_datami[d_w][k_w] = zrealizowane_z_datami[d_w].get(k_w, 0) + int(w['ilosc'])

    # Bezpieczne obliczanie postępu dla wybranej daty progress
    plan_wybrany_prog = plan_produkcyjny.get(wybrana_data_progress, {})
    suma_planu_calkowita_prog = sum(int(v) for v in plan_wybrany_prog.values())
    
    wykonane_dla_wybranej_daty_prog = {}
    for w in wpisy_baza_wszystkie:
        if w['data'] == wybrana_data_progress:
            kod_w = w['kod']
            wykonane_dla_wybranej_daty_prog[kod_w] = wykonane_dla_wybranej_daty_prog.get(kod_w, 0) + int(w['ilosc'])
            
    suma_wykonana_calkowita_prog = sum(int(v) for v in wykonane_dla_wybranej_daty_prog.values())
    pozostalo_calkowite_prog = max(suma_planu_calkowita_prog - suma_wykonana_calkowita_prog, 0)
    procent_calkowity_prog = round(suma_wykonana_calkowita_prog / suma_planu_calkowita_prog * 100) if suma_planu_calkowita_prog > 0 else 0

    wpisy_baza = [w for w in wpisy_baza_wszystkie if w['data'] == wybrana_data]
    wpisy_dla_linii = [w for w in wpisy_baza if w['linia'] == linia and w['zmiana'] == wypr_zmiana]
    suma_sztuk = sum(int(w['ilosc']) for w in wpisy_dla_linii)

    # Statystyki wydajności dla wybranej daty w zakładce wydajności
    wpisy_dla_wydajnosci = [w for w in wpisy_baza_wszystkie if w['data'] == wybrana_data_wydajnosc]
    statystyki_wydajnosci_wybrana = {}
    for item in wpisy_dla_wydajnosci:
        key = f"L{item['linia']}_Z{item['zmiana']}"
        if key not in statystyki_wydajnosci_wybrana:
            statystyki_wydajnosci_wybrana[key] = {
                "linia": item['linia'],
                "zmiana": item['zmiana'],
                "suma_sztuk": 0,
                "czas_minut": 0
            }
        statystyki_wydajnosci_wybrana[key]["suma_sztuk"] += int(item['ilosc'])
        try:
            t1 = datetime.strptime(item['czas_start'], '%H:%M')
            t2 = datetime.strptime(item['czas_koniec'], '%H:%M')
            diff = (t2 - t1).total_seconds() / 60
            if diff > 0:
                statystyki_wydajnosci_wybrana[key]["czas_minut"] += diff
        except Exception:
            pass

    for key, stat in statystyki_wydajnosci_wybrana.items():
        godziny = stat["czas_minut"] / 60
        stat["czas_godz"] = round(godziny, 2)
        if godziny > 0:
            stat["sztuk_na_godzine"] = round(stat["suma_sztuk"] / godziny, 1)
        else:
            stat["sztuk_na_godzine"] = 0

    uzytkownicy_baza_dict = {u.username: {"pelna_nazwa": u.pelna_nazwa, "rola": u.rola, "dostep": json.loads(u.dostep)} for u in UzytkownikModel.query.all()}

    return render_template_string(
        HTML_TEMPLATE,
        sesja_dane=sesja_dane,
        t=t,
        lang=lang,
        linia=linia,
        wypr_zmiana=wypr_zmiana,
        dzis=dzis,
        wybrana_data=wybrana_data,
        wybrana_data_planu=wybrana_data_planu,
        wybrana_data_progress=wybrana_data_progress,
        wybrana_data_wydajnosc=wybrana_data_wydajnosc,
        dostepne_daty_planu=dostepne_daty_planu,
        dostepne_daty_wpisow=dostepne_daty_wpisow,
        dostepne_daty_wpisow_lub_planu=dostepne_daty_wpisow_lub_planu,
        min_data=min_data,
        domyslna_partia=domyslna_partia,
        produkty_baza=produkty_baza,
        produkty_json=json.dumps(produkty_baza),
        plan_json=json.dumps(plan_produkcyjny),
        plan_produkcyjny=plan_produkcyjny,
        zrealizowane_json=json.dumps(zrealizowane_z_datami),
        plan_wybrany_prog=plan_wybrany_prog,
        wykonane_dla_wybranej_daty_prog=wykonane_dla_wybranej_daty_prog,
        suma_planu_calkowita_prog=suma_planu_calkowita_prog,
        suma_wykonana_calkowita_prog=suma_wykonana_calkowita_prog,
        pozostalo_calkowite_prog=pozostalo_calkowite_prog,
        procent_calkowity_prog=procent_calkowity_prog,
        wpisy_baza_wszystkie=wpisy_baza_wszystkie,
        wpisy_dla_linii=wpisy_dla_linii,
        suma_sztuk=suma_sztuk,
        statystyki_wydajnosci_wybrana=statystyki_wydajnosci_wybrana,
        user=user,
        role=role,
        pelna_nazwa=pelna_nazwa,
        dostep=dostep,
        uzytkownicy_baza=uzytkownicy_baza_dict,
        error_msg=error_msg
    )

@app.route('/wybierz-zmiane', methods=['POST'])
def wybierz_zmiane():
    session['wybrana_zmiana'] = request.form.get('zmiana')
    return redirect(url_for('index'))

@app.route('/reset-zmiana')
def reset_zmiana():
    session.pop('wybrana_zmiana', None)
    return redirect(url_for('index'))

@app.route('/dodaj', methods=['POST'])
def dodaj():
    if 'user' not in session:
        return redirect(url_for('login'))
    
    dzis = datetime.now().strftime('%Y-%m-%d')
    role = session.get('rola')
    data_prod = request.form.get('data', dzis) if role == 'admin' else dzis
    linia = request.args.get('linia', '1')

    czas_start = request.form.get('czas_start')
    czas_koniec = request.form.get('czas_koniec')

    if czas_start and czas_koniec and czas_koniec <= czas_start:
        session['error_msg'] = "Czas zakończenia palety nie może być wcześniejszy lub równy czasowi rozpoczęcia!"
        return redirect(url_for('index', linia=linia, wybrana_data=data_prod))

    kod = request.form.get('kod')
    artykul = ArtykulModel.query.get(kod)
    nazwa_art = artykul.nazwa if artykul else ""
    odbiorca_art = artykul.odbiorca if artykul else ""
    
    nowy_wpis = WpisProdukcyjnyModel(
        data=data_prod,
        linia=linia,
        zmiana=session.get('wybrana_zmiana'),
        kod=kod,
        nazwa=nazwa_art,
        odbiorca=odbiorca_art,
        ilosc=int(request.form.get('ilosc')),
        data_przyd=request.form.get('data_przyd'),
        partia_prod=request.form.get('partia_prod'),
        etykieta_gora=request.form.get('etykieta_gora'),
        etykieta_dol=request.form.get('etykieta_dol'),
        czas_start=czas_start,
        czas_koniec=czas_koniec,
        potwierdzenie_kontroli=request.form.get('potwierdzenie_kontroli'),
        potwierdzenie_oczyszczenia=request.form.get('potwierdzenie_oczyszczenia'),
        operator=session.get('pelna_nazwa')
    )
    db.session.add(nowy_wpis)
    db.session.commit()
    return redirect(url_for('index', linia=linia, wybrana_data=data_prod))

@app.route('/usun/<int:wpis_id>')
def usun(wpis_id):
    if 'user' not in session:
        return redirect(url_for('login'))
    
    dzis = datetime.now().strftime('%Y-%m-%d')
    role = session.get('rola')
    data_p = request.args.get('data', dzis)
    
    if role != 'admin' and data_p != dzis:
        return "Brak uprawnień!"

    wpis = WpisProdukcyjnyModel.query.get(wpis_id)
    if wpis:
        db.session.delete(wpis)
        db.session.commit()
    return redirect(url_for('index', linia=session.get('aktywna_linia', '1'), wybrana_data=data_p))

@app.route('/dodaj-artykul', methods=['POST'])
def dodaj_artykul():
    if session.get('rola') != 'admin':
        return "Brak uprawnień!"
    kod = request.form.get('kod').strip()
    nazwa = request.form.get('nazwa').strip()
    odbiorca = request.form.get('odbiorca').strip()
    if kod and nazwa:
        istniejacy = ArtykulModel.query.get(kod)
        if istniejacy:
            istniejacy.nazwa = nazwa
            istniejacy.odbiorca = odbiorca
        else:
            nowy = ArtykulModel(kod=kod, nazwa=nazwa, odbiorca=odbiorca)
            db.session.add(nowy)
        db.session.commit()
    return redirect(url_for('index', linia='db_art'))

@app.route('/usun-artykul/<kod>')
def usun_artykul(kod):
    if session.get('rola') != 'admin':
        return "Brak uprawnień!"
    art = ArtykulModel.query.get(kod)
    if art:
        db.session.delete(art)
        db.session.commit()
    return redirect(url_for('index', linia='db_art'))

@app.route('/zapisz-plan', methods=['POST'])
def zapisz_plan():
    if session.get('rola') != 'admin':
        return "Brak uprawnień!"
    data_planu = request.form.get('data_planu')
    surowy = request.form.get('surowy_plan')
    if data_planu and surowy:
        wiersze = surowy.split('\n')
        for w in wiersze:
            czesci = w.split('\t')
            if len(czesci) >= 2:
                kod = czesci[0].strip()
                try:
                    ilosc = int(czesci[1].strip().replace(' ', ''))
                    istniejacy_plan = PlanModel.query.filter_by(data=data_planu, kod=kod).first()
                    if istniejacy_plan:
                        istniejacy_plan.ilosc = ilosc
                    else:
                        nowy_p = PlanModel(data=data_planu, kod=kod, ilosc=ilosc)
                        db.session.add(nowy_p)
                except ValueError:
                    pass
        db.session.commit()
    return redirect(url_for('index', linia='plan', wybrana_data_planu=data_planu))

@app.route('/usun-z-planu')
def usun_z_planu():
    if session.get('rola') != 'admin':
        return "Brak uprawnień!"
    data_p = request.args.get('data')
    kod = request.args.get('kod')
    plan_poz = PlanModel.query.filter_by(data=data_p, kod=kod).first()
    if plan_poz:
        db.session.delete(plan_poz)
        db.session.commit()
    return redirect(url_for('index', linia='plan', wybrana_data_planu=data_p))

@app.route('/admin/dodaj-uzytkownika', methods=['POST'])
def admin_dodaj_uzytkownika():
    if session.get('rola') != 'admin':
        return "Brak uprawnień!"
    username = request.form.get('username').strip()
    password = request.form.get('password').strip()
    pelna_nazwa = request.form.get('pelna_nazwa').strip()
    rola = request.form.get('rola')
    dostep = request.form.getlist('dostep')
    if username and password and pelna_nazwa:
        usr = UzytkownikModel.query.get(username)
        if usr:
            usr.haslo = password
            usr.pelna_nazwa = pelna_nazwa
            usr.rola = rola
            usr.dostep = json.dumps(dostep)
        else:
            nowy_u = UzytkownikModel(username=username, haslo=password, rola=rola, pelna_nazwa=pelna_nazwa, dostep=json.dumps(dostep))
            db.session.add(nowy_u)
        db.session.commit()
    return redirect(url_for('index', linia='admin'))

@app.route('/admin/usun-uzytkownika/<username>')
def admin_usun_uzytkownika(username):
    if session.get('rola') != 'admin':
        return "Brak uprawnień!"
    if username != 'admin':
        usr = UzytkownikModel.query.get(username)
        if usr:
            db.session.delete(usr)
            db.session.commit()
    return redirect(url_for('index', linia='admin'))

@app.route('/generate', methods=['POST'])
def generate():
    if 'user' not in session:
        return redirect(url_for('login'))
    
    dzis = datetime.now().strftime('%Y-%m-%d')
    role = session.get('rola')
    wybrana_data = request.args.get('data', dzis) if role == 'admin' else dzis
    
    wpisy_baza = WpisProdukcyjnyModel.query.filter_by(data=wybrana_data).all()
    wpisy_dla_dnia = [{
        "linia": w.linia, "zmiana": w.zmiana, "kod": w.kod,
        "nazwa": w.nazwa, "odbiorca": w.odbiorca, "ilosc": w.ilosc,
        "data_przyd": w.data_przyd, "partia_prod": w.partia_prod,
        "etykieta_gora": w.etykieta_gora, "etykieta_dol": w.etykieta_dol,
        "czas_start": w.czas_start, "czas_koniec": w.czas_koniec,
        "operator": w.operator
    } for w in wpisy_baza]

    if not wpisy_dla_dnia:
        return f"Brak wpisów do wygenerowania dla daty {wybrana_data}!"

    wb = openpyxl.Workbook()
    
    # ----------------------------------------------------
    # 1. ARKUSZ: RAPORT WYDAJNOŚCI (TABELE OBOK SIEBIE)
    # ----------------------------------------------------
    ws_wyd = wb.active
    ws_wyd.title = "Raport Wydajności"
    ws_wyd.views.sheetView[0].showGridLines = True

    font_title = Font(name="Calibri", size=15, bold=True, color="1F497D")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_normal = Font(name="Calibri", size=10)
    
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )

    ws_wyd['A1'] = f"SZCZEGÓŁOWY RAPORT WYDAJNOŚCI GODZINOWEJ - {wybrana_data}"
    ws_wyd['A1'].font = font_title

    struktura_wyd = {"1": {}, "2": {}}
    for item in wpisy_dla_dnia:
        zm = str(item['zmiana'])
        ln = str(item['linia'])
        if zm not in struktura_wyd:
            struktura_wyd[zm] = {}
        if ln not in struktura_wyd[zm]:
            struktura_wyd[zm][ln] = []
        struktura_wyd[zm][ln].append(item)

    start_col = 1
    for zmiana_nr in ["1", "2"]:
        linia_dict = struktura_wyd.get(zmiana_nr, {})
        if not any(linia_dict.values()):
            continue

        ws_wyd.cell(row=3, column=start_col, value=f"⭐ ZMIANA {zmiana_nr}").font = Font(name="Calibri", size=12, bold=True, color="1F497D")

        for l_num in ['1', '2', '3', '4']:
            wpisy_l = linia_dict.get(l_num, [])
            if not wpisy_l:
                continue

            ws_wyd.cell(row=4, column=start_col, value=f"📅 Linia {l_num} | Zmiana {zmiana_nr}").font = Font(name="Calibri", size=10, bold=True, color="1F497D")

            h_tab = ["Godzina rozpoczęcia", "Godzina zakończenia", "Wyprodukowana ilość sztuk"]
            for c_offset, h_t in enumerate(h_tab):
                cell = ws_wyd.cell(row=5, column=start_col + c_offset, value=h_t)
                cell.font = font_header
                cell.fill = fill_header
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = thin_border
            ws_wyd.row_dimensions[5].height = 25

            r_idx = 6
            suma_l_sztuk = 0
            for w_g in wpisy_l:
                g_start = w_g['czas_start']
                g_koniec = w_g['czas_koniec']
                szt = int(w_g['ilosc'])
                suma_l_sztuk += szt

                c1 = ws_wyd.cell(row=r_idx, column=start_col, value=g_start)
                c2 = ws_wyd.cell(row=r_idx, column=start_col + 1, value=g_koniec)
                c3 = ws_wyd.cell(row=r_idx, column=start_col + 2, value=szt)
                
                for c_cell in [c1, c2, c3]:
                    c_cell.font = font_normal
                    c_cell.border = thin_border
                    c_cell.alignment = Alignment(horizontal="center", vertical="center")
                r_idx += 1

            sum_l_label = ws_wyd.cell(row=r_idx, column=start_col, value="Suma łączna:")
            sum_l_label.font = font_bold
            sum_l_label.border = thin_border
            sum_l_label.alignment = Alignment(horizontal="right", vertical="center")

            ws_wyd.cell(row=r_idx, column=start_col + 1, value="-").border = thin_border
            ws_wyd.cell(row=r_idx, column=start_col + 1).alignment = Alignment(horizontal="center", vertical="center")

            sum_l_val = ws_wyd.cell(row=r_idx, column=start_col + 2, value=suma_l_sztuk)
            sum_l_val.font = font_bold
            sum_l_val.border = thin_border
            sum_l_val.alignment = Alignment(horizontal="center", vertical="center")
            sum_l_val.fill = PatternFill(start_color="EAECEE", end_color="EAECEE", fill_type="solid")

            start_col += 4

    for col in ws_wyd.columns:
        max_l = max(len(str(cell.value or '')) for cell in col)
        ws_wyd.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = max(max_l + 3, 18)

    # ----------------------------------------------------
    # 2. ARKUSZE DLA KAŻDEJ LINII
    # ----------------------------------------------------
    linie_numery = ['1', '2', '3', '4']
    wybrana_zmiana = session.get('wybrana_zmiana')

    for l_nr in linie_numery:
        wpisy_linii = [w for w in wpisy_dla_dnia if str(w['linia']) == l_nr and (not wybrana_zmiana or str(w['zmiana']) == str(wybrana_zmiana))]
        if not wpisy_linii:
            continue

        ws_linia = wb.create_sheet(title=f"Linia {l_nr}")
        ws_linia.views.sheetView[0].showGridLines = True

        logo_path = 'logo.png'
        if os.path.exists(logo_path):
            try:
                img = Image(logo_path)
                img.width = 110
                img.height = 40
                ws_linia.add_image(img, 'A1')
            except Exception:
                pass

        ws_linia.cell(row=1, column=6, value="Nr dok: CP4/F01").font = Font(name="Calibri", size=9, bold=True)
        ws_linia.cell(row=1, column=8, value="Nr wydania: 3").font = Font(name="Calibri", size=9, bold=True)
        ws_linia.cell(row=1, column=10, value="Data wydania: 18.11.2026").font = Font(name="Calibri", size=9, bold=True)
        ws_linia.cell(row=1, column=12, value="Sporządził: Barbara Łogin").font = Font(name="Calibri", size=9, bold=True)

        thick_black = Side(style='medium', color='000000')
        thin_black = Side(style='thin', color='000000')

        for c_idx in range(6, 13):
            ws_linia.cell(row=1, column=c_idx).border = Border(
                top=thick_black, bottom=thick_black,
                left=thick_black if c_idx == 6 else None,
                right=thick_black if c_idx == 12 else None
            )

        d_z = datetime.strptime(wybrana_data, '%Y-%m-%d')
        dzien_roku = d_z.timetuple().tm_yday

        ws_linia.cell(row=5, column=1, value="Data produkcji:").font = font_bold
        ws_linia.cell(row=5, column=2, value=wybrana_data).font = font_normal
        ws_linia.cell(row=5, column=4, value="Dzień roku:").font = font_bold
        ws_linia.cell(row=5, column=5, value=dzien_roku).font = font_normal
        ws_linia.cell(row=5, column=6, value="Zmiana:").font = font_bold
        ws_linia.cell(row=5, column=7, value=wybrana_zmiana or "1").font = font_normal
        ws_linia.cell(row=5, column=8, value="Nr linii:").font = font_bold
        ws_linia.cell(row=5, column=9, value=l_nr).font = font_normal

        for c_idx in range(1, 10):
            ws_linia.cell(row=5, column=c_idx).border = Border(
                top=thick_black, bottom=thick_black,
                left=thick_black if c_idx == 1 else thin_black,
                right=thick_black if c_idx == 9 else thin_black
            )

        ws_linia.cell(row=7, column=1, value="Kontrola pracowników przed przystąpieniem do pracy (jeśli N - powód):").font = Font(name="Calibri", size=9)
        ws_linia.cell(row=7, column=12, value="+").font = font_bold
        ws_linia.cell(row=8, column=1, value="Sprawdzenie sprawności/czystości urządzeń przed rozpoczęciem pracy:").font = Font(name="Calibri", size=9)
        ws_linia.cell(row=8, column=12, value="+").font = font_bold

        for r_idx in [7, 8]:
            for c_idx in range(1, 13):
                ws_linia.cell(row=r_idx, column=c_idx).border = Border(
                    top=thick_black if r_idx == 7 else None,
                    bottom=thick_black if r_idx == 8 else None,
                    left=thick_black if c_idx == 1 else None,
                    right=thick_black if c_idx == 12 else None
                )

        ws_linia['A3'] = "ROZCHÓD WEWNĘTRZNY - OCENA PRODUKTU - WYRÓB GOTOWY"
        ws_linia['A3'].font = Font(name="Calibri", size=13, bold=True, color="1F497D")

        ws_linia.merge_cells('F10:H10')
        ws_linia.cell(row=10, column=6, value="Ocena i zwolnienie partii produktu*\nCP4 Wygląd produktu, prawidłowa, czytelna etykieta***")

        for r in range(10, 12):
            for c in range(1, 15):
                cell_curr = ws_linia.cell(row=r, column=c)
                cell_curr.fill = fill_header
                cell_curr.border = Border(
                    top=thick_black if r == 10 else thin_black,
                    bottom=thick_black if r == 11 else thin_black,
                    left=thick_black if c == 1 else thin_black,
                    right=thick_black if c == 14 else thin_black
                )
                cell_curr.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
                cell_curr.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ws_linia.cell(row=10, column=1, value="Kod")
        ws_linia.cell(row=10, column=2, value="Nazwa produktu")
        ws_linia.cell(row=10, column=3, value="Odbiorca")
        ws_linia.cell(row=10, column=4, value="Data przydatności / Nr partii")
        ws_linia.cell(row=10, column=5, value="Ilość")
        
        ws_linia.cell(row=11, column=6, value="P")
        ws_linia.cell(row=11, column=7, value="S")
        ws_linia.cell(row=11, column=8, value="K")
        
        ws_linia.cell(row=10, column=9, value="Nr partii etykiety GÓRNEJ/OKUB")
        ws_linia.cell(row=10, column=10, value="Nr partii etykiety DOLNEJ")
        ws_linia.cell(row=10, column=11, value="Godz. rozpoczęcia")
        ws_linia.cell(row=10, column=12, value="Godz. zakończenia")
        ws_linia.cell(row=10, column=13, value="Prawidłowe oczyszczenie linii z pozostałych etykiet, banderol, kartonów przed rozpoczęciem produkcji kolejnego produktu*")
        ws_linia.cell(row=10, column=14, value="Operator")

        ws_linia.row_dimensions[10].height = 45
        ws_linia.row_dimensions[11].height = 20

        curr_r = 12
        for item in wpisy_linii:
            w_dane = [
                str(item['kod']),
                item['nazwa'],
                item['odbiorca'],
                f"{item['data_przyd']} / {item['partia_prod']}",
                int(item['ilosc']),
                "+", "+", "+",
                item['etykieta_gora'],
                item['etykieta_dol'],
                item['czas_start'],
                item['czas_koniec'],
                " + ",
                item['operator']
            ]

            for col_idx, val in enumerate(w_dane, start=1):
                cell = ws_linia.cell(row=curr_r, column=col_idx, value=val)
                cell.font = font_normal
                cell.border = Border(
                    top=thin_black,
                    bottom=thick_black if len(wpisy_linii) == (curr_r - 11) else thin_black,
                    left=thick_black if col_idx == 1 else thin_black,
                    right=thick_black if col_idx == 14 else thin_black
                )
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            ws_linia.row_dimensions[curr_r].height = 24
            curr_r += 1

        foot_r = curr_r + 2
        ws_linia.cell(row=foot_r, column=1, value="***kontrola na początku P, w środku S i na końcu K produkcji danego produktu").font = Font(name="Calibri", size=8, italic=True)
        ws_linia.cell(row=foot_r+1, column=1, value="dodatkowo (3 pomiary) przy zmianie rolki etykiet/kartonu, po awarii sprzętu").font = Font(name="Calibri", size=8, italic=True)
        ws_linia.cell(row=foot_r+3, column=1, value="Weryfikacja brygadzisty / kierownika (podpis): ........................................................................").font = font_bold

        kolumny_szerokosci = {
            'A': 12, 'B': 26, 'C': 16, 'D': 24, 'E': 10, 
            'F': 6, 'G': 6, 'H': 6, 'I': 18, 'J': 18, 'K': 14, 'L': 14, 'M': 24, 'N': 20
        }
        for let, szer in kolumny_szerokosci.items():
            ws_linia.column_dimensions[let].width = szer

    output_path = f"Raport_Kompletny_{wybrana_data}.xlsx"
    wb.save(output_path)
    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(port=5000, debug=True)
