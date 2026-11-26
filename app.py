from flask import Flask, render_template_string, request, redirect, url_for, send_file, session, jsonify
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime, timedelta
import os
import json
import io
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text

app = Flask(__name__)
app.secret_key = 'tajny_klucz_sesji_produkcyjnej'

# Bezpieczna konfiguracja bazy danych (Render / Supabase)
SUPABASE_URI = os.getenv(
    "DATABASE_URL",
    "postgresql://postgres.dbgjiavdnlgnvqnrkykx:Prod_Glo_2026@aws-0-eu-west-1.pooler.supabase.com:6543/postgres"
)
app.config['SQLALCHEMY_DATABASE_URI'] = SUPABASE_URI
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# --- MODELE BAZODANOWE (JEDNA WSPÓLNA BAZA POSTGRESQL) ---

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
    dostep = db.Column(db.Text, nullable=False)

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

# --- WSPÓLNA BAZA PRACOWNIKÓW I MODELE MODUŁU PRODUKCYJNEGO ---

class PracownikQrModel(db.Model):
    __tablename__ = 'pracownicy_qr_baza'
    emp_id = db.Column(db.String(50), primary_key=True)
    imie_nazwisko = db.Column(db.String(150), nullable=False)
    stanowisko = db.Column(db.String(120), nullable=False, default="")
    kolejnosc = db.Column(db.Integer, nullable=False, default=0)

class LiniaMonitorStanModel(db.Model):
    __tablename__ = 'linie_monitor_stan'
    linia_id = db.Column(db.Integer, primary_key=True)
    status = db.Column(db.String(20), nullable=False, default='IDLE')
    start_time_iso = db.Column(db.String(50), nullable=True)
    reason = db.Column(db.String(250), nullable=False, default='')
    started_by_login = db.Column(db.String(80), nullable=True)

class LiniaMonitorZdarzenieModel(db.Model):
    __tablename__ = 'linie_monitor_zdarzenia'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    linia_id = db.Column(db.Integer, nullable=False)
    event_type = db.Column(db.String(20), nullable=False)
    start_iso = db.Column(db.String(50), nullable=False)
    end_iso = db.Column(db.String(50), nullable=False)
    duration_sec = db.Column(db.Integer, nullable=False, default=0)
    reason = db.Column(db.String(250), nullable=False, default='-')
    iso_date = db.Column(db.String(20), nullable=False)
    started_by_login = db.Column(db.String(80), nullable=True)
    ended_by_login = db.Column(db.String(80), nullable=True)

class PracownikSesjaAktywnaModel(db.Model):
    __tablename__ = 'pracownicy_sesje_aktywne'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    emp_id = db.Column(db.String(50), nullable=False)
    emp_name = db.Column(db.String(150), nullable=False)
    line_id = db.Column(db.Integer, nullable=False)
    start_time_iso = db.Column(db.String(50), nullable=False)
    iso_date = db.Column(db.String(20), nullable=False)

class PracownikSesjaHistoriaModel(db.Model):
    __tablename__ = 'pracownicy_sesje_historia'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    emp_id = db.Column(db.String(50), nullable=False)
    emp_name = db.Column(db.String(150), nullable=False)
    line_id = db.Column(db.Integer, nullable=False)
    start_time_iso = db.Column(db.String(50), nullable=False)
    end_time_iso = db.Column(db.String(50), nullable=False)
    duration_sec = db.Column(db.Integer, nullable=False, default=0)
    iso_date = db.Column(db.String(20), nullable=False)

DEFAULT_EMPLOYEES_LIST = [
    ("EMPO070", "Giraldo Parra Diana Marcela"),
    ("EMPO037", "Logreira Marriaga Emanuel David"),
    ("EMPO022", "Perez Redondo Fredy Rafael"),
    ("EMPO101", "Burgos Cadena Alfredo Roberto"),
    ("EMPO023", "Masłowski Maciej"),
    ("EMPE002", "Kobets Maksym"),
    ("EMPO118", "Fuentes Sosa David Enrique"),
    ("EMPO120", "Tamarska Halyna"),
    ("EMPO029", "Rakitina Tamara"),
    ("EMPO107", "Sanchez Ramirez Cristopher Leandro"),
    ("EMPO062", "Artiaga DeLa Rosa Norana"),
    ("EMPO071", "Kyrychenko Viktoriia"),
    ("EMPO052", "Moreno Malambo Juan David"),
    ("EMPO014", "Hetnarowicz Dorota"),
    ("EMPO011", "Munoz Vera Daniel Mauricio"),
    ("EMPO051", "Espinosa Corrales Jonathan Alexis"),
    ("EMPO004", "Correa Agustin Joel"),
    ("EMPO122", "Siriachenko Olena"),
    ("EMPO038", "Ortiz Montes Andres Felipe"),
    ("EMPO055", "Avendano Ordonez Vanessa"),
    ("EMPO031", "Padilla Fernandez Melarith Paola"),
    ("EMPO024", "Poberezhna Iryna"),
    ("EMPO046", "Sobolevskyi Yevhenii"),
    ("EMPO084", "Nieto Camargo Diana Catalina"),
    ("EMP022", "Marszeniuk Kamil"),
    ("EMPO032", "Pardo Luna Emna Katerine"),
    ("EMPO103", "Potes Riascos Jhonny Alexander"),
    ("EMPO077", "Guerrero Palacios Fenis"),
    ("EMPO044", "Bielinskyi Artem"),
    ("EMP016", "Jaśkiewicz Karolina"),
    ("EMPO039", "Pineda Quintana Yina Valeria"),
    ("EMPO008", "Lashtur Anastasiia"),
    ("EMPO091", "Sobolevskyi Oleksandr"),
    ("EMPO086", "Sanchez Mancera Julian Andres"),
    ("EMPO030", "Sadowski Janusz"),
    ("EMPO108", "Galvis Rojas Jackson Jesus"),
    ("EMPO043", "Chmykhun Hennadii"),
    ("EMP027", "Pieróg Karol"),
    ("EMPO119", "Quinones Ferreira Heidy Maria"),
    ("EMP006", "Bator Urszula"),
    ("EMPO064", "Herrera Martinez Luz Carime"),
    ("EMPO109", "Suaza Quintero Marlon Adrian"),
    ("EMP017", "Kirko Justyna"),
    ("EMPO088", "Tsikhanovich Iryna"),
    ("EMPO040", "Torres Muriel Natalia Carolina"),
    ("EMPO020", "Llanos Jimenez Javier"),
    ("EMPO082", "Sanchez Cortes Devinson Andres"),
    ("EMP011", "Chomontek Mariusz"),
    ("EMPO049", "Zabaleta Sanchez Jhusty Sther"),
    ("EMPO001", "Campo Cueto Ninoska"),
    ("EMPO094", "Rivera Garcia Jose Manuel"),
    ("EMPO116", "Barroso Ferrer Zunilda"),
    ("EMP015", "Janicki Jacek"),
    ("EMPO076", "Rivera Reyes Andres Felipe"),
    ("EMPO018", "Munive Monsalvo Camilo Antonio"),
    ("EMP032", "Sawko Marcin"),
    ("EMPO112", "Moreno Diaz Andres David"),
    ("EMPO021", "Parra Calderon Brayan Fernando"),
    ("EMP003", "Artem (Operator) L1"),
    ("EMPO072", "Morales Barajas Fabian Dayanny"),
    ("EMPO115", "Valencia Ortiz Jakeline"),
    ("EMP018", "Kobus Piotr"),
    ("EMP031", "Sarzhevskyi Yevhen"),
    ("EMP024", "Musiał Damian"),
    ("EMPO012", "Denysova Ruslana"),
    ("EMPO041", "Dubovyk Iryna"),
    ("EMP002", "Mariusz Siwiński"),
    ("EMPO069", "Rojas Avila Jaeth Maryory"),
    ("EMPO067", "Osorio Suaza Angie Lorena"),
    ("EMPT001", "Srokosz Joanna"),
    ("EMPO027", "Carrero Rangel John Jairo"),
    ("EMPO073", "Reyes Castaneda Maikol Alexander"),
    ("EMP008", "Bogdanowicz Monika"),
    ("EMPO100", "Shkolna Anastasiia"),
    ("EMPO092", "Alvarez Begazo Alison Elvira"),
    ("EMPO063", "Avila Pena Omaira Viviana"),
    ("EMP028", "Polut Robert"),
    ("EMP001", "Janusz Sadowski"),
    ("EMP010", "Bola Jeremiasz"),
    ("EMPO117", "Ferreira Jordan Maria Nereida"),
    ("EMPO053", "Patino Plata Yosert Felipe"),
    ("EMP005", "Artem Bielynski (operator)"),
    ("EMPO033", "Pastrana Jimenez Eva Sandrid"),
    ("EMPO080", "Herrera Yesenia Varon"),
    ("EMPO098", "Chmykhun Vira"),
    ("EMPO048", "Valencia Pineros Mery Giseth"),
    ("EMPO003", "Alfonso Avila Brayan Esneider"),
    ("EMPO012", "Tovar Luna Sebastian"),
    ("EMP013", "Dudek Dorota"),
    ("EMPO061", "Angulo Mendivil Jocabell Yolianys"),
    ("EMPO057", "Rodriguez Malvacea Iris Coromoto"),
    ("EMPO096", "Toledo Chaparro Yeferson Alexander"),
    ("EMPO056", "Bacilio Somoza Faustina"),
    ("EMPO110", "Florez Guerrero Idamis"),
    ("EMPO083", "Bacca Cruz Sirley Michell"),
    ("EMPO002", "Podoloschii Victor"),
    ("EMPE001", "Dubovyk Leonid"),
    ("EMPO010", "Sanchez Perez Juan Carlos"),
    ("EMPO045", "Ponomarenko Nazar"),
    ("EMP004", "Maciej Bąk (Operator)"),
    ("EMPO009", "Kryszewska (Haidash) Nataliia"),
    ("EMPO059", "Petryk Olena"),
    ("EMPO050", "Ramirez Jimenez Alex Ivan"),
    ("EMPO054", "Arcon Caceres Dana Marcela"),
    ("EMPO079", "Bustillo Bello Diego Armando"),
    ("EMPO075", "Bielova Olena"),
    ("EMPO078", "Serrano Ariza Diego Alejandro"),
    ("EMPO105", "Sabluk Inna"),
    ("EMPO025", "Nguyen Bao Ngan"),
    ("EMPO047", "Ramos Nieto Eduardo Luis"),
    ("EMP033", "Semmerling Andrzej"),
    ("EMPO066", "Munoz Talero Paula Mayerli"),
    ("EMPO102", "Santos Lora Valentina"),
    ("EMPO058", "Alvarez Calderon Bacilio Kiomi Mayra"),
    ("EMPO015", "Roa Delgado Angie Julieth"),
    ("EMPO036", "Fernandez Aranda Sandra Iris"),
    ("EMPO019", "Quiroga Silva Luis Edgar"),
    ("EMPO042", "Ponomarenko Oksana"),
    ("EMPO106", "Del Rio Lopez Sandra"),
    ("EMPO099", "Lazko Lidiia"),
    ("EMPE003", "Sosna Zinovia"),
    ("EMPO089", "Novikova Mariia"),
    ("EMPO093", "Bautista Quesada Breyder Jesid"),
    ("EMPO123", "Skunts Valerii"),
    ("EMPO090", "Petsko Oleksandr"),
    ("EMPO113", "Lasprilla Molina Victor Manuel"),
    ("EMPO087", "Netylkina Iryna"),
    ("EMPO085", "Ortega Ayala Gerson Andres"),
    ("EMP026", "Pęski Marek"),
    ("EMPO121", "Kantymyr Tetiana"),
    ("EMPO124", "Rodriguez Enciso Luis Eduardo"),
    ("EMP019", "Korzeniewski Piotr"),
    ("EMP007", "Bielańczuk Dawid"),
    ("EMPO095", "Bokova Nadiia"),
    ("EMPKJ001", "Kosicka Krystyna"),
    ("EMPO026", "Ardila Estupinan Catalina Julieth"),
    ("EMPO068", "Retamoza Castillo Francisco Javier"),
    ("EMPO114", "Duarte luna Monica Bibiana"),
    ("EMPO006", "Volodina Anastasiia"),
    ("EMPO016", "Sanguino Aparicio Marlin Danitza"),
    ("EMPO028", "Garcia Suarez Hellen Patricia"),
    ("EMP034", "Szereszewski Andrzej"),
    ("EMPO104", "Abakumau Vadzim"),
    ("EMPO060", "Abella Fernandez Johann Sebastian"),
    ("EMPO035", "Murillo Casas Javier Alejandro"),
    ("EMPO097", "Ibragimov Shukhrat"),
    ("EMPO081", "Hrytsiuk Mariia GW"),
    ("EMPO007", "Zabaleta Fernandez Keiber Luis"),
    ("EMPO111", "Romero Guzman Yarelis"),
    ("EMPO034", "Kaliusenko Artem"),
    ("EMPO065", "Iakuba Vladyslav"),
    ("EMPO074", "Valencia Munoz Wilmar Humberto"),
    ("EMPO005", "Khliebas Nataliia"),
    ("EMP021", "Łużyński Piotr"),
    ("EMPKJ002", "Witaszek Aleksandra"),
    ("EMPO017", "Larios Tuiran Andrea Carolina"),
    ("EMP020", "Kosior Roman"),
    ("EMP009", "Bogucki Robert"),
    ("EMPO013", "Cerquera Rodriguez David Ricardo"),
    ("EMP025", "Niewiadomski Arkadiusz"),
]

with app.app_context():
    db.create_all()

    try:
        with db.engine.begin() as conn:
            conn.execute(text("ALTER TABLE linie_monitor_stan ADD COLUMN IF NOT EXISTS started_by_login VARCHAR(80)"))
            conn.execute(text("ALTER TABLE linie_monitor_zdarzenia ADD COLUMN IF NOT EXISTS started_by_login VARCHAR(80)"))
            conn.execute(text("ALTER TABLE linie_monitor_zdarzenia ADD COLUMN IF NOT EXISTS ended_by_login VARCHAR(80)"))
    except Exception:
        pass

    admin_user = UzytkownikModel.query.filter_by(username="admin").first()
    if admin_user:
        dostepy = json.loads(admin_user.dostep)
        zmieniono = False
        for req_perm in ["raporty", "mod_koncowki", "mod_produkcja"]:
            if req_perm not in dostepy:
                dostepy.append(req_perm)
                zmieniono = True
        if zmieniono:
            admin_user.dostep = json.dumps(dostepy)
            db.session.commit()

    if UzytkownikModel.query.count() == 0:
        domyslni_uzyt = {
            "admin": {
                "haslo": "admin123",
                "rola": "admin",
                "pelna_nazwa": "Administrator Systemu",
                "dostep": ["mod_koncowki", "mod_produkcja", "1", "2", "3", "4", "plan", "progress", "wydajnosc", "raporty", "db_art", "admin"]
            },
            "tetiana": {
                "haslo": "operator123",
                "rola": "operator",
                "pelna_nazwa": "Tetiana Kantymyr",
                "dostep": ["mod_koncowki", "1", "2", "3", "4", "progress", "wydajnosc", "raporty"]
            },
            "iryna": {
                "haslo": "operator123",
                "rola": "operator",
                "pelna_nazwa": "Poberezhna Iryna",
                "dostep": ["mod_koncowki", "1", "2", "3", "4", "progress", "wydajnosc", "raporty"]
            }
        }
        for usr, info in domyslni_uzyt.items():
            db.session.add(UzytkownikModel(
                username=usr, haslo=info["haslo"], rola=info["rola"],
                pelna_nazwa=info["pelna_nazwa"], dostep=json.dumps(info["dostep"])
            ))

        domyslnie_produkty = {
            "1354": {"nazwa": "Hähnchen Teriyaki 190g", "odbiorca": "Aldi SUD"},
            "1378": {"nazwa": "Hähnchen Sweet Chili 190g", "odbiorca": "Aldi SUD"},
            "1422": {"nazwa": "Falafel-Hummus 190g", "odbiorca": "Aldi SUD"},
            "1553": {"nazwa": "Kanapka z jajkiem i pastą jajeczną", "odbiorca": "Perla PL"},
            "2690": {"nazwa": "Thunfisch 190g", "odbiorca": "Aldi SUD"},
            "3101": {"nazwa": "G&G Paprykowy 190g", "odbiorca": "EDEKA"}
        }
        for kod, info in domyslnie_produkty.items():
            db.session.add(ArtykulModel(kod=kod, nazwa=info["nazwa"], odbiorca=info["odbiorca"]))

        dzis_str = datetime.now().strftime('%Y-%m-%d')
        for kod, szt in {"1354": 2500, "2690": 1500, "3101": 2000}.items():
            db.session.add(PlanModel(data=dzis_str, kod=kod, ilosc=szt))

        db.session.commit()

    if PracownikQrModel.query.count() == 0:
        for idx, (eid, ename) in enumerate(DEFAULT_EMPLOYEES_LIST):
            if not PracownikQrModel.query.get(eid):
                db.session.add(PracownikQrModel(
                    emp_id=eid, imie_nazwisko=ename, stanowisko="", kolejnosc=idx
                ))
        db.session.commit()

    for lid in [1, 2, 3, 4, 5, 6]:
        if not LiniaMonitorStanModel.query.get(lid):
            db.session.add(LiniaMonitorStanModel(linia_id=lid, status='IDLE', start_time_iso=None, reason='', started_by_login=None))
    db.session.commit()

# --- FUNKCJE POMOCNICZE ---

def oblicz_czas_bez_nakladania(przedzialy):
    if not przedzialy:
        return 0.0
    posortowane = sorted(przedzialy, key=lambda x: x[0])
    polaczone = [list(posortowane[0])]
    for start, end in posortowane[1:]:
        ostatni = polaczone[-1]
        if start <= ostatni[1]:
            ostatni[1] = max(ostatni[1], end)
        else:
            polaczone.append([start, end])
    return sum(max(0.0, b - a) for a, b in polaczone)

def czy_ma_dostep_do_modulu(dostep_lista, rola, modul):
    if rola == 'admin':
        return True
    if modul == 'koncowki':
        return 'mod_koncowki' in dostep_lista or any(k in dostep_lista for k in ['1', '2', '3', '4'])
    if modul == 'produkcja':
        return 'mod_produkcja' in dostep_lista
    return False

def parse_iso_dt(iso_str):
    if not iso_str:
        return None
    try:
        return datetime.fromisoformat(iso_str.replace('Z', '+00:00')).replace(tzinfo=None)
    except Exception:
        return None

def okresl_dzien_roboczy_i_zmiane(iso_str):
    dt = parse_iso_dt(iso_str)
    if not dt:
        return None, '1'
    if 4 <= dt.hour < 14:
        return dt.strftime('%Y-%m-%d'), '1'
    if dt.hour >= 14:
        return dt.strftime('%Y-%m-%d'), '2'
    dzien_poprzedni = (dt - timedelta(days=1)).strftime('%Y-%m-%d')
    return dzien_poprzedni, '2'

def format_sec_hhmmss(seconds):
    seconds = max(0, int(seconds or 0))
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"

def build_monitor_snapshot(target_date=None):
    today_iso = target_date or datetime.now().strftime('%Y-%m-%d')
    now_dt = datetime.now()

    stany_db = {s.linia_id: s for s in LiniaMonitorStanModel.query.all()}
    zdarzenia_db = LiniaMonitorZdarzenieModel.query.all()
    aktywne_db = PracownikSesjaAktywnaModel.query.all()
    historia_emp_db = PracownikSesjaHistoriaModel.query.all()

    lines_payload = {}
    for lid in [1, 2, 3, 4, 5, 6]:
        s = stany_db.get(lid)
        status = s.status if s else 'IDLE'
        start_iso = s.start_time_iso if s else None
        reason = s.reason if s else ''
        started_by = s.started_by_login if s else None

        current_elapsed = 0
        start_dt = parse_iso_dt(start_iso)
        if status != 'IDLE' and start_dt:
            current_elapsed = max(0, int((now_dt - start_dt).total_seconds()))

        daily_work_sec = 0
        daily_stop_sec = 0
        for h in zdarzenia_db:
            if h.linia_id == lid and h.iso_date == today_iso:
                if h.event_type == 'PRACA':
                    daily_work_sec += int(h.duration_sec or 0)
                elif h.event_type == 'POSTÓJ':
                    daily_stop_sec += int(h.duration_sec or 0)

        if status == 'RUNNING':
            daily_work_sec += current_elapsed
        elif status == 'STOPPED':
            daily_stop_sec += current_elapsed

        total_emp_sec = 0
        for eh in historia_emp_db:
            if eh.line_id == lid and eh.iso_date == today_iso:
                total_emp_sec += int(eh.duration_sec or 0)

        line_active_emps = []
        for act in aktywne_db:
            if act.line_id == lid:
                act_dt = parse_iso_dt(act.start_time_iso)
                act_elapsed = max(0, int((now_dt - act_dt).total_seconds())) if act_dt else 0
                total_emp_sec += act_elapsed
                line_active_emps.append({
                    "id": act.id,
                    "empId": act.emp_id,
                    "empName": act.emp_name,
                    "startTimeIso": act.start_time_iso
                })

        line_events = [
            {
                "id": h.id,
                "lineId": h.linia_id,
                "type": h.event_type,
                "startIso": h.start_iso,
                "endIso": h.end_iso,
                "durationSec": h.duration_sec,
                "reason": h.reason,
                "isoDate": h.iso_date,
                "startedByLogin": h.started_by_login,
                "endedByLogin": h.ended_by_login
            }
            for h in zdarzenia_db if h.linia_id == lid
        ]
        line_events.sort(key=lambda x: x["startIso"], reverse=True)

        lines_payload[str(lid)] = {
            "lineId": lid,
            "status": status,
            "startTimeIso": start_iso,
            "reason": reason,
            "startedByLogin": started_by,
            "currentElapsedSec": current_elapsed,
            "dailyWorkSec": daily_work_sec,
            "dailyStopSec": daily_stop_sec,
            "totalEmpSec": total_emp_sec,
            "activeEmployees": line_active_emps,
            "recentEvents": line_events[:3]
        }

    return {
        "serverNowIso": now_dt.isoformat(),
        "todayIso": today_iso,
        "lines": lines_payload
    }

# --- WIELOJĘZYCZNY SŁOWNIK (PL / EN / UK / ES) ---

T = {
    "pl": {
        "tytul": "Panel Realizacji Produkcji i Raportowania",
        "wybierz_zm": "Wybierz zmianę roboczą:",
        "rozpocznij": "Rozpocznij pracę",
        "zmiana": "Aktualna zmiana",
        "zmien_zm": "Zmień zmianę",
        "zalog_op": "Zalogowany operator",
        "linia_1": "Linia 1", "linia_2": "Linia 2", "linia_3": "Linia 3", "linia_4": "Linia 4",
        "plan_tab": "Plan Produkcyjny", "prog_tab": "Postęp / Realizacja", "wydajnosc_tab": "Wydajność (Szt/h)",
        "raporty_tab": "Centrum Raportów", "db_art_tab": "Baza Artykułów", "admin_tab": "Panel Administratora",
        "zarejestruj": "Zarejestruj wyprodukowaną partię dla",
        "data_prod": "Data produkcji:", "kod_art": "Kod artykułu:", "wybierz_kod": "-- Wybierz kod artykułu --",
        "ilosc_szt": "Ilość sztuk (na koniec palety):", "wpisz_ilosc": "wpisz ilość na koniec...", "nazwa_ind": "Nazwa indeksu:",
        "odbiorca": "Odbiorca:", "data_przyd": "Data przydatności (min. +4 dni):", "nr_partii": "Nr partii produktu:",
        "ety_gora": "Etykieta GÓRNA:", "ety_dol": "Etykieta DOLNA:", "operator": "Operator:",
        "dodaj_btn": "Zarejestruj wyprodukowaną paletę / partię", "realizacja_biez": "Realizacja na bieżąco",
        "brak_wpisow": "Brak wpisów dla tej linii i daty.", "lacznie": "Łącznie wyprodukowano na linii:",
        "wklej_plan": "Zarządzanie Planem Produkcyjnym",
        "instrukcja_plan": "Wybierz datę planu, wklej dane z Excela (Kod [Tab] Ilość):",
        "zapisz_plan_btn": "Zapisz plan dla wybranej daty", "zarzadzanie_art": "Zarządzanie Bazą Artykułów",
        "dodaj_art_btn": "Zapisz artykuł", "istniejace_art": "Istniejące Artykuły", "usun": "[Usuń]",
        "logowanie_tytul": "Logowanie do Systemu Produkcyjnego", "wybierz_modul_lbl": "Wybierz moduł systemu:",
        "mod_koncowki_nazwa": "📦 Końcówki Linii (Kompletacja / Palety / Raporty)",
        "mod_produkcja_nazwa": "⚙️ Moduł Produkcyjny (Linie 1–6 / Pracownicy / QR)",
        "login": "Login:", "haslo": "Hasło:", "zaloguj": "Zaloguj się",
        "brak_dostepu": "Brak uprawnień do tej sekcji lub modułu!", "wyloguj": "Wyloguj",
        "zarzadzanie_uzyt": "Zarządzanie Użytkownikami i Uprawnieniami",
        "bhp_tytul": "Wymagane potwierdzenia BHP i Kontroli:",
        "bhp_1": "Potwierdzam przeprowadzenie kontroli pracowników i czystości urządzeń przed przystąpieniem do pracy.",
        "bhp_2": "Potwierdzam prawidłowe oczyszczenie linii z pozostałych etykiet, banderol i kartonów przed rozpoczęciem produkcji kolejnego produktu.",
        "czas_pracy_tytul": "Czas pracy (WYMAGANE):",
        "czas_opis": "Kliknij Start Palety na początku (nie musisz podawać ilości). Koniec i ilość uzupełnisz na końcu.",
        "start_palety": "Start Palety", "koniec_palety": "Koniec Palety",
        "czas_start_lbl": "Czas Startu (Wymagane):", "czas_koniec_lbl": "Czas Końca (Wymagane):",
        "postep_tytul": "Łączny Postęp Produkcji",
        "podsumowanie_dnia": "Podsumowanie Dnia (Wszystkie Artykuły)",
        "realizacja_podstawowa": "Realizacja planu podstawowego (bez nadwyżek):",
        "realizacja_z_nadwyzka": "Realizacja łącznie z nadwyżkami:",
        "do_zrobienia": "Do zrobienia łącznie:", "wykonano": "Wykonano:", "pozostalo": "Pozostało:",
        "postep_art": "Postęp według poszczególnych artykułów:",
        "brak_planu_dzis": "Brak planu produkcyjnego na ten dzień.",
        "raport_wyd_tytul": "Raport Wydajności (Sztuk na Godzinę)",
        "raport_wyd_opis": "Poniższy raport przedstawia podział na poszczególne linie i zmiany (bez podwójnego liczenia nakładających się palet) oraz średnią produktywność w przeliczeniu na 1 godzinę.",
        "th_linia": "Linia", "th_zmiana": "Zmiana", "th_wyprodukowano": "Wyprodukowano sztuk", "th_czas_pracy": "Czas pracy (godziny)", "th_wydajnosc": "Wydajność (szt/h)",
        "laczna_suma_wypr": "Łączna suma wyprodukowana:", "laczny_czas": "Łączny czas pracy:", "srednia_wyd": "Średnia wydajność:",
        "brak_czasow": "Brak zarejestrowanych czasów pracy i palet dla wybranej daty.",
        "wybierz_dzien_plan": "Wybierz dzień planu:", "podglad_planu": "Podgląd zaplanowanej produkcji:", "brak_planow": "Brak zdefiniowanych planów produkcyjnych.",
        "dodaj_nowego_uzyt": "Dodaj / zaktualizuj użytkownika (wspólna baza loginów i haseł)",
        "imie_nazwisko": "Imię i Nazwisko operatora:", "rola": "Rola:",
        "uprawnienia_lbl": "Uprawnienia (moduły i zakładki):", "zapisz_uzyt_btn": "Zapisz użytkownika", "zarejestrowani_uzyt": "Zarejestrowani użytkownicy",
        "systemowy": "(systemowy)", "kol_kod": "Kod", "kol_nazwa": "Nazwa", "kol_odbiorca": "Odbiorca", "kol_akcja": "Akcja",
        "wybierz_dzien_hist": "Wybierz dzień do podglądu historii i generowania raportu (Admin):",
        "data_raportu_hist": "Data raportu / historii:", "tabela_nr": "#", "tabela_ilosc": "Ilość", "tabela_czas": "Czas pracy",
        "tabela_bhp": "Potwierdzenia BHP", "tabela_szczegoly": "Szczegóły",
        "nadwyzka": "Nadwyżka", "wybierz_dzien_prog": "Wybierz dzień do podglądu postępu:",
        "ph_kod": "np. 9999", "ph_nazwa": "Nazwa produktu", "ph_odbiorca": "Odbiorca", "ph_ilosc": "wpisz ilość na koniec...",
        "ph_etygora": "Wpisz etykietę górną...", "ph_etydol": "Wpisz etykietę dolną..."
    },
    "en": {
        "tytul": "Production Execution & Reporting Panel",
        "wybierz_zm": "Select work shift:",
        "rozpocznij": "Start shift",
        "zmiana": "Current shift",
        "zmien_zm": "Change shift",
        "zalog_op": "Logged-in operator",
        "linia_1": "Line 1", "linia_2": "Line 2", "linia_3": "Line 3", "linia_4": "Line 4",
        "plan_tab": "Production Plan", "prog_tab": "Progress / Execution", "wydajnosc_tab": "Efficiency (Pcs/h)",
        "raporty_tab": "Reports Hub", "db_art_tab": "Items Database", "admin_tab": "Admin Panel",
        "zarejestruj": "Register produced batch for",
        "data_prod": "Production date:", "kod_art": "Item code:", "wybierz_kod": "-- Select item code --",
        "ilosc_szt": "Quantity (end of pallet):", "wpisz_ilosc": "enter end quantity...", "nazwa_ind": "Item name:",
        "odbiorca": "Customer:", "data_przyd": "Expiry date (min. +4 days):", "nr_partii": "Batch number:",
        "ety_gora": "TOP label:", "ety_dol": "BOTTOM label:", "operator": "Operator:",
        "dodaj_btn": "Register produced pallet / batch", "realizacja_biez": "Live execution",
        "brak_wpisow": "No entries for this line and date.", "lacznie": "Total produced on line:",
        "wklej_plan": "Production Plan Management",
        "instrukcja_plan": "Select plan date, paste data from Excel (Code [Tab] Qty):",
        "zapisz_plan_btn": "Save plan for selected date", "zarzadzanie_art": "Items Database Management",
        "dodaj_art_btn": "Save item", "istniejace_art": "Existing Items", "usun": "[Delete]",
        "logowanie_tytul": "Production System Login", "wybierz_modul_lbl": "Select system module:",
        "mod_koncowki_nazwa": "📦 Line Ends (Packing / Pallets / Reports)",
        "mod_produkcja_nazwa": "⚙️ Production Module (Lines 1–6 / Workers / QR)",
        "login": "Username:", "haslo": "Password:", "zaloguj": "Log in",
        "brak_dostepu": "No permissions for this section or module!", "wyloguj": "Log out",
        "zarzadzanie_uzyt": "Users & Permissions Management",
        "bhp_tytul": "Required HSE & Quality Checks:",
        "bhp_1": "I confirm worker check and equipment cleanliness before starting work.",
        "bhp_2": "I confirm proper line clearance from remaining labels, seals and cartons before starting next product.",
        "czas_pracy_tytul": "Working time (REQUIRED):",
        "czas_opis": "Click Pallet Start at the beginning. Enter end time and quantity at the end.",
        "start_palety": "Pallet Start", "koniec_palety": "Pallet End",
        "czas_start_lbl": "Start Time (Required):", "czas_koniec_lbl": "End Time (Required):",
        "postep_tytul": "Total Production Progress",
        "podsumowanie_dnia": "Daily Summary (All Items)",
        "realizacja_podstawowa": "Base plan completion (without surplus):",
        "realizacja_z_nadwyzka": "Completion including surplus:",
        "do_zrobienia": "Total planned:", "wykonano": "Completed:", "pozostalo": "Remaining:",
        "postep_art": "Progress by item:",
        "brak_planu_dzis": "No production plan for this day.",
        "raport_wyd_tytul": "Efficiency Report (Pieces per Hour)",
        "raport_wyd_opis": "Breakdown by line and shift (without double-counting overlapping pallets) and average hourly output.",
        "th_linia": "Line", "th_zmiana": "Shift", "th_wyprodukowano": "Produced pcs", "th_czas_pracy": "Working time (hours)", "th_wydajnosc": "Efficiency (pcs/h)",
        "laczna_suma_wypr": "Total produced:", "laczny_czas": "Total working time:", "srednia_wyd": "Average efficiency:",
        "brak_czasow": "No recorded times or pallets for selected date.",
        "wybierz_dzien_plan": "Select plan date:", "podglad_planu": "Planned production preview:", "brak_planow": "No production plans defined.",
        "dodaj_nowego_uzyt": "Add / update user (shared credentials database)",
        "imie_nazwisko": "Operator full name:", "rola": "Role:",
        "uprawnienia_lbl": "Permissions (modules and tabs):", "zapisz_uzyt_btn": "Save user", "zarejestrowani_uzyt": "Registered users",
        "systemowy": "(system)", "kol_kod": "Code", "kol_nazwa": "Name", "kol_odbiorca": "Customer", "kol_akcja": "Action",
        "wybierz_dzien_hist": "Select day to view history and generate report (Admin):",
        "data_raportu_hist": "Report / history date:", "tabela_nr": "#", "tabela_ilosc": "Qty", "tabela_czas": "Work time",
        "tabela_bhp": "HSE Checks", "tabela_szczegoly": "Details",
        "nadwyzka": "Surplus", "wybierz_dzien_prog": "Select day to view progress:",
        "ph_kod": "e.g. 9999", "ph_nazwa": "Product name", "ph_odbiorca": "Customer", "ph_ilosc": "enter end quantity...",
        "ph_etygora": "Enter top label...", "ph_etydol": "Enter bottom label..."
    },
    "uk": {
        "tytul": "Панель Виконання Виробництва та Звітів",
        "wybierz_zm": "Оберіть робочу зміну:",
        "rozpocznij": "Розпочати зміну",
        "zmiana": "Поточна зміна",
        "zmien_zm": "Змінити зміну",
        "zalog_op": "Авторизований оператор",
        "linia_1": "Лінія 1", "linia_2": "Лінія 2", "linia_3": "Лінія 3", "linia_4": "Лінія 4",
        "plan_tab": "План Виробництва", "prog_tab": "Прогрес / Виконання", "wydajnosc_tab": "Продуктивність (Шт/год)",
        "raporty_tab": "Центр Звітів", "db_art_tab": "База Артикулів", "admin_tab": "Панель Адміністратора",
        "zarejestruj": "Зареєструвати партію для",
        "data_prod": "Дата виробництва:", "kod_art": "Код артикулу:", "wybierz_kod": "-- Оберіть код артикулу --",
        "ilosc_szt": "Кількість штук (на кінець палети):", "wpisz_ilosc": "введіть кількість...", "nazwa_ind": "Назва артикулу:",
        "odbiorca": "Клієнт:", "data_przyd": "Термін придатності (мін. +4 дні):", "nr_partii": "Номер партії:",
        "ety_gora": "Етикетка ВЕРХНЯ:", "ety_dol": "Етикетка НИЖНЯ:", "operator": "Оператор:",
        "dodaj_btn": "Зареєструвати палету / партію", "realizacja_biez": "Поточне виконання",
        "brak_wpisow": "Немає записів для цієї лінії та дати.", "lacznie": "Всього вироблено на лінії:",
        "wklej_plan": "Керування Планом Виробництва",
        "instrukcja_plan": "Оберіть дату плану, вставте дані з Excel (Код [Tab] Кількість):",
        "zapisz_plan_btn": "Зберегти план на обрану дату", "zarzadzanie_art": "Керування Базою Артикулів",
        "dodaj_art_btn": "Зберегти артикул", "istniejace_art": "Наявні Артикули", "usun": "[Видалити]",
        "logowanie_tytul": "Вхід до Виробничої Системи", "wybierz_modul_lbl": "Оберіть модуль системи:",
        "mod_koncowki_nazwa": "📦 Кінці Ліній (Пакування / Палети / Звіти)",
        "mod_produkcja_nazwa": "⚙️ Виробничий Модуль (Лінії 1–6 / Працівники / QR)",
        "login": "Логін:", "haslo": "Пароль:", "zaloguj": "Увійти",
        "brak_dostepu": "Немає прав доступу до цього розділу або модуля!", "wyloguj": "Вийти",
        "zarzadzanie_uzyt": "Керування Користувачами та Правами",
        "bhp_tytul": "Обов'язкові перевірки ОП та Контролю:",
        "bhp_1": "Підтверджую перевірку працівників та чистоти обладнання перед початком роботи.",
        "bhp_2": "Підтверджую очищення лінії від залишків етикеток, стрічок та коробок перед наступним продуктом.",
        "czas_pracy_tytul": "Час роботи (ОБОВ'ЯЗКОВО):",
        "czas_opis": "Натисніть Старт Палети на початку. Кінець і кількість введіть наприкінці.",
        "start_palety": "Старт Палети", "koniec_palety": "Кінець Палети",
        "czas_start_lbl": "Час Старту (Обов'язково):", "czas_koniec_lbl": "Час Кінця (Обов'язково):",
        "postep_tytul": "Загальний Прогрес Виробництва",
        "podsumowanie_dnia": "Підсумок Дня (Усі Артикули)",
        "realizacja_podstawowa": "Виконання базового плану (без надлишків):",
        "realizacja_z_nadwyzka": "Виконання разом з надлишками:",
        "do_zrobienia": "Заплановано всього:", "wykonano": "Виконано:", "pozostalo": "Залишилось:",
        "postep_art": "Прогрес за артикулами:",
        "brak_planu_dzis": "Немає плану виробництва на цей день.",
        "raport_wyd_tytul": "Звіт Продуктивності (Штук за Годину)",
        "raport_wyd_opis": "Розподіл за лініями та змінами (без подвійного підрахунку палет) та середня продуктивність за годину.",
        "th_linia": "Лінія", "th_zmiana": "Зміна", "th_wyprodukowano": "Вироблено шт", "th_czas_pracy": "Час роботи (год)", "th_wydajnosc": "Продуктивність (шт/год)",
        "laczna_suma_wypr": "Загальна сума виробленого:", "laczny_czas": "Загальний час роботи:", "srednia_wyd": "Середня продуктивність:",
        "brak_czasow": "Немає зафіксованих часів або палет за обрану дату.",
        "wybierz_dzien_plan": "Оберіть дату плану:", "podglad_planu": "Перегляд запланованого виробництва:", "brak_planow": "Планів виробництва не визначено.",
        "dodaj_nowego_uzyt": "Додати / оновити користувача (спільна база логінів та паролів)",
        "imie_nazwisko": "Ім'я та Прізвище оператора:", "rola": "Роль:",
        "uprawnienia_lbl": "Права доступу (модулі та вкладки):", "zapisz_uzyt_btn": "Зберегти користувача", "zarejestrowani_uzyt": "Зареєстровані користувачі",
        "systemowy": "(системний)", "kol_kod": "Код", "kol_nazwa": "Назва", "kol_odbiorca": "Клієнт", "kol_akcja": "Дія",
        "wybierz_dzien_hist": "Оберіть день для перегляду історії та звіту (Адмін):",
        "data_raportu_hist": "Дата звіту / історії:", "tabela_nr": "#", "tabela_ilosc": "Кількість", "tabela_czas": "Час роботи",
        "tabela_bhp": "Перевірки ОП", "tabela_szczegoly": "Деталі",
        "nadwyzka": "Надлишок", "wybierz_dzien_prog": "Оберіть день для перегляду прогресу:",
        "ph_kod": "напр. 9999", "ph_nazwa": "Назва продукту", "ph_odbiorca": "Клієнт", "ph_ilosc": "введіть кількість...",
        "ph_etygora": "Введіть верхню етикетку...", "ph_etydol": "Введіть нижню етикетку..."
    },
    "es": {
        "tytul": "Panel de Ejecución de Producción e Informes",
        "wybierz_zm": "Selecciona el turno de trabajo:",
        "rozpocznij": "Iniciar turno",
        "zmiana": "Turno actual",
        "zmien_zm": "Cambiar turno",
        "zalog_op": "Operador conectado",
        "linia_1": "Línea 1", "linia_2": "Línea 2", "linia_3": "Línea 3", "linia_4": "Línea 4",
        "plan_tab": "Plan de Producción", "prog_tab": "Progreso / Ejecución", "wydajnosc_tab": "Rendimiento (Uds/h)",
        "raporty_tab": "Centro de Informes", "db_art_tab": "Base de Artículos", "admin_tab": "Panel de Administrador",
        "zarejestruj": "Registrar lote producido para",
        "data_prod": "Fecha de producción:", "kod_art": "Código de artículo:", "wybierz_kod": "-- Selecciona código de artículo --",
        "ilosc_szt": "Cantidad (fin de palé):", "wpisz_ilosc": "escribe cantidad final...", "nazwa_ind": "Nombre del artículo:",
        "odbiorca": "Cliente:", "data_przyd": "Fecha de caducidad (mín. +4 días):", "nr_partii": "Nº de lote:",
        "ety_gora": "Etiqueta SUPERIOR:", "ety_dol": "Etiqueta INFERIOR:", "operator": "Operador:",
        "dodaj_btn": "Registrar palé / lote producido", "realizacja_biez": "Ejecución en tiempo real",
        "brak_wpisow": "No hay registros para esta línea y fecha.", "lacznie": "Total producido en la línea:",
        "wklej_plan": "Gestión del Plan de Producción",
        "instrukcja_plan": "Selecciona la fecha del plan y pega datos de Excel (Código [Tab] Cantidad):",
        "zapisz_plan_btn": "Guardar plan para la fecha seleccionada", "zarzadzanie_art": "Gestión de Base de Artículos",
        "dodaj_art_btn": "Guardar artículo", "istniejace_art": "Artículos Existentes", "usun": "[Eliminar]",
        "logowanie_tytul": "Inicio de Sesión del Sistema de Producción", "wybierz_modul_lbl": "Selecciona el módulo del sistema:",
        "mod_koncowki_nazwa": "📦 Finales de Línea (Empaque / Palés / Informes)",
        "mod_produkcja_nazwa": "⚙️ Módulo de Producción (Líneas 1–6 / Operarios / QR)",
        "login": "Usuario:", "haslo": "Contraseña:", "zaloguj": "Iniciar sesión",
        "brak_dostepu": "¡Sin permisos para esta sección o módulo!", "wyloguj": "Cerrar sesión",
        "zarzadzanie_uzyt": "Gestión de Usuarios y Permisos",
        "bhp_tytul": "Controles Obligatorios de PRL y Calidad:",
        "bhp_1": "Confirmo el control del personal y la limpieza de los equipos antes de iniciar el trabajo.",
        "bhp_2": "Confirmo la correcta limpieza de la línea de etiquetas, precintos y cajas restantes antes del siguiente producto.",
        "czas_pracy_tytul": "Tiempo de trabajo (OBLIGATORIO):",
        "czas_opis": "Pulsa Inicio de Palé al empezar. Introduce la hora final y cantidad al terminar.",
        "start_palety": "Inicio de Palé", "koniec_palety": "Fin de Palé",
        "czas_start_lbl": "Hora de Inicio (Obligatorio):", "czas_koniec_lbl": "Hora de Fin (Obligatorio):",
        "postep_tytul": "Progreso Total de Producción",
        "podsumowanie_dnia": "Resumen del Día (Todos los Artículos)",
        "realizacja_podstawowa": "Cumplimiento del plan base (sin excedentes):",
        "realizacja_z_nadwyzka": "Cumplimiento incluyendo excedentes:",
        "do_zrobienia": "Planificado total:", "wykonano": "Realizado:", "pozostalo": "Restante:",
        "postep_art": "Progreso por artículo:",
        "brak_planu_dzis": "No hay plan de producción para este día.",
        "raport_wyd_tytul": "Informe de Rendimiento (Unidades por Hora)",
        "raport_wyd_opis": "Desglose por línea y turno (sin duplicar palés superpuestos) y productividad media por hora.",
        "th_linia": "Línea", "th_zmiana": "Turno", "th_wyprodukowano": "Uds. producidas", "th_czas_pracy": "Tiempo de trabajo (horas)", "th_wydajnosc": "Rendimiento (uds/h)",
        "laczna_suma_wypr": "Total producido:", "laczny_czas": "Tiempo total de trabajo:", "srednia_wyd": "Rendimiento medio:",
        "brak_czasow": "No hay tiempos ni palés registrados para la fecha seleccionada.",
        "wybierz_dzien_plan": "Selecciona fecha del plan:", "podglad_planu": "Vista previa de producción planificada:", "brak_planow": "No hay planes de producción definidos.",
        "dodaj_nowego_uzyt": "Añadir / actualizar usuario (base compartida de usuarios y contraseñas)",
        "imie_nazwisko": "Nombre y apellidos del operador:", "rola": "Rol:",
        "uprawnienia_lbl": "Permisos (módulos y pestañas):", "zapisz_uzyt_btn": "Guardar usuario", "zarejestrowani_uzyt": "Usuarios registrados",
        "systemowy": "(sistema)", "kol_kod": "Código", "kol_nazwa": "Nombre", "kol_odbiorca": "Cliente", "kol_akcja": "Acción",
        "wybierz_dzien_hist": "Selecciona día para ver historial y generar informe (Admin):",
        "data_raportu_hist": "Fecha de informe / historial:", "tabela_nr": "#", "tabela_ilosc": "Cant.", "tabela_czas": "Tiempo",
        "tabela_bhp": "Controles PRL", "tabela_szczegoly": "Detalles",
        "nadwyzka": "Excedente", "wybierz_dzien_prog": "Selecciona día para ver progreso:",
        "ph_kod": "ej. 9999", "ph_nazwa": "Nombre del producto", "ph_odbiorca": "Cliente", "ph_ilosc": "escribe cantidad final...",
        "ph_etygora": "Escribe etiqueta superior...", "ph_etydol": "Escribe etiqueta inferior..."
    }
}

def get_lang_dict():
    lang = session.get('lang', 'pl')
    if lang not in T:
        lang = 'pl'
    return T[lang], lang

LOGIN_TEMPLATE = """
<!DOCTYPE html>
<html lang="{{ lang }}">
<head>
    <meta charset="UTF-8">
    <title>{{ t.logowanie_tytul }}</title>
    <style>
        body { font-family: Arial, sans-serif; background: #0f172a; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; padding: 15px; box-sizing: border-box; }
        .login-box { background: white; padding: 30px 34px; border-radius: 12px; box-shadow: 0 12px 28px rgba(0,0,0,0.35); width: 460px; max-width: 100%; }
        .top-lang { display: flex; justify-content: flex-end; gap: 6px; margin-bottom: 10px; flex-wrap: wrap; }
        .lang-pill { padding: 4px 10px; border-radius: 6px; font-size: 12px; font-weight: bold; text-decoration: none; border: 1px solid #cbd5e1; color: #334155; background: #f8fafc; }
        .lang-pill.active { background: #0284c7; color: white; border-color: #0284c7; }
        h3 { text-align: center; color: #0f172a; margin-top: 0; margin-bottom: 18px; font-size: 1.35rem; }
        label { font-weight: bold; display: block; margin-top: 14px; color: #334155; font-size: 13px; }
        input[type="text"], input[type="password"] { width: 100%; padding: 10px 12px; margin-top: 5px; border: 1px solid #cbd5e1; border-radius: 7px; box-sizing: border-box; font-size: 14px; }
        input:focus { outline: none; border-color: #0284c7; box-shadow: 0 0 0 3px rgba(2,132,199,0.18); }
        .module-grid { display: flex; flex-direction: column; gap: 10px; margin-top: 8px; }
        .module-card { border: 2px solid #e2e8f0; border-radius: 9px; padding: 12px 14px; cursor: pointer; display: flex; align-items: center; gap: 12px; transition: all 0.15s ease; background: #f8fafc; margin-top: 0; font-weight: normal; }
        .module-card:hover { border-color: #0284c7; background: #f0f9ff; }
        .module-card input[type="radio"] { width: 18px; height: 18px; margin: 0; cursor: pointer; }
        .module-title { font-weight: bold; font-size: 14px; color: #0f172a; }
        .module-desc { font-size: 12px; color: #64748b; margin-top: 2px; }
        button { background: #0284c7; color: white; border: none; padding: 12px; margin-top: 20px; cursor: pointer; width: 100%; font-size: 15px; border-radius: 7px; font-weight: bold; }
        button:hover { background: #0369a1; }
        .error { background: #fef2f2; color: #b91c1c; border: 1px solid #fecaca; padding: 10px; border-radius: 7px; font-size: 13px; text-align: center; margin-bottom: 12px; font-weight: bold; }
    </style>
</head>
<body>
<div class="login-box">
    <div class="top-lang">
        <a href="/set-lang/pl" class="lang-pill {{ 'active' if lang == 'pl' else '' }}">🇵🇱 PL</a>
        <a href="/set-lang/en" class="lang-pill {{ 'active' if lang == 'en' else '' }}">🇬🇧 EN</a>
        <a href="/set-lang/uk" class="lang-pill {{ 'active' if lang == 'uk' else '' }}">🇺🇦 UK</a>
        <a href="/set-lang/es" class="lang-pill {{ 'active' if lang == 'es' else '' }}">🇪🇸 ES</a>
    </div>
    <h3>🔐 {{ t.logowanie_tytul }}</h3>
    {% if error %}
        <div class="error">{{ error }}</div>
    {% endif %}
    <form method="POST" action="/login">
        <label>{{ t.wybierz_modul_lbl }}</label>
        <div class="module-grid">
            <label class="module-card">
                <input type="radio" name="modul" value="koncowki" {{ 'checked' if wybrany_modul != 'produkcja' else '' }}>
                <div>
                    <div class="module-title">{{ t.mod_koncowki_nazwa }}</div>
                    <div class="module-desc">Kompletacja / palety, plan, postęp, wydajność, centrum raportów</div>
                </div>
            </label>
            <label class="module-card">
                <input type="radio" name="modul" value="produkcja" {{ 'checked' if wybrany_modul == 'produkcja' else '' }}>
                <div>
                    <div class="module-title">{{ t.mod_produkcja_nazwa }}</div>
                    <div class="module-desc">Monitorowanie linii 1–6, postój, skan QR pracowników, plakietki 54×85 mm</div>
                </div>
            </label>
        </div>

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

PRODUKCJA_MODULE_TEMPLATE = """
<!DOCTYPE html>
<html lang="{{ lang }}">
<head>
  <meta charset="UTF-8">
  <title>{{ t.mod_produkcja_nazwa }}</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
  <style>
    * { box-sizing: border-box; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    body { background-color: #f1f5f9; margin: 0; padding: 20px; color: #1e293b; }
    .container { max-width: 1320px; margin: auto; }

    .header-bar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; background: white; padding: 14px 22px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); flex-wrap: wrap; gap: 10px; }
    .header-meta { font-size: 0.9rem; color: #334155; display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
    .mod-badge { background: #0284c7; color: white; padding: 3px 9px; border-radius: 6px; font-size: 0.78rem; font-weight: bold; }
    .header-buttons { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }

    .lang-box { display: inline-flex; gap: 4px; background: #f1f5f9; padding: 4px; border-radius: 8px; border: 1px solid #cbd5e1; }
    .lang-btn { padding: 5px 8px; border-radius: 6px; font-size: 0.78rem; font-weight: bold; text-decoration: none; color: #334155; }
    .lang-btn.active { background: #0f172a; color: white; }

    .btn-switch { background-color: #16a34a; color: white; padding: 9px 14px; border: none; border-radius: 8px; font-weight: bold; cursor: pointer; font-size: 0.88rem; text-decoration: none; display: inline-block; }
    .btn-switch:hover { background-color: #15803d; }
    .btn-report { background-color: #ea580c; color: white; padding: 9px 14px; border: none; border-radius: 8px; font-weight: bold; cursor: pointer; font-size: 0.88rem; text-decoration: none; display: inline-block; }
    .btn-report:hover { background-color: #c2410c; }
    .btn-logout { background-color: #dc2626; color: white; padding: 9px 14px; border: none; border-radius: 8px; font-weight: bold; cursor: pointer; font-size: 0.88rem; text-decoration: none; display: inline-block; }
    .btn-logout:hover { background-color: #b91c1c; }

    .subtabs { display: flex; gap: 8px; margin-bottom: 18px; flex-wrap: wrap; }
    .subtab-btn { padding: 10px 16px; border-radius: 8px; border: 1px solid #cbd5e1; background: white; color: #334155; font-weight: bold; cursor: pointer; font-size: 0.9rem; }
    .subtab-btn.active { background: #0f172a; color: white; border-color: #0f172a; }
    .subtab-btn.admin-sub { background: #fee2e2; color: #991b1b; border-color: #fca5a5; }
    .subtab-btn.admin-sub.active { background: #dc2626; color: white; border-color: #dc2626; }

    .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(360px, 1fr)); gap: 20px; margin-bottom: 25px; }

    .card { background: white; border-radius: 12px; padding: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.06); transition: background-color 0.25s, border-color 0.25s; border: 2px solid transparent; }
    .card.running { background-color: #f0fdf4; border-color: #10b981; }
    .card.stopped { background-color: #fef2f2; border-color: #ef4444; }
    .card.idle { background-color: #f8fafc; border-color: #cbd5e1; }

    .card-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }
    .card-title { font-size: 1.25rem; font-weight: bold; margin: 0; }
    .status-badge { padding: 4px 10px; border-radius: 20px; font-weight: bold; font-size: 0.8rem; }
    .status-running { background-color: #d1fae5; color: #047857; }
    .status-stopped { background-color: #fee2e2; color: #b91c1c; }
    .status-idle { background-color: #e2e8f0; color: #475569; }

    .timers-container { display: flex; justify-content: space-around; margin: 10px 0 8px 0; background: rgba(255,255,255,0.8); padding: 10px; border-radius: 8px; border: 1px solid rgba(0,0,0,0.08); }
    .timer-box { text-align: center; }
    .timer-label { font-size: 0.7rem; text-transform: uppercase; font-weight: 700; color: #64748b; margin-bottom: 2px; }
    .timer-val { font-size: 1.2rem; font-family: monospace; font-weight: bold; color: #0f172a; }
    .timer-val.active-work { color: #059669; }
    .timer-val.active-stop { color: #dc2626; }

    .daily-timers-container { display: flex; justify-content: space-around; margin-bottom: 8px; background: rgba(255,255,255,0.8); padding: 8px; border-radius: 8px; border: 1px solid rgba(0,0,0,0.08); }
    .daily-timer-box { text-align: center; }
    .daily-timer-label { font-size: 0.68rem; text-transform: uppercase; font-weight: bold; color: #475569; margin-bottom: 2px; }
    .daily-timer-val { font-size: 1.05rem; font-family: monospace; font-weight: bold; }
    .daily-work-val { color: #047857; }
    .daily-stop-val { color: #b91c1c; }

    .total-emp-time-box { background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 8px; padding: 10px 14px; margin-bottom: 12px; display: flex; justify-content: space-between; align-items: center; }
    .total-emp-time-label { font-size: 0.85rem; font-weight: bold; color: #1e40af; }
    .total-emp-time-val { font-size: 1.15rem; font-family: monospace; font-weight: bold; color: #1d4ed8; }

    .emp-section { background: rgba(255,255,255,0.88); border: 1px solid rgba(0,0,0,0.08); border-radius: 8px; padding: 10px; margin-bottom: 12px; }
    .emp-section-title { font-size: 0.8rem; font-weight: bold; color: #0369a1; margin-bottom: 6px; }
    .emp-list { max-height: 95px; overflow-y: auto; font-size: 0.8rem; }
    .emp-item { display: flex; justify-content: space-between; align-items: center; background: white; padding: 4px 8px; border-radius: 4px; margin-bottom: 4px; border: 1px solid #e0f2fe; }
    .btn-remove-emp { background: #ef4444; color: white; border: none; border-radius: 4px; padding: 2px 7px; font-size: 0.72rem; cursor: pointer; }

    .scan-box { display: flex; gap: 5px; margin-bottom: 10px; }
    select, input { padding: 8px; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 0.85rem; width: 100%; }

    .controls { display: flex; flex-direction: column; gap: 8px; }
    .btn-group { display: flex; flex-wrap: wrap; gap: 6px; }
    button.action-btn { flex: 1 1 auto; padding: 8px 6px; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; font-size: 0.8rem; white-space: nowrap; }
    .btn-start { background-color: #10b981; color: white; }
    .btn-stop { background-color: #ef4444; color: white; }
    .btn-end { background-color: #475569; color: white; }
    .btn-disabled { opacity: 0.4; cursor: not-allowed !important; }

    .reason-box { background: #fff5f5; border: 1px solid #fecaca; border-radius: 8px; padding: 10px; margin-top: 5px; }
    .reason-box label { font-size: 0.75rem; font-weight: bold; color: #991b1b; display: block; margin-bottom: 4px; }

    .history { margin-top: 10px; border-top: 1px solid rgba(0,0,0,0.1); padding-top: 8px; max-height: 100px; overflow-y: auto; font-size: 0.78rem; }
    .history-item { margin-bottom: 4px; padding-bottom: 4px; border-bottom: 1px dashed rgba(0,0,0,0.1); }

    .panel-card { background: white; padding: 22px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); margin-bottom: 25px; }
    .qr-result { display: flex; align-items: center; gap: 20px; margin-top: 15px; background: #f8fafc; padding: 15px; border-radius: 8px; flex-wrap: wrap; border: 1px solid #e2e8f0; }

    .badge-54x85 {
      width: 54mm; height: 85mm; background: #ffffff; border: 1px dashed #94a3b8;
      border-radius: 3mm; padding: 4mm 3mm; display: flex; flex-direction: column;
      align-items: center; justify-content: space-between; text-align: center;
      page-break-inside: avoid; break-inside: avoid;
    }
    .badge-header { font-size: 7.5pt; font-weight: 700; letter-spacing: 0.5px; text-transform: uppercase; color: #475569; border-bottom: 1px solid #e2e8f0; width: 100%; padding-bottom: 1.5mm; }
    .badge-qr-box { display: flex; align-items: center; justify-content: center; margin: 1mm 0; }
    .badge-qr-box img, .badge-qr-box canvas { width: 34mm !important; height: 34mm !important; }
    .badge-id { font-size: 10.5pt; font-weight: 800; color: #0f172a; font-family: monospace; margin-top: 1mm; }
    .badge-name { font-size: 9pt; font-weight: 700; color: #1e293b; line-height: 1.25; max-height: 12mm; overflow: hidden; word-break: break-word; }

    .all-badges-container { margin-top: 20px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 16px; }
    .all-badges-toolbar { display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; margin-bottom: 14px; }
    .badges-grid { display: flex; flex-wrap: wrap; gap: 4mm; justify-content: flex-start; }

    .sortable-th { cursor: pointer; user-select: none; }
    .sortable-th:hover { background-color: #e2e8f0; }

    table { width: 100%; border-collapse: collapse; margin-top: 10px; }
    th, td { text-align: left; padding: 8px 10px; border-bottom: 1px solid #e2e8f0; font-size: 0.85rem; }
    th { background-color: #f1f5f9; color: #475569; }

    .row-flex { display: flex; gap: 12px; flex-wrap: wrap; align-items: flex-end; }
    .row-flex > div { flex: 1; min-width: 170px; }

    @media print {
      @page { margin: 6mm; }
      body * { visibility: hidden; }
      body.print-single #printableSingleBadge, body.print-single #printableSingleBadge * { visibility: visible; }
      body.print-single #printableSingleBadge { position: fixed; left: 0; top: 0; }
      body.print-all #printableAllBadges, body.print-all #printableAllBadges * { visibility: visible; }
      body.print-all #printableAllBadges { position: absolute; left: 0; top: 0; width: 100%; background: white !important; border: none !important; padding: 0 !important; margin: 0 !important; }
      body.print-all .all-badges-toolbar { display: none !important; }
      body.print-all .badges-grid { display: flex; flex-wrap: wrap; gap: 3mm; }
      .badge-54x85 { border: 0.25mm solid #cbd5e1; }
    }
  </style>
</head>
<body>
<div class="container">
  <div class="header-bar">
    <div class="header-meta">
      <span class="mod-badge">{{ t.mod_produkcja_nazwa }}</span>
      <span><b>{{ t.zalog_op }}:</b> {{ pelna_nazwa }} ({{ role }})</span>
      <span>| <b>Auto-refresh:</b> <span style="color:#059669; font-weight:bold;">4s</span></span>
    </div>
    <div class="header-buttons">
      <div class="lang-box">
        <a href="/set-lang/pl" class="lang-btn {{ 'active' if lang == 'pl' else '' }}">PL</a>
        <a href="/set-lang/en" class="lang-btn {{ 'active' if lang == 'en' else '' }}">EN</a>
        <a href="/set-lang/uk" class="lang-btn {{ 'active' if lang == 'uk' else '' }}">UK</a>
        <a href="/set-lang/es" class="lang-btn {{ 'active' if lang == 'es' else '' }}">ES</a>
      </div>
      {% if role == 'admin' or can_access_koncowki %}
        <a href="/przelacz-modul/koncowki" class="btn-switch">🔀 {{ t.mod_koncowki_nazwa.split('(')[0] }}</a>
      {% endif %}
      {% if role == 'admin' or 'raporty' in dostep %}
        <a href="/przelacz-modul/koncowki?linia=raporty" class="btn-report">📈 {{ t.raporty_tab }}</a>
      {% endif %}
      <a href="/logout" class="btn-logout">🔒 {{ t.wyloguj }}</a>
    </div>
  </div>

  <div class="subtabs">
    <button type="button" class="subtab-btn active" id="tabBtnLines" onclick="switchSubtab('lines')">🏭 {{ t.mod_produkcja_nazwa.split('(')[0] }}</button>
    <button type="button" class="subtab-btn" id="tabBtnEmployees" onclick="switchSubtab('employees')">👤 {{ t.zarejestrowani_uzyt.split()[0] }} & QR</button>
    {% if role == 'admin' %}
      <button type="button" class="subtab-btn admin-sub" id="tabBtnAdmin" onclick="switchSubtab('admin')">🔐 {{ t.admin_tab }}</button>
    {% endif %}
  </div>

  <div id="subtabLines">
    <div class="grid" id="linesContainer"></div>
  </div>

  <div id="subtabEmployees" style="display:none;">
    <div class="panel-card">
      <h3 style="margin-top:0;">👤 Baza Pracowników & Generator Kodów QR (54 mm × 85 mm)</h3>
      <p style="font-size:0.85rem; color:#475569; margin-top:0;">Kody QR są generowane z ID pracownika — istniejące wydrukowane plakietki pozostają w pełni kompatybilne. Stanowisko możesz dodać lub zmienić przez przycisk <b>Edytuj</b>.</p>

      <div class="row-flex" style="margin-bottom: 18px;">
        <div>
          <label style="font-size:0.78rem; font-weight:bold; color:#334155;">ID Pracownika</label>
          <input type="text" id="newEmpId" placeholder="np. EMPO070">
        </div>
        <div>
          <label style="font-size:0.78rem; font-weight:bold; color:#334155;">{{ t.imie_nazwisko }}</label>
          <input type="text" id="newEmpName" placeholder="{{ t.imie_nazwisko }}">
        </div>
        <div>
          <label style="font-size:0.78rem; font-weight:bold; color:#334155;">Stanowisko</label>
          <input type="text" id="newEmpRole" placeholder="np. Operator linii, Pakowacz...">
        </div>
        <div style="flex: 0 0 auto; display: flex; gap: 8px;">
          <button type="button" id="btnSaveEmp" onclick="registerEmployee()" style="background: #0284c7; color: white; border: none; padding: 9px 15px; border-radius: 6px; font-weight: bold; cursor: pointer;">Zapisz i QR</button>
          <button type="button" id="btnCancelEdit" onclick="cancelEdit()" style="background: #64748b; color: white; border: none; padding: 9px 12px; border-radius: 6px; font-weight: bold; cursor: pointer; display: none;">Anuluj</button>
          <button type="button" onclick="generateAllQRs()" style="background: #0f766e; color: white; border: none; padding: 9px 15px; border-radius: 6px; font-weight: bold; cursor: pointer;">🏷️ Wszystkie QR</button>
        </div>
      </div>

      <div id="qrResultContainer" class="qr-result" style="display: none;">
        <div id="printableSingleBadge">
          <div class="badge-54x85" id="singleBadgeCard">
            <div class="badge-header">KARTA PRACOWNIKA • 54×85 mm</div>
            <div class="badge-qr-box" id="qrcode"></div>
            <div>
              <div class="badge-id" id="qrEmpId"></div>
              <div class="badge-name" id="qrEmpName"></div>
            </div>
          </div>
        </div>
        <div style="display: flex; flex-direction: column; gap: 8px;">
          <button type="button" onclick="printSingleBadge()" style="background: #475569; color: white; border: none; padding: 10px 16px; border-radius: 6px; cursor: pointer; font-weight: bold;">🖨️ Drukuj Plakietkę (54×85 mm)</button>
          <button type="button" onclick="document.getElementById('qrResultContainer').style.display='none'" style="background: #e2e8f0; color: #334155; border: none; padding: 8px 14px; border-radius: 6px; cursor: pointer; font-size: 0.85rem;">Ukryj podgląd</button>
        </div>
      </div>

      <div id="printableAllBadges" class="all-badges-container" style="display: none;">
        <div class="all-badges-toolbar">
          <div>
            <strong>🏷️ Wszystkie Plakietki QR (54 mm × 85 mm)</strong>
            <span id="allBadgesCount" style="margin-left: 8px; color: #475569; font-size: 0.85rem;"></span>
          </div>
          <div style="display: flex; gap: 8px;">
            <button type="button" onclick="printAllBadges()" style="background: #0284c7; color: white; border: none; padding: 8px 14px; border-radius: 6px; cursor: pointer; font-weight: bold;">🖨️ Drukuj Wszystkie Plakietki</button>
            <button type="button" onclick="closeAllQRs()" style="background: #64748b; color: white; border: none; padding: 8px 12px; border-radius: 6px; cursor: pointer; font-size: 0.85rem;">Zamknij</button>
          </div>
        </div>
        <div id="allBadgesGrid" class="badges-grid"></div>
      </div>

      <div style="margin-top: 20px; border-top: 2px solid #f1f5f9; padding-top: 15px;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; flex-wrap: wrap; gap: 10px;">
          <h4 style="margin: 0; color: #334155;">📋 Lista Pracowników (<span id="empCountLabel">0</span>)</h4>

          <div style="display: flex; gap: 10px; align-items: center; flex-wrap: wrap;">
            <input type="text" id="filterPositionInput" placeholder="Filtruj po stanowisku..." oninput="renderRegisteredEmployeesTable()" style="padding: 6px 10px; border-radius: 6px; border: 1px solid #cbd5e1; font-size: 0.85rem; width: 190px;">

            <select id="filterSessionSelect" onchange="renderRegisteredEmployeesTable()" style="padding: 6px 10px; border-radius: 6px; border: 1px solid #cbd5e1; font-size: 0.85rem; width: auto;">
              <option value="all">Wszyscy pracownicy</option>
              <option value="active">Obecnie na linii</option>
              <option value="inactive">Poza linią</option>
            </select>

            <select id="sortSelect" onchange="changeSortMode(this.value)" style="padding: 6px 10px; border-radius: 6px; border: 1px solid #cbd5e1; font-size: 0.85rem; width: auto;">
              <option value="name-asc">Imię i Nazwisko (A-Z)</option>
              <option value="name-desc">Imię i Nazwisko (Z-A)</option>
              <option value="id-asc">ID Pracownika (Rosnąco)</option>
              <option value="id-desc">ID Pracownika (Malejąco)</option>
              <option value="pos-asc">Stanowisko (A-Z)</option>
              <option value="default">Kolejność domyślna</option>
            </select>

            <input type="text" id="searchEmpInput" placeholder="Szukaj ID lub nazwiska..." oninput="renderRegisteredEmployeesTable()" style="padding: 6px 10px; border-radius: 6px; border: 1px solid #cbd5e1; font-size: 0.85rem; width: 210px;">
          </div>
        </div>

        <table>
          <thead>
            <tr>
              <th class="sortable-th" onclick="toggleSort('id')">ID Pracownika <span id="sort-icon-id">↕️</span></th>
              <th class="sortable-th" onclick="toggleSort('name')">{{ t.imie_nazwisko }} <span id="sort-icon-name">↕</span></th>
              <th class="sortable-th" onclick="toggleSort('pos')">Stanowisko <span id="sort-icon-pos">↕</span></th>
              <th>Status na linii</th>
              <th style="text-align: right; padding-right: 15px;">{{ t.kol_akcja }}</th>
            </tr>
          </thead>
          <tbody id="registeredEmpTableBody"></tbody>
        </table>
      </div>
    </div>
  </div>

  {% if role == 'admin' %}
  <div id="subtabAdmin" style="display:none;">
    <div class="panel-card" style="background: #fef2f2; border: 1px solid #fecaca;">
      <h3 style="margin-top:0; color: #991b1b;">🔐 {{ t.zarzadzanie_uzyt }}</h3>
      <form method="POST" action="/admin/dodaj-uzytkownika" style="background: white; padding: 16px; border-radius: 8px; margin-bottom: 20px;">
        <input type="hidden" name="powrot_modul" value="produkcja">
        <div class="row-flex">
          <div>
            <label style="font-size:0.8rem; font-weight:bold;">{{ t.login }}</label>
            <input type="text" name="username" placeholder="np. jan_kowalski" required>
          </div>
          <div>
            <label style="font-size:0.8rem; font-weight:bold;">{{ t.haslo }}</label>
            <input type="password" name="password" placeholder="Nowe hasło" required>
          </div>
          <div>
            <label style="font-size:0.8rem; font-weight:bold;">{{ t.imie_nazwisko }}</label>
            <input type="text" name="pelna_nazwa" placeholder="np. Jan Kowalski" required>
          </div>
          <div>
            <label style="font-size:0.8rem; font-weight:bold;">{{ t.rola }}</label>
            <select name="rola">
              <option value="operator">operator</option>
              <option value="admin">admin</option>
            </select>
          </div>
        </div>

        <div style="margin-top: 12px; display: flex; gap: 14px; flex-wrap: wrap; font-size: 0.85rem;">
          <label style="display:flex; align-items:center; gap:5px; font-weight:bold;"><input type="checkbox" name="dostep" value="mod_produkcja" checked style="width:auto;"> Moduł Produkcyjny</label>
          <label style="display:flex; align-items:center; gap:5px; font-weight:bold;"><input type="checkbox" name="dostep" value="mod_koncowki" checked style="width:auto;"> Moduł Końcówki Linii</label>
          <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="1" checked style="width:auto;"> {{ t.linia_1 }}</label>
          <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="2" checked style="width:auto;"> {{ t.linia_2 }}</label>
          <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="3" checked style="width:auto;"> {{ t.linia_3 }}</label>
          <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="4" checked style="width:auto;"> {{ t.linia_4 }}</label>
          <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="plan" style="width:auto;"> {{ t.plan_tab }}</label>
          <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="progress" checked style="width:auto;"> {{ t.prog_tab }}</label>
          <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="wydajnosc" checked style="width:auto;"> {{ t.wydajnosc_tab }}</label>
          <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="raporty" checked style="width:auto;"> {{ t.raporty_tab }}</label>
          <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="db_art" style="width:auto;"> {{ t.db_art_tab }}</label>
          <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="admin" style="width:auto;"> {{ t.admin_tab }}</label>
        </div>

        <button type="submit" style="background: #dc2626; color: white; border: none; padding: 10px 18px; border-radius: 6px; font-weight: bold; cursor: pointer; margin-top: 14px;">💾 {{ t.zapisz_uzyt_btn }}</button>
      </form>

      <div style="background: white; padding: 14px; border-radius: 8px;">
        <h4 style="margin-top:0;">{{ t.zarejestrowani_uzyt }}:</h4>
        <table>
          <thead>
            <tr><th>{{ t.login }}</th><th>{{ t.imie_nazwisko }}</th><th>{{ t.rola }}</th><th>{{ t.uprawnienia_lbl.split()[0] }}</th><th>{{ t.kol_akcja }}</th></tr>
          </thead>
          <tbody>
            {% for usr, info in uzytkownicy_baza.items() %}
            <tr>
              <td><b>{{ usr }}</b></td>
              <td>{{ info.pelna_nazwa }}</td>
              <td>{{ info.rola }}</td>
              <td>{{ info.dostep | join(', ') }}</td>
              <td>
                {% if usr != 'admin' %}
                  <a href="/admin/usun-uzytkownika/{{ usr }}?powrot_modul=produkcja" style="color: #dc2626; text-decoration: none; font-weight: bold;">{{ t.usun }}</a>
                {% else %}
                  <i>{{ t.systemowy }}</i>
                {% endif %}
              </td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
    </div>
  </div>
  {% endif %}
</div>

<script>
  const lines = [1, 2, 3, 4, 5, 6];
  const defaultReasons = ["Awaria mechaniczna", "Brak materiału", "Przezbrojenie", "Przerwa pracownicza", "Problem z jakością", "Inny"];

  let employees = {{ employees_json | safe }};
  let snapshotData = {{ initial_snapshot_json | safe }};

  let editingEmpId = null;
  let currentSortMode = 'name-asc';
  let serverOffsetMs = 0;

  if (snapshotData && snapshotData.serverNowIso) {
    serverOffsetMs = new Date(snapshotData.serverNowIso).getTime() - Date.now();
  }

  function getSyncedNow() {
    return new Date(Date.now() + serverOffsetMs);
  }

  function formatTime(seconds) {
    seconds = Math.max(0, Math.floor(seconds || 0));
    const h = Math.floor(seconds / 3600);
    const m = Math.floor((seconds % 3600) / 60);
    const s = seconds % 60;
    return (h < 10 ? '0' + h : h) + ':' + (m < 10 ? '0' + m : m) + ':' + (s < 10 ? '0' + s : s);
  }

  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text == null ? '' : String(text);
    return div.innerHTML;
  }

  function getStatusBadgeClass(status) {
    if (status === 'RUNNING') return 'status-running';
    if (status === 'STOPPED') return 'status-stopped';
    return 'status-idle';
  }

  function getStatusText(status) {
    if (status === 'RUNNING') return 'W PRACY';
    if (status === 'STOPPED') return 'POSTÓJ';
    return 'ZAKOŃCZONO';
  }

  function switchSubtab(tab) {
    document.getElementById('subtabLines').style.display = tab === 'lines' ? 'block' : 'none';
    document.getElementById('subtabEmployees').style.display = tab === 'employees' ? 'block' : 'none';
    const adminEl = document.getElementById('subtabAdmin');
    if (adminEl) adminEl.style.display = tab === 'admin' ? 'block' : 'none';

    document.getElementById('tabBtnLines').classList.toggle('active', tab === 'lines');
    document.getElementById('tabBtnEmployees').classList.toggle('active', tab === 'employees');
    const adminBtn = document.getElementById('tabBtnAdmin');
    if (adminBtn) adminBtn.classList.toggle('active', tab === 'admin');
  }

  function renderLinesSkeleton() {
    const container = document.getElementById('linesContainer');
    container.innerHTML = '';
    lines.forEach(function(id) {
      const card = document.createElement('div');
      card.className = 'card idle';
      card.id = 'card-' + id;
      card.innerHTML =
        '<div class="card-header"><h2 class="card-title">Linia ' + id + '</h2><span class="status-badge status-idle" id="badge-' + id + '">ZAKOŃCZONO</span></div>' +
        '<div class="timers-container">' +
          '<div class="timer-box"><div class="timer-label">OBECNA PRACA</div><div class="timer-val" id="work-timer-' + id + '">00:00:00</div></div>' +
          '<div class="timer-box"><div class="timer-label">OBECNY POSTÓJ</div><div class="timer-val" id="stop-timer-' + id + '">00:00:00</div></div>' +
        '</div>' +
        '<div class="daily-timers-container">' +
          '<div class="daily-timer-box"><div class="daily-timer-label">ŁĄCZNA PRACA DZISIAJ</div><div class="daily-timer-val daily-work-val" id="daily-work-timer-' + id + '">00:00:00</div></div>' +
          '<div class="daily-timer-box"><div class="daily-timer-label">ŁĄCZNY POSTÓJ DZISIAJ</div><div class="daily-timer-val daily-stop-val" id="daily-stop-timer-' + id + '">00:00:00</div></div>' +
        '</div>' +
        '<div class="total-emp-time-box">' +
          '<span class="total-emp-time-label">👥 Suma Czasu Pracowników:</span>' +
          '<span class="total-emp-time-val" id="total-emp-timer-' + id + '">00:00:00</span>' +
        '</div>' +
        '<div class="emp-section">' +
          '<div class="emp-section-title">🔍 Skaner ID Pracownika</div>' +
          '<div class="scan-box">' +
            '<input type="text" id="manual-emp-' + id + '" placeholder="Zeskanuj lub wpisz ID..." onkeypress="handleScanKeyPress(event, ' + id + ')">' +
            '<button type="button" onclick="assignEmpManual(' + id + ')" style="background:#0284c7; color:white; border:none; border-radius:6px; cursor:pointer; padding:0 12px; font-weight:bold;">OK</button>' +
          '</div>' +
          '<div class="emp-list" id="emp-list-' + id + '"></div>' +
        '</div>' +
        '<div class="controls">' +
          '<div class="btn-group">' +
            '<button type="button" class="action-btn btn-start" id="btn-start-' + id + '" onclick="lineAction(' + id + ', \\'start\\')">▶ Start</button>' +
            '<button type="button" class="action-btn btn-stop" id="btn-stop-' + id + '" onclick="handleStopClick(' + id + ')">⏹ Postój</button>' +
            '<button type="button" class="action-btn btn-end" id="btn-end-' + id + '" onclick="lineAction(' + id + ', \\'end\\')">🏁 Koniec</button>' +
          '</div>' +
          '<div class="reason-box" id="reason-box-' + id + '">' +
            '<label>Przyczyna postoju (wymagana przed zatrzymaniem):</label>' +
            '<select id="reason-select-' + id + '" style="margin-bottom: 4px;">' +
              '<option value="">-- Wybierz przyczynę --</option>' +
              defaultReasons.map(function(r) { return '<option value="' + r + '">' + r + '</option>'; }).join('') +
            '</select>' +
            '<input type="text" id="custom-reason-' + id + '" placeholder="Wpisz własną przyczynę...">' +
          '</div>' +
        '</div>' +
        '<div class="history"><strong>Ostatnie zdarzenia:</strong><div id="history-list-' + id + '"></div></div>';
      container.appendChild(card);
    });
  }

  function applySnapshotToUI(snap) {
    if (!snap || !snap.lines) return;
    snapshotData = snap;
    if (snap.serverNowIso) {
      serverOffsetMs = new Date(snap.serverNowIso).getTime() - Date.now();
    }

    lines.forEach(function(id) {
      const lData = snap.lines[String(id)];
      if (!lData) return;

      const card = document.getElementById('card-' + id);
      if (card) card.className = 'card ' + String(lData.status || 'IDLE').toLowerCase();

      const badge = document.getElementById('badge-' + id);
      if (badge) {
        badge.className = 'status-badge ' + getStatusBadgeClass(lData.status);
        badge.innerText = getStatusText(lData.status);
      }

      const btnStart = document.getElementById('btn-start-' + id);
      const btnEnd = document.getElementById('btn-end-' + id);
      if (btnStart) btnStart.className = 'action-btn btn-start ' + (lData.status === 'RUNNING' ? 'btn-disabled' : '');
      if (btnEnd) btnEnd.className = 'action-btn btn-end ' + (lData.status === 'IDLE' ? 'btn-disabled' : '');

      const selectEl = document.getElementById('reason-select-' + id);
      const customEl = document.getElementById('custom-reason-' + id);
      if (selectEl && document.activeElement !== selectEl && lData.reason && defaultReasons.includes(lData.reason)) {
        selectEl.value = lData.reason;
      }
      if (customEl && document.activeElement !== customEl && lData.reason && !defaultReasons.includes(lData.reason)) {
        customEl.value = lData.reason;
      }

      const empListEl = document.getElementById('emp-list-' + id);
      if (empListEl) {
        const activeEmps = lData.activeEmployees || [];
        if (activeEmps.length === 0) {
          empListEl.innerHTML = '<div style="color:#94a3b8; text-align:center; padding:4px;">Brak zalogowanych pracowników</div>';
        } else {
          empListEl.innerHTML = activeEmps.map(function(s) {
            return '<div class="emp-item">' +
              '<span><strong>' + escapeHtml(s.empId) + '</strong>: ' + escapeHtml(s.empName) + '</span>' +
              '<button type="button" class="btn-remove-emp" onclick="removeEmpSession(\\'' + escapeHtml(s.empId) + '\\', ' + id + ')">Wymelduj</button>' +
            '</div>';
          }).join('');
        }
      }

      const histEl = document.getElementById('history-list-' + id);
      if (histEl) {
        const events = lData.recentEvents || [];
        if (events.length === 0) {
          histEl.innerHTML = '<div style="color:#94a3b8;">Brak zdarzeń</div>';
        } else {
          histEl.innerHTML = events.map(function(h) {
            const tStr = new Date(h.startIso).toLocaleTimeString();
            const who = (h.startedByLogin || h.endedByLogin) ? ' [' + escapeHtml(h.startedByLogin || '-') + ' → ' + escapeHtml(h.endedByLogin || '-') + ']' : '';
            return '<div class="history-item">[' + tStr + '] <strong>' + escapeHtml(h.type) + '</strong> (' + formatTime(h.durationSec) + ')' + who + (h.reason && h.reason !== '-' ? ' - ' + escapeHtml(h.reason) : '') + '</div>';
          }).join('');
        }
      }
    });

    tickLiveTimers();
    renderRegisteredEmployeesTable();
  }

  function tickLiveTimers() {
    if (!snapshotData || !snapshotData.lines) return;
    const nowMs = getSyncedNow().getTime();
    const snapAgeSec = snapshotData.serverNowIso
      ? Math.max(0, Math.floor((nowMs - new Date(snapshotData.serverNowIso).getTime()) / 1000))
      : 0;

    lines.forEach(function(id) {
      const lData = snapshotData.lines[String(id)];
      if (!lData) return;

      let liveElapsed = 0;
      if (lData.status !== 'IDLE' && lData.startTimeIso) {
        liveElapsed = Math.max(0, Math.floor((nowMs - new Date(lData.startTimeIso).getTime()) / 1000));
      }

      const workEl = document.getElementById('work-timer-' + id);
      const stopEl = document.getElementById('stop-timer-' + id);
      if (workEl && stopEl) {
        if (lData.status === 'RUNNING') {
          workEl.innerText = formatTime(liveElapsed);
          workEl.className = 'timer-val active-work';
          stopEl.innerText = '00:00:00';
          stopEl.className = 'timer-val';
        } else if (lData.status === 'STOPPED') {
          stopEl.innerText = formatTime(liveElapsed);
          stopEl.className = 'timer-val active-stop';
          workEl.innerText = '00:00:00';
          workEl.className = 'timer-val';
        } else {
          workEl.innerText = '00:00:00';
          workEl.className = 'timer-val';
          stopEl.innerText = '00:00:00';
          stopEl.className = 'timer-val';
        }
      }

      const liveDailyWork = (lData.dailyWorkSec || 0) + (lData.status === 'RUNNING' ? Math.max(0, liveElapsed - (lData.currentElapsedSec || 0)) : 0);
      const liveDailyStop = (lData.dailyStopSec || 0) + (lData.status === 'STOPPED' ? Math.max(0, liveElapsed - (lData.currentElapsedSec || 0)) : 0);

      const dwEl = document.getElementById('daily-work-timer-' + id);
      const dsEl = document.getElementById('daily-stop-timer-' + id);
      if (dwEl) dwEl.innerText = formatTime(liveDailyWork);
      if (dsEl) dsEl.innerText = formatTime(liveDailyStop);

      const teEl = document.getElementById('total-emp-timer-' + id);
      if (teEl) {
        const activeCount = (lData.activeEmployees || []).length;
        teEl.innerText = formatTime((lData.totalEmpSec || 0) + (activeCount * snapAgeSec));
      }
    });
  }

  async function refreshFromServer() {
    try {
      const resp = await fetch('/api/produkcja/snapshot?_=' + Date.now(), { cache: 'no-store' });
      if (resp.ok) {
        const data = await resp.json();
        applySnapshotToUI(data);
      }
    } catch (e) {}
  }

  async function lineAction(lineId, action, reason = '') {
    const resp = await fetch('/api/produkcja/linia-akcja', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ lineId: lineId, action: action, reason: reason })
    });
    const data = await resp.json();
    if (!resp.ok) {
      alert(data.error || 'Błąd zapisu stanu linii.');
      return;
    }
    if (action === 'start') {
      const sEl = document.getElementById('reason-select-' + lineId);
      const cEl = document.getElementById('custom-reason-' + lineId);
      if (sEl) sEl.value = '';
      if (cEl) cEl.value = '';
    }
    applySnapshotToUI(data.snapshot);
  }

  function handleStopClick(lineId) {
    const selectEl = document.getElementById('reason-select-' + lineId);
    const customEl = document.getElementById('custom-reason-' + lineId);
    const reason = (customEl && customEl.value.trim()) || (selectEl ? selectEl.value : '');
    if (!reason) {
      alert("Wybierz lub wpisz przyczynę postoju z pola poniżej!");
      if (selectEl) selectEl.focus();
      return;
    }
    lineAction(lineId, 'stop', reason);
  }

  function handleScanKeyPress(e, lineId) {
    if (e.key === 'Enter') assignEmpManual(lineId);
  }

  async function assignEmpManual(lineId) {
    const inputEl = document.getElementById('manual-emp-' + lineId);
    const empId = (inputEl.value || '').trim();
    if (!empId) return;

    let customName = '';
    if (!employees[empId]) {
      if (confirm("Nieznany pracownik ID: " + empId + ". Czy chcesz go dodać do wspólnej bazy?")) {
        const name = prompt("Podaj Imię i Nazwisko pracownika:");
        if (!name || !name.trim()) return;
        customName = name.trim();
      } else {
        return;
      }
    }

    const resp = await fetch('/api/produkcja/skan-pracownika', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ lineId: lineId, empId: empId, newName: customName })
    });
    const data = await resp.json();
    if (!resp.ok) {
      alert(data.error || 'Błąd przypisywania pracownika.');
      return;
    }
    if (data.employees) employees = data.employees;
    inputEl.value = '';
    applySnapshotToUI(data.snapshot);
  }

  async function removeEmpSession(empId, lineId) {
    const resp = await fetch('/api/produkcja/wymelduj-pracownika', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ lineId: lineId, empId: empId })
    });
    const data = await resp.json();
    if (resp.ok && data.snapshot) {
      applySnapshotToUI(data.snapshot);
    }
  }

  async function registerEmployee() {
    const idInput = document.getElementById('newEmpId');
    const nameInput = document.getElementById('newEmpName');
    const roleInput = document.getElementById('newEmpRole');

    const empId = idInput.value.trim();
    const imieNazwisko = nameInput.value.trim();
    const stanowisko = roleInput.value.trim();

    if (!empId || !imieNazwisko) {
      alert("Wypełnij ID oraz Imię i Nazwisko pracownika!");
      return;
    }

    const resp = await fetch('/api/pracownicy/zapisz', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        empId: empId,
        imieNazwisko: imieNazwisko,
        stanowisko: stanowisko,
        isEdit: Boolean(editingEmpId)
      })
    });
    const data = await resp.json();
    if (!resp.ok) {
      alert(data.error || 'Błąd zapisu pracownika.');
      return;
    }
    employees = data.employees;
    renderRegisteredEmployeesTable();
    generateQR(empId, imieNazwisko);

    if (!editingEmpId) {
      idInput.value = '';
      nameInput.value = '';
      roleInput.value = '';
    } else {
      cancelEdit();
    }
  }

  function generateQR(id, name) {
    const rec = employees[id];
    const empName = name !== undefined ? name : (rec ? rec.imie_nazwisko : '');
    const container = document.getElementById('qrcode');
    container.innerHTML = '';
    new QRCode(container, { text: id, width: 160, height: 160, correctLevel: QRCode.CorrectLevel.M });
    document.getElementById('qrEmpId').textContent = id;
    document.getElementById('qrEmpName').textContent = empName;
    document.getElementById('qrResultContainer').style.display = 'flex';
  }

  function generateAllQRs() {
    const keys = getSortedFilteredEmployeeKeys();
    const grid = document.getElementById('allBadgesGrid');
    const container = document.getElementById('printableAllBadges');
    const countEl = document.getElementById('allBadgesCount');

    if (keys.length === 0) {
      alert("Brak pracowników spełniających kryteria filtrów.");
      return;
    }

    grid.innerHTML = '';
    countEl.textContent = '(' + keys.length + ' plakietek)';
    container.style.display = 'block';

    keys.forEach(function(id) {
      const badge = document.createElement('div');
      badge.className = 'badge-54x85';

      const header = document.createElement('div');
      header.className = 'badge-header';
      header.textContent = 'KARTA PRACOWNIKA • 54×85 mm';

      const qrBox = document.createElement('div');
      qrBox.className = 'badge-qr-box';

      const infoBox = document.createElement('div');
      const idEl = document.createElement('div');
      idEl.className = 'badge-id';
      idEl.textContent = id;

      const nameEl = document.createElement('div');
      nameEl.className = 'badge-name';
      nameEl.textContent = employees[id].imie_nazwisko;

      infoBox.appendChild(idEl);
      infoBox.appendChild(nameEl);
      badge.appendChild(header);
      badge.appendChild(qrBox);
      badge.appendChild(infoBox);
      grid.appendChild(badge);

      new QRCode(qrBox, { text: id, width: 160, height: 160, correctLevel: QRCode.CorrectLevel.M });
    });

    container.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function closeAllQRs() {
    document.getElementById('printableAllBadges').style.display = 'none';
    document.getElementById('allBadgesGrid').innerHTML = '';
  }

  function printSingleBadge() {
    document.body.classList.remove('print-all');
    document.body.classList.add('print-single');
    window.print();
    setTimeout(function() { document.body.classList.remove('print-single'); }, 300);
  }

  function printAllBadges() {
    document.body.classList.remove('print-single');
    document.body.classList.add('print-all');
    window.print();
    setTimeout(function() { document.body.classList.remove('print-all'); }, 300);
  }

  function editEmployee(id) {
    const rec = employees[id];
    if (!rec) return;
    editingEmpId = id;
    document.getElementById('newEmpId').value = id;
    document.getElementById('newEmpId').disabled = true;
    document.getElementById('newEmpName').value = rec.imie_nazwisko || '';
    document.getElementById('newEmpRole').value = rec.stanowisko || '';
    document.getElementById('btnSaveEmp').innerText = '💾 Zapisz zmiany';
    document.getElementById('btnCancelEdit').style.display = 'inline-block';
    document.getElementById('newEmpRole').focus();
  }

  function cancelEdit() {
    editingEmpId = null;
    const idInput = document.getElementById('newEmpId');
    idInput.value = '';
    idInput.disabled = false;
    document.getElementById('newEmpName').value = '';
    document.getElementById('newEmpRole').value = '';
    document.getElementById('btnSaveEmp').innerText = 'Zapisz i QR';
    document.getElementById('btnCancelEdit').style.display = 'none';
  }

  async function deleteEmployee(id) {
    const rec = employees[id];
    if (!rec) return;
    if (confirm("Czy na pewno chcesz usunąć pracownika " + id + " (" + rec.imie_nazwisko + ") ze wspólnej bazy?")) {
      const resp = await fetch('/api/pracownicy/usun', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ empId: id })
      });
      const data = await resp.json();
      if (resp.ok && data.employees) {
        employees = data.employees;
        renderRegisteredEmployeesTable();
      }
    }
  }

  function getActiveEmpMap() {
    const map = {};
    if (!snapshotData || !snapshotData.lines) return map;
    lines.forEach(function(lid) {
      const lData = snapshotData.lines[String(lid)];
      if (lData && lData.activeEmployees) {
        lData.activeEmployees.forEach(function(act) { map[act.empId] = lid; });
      }
    });
    return map;
  }

  function changeSortMode(mode) {
    currentSortMode = mode;
    renderRegisteredEmployeesTable();
  }

  function toggleSort(field) {
    if (field === 'id') currentSortMode = currentSortMode === 'id-asc' ? 'id-desc' : 'id-asc';
    else if (field === 'name') currentSortMode = currentSortMode === 'name-asc' ? 'name-desc' : 'name-asc';
    else if (field === 'pos') currentSortMode = currentSortMode === 'pos-asc' ? 'default' : 'pos-asc';
    document.getElementById('sortSelect').value = currentSortMode;
    renderRegisteredEmployeesTable();
  }

  function getSortedFilteredEmployeeKeys() {
    const searchVal = (document.getElementById('searchEmpInput')?.value || '').toLowerCase().trim();
    const posFilterVal = (document.getElementById('filterPositionInput')?.value || '').toLowerCase().trim();
    const sessionFilter = document.getElementById('filterSessionSelect')?.value || 'all';
    const activeMap = getActiveEmpMap();

    let keys = Object.keys(employees);

    if (searchVal) {
      keys = keys.filter(function(k) {
        const rec = employees[k];
        return k.toLowerCase().includes(searchVal) || String(rec.imie_nazwisko || '').toLowerCase().includes(searchVal);
      });
    }
    if (posFilterVal) {
      keys = keys.filter(function(k) {
        return String(employees[k].stanowisko || '').toLowerCase().includes(posFilterVal);
      });
    }
    if (sessionFilter === 'active') keys = keys.filter(function(k) { return Boolean(activeMap[k]); });
    else if (sessionFilter === 'inactive') keys = keys.filter(function(k) { return !activeMap[k]; });

    keys.sort(function(a, b) {
      const recA = employees[a];
      const recB = employees[b];
      if (currentSortMode === 'name-asc') return String(recA.imie_nazwisko || '').localeCompare(String(recB.imie_nazwisko || ''));
      if (currentSortMode === 'name-desc') return String(recB.imie_nazwisko || '').localeCompare(String(recA.imie_nazwisko || ''));
      if (currentSortMode === 'id-asc') return a.localeCompare(b);
      if (currentSortMode === 'id-desc') return b.localeCompare(a);
      if (currentSortMode === 'pos-asc') return String(recA.stanowisko || '').localeCompare(String(recB.stanowisko || ''));
      return (recA.kolejnosc || 0) - (recB.kolejnosc || 0);
    });

    return keys;
  }

  function renderRegisteredEmployeesTable() {
    const tbody = document.getElementById('registeredEmpTableBody');
    if (!tbody) return;
    const keys = getSortedFilteredEmployeeKeys();
    const activeMap = getActiveEmpMap();
    const countLabel = document.getElementById('empCountLabel');
    if (countLabel) countLabel.textContent = keys.length;

    tbody.innerHTML = '';
    if (keys.length === 0) {
      tbody.innerHTML = '<tr><td colspan="5" style="text-align:center; color:#94a3b8; padding:15px;">Brak pracowników spełniających kryteria</td></tr>';
      return;
    }

    keys.forEach(function(id) {
      const rec = employees[id];
      const tr = document.createElement('tr');

      const tdId = document.createElement('td');
      tdId.innerHTML = '<strong>' + escapeHtml(id) + '</strong>';

      const tdName = document.createElement('td');
      tdName.textContent = rec.imie_nazwisko || '';

      const tdPos = document.createElement('td');
      tdPos.innerHTML = rec.stanowisko ? escapeHtml(rec.stanowisko) : '<span style="color:#94a3b8; font-style:italic;">nie uzupełniono</span>';

      const tdStatus = document.createElement('td');
      if (activeMap[id]) {
        tdStatus.innerHTML = '<span style="background:#d1fae5; color:#047857; padding:2px 8px; border-radius:12px; font-size:0.75rem; font-weight:bold;">Na Linii ' + activeMap[id] + '</span>';
      } else {
        tdStatus.innerHTML = '<span style="color:#64748b; font-size:0.78rem;">Poza linią</span>';
      }

      const tdActions = document.createElement('td');
      tdActions.style.textAlign = 'right';

      const btnQr = document.createElement('button');
      btnQr.type = 'button';
      btnQr.textContent = 'QR';
      btnQr.style.cssText = 'background:#0284c7; color:white; border:none; padding:4px 8px; border-radius:4px; cursor:pointer; margin-right:4px; font-size:0.75rem;';
      btnQr.onclick = function() { generateQR(id, rec.imie_nazwisko); };

      const btnEdit = document.createElement('button');
      btnEdit.type = 'button';
      btnEdit.textContent = 'Edytuj';
      btnEdit.style.cssText = 'background:#d97706; color:white; border:none; padding:4px 8px; border-radius:4px; cursor:pointer; margin-right:4px; font-size:0.75rem;';
      btnEdit.onclick = function() { editEmployee(id); };

      const btnDel = document.createElement('button');
      btnDel.type = 'button';
      btnDel.textContent = 'Usuń';
      btnDel.style.cssText = 'background:#ef4444; color:white; border:none; padding:4px 8px; border-radius:4px; cursor:pointer; font-size:0.75rem;';
      btnDel.onclick = function() { deleteEmployee(id); };

      tdActions.appendChild(btnQr);
      tdActions.appendChild(btnEdit);
      tdActions.appendChild(btnDel);

      tr.appendChild(tdId);
      tr.appendChild(tdName);
      tr.appendChild(tdPos);
      tr.appendChild(tdStatus);
      tr.appendChild(tdActions);
      tbody.appendChild(tr);
    });
  }

  window.addEventListener('DOMContentLoaded', function() {
    renderLinesSkeleton();
    applySnapshotToUI(snapshotData);
    setInterval(tickLiveTimers, 1000);
    setInterval(refreshFromServer, 4000);
  });
</script>
</body>
</html>
"""

KONCOWKI_TEMPLATE = """
<!DOCTYPE html>
<html lang="{{ lang }}">
<head>
    <meta charset="UTF-8">
    <title>{{ t.tytul }}</title>
    <style>
        body { font-family: Arial, sans-serif; background: #f0f2f5; margin: 0; padding: 20px; }
        .container { max-width: 1250px; background: white; padding: 25px; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); margin: auto; }
        h2, h3 { color: #333; text-align: center; }
        .header-bar { display: flex; justify-content: space-between; align-items: center; background: #e9ecef; padding: 10px 15px; border-radius: 6px; margin-bottom: 20px; flex-wrap: wrap; gap: 10px; }
        .lang-box { display: inline-flex; gap: 4px; background: white; padding: 3px 6px; border-radius: 6px; border: 1px solid #ccc; }
        .lang-btn { padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: bold; text-decoration: none; color: #333; }
        .lang-btn.active { background: #007bff; color: white; }
        .tabs { display: flex; gap: 5px; margin-bottom: 20px; border-bottom: 2px solid #ddd; padding-bottom: 10px; flex-wrap: wrap; }
        .tabs a { padding: 10px 18px; background: #e2e6ea; color: #333; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 14px; }
        .tabs a.active { background: #007bff; color: white; }
        .tabs a.plan-tab { background: #17a2b8; color: white; }
        .tabs a.plan-tab.active { background: #117a8b; color: white; }
        .tabs a.progress-tab { background: #28a745; color: white; }
        .tabs a.progress-tab.active { background: #1e7e34; color: white; }
        .tabs a.wydajnosc-tab { background: #6f42c1; color: white; }
        .tabs a.wydajnosc-tab.active { background: #5a32a3; color: white; }
        .tabs a.raporty-tab { background: #fd7e14; color: white; }
        .tabs a.raporty-tab.active { background: #e8590c; color: white; }
        .tabs a.db-tab { background: #ffc107; color: #333; }
        .tabs a.db-tab.active { background: #d39e00; color: white; }
        .tabs a.admin-tab { background: #dc3545; color: white; }
        .tabs a.admin-tab.active { background: #bd2130; color: white; }
        label { font-weight: bold; display: block; margin-top: 10px; color: #444; font-size: 13px; }
        input, select, textarea { width: 100%; padding: 8px; margin-top: 4px; border: 1px solid #ccc; border-radius: 5px; box-sizing: border-box; font-size: 14px; }
        .row { display: flex; gap: 15px; flex-wrap: wrap; }
        .row > div { flex: 1; min-width: 180px; }
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
        .raporty-section { background: #fff4e6; padding: 20px; border-radius: 8px; border: 1px solid #ffd8a8; }
        .db-section { background: #fff3cd; padding: 20px; border-radius: 8px; border: 1px solid #ffeeba; }
        .admin-section { background: #f8d7da; padding: 20px; border-radius: 8px; border: 1px solid #f5c6cb; }
        .error-banner { background: #f8d7da; color: #721c24; padding: 12px; border-radius: 6px; margin-bottom: 15px; border: 1px solid #f5c6cb; font-weight: bold; text-align: center; }

        .progress-table { width: 100%; border-collapse: collapse; margin-top: 15px; background: white; font-size: 13px; }
        .progress-table th, .progress-table td { border: 1px solid #cbd5e1; padding: 8px 10px; text-align: left; }
        .progress-table th { background-color: #f1f5f9; cursor: pointer; user-select: none; }
        .progress-table th:hover { background-color: #e2e8f0; }
        .filter-input { width: 100%; padding: 5px 6px; margin-top: 5px; font-size: 12px; border: 1px solid #94a3b8; border-radius: 4px; box-sizing: border-box; font-weight: normal; cursor: text; }

        .wyd-layout { display: flex; gap: 20px; align-items: flex-start; flex-wrap: wrap; }
        .wyd-table-col { flex: 3; min-width: 550px; }
        .wyd-side-summary { flex: 1; min-width: 240px; background: white; border: 2px solid #6f42c1; border-radius: 8px; padding: 16px; box-shadow: 0 2px 6px rgba(0,0,0,0.05); }
        .wyd-stat-row { margin-bottom: 12px; padding-bottom: 10px; border-bottom: 1px solid #eee; }
        .wyd-stat-row:last-child { border-bottom: none; margin-bottom: 0; padding-bottom: 0; }

        @media print {
            .header-bar, .tabs, form, button, .filter-input { display: none !important; }
            body { background: white; padding: 0; }
            .container { box-shadow: none; max-width: 100%; padding: 0; }
        }
    </style>
    <script>
        const baza = {{ produkty_json | safe }};
        const calyPlan = {{ plan_json | safe }};
        const zrealizowanoDzien = {{ zrealizowane_json | safe }};
        const aktualnaLinia = "{{ linia }}";
        const draftKey = "paleta_draft_linia_" + aktualnaLinia;

        function zapiszSzkic() {
            if (!['1', '2', '3', '4'].includes(aktualnaLinia)) return;
            const draft = {
                kod: document.getElementById('kodSelect')?.value || "",
                ilosc: document.querySelector('input[name="ilosc"]')?.value || "",
                data_przyd: document.querySelector('input[name="data_przyd"]')?.value || "",
                etykieta_gora: document.querySelector('input[name="etykieta_gora"]')?.value || "",
                etykieta_dol: document.querySelector('input[name="etykieta_dol"]')?.value || "",
                czas_start: document.getElementById('czasStart')?.value || "",
                czas_koniec: document.getElementById('czasKoniec')?.value || ""
            };
            localStorage.setItem(draftKey, JSON.stringify(draft));
        }

        function wczytajSzkic() {
            if (!['1', '2', '3', '4'].includes(aktualnaLinia)) return;
            const raw = localStorage.getItem(draftKey);
            if (!raw) return;
            try {
                const draft = JSON.parse(raw);
                if (draft.kod && document.getElementById('kodSelect')) document.getElementById('kodSelect').value = draft.kod;
                if (draft.ilosc && document.querySelector('input[name="ilosc"]')) document.querySelector('input[name="ilosc"]').value = draft.ilosc;
                if (draft.data_przyd && document.querySelector('input[name="data_przyd"]')) document.querySelector('input[name="data_przyd"]').value = draft.data_przyd;
                if (draft.etykieta_gora && document.querySelector('input[name="etykieta_gora"]')) document.querySelector('input[name="etykieta_gora"]').value = draft.etykieta_gora;
                if (draft.etykieta_dol && document.querySelector('input[name="etykieta_dol"]')) document.querySelector('input[name="etykieta_dol"]').value = draft.etykieta_dol;
                if (draft.czas_start && document.getElementById('czasStart')) document.getElementById('czasStart').value = draft.czas_start;
                if (draft.czas_koniec && document.getElementById('czasKoniec')) document.getElementById('czasKoniec').value = draft.czas_koniec;
                aktualizujDane();
            } catch (e) {}
        }

        function wyczyscSzkicPoDodaniu() {
            localStorage.removeItem(draftKey);
        }

        function aktualizujDane() {
            const kod = document.getElementById('kodSelect')?.value;
            const nazwaInput = document.getElementById('nazwaInput');
            const odbiorcaInput = document.getElementById('odbiorcaInput');
            const infoBox = document.getElementById('planInfoBox');
            if (!nazwaInput || !odbiorcaInput || !infoBox) return;

            const dataProd = document.getElementById('dataProdInput')?.value || "{{ dzis }}";
            if (kod && baza[kod]) {
                nazwaInput.value = baza[kod].nazwa;
                odbiorcaInput.value = baza[kod].odbiorca;
                const planNaDzien = calyPlan[dataProd] || {};
                if (planNaDzien[kod] !== undefined) {
                    const planKwota = planNaDzien[kod];
                    const wykonaneKwota = (zrealizowanoDzien[dataProd] && zrealizowanoDzien[dataProd][kod]) ? zrealizowanoDzien[dataProd][kod] : 0;
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
            zapiszSzkic();
        }

        function filtrKolumnTabeli() {
            const fKod = (document.getElementById('filterKod')?.value || '').toLowerCase();
            const fNazwa = (document.getElementById('filterNazwa')?.value || '').toLowerCase();
            const fOdbiorca = (document.getElementById('filterOdbiorca')?.value || '').toLowerCase();
            const fPlan = (document.getElementById('filterPlan')?.value || '').toLowerCase();
            const fWykonano = (document.getElementById('filterWykonano')?.value || '').toLowerCase();
            const fPozostalo = (document.getElementById('filterPozostalo')?.value || '').toLowerCase();
            const fNadwyzka = (document.getElementById('filterNadwyzka')?.value || '').toLowerCase();
            const fProcent = (document.getElementById('filterProcent')?.value || '').toLowerCase();

            const rows = document.getElementById('progressTableBody')?.getElementsByTagName('tr') || [];
            for (let i = 0; i < rows.length; i++) {
                const match =
                    (rows[i].getAttribute('data-kod') || '').toLowerCase().includes(fKod) &&
                    (rows[i].getAttribute('data-nazwa') || '').toLowerCase().includes(fNazwa) &&
                    (rows[i].getAttribute('data-odbiorca') || '').toLowerCase().includes(fOdbiorca) &&
                    (rows[i].getAttribute('data-plan') || '').toLowerCase().includes(fPlan) &&
                    (rows[i].getAttribute('data-wykonano') || '').toLowerCase().includes(fWykonano) &&
                    (rows[i].getAttribute('data-pozostalo') || '').toLowerCase().includes(fPozostalo) &&
                    (rows[i].getAttribute('data-nadwyzka') || '').toLowerCase().includes(fNadwyzka) &&
                    (rows[i].getAttribute('data-procent') || '').toLowerCase().includes(fProcent);
                rows[i].style.display = match ? "" : "none";
            }
        }

        function filtrWydajnosci() {
            const fLinia = (document.getElementById('filterWydLinia')?.value || '').toLowerCase();
            const fZmiana = (document.getElementById('filterWydZmiana')?.value || '').toLowerCase();
            const fSztuki = (document.getElementById('filterWydSztuki')?.value || '').toLowerCase();
            const fCzas = (document.getElementById('filterWydCzas')?.value || '').toLowerCase();
            const fWyd = (document.getElementById('filterWydSztH')?.value || '').toLowerCase();

            const tbody = document.getElementById('wydajnoscTableBody');
            if (!tbody) return;

            const rows = tbody.getElementsByTagName('tr');
            let sumaSztuk = 0;
            let sumaCzas = 0;

            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                if (!row.hasAttribute('data-sztuki')) continue;

                const dLinia = (row.getAttribute('data-linia') || '').toLowerCase();
                const dZmiana = (row.getAttribute('data-zmiana') || '').toLowerCase();
                const dSztukiRaw = parseFloat(row.getAttribute('data-sztuki') || '0') || 0;
                const dCzasRaw = parseFloat(row.getAttribute('data-czas') || '0') || 0;
                const dWydRaw = (row.getAttribute('data-wyd') || '').toLowerCase();

                const match =
                    dLinia.includes(fLinia) &&
                    dZmiana.includes(fZmiana) &&
                    String(dSztukiRaw).includes(fSztuki) &&
                    String(dCzasRaw).includes(fCzas) &&
                    dWydRaw.includes(fWyd);

                if (match) {
                    row.style.display = "";
                    sumaSztuk += dSztukiRaw;
                    sumaCzas += dCzasRaw;
                } else {
                    row.style.display = "none";
                }
            }

            if (document.getElementById('wydSumaSztuk')) document.getElementById('wydSumaSztuk').innerText = sumaSztuk + " szt.";
            if (document.getElementById('wydSumaCzas')) document.getElementById('wydSumaCzas').innerText = sumaCzas.toFixed(2) + " h";
            if (document.getElementById('wydSrednia')) {
                const avg = sumaCzas > 0 ? (sumaSztuk / sumaCzas).toFixed(1) : "0.0";
                document.getElementById('wydSrednia').innerText = avg + " szt/h";
            }
        }

        let sortDirections = {};
        function sortTable(columnIndex, isNumeric = false) {
            const tbody = document.getElementById("progressTableBody");
            if (!tbody) return;
            const rowsArray = Array.from(tbody.querySelectorAll("tr"));
            let dir = sortDirections[columnIndex] === 'asc' ? 'desc' : 'asc';
            sortDirections[columnIndex] = dir;
            rowsArray.sort((rowA, rowB) => {
                let cellA = rowA.children[columnIndex].innerText.trim();
                let cellB = rowB.children[columnIndex].innerText.trim();
                if (isNumeric) {
                    cellA = parseFloat(cellA.replace(/[^0-9.-]+/g,"")) || 0;
                    cellB = parseFloat(cellB.replace(/[^0-9.-]+/g,"")) || 0;
                    return dir === 'asc' ? cellA - cellB : cellB - cellA;
                }
                return dir === 'asc' ? cellA.localeCompare(cellB) : cellB.localeCompare(cellA);
            });
            rowsArray.forEach(row => tbody.appendChild(row));
        }

        let sortWydDirections = {};
        function sortWydTable(columnIndex, isNumeric = false) {
            const tbody = document.getElementById("wydajnoscTableBody");
            if (!tbody) return;
            const rowsArray = Array.from(tbody.querySelectorAll("tr")).filter(r => r.hasAttribute('data-sztuki'));
            let dir = sortWydDirections[columnIndex] === 'asc' ? 'desc' : 'asc';
            sortWydDirections[columnIndex] = dir;
            rowsArray.sort((rowA, rowB) => {
                let cellA = rowA.children[columnIndex].innerText.trim();
                let cellB = rowB.children[columnIndex].innerText.trim();
                if (isNumeric) {
                    cellA = parseFloat(cellA.replace(/[^0-9.-]+/g,"")) || 0;
                    cellB = parseFloat(cellB.replace(/[^0-9.-]+/g,"")) || 0;
                    return dir === 'asc' ? cellA - cellB : cellB - cellA;
                }
                return dir === 'asc' ? cellA.localeCompare(cellB) : cellB.localeCompare(cellA);
            });
            rowsArray.forEach(row => tbody.appendChild(row));
        }

        window.onload = function() {
            wczytajSzkic();
            filtrWydajnosci();
            const dataInput = document.getElementById('dataProdInput');
            if (dataInput) dataInput.addEventListener('change', aktualizujDane);
            ['ilosc', 'data_przyd', 'etykieta_gora', 'etykieta_dol'].forEach(name => {
                const el = document.querySelector('input[name="' + name + '"]');
                if (el) {
                    el.addEventListener('input', zapiszSzkic);
                    el.addEventListener('change', zapiszSzkic);
                }
            });
        };

        function walidujFormularz(event) {
            const iloscVal = parseInt(document.querySelector('input[name="ilosc"]')?.value || "0", 10);
            if (!iloscVal || iloscVal < 1) {
                alert("Błąd: Podaj ilość sztuk na koniec palety (minimum 1)!");
                event.preventDefault();
                return false;
            }
            const start = document.getElementById('czasStart').value;
            const koniec = document.getElementById('czasKoniec').value;
            if (start && koniec && koniec <= start) {
                alert("Błąd: Czas zakończenia palety nie może być wcześniejszy lub równy czasowi rozpoczęcia!");
                event.preventDefault();
                return false;
            }
            wyczyscSzkicPoDodaniu();
            return true;
        }
    </script>
</head>
<body>
<div class="container">
    <div class="header-bar">
        <div>
            <b>{{ t.mod_koncowki_nazwa.split('(')[0] }}</b> |
            <b>{{ t.zalog_op }}:</b> {{ pelna_nazwa }} ({{ role }}) |
            {% if wypr_zmiana %}
                <b>{{ t.zmiana }}:</b> {{ wypr_zmiana }} |
                <a href="/reset-zmiana" style="color: #dc3545; font-weight: bold; text-decoration: none;">🔄 {{ t.zmien_zm }}</a>
            {% endif %}
        </div>
        <div style="display: flex; gap: 8px; align-items: center; flex-wrap: wrap;">
            <div class="lang-box">
                <a href="/set-lang/pl" class="lang-btn {{ 'active' if lang == 'pl' else '' }}">PL</a>
                <a href="/set-lang/en" class="lang-btn {{ 'active' if lang == 'en' else '' }}">EN</a>
                <a href="/set-lang/uk" class="lang-btn {{ 'active' if lang == 'uk' else '' }}">UK</a>
                <a href="/set-lang/es" class="lang-btn {{ 'active' if lang == 'es' else '' }}">ES</a>
            </div>
            {% if role == 'admin' or can_access_produkcja %}
                <a href="/przelacz-modul/produkcja" style="background: #0284c7; color: white; padding: 6px 12px; border-radius: 4px; text-decoration: none; font-size: 13px; font-weight: bold;">🔀 {{ t.mod_produkcja_nazwa.split('(')[0] }}</a>
            {% endif %}
            <a href="/logout" style="background: #dc3545; color: white; padding: 6px 12px; border-radius: 4px; text-decoration: none; font-size: 13px; font-weight: bold;">{{ t.wyloguj }}</a>
        </div>
    </div>

    {% if not wypr_zmiana %}
        <div style="text-align: center; padding: 40px;">
            <h2>{{ t.wybierz_zm }}</h2>
            <form method="POST" action="/wybierz-zmiane">
                <select name="zmiana" style="max-width: 300px; margin: 0 auto 15px auto;">
                    <option value="1">Zmiana 1 (04:00 / 06:00 – 14:00)</option>
                    <option value="2">Zmiana 2 (14:00 – 22:00 / max 02:00)</option>
                </select>
                <button type="submit" style="max-width: 300px; margin: 0 auto;">{{ t.rozpocznij }}</button>
            </form>
        </div>
    {% else %}
        <h2>🏭 {{ t.tytul }}</h2>

        <div class="tabs">
            {% if '1' in dostep %}<a href="/?linia=1" class="{{ 'active' if linia == '1' else '' }}">{{ t.linia_1 }}</a>{% endif %}
            {% if '2' in dostep %}<a href="/?linia=2" class="{{ 'active' if linia == '2' else '' }}">{{ t.linia_2 }}</a>{% endif %}
            {% if '3' in dostep %}<a href="/?linia=3" class="{{ 'active' if linia == '3' else '' }}">{{ t.linia_3 }}</a>{% endif %}
            {% if '4' in dostep %}<a href="/?linia=4" class="{{ 'active' if linia == '4' else '' }}">{{ t.linia_4 }}</a>{% endif %}
            {% if 'plan' in dostep %}<a href="/?linia=plan" class="plan-tab {{ 'active' if linia == 'plan' else '' }}">{{ t.plan_tab }}</a>{% endif %}
            {% if 'progress' in dostep %}<a href="/?linia=progress" class="progress-tab {{ 'active' if linia == 'progress' else '' }}">{{ t.prog_tab }}</a>{% endif %}
            {% if 'wydajnosc' in dostep %}<a href="/?linia=wydajnosc" class="wydajnosc-tab {{ 'active' if linia == 'wydajnosc' else '' }}">{{ t.wydajnosc_tab }}</a>{% endif %}
            {% if 'raporty' in dostep %}<a href="/?linia=raporty" class="raporty-tab {{ 'active' if linia == 'raporty' else '' }}">📈 {{ t.raporty_tab }}</a>{% endif %}
            {% if 'db_art' in dostep %}<a href="/?linia=db_art" class="db-tab {{ 'active' if linia == 'db_art' else '' }}">{{ t.db_art_tab }}</a>{% endif %}
            {% if 'admin' in dostep %}<a href="/?linia=admin" class="admin-tab {{ 'active' if linia == 'admin' else '' }}">{{ t.admin_tab }}</a>{% endif %}
        </div>

        {% if error_msg %}<div class="error-banner">⚠️ {{ error_msg }}</div>{% endif %}

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
                        <label><b>{{ t.wybierz_dzien_plan }}</b></label>
                        <select name="wybrana_data_planu" onchange="this.form.submit()">
                            {% for data_p in dostepne_daty_planu %}
                                <option value="{{ data_p }}" {{ 'selected' if data_p == wybrana_data_planu else '' }}>📅 {{ data_p }}</option>
                            {% endfor %}
                        </select>
                    </div>
                </form>

                {% if wybrana_data_planu and wybrana_data_planu in plan_produkcyjny %}
                    <div style="background: white; padding: 15px; border-radius: 6px; margin-bottom: 10px; border: 1px solid #c8e1b9;">
                        <b>📅 {{ wybrana_data_planu }}</b>
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
                <div style="display: flex; justify-content: space-between; align-items: flex-end; gap: 15px; margin-bottom: 20px; flex-wrap: wrap;">
                    <form method="GET" action="/" style="flex: 1; max-width: 300px; margin: 0;">
                        <input type="hidden" name="linia" value="progress">
                        <label><b>{{ t.wybierz_dzien_prog }}</b></label>
                        <select name="wybrana_data_progress" onchange="this.form.submit()">
                            {% for data_d in dostepne_daty_wpisow_lub_planu %}
                                <option value="{{ data_d }}" {{ 'selected' if data_d == wybrana_data_progress else '' }}>📅 {{ data_d }}</option>
                            {% endfor %}
                        </select>
                    </form>
                    <div>
                        <button type="button" onclick="window.print()" style="background: #007bff; width: auto; padding: 10px 18px; margin: 0; font-size: 13px;">🖨️ PDF / Print</button>
                    </div>
                </div>

                <div style="background: white; padding: 20px; border-radius: 8px; margin-bottom: 25px; border: 2px solid #28a745; box-shadow: 0 2px 6px rgba(0,0,0,0.05);">
                    <div style="font-size: 16px; font-weight: bold; margin-bottom: 10px; color: #155724; display: flex; justify-content: space-between; flex-wrap: wrap; gap: 8px;">
                        <span>🚀 {{ t.podsumowanie_dnia }} ({{ wybrana_data_progress }})</span>
                        <span>{{ t.realizacja_podstawowa }} <b>{{ procent_podstawowy_prog }}%</b> | {{ t.realizacja_z_nadwyzka }} <b>{{ procent_calkowity_prog }}%</b></span>
                    </div>

                    <div style="margin-bottom: 6px; font-size: 12px; color: #004085; font-weight: bold;">
                        {{ t.realizacja_podstawowa }} {{ suma_wkladu_plan_prog }} / {{ suma_planu_calkowita_prog }} szt. ({{ procent_podstawowy_prog }}%)
                    </div>
                    <div style="background: #e9ecef; border-radius: 8px; height: 22px; overflow: hidden; border: 1px solid #ced4da; margin-bottom: 12px;">
                        <div style="background: {{ '#28a745' if procent_podstawowy_prog >= 100 else '#007bff' }}; width: {{ [procent_podstawowy_prog, 100] | min }}%; height: 100%; display: flex; align-items: center; justify-content: center; color: white; font-weight: bold; font-size: 12px;">
                            {{ procent_podstawowy_prog }}%
                        </div>
                    </div>

                    <div style="margin-bottom: 6px; font-size: 12px; color: #155724; font-weight: bold;">
                        {{ t.realizacja_z_nadwyzka }} {{ suma_wykonana_calkowita_prog }} / {{ suma_planu_calkowita_prog }} szt. ({{ procent_calkowity_prog }}%)
                    </div>
                    <div style="background: #e9ecef; border-radius: 8px; height: 22px; overflow: hidden; border: 1px solid #ced4da; margin-bottom: 12px;">
                        <div style="background: {{ '#16a34a' if procent_calkowity_prog >= 100 else '#17a2b8' }}; width: {{ [procent_calkowity_prog, 100] | min }}%; height: 100%; display: flex; align-items: center; justify-content: center; color: white; font-weight: bold; font-size: 12px;">
                            {{ procent_calkowity_prog }}%
                        </div>
                    </div>

                    <div style="font-size: 14px; color: #333; display: flex; justify-content: space-between; flex-wrap: wrap; gap: 10px;">
                        <span>📋 {{ t.do_zrobienia }} <b>{{ suma_planu_calkowita_prog }} szt.</b></span>
                        <span>✅ {{ t.wykonano }} <b>{{ suma_wykonana_calkowita_prog }} szt.</b></span>
                        <span>⚠️ {{ t.pozostalo }} <b style="color: #dc3545;">{{ pozostalo_calkowite_prog }} szt.</b></span>
                        <span>➕ {{ t.nadwyzka }}: <b style="color: #16a34a;">{{ nadwyzka_calkowita_prog }} szt.</b></span>
                    </div>
                </div>

                <h4 style="margin-top: 20px; color: #333;">{{ t.postep_art }}</h4>
                <div style="max-height: 500px; overflow-y: auto; background: white; border-radius: 6px; border: 1px solid #cbd5e1;">
                    <table class="progress-table" id="progressTable">
                        <thead>
                            <tr>
                                <th onclick="sortTable(0, false)">{{ t.kol_kod }} ↕️<input type="text" id="filterKod" class="filter-input" onkeyup="filtrKolumnTabeli()" onclick="event.stopPropagation()" placeholder="..."></th>
                                <th onclick="sortTable(1, false)">{{ t.kol_nazwa }} ↕️<input type="text" id="filterNazwa" class="filter-input" onkeyup="filtrKolumnTabeli()" onclick="event.stopPropagation()" placeholder="..."></th>
                                <th onclick="sortTable(2, false)">{{ t.kol_odbiorca }} ↕️<input type="text" id="filterOdbiorca" class="filter-input" onkeyup="filtrKolumnTabeli()" onclick="event.stopPropagation()" placeholder="..."></th>
                                <th style="width: 90px;" onclick="sortTable(3, true)">Plan ↕️<input type="text" id="filterPlan" class="filter-input" onkeyup="filtrKolumnTabeli()" onclick="event.stopPropagation()" placeholder="..."></th>
                                <th style="width: 90px;" onclick="sortTable(4, true)">{{ t.tabela_ilosc }} ↕️<input type="text" id="filterWykonano" class="filter-input" onkeyup="filtrKolumnTabeli()" onclick="event.stopPropagation()" placeholder="..."></th>
                                <th style="width: 90px;" onclick="sortTable(5, true)">{{ t.pozostalo }} ↕️<input type="text" id="filterPozostalo" class="filter-input" onkeyup="filtrKolumnTabeli()" onclick="event.stopPropagation()" placeholder="..."></th>
                                <th style="width: 90px;" onclick="sortTable(6, true)">{{ t.nadwyzka }} ↕️<input type="text" id="filterNadwyzka" class="filter-input" onkeyup="filtrKolumnTabeli()" onclick="event.stopPropagation()" placeholder="..."></th>
                                <th style="width: 120px;" onclick="sortTable(7, true)">% ↕️<input type="text" id="filterProcent" class="filter-input" onkeyup="filtrKolumnTabeli()" onclick="event.stopPropagation()" placeholder="..."></th>
                            </tr>
                        </thead>
                        <tbody id="progressTableBody">
                            {% for kod, plan_val in plan_wybrany_prog.items() %}
                                {% set wykonane = wykonane_dla_wybranej_daty_prog.get(kod, 0) %}
                                {% set left_amt = [plan_val - wykonane, 0] | max %}
                                {% set nadwyzka_amt = [wykonane - plan_val, 0] | max %}
                                {% set procent = (wykonane / plan_val * 100) | round | int if plan_val > 0 else 0 %}
                                {% set prod_info = produkty_baza.get(kod, {}) %}
                                <tr data-kod="{{ kod }}" data-nazwa="{{ prod_info.get('nazwa', 'Nieznany') }}" data-odbiorca="{{ prod_info.get('odbiorca', '-') }}" data-plan="{{ plan_val }}" data-wykonano="{{ wykonane }}" data-pozostalo="{{ left_amt }}" data-nadwyzka="{{ nadwyzka_amt }}" data-procent="{{ procent }}%">
                                    <td><b>{{ kod }}</b></td>
                                    <td>{{ prod_info.get('nazwa', 'Nieznany') }}</td>
                                    <td>{{ prod_info.get('odbiorca', '-') }}</td>
                                    <td>{{ plan_val }} szt.</td>
                                    <td><b>{{ wykonane }}</b> szt.</td>
                                    <td><b style="color: #dc3545;">{{ left_amt }}</b> szt.</td>
                                    <td><b style="color: #16a34a;">{{ nadwyzka_amt }}</b> szt.</td>
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
                                <tr><td colspan="8" style="text-align: center; color: #777; padding: 20px;">{{ t.brak_planu_dzis }}</td></tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        {% elif linia == 'wydajnosc' and 'wydajnosc' in dostep %}
            <div class="wydajnosc-section">
                <h3 style="margin-top: 0; color: #5a32a3;">⏱️ {{ t.raport_wyd_tytul }}</h3>
                <p style="font-size: 13px; color: #555;">{{ t.raport_wyd_opis }}</p>

                <form method="GET" action="/" style="margin-bottom: 20px; display: flex; gap: 10px; align-items: flex-end;">
                    <input type="hidden" name="linia" value="wydajnosc">
                    <div style="flex: 1; max-width: 300px;">
                        <label><b>{{ t.data_raportu_hist }}</b></label>
                        <select name="wybrana_data_wydajnosc" onchange="this.form.submit()">
                            {% for data_d in dostepne_daty_wpisow %}
                                <option value="{{ data_d }}" {{ 'selected' if data_d == wybrana_data_wydajnosc else '' }}>📅 {{ data_d }}</option>
                            {% endfor %}
                        </select>
                    </div>
                </form>

                <div class="wyd-layout">
                    <div class="wyd-table-col">
                        <table class="progress-table" id="wydajnoscTable" style="margin-top: 0;">
                            <thead>
                                <tr>
                                    <th onclick="sortWydTable(0, false)">{{ t.th_linia }} ↕️<input type="text" id="filterWydLinia" class="filter-input" onkeyup="filtrWydajnosci()" onclick="event.stopPropagation()" placeholder="..."></th>
                                    <th onclick="sortWydTable(1, false)">{{ t.th_zmiana }} ↕️<input type="text" id="filterWydZmiana" class="filter-input" onkeyup="filtrWydajnosci()" onclick="event.stopPropagation()" placeholder="..."></th>
                                    <th onclick="sortWydTable(2, true)">{{ t.th_wyprodukowano }} ↕️<input type="text" id="filterWydSztuki" class="filter-input" onkeyup="filtrWydajnosci()" onclick="event.stopPropagation()" placeholder="..."></th>
                                    <th onclick="sortWydTable(3, true)">{{ t.th_czas_pracy }} ↕️<input type="text" id="filterWydCzas" class="filter-input" onkeyup="filtrWydajnosci()" onclick="event.stopPropagation()" placeholder="..."></th>
                                    <th onclick="sortWydTable(4, true)">{{ t.th_wydajnosc }} ↕️<input type="text" id="filterWydSztH" class="filter-input" onkeyup="filtrWydajnosci()" onclick="event.stopPropagation()" placeholder="..."></th>
                                </tr>
                            </thead>
                            <tbody id="wydajnoscTableBody">
                                {% for stat in statystyki_wydajnosci_lista %}
                                <tr data-linia="Linia {{ stat.linia }}" data-zmiana="Zmiana {{ stat.zmiana }}" data-sztuki="{{ stat.suma_sztuk }}" data-czas="{{ stat.czas_godz }}" data-wyd="{{ stat.sztuk_na_godzine }}">
                                    <td><b>Linia {{ stat.linia }}</b></td>
                                    <td>Zmiana {{ stat.zmiana }}</td>
                                    <td><b>{{ stat.suma_sztuk }} szt.</b></td>
                                    <td>{{ stat.czas_godz }} h</td>
                                    <td><b style="color: #6f42c1; font-size: 15px;">{{ stat.sztuk_na_godzine }} szt/h</b></td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>

                    <div class="wyd-side-summary">
                        <h4 style="margin-top: 0; color: #5a32a3; border-bottom: 1px solid #e2e8f0; padding-bottom: 8px;">📌 {{ t.podsumowanie_dnia.split()[0] }}</h4>
                        <div class="wyd-stat-row">
                            <div style="font-size: 12px; color: #555;">{{ t.laczna_suma_wypr }}</div>
                            <div id="wydSumaSztuk" style="font-size: 20px; font-weight: bold; color: #28a745; margin-top: 2px;">0 szt.</div>
                        </div>
                        <div class="wyd-stat-row">
                            <div style="font-size: 12px; color: #555;">{{ t.laczny_czas }}</div>
                            <div id="wydSumaCzas" style="font-size: 15px; font-weight: bold; color: #333; margin-top: 2px;">0.00 h</div>
                        </div>
                        <div class="wyd-stat-row">
                            <div style="font-size: 12px; color: #555;">{{ t.srednia_wyd }}</div>
                            <div id="wydSrednia" style="font-size: 16px; font-weight: bold; color: #6f42c1; margin-top: 2px;">0.0 szt/h</div>
                        </div>
                    </div>
                </div>
            </div>
        {% elif linia == 'raporty' and 'raporty' in dostep %}
            <div class="raporty-section">
                <h3 style="margin-top: 0; color: #d9480f;">📈 {{ t.raporty_tab }}</h3>
                <p style="font-size: 13px; color: #555;">{{ t.data_raportu_hist }}</p>
                <div style="background: white; padding: 22px; border-radius: 8px; border: 1px solid #ffd8a8; max-width: 780px;">
                    <form action="/generate" method="GET">
                        <label><b>{{ t.data_raportu_hist }}</b></label>
                        <input type="date" name="data" value="{{ wybrana_data }}" style="margin-bottom: 15px; max-width: 280px;" required>

                        <label><b>{{ t.raporty_tab }}:</b></label>
                        <select name="typ_raportu" style="margin-bottom: 20px;">
                            <optgroup label="Końcówki Linii / Rozchód i Plan">
                                <option value="rozchod_zmiana_1">📦 Raport Rozchodu / Produkcji — ZMIANA 1</option>
                                <option value="rozchod_zmiana_2">📦 Raport Rozchodu / Produkcji — ZMIANA 2</option>
                                <option value="rozchod_osobne_arkusze">📦 Raport Rozchodu — ZMIANA 1 + ZMIANA 2 (Z1 i Z2)</option>
                                <option value="postep_realizacja">📊 Raport Postępu i Realizacji Planu</option>
                                <option value="wydajnosc">⏱️ Raport Wydajności Linii i Zmian (szt/h)</option>
                            </optgroup>
                            <optgroup label="Moduł Produkcyjny (Linie 1–6, Pracownicy, Roboczogodziny)">
                                <option value="wydajnosc_rbh">⚡ Raport Wydajności na Roboczogodzinę (Linia 1–6, Z1–Z2, szt./rbh)</option>
                                <option value="monitoring_linie_pracownicy">👥 Raport Monitoringu Linii 1–6, Postojów i Czasu Pracowników</option>
                            </optgroup>
                        </select>

                        <button type="submit" style="background: #fd7e14; color: white;">📥 XLSX</button>
                    </form>
                </div>
            </div>
        {% elif linia == 'db_art' and 'db_art' in dostep %}
            <div class="db-section">
                <h3 style="margin-top: 0; color: #856404;">📦 {{ t.zarzadzanie_art }}</h3>
                <form action="/dodaj-artykul" method="POST" style="margin-bottom: 25px; background: white; padding: 15px; border-radius: 6px;">
                    <div class="row">
                        <div><label>{{ t.kol_kod }}:</label><input type="text" name="kod" placeholder="{{ t.ph_kod }}" required></div>
                        <div><label>{{ t.kol_nazwa }}:</label><input type="text" name="nazwa" placeholder="{{ t.ph_nazwa }}" required></div>
                        <div><label>{{ t.kol_odbiorca }}:</label><input type="text" name="odbiorca" placeholder="{{ t.ph_odbiorca }}" required></div>
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
                        <div><label>{{ t.login }}</label><input type="text" name="username" placeholder="np. anna" required></div>
                        <div><label>{{ t.haslo }}</label><input type="password" name="password" placeholder="hasło" required></div>
                        <div><label>{{ t.imie_nazwisko }}</label><input type="text" name="pelna_nazwa" placeholder="np. Anna Nowak" required></div>
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
                        <label style="display:flex; align-items:center; gap:5px; font-weight:bold;"><input type="checkbox" name="dostep" value="mod_koncowki" checked style="width:auto; margin:0;"> Moduł Końcówki Linii</label>
                        <label style="display:flex; align-items:center; gap:5px; font-weight:bold;"><input type="checkbox" name="dostep" value="mod_produkcja" checked style="width:auto; margin:0;"> Moduł Produkcyjny</label>
                        <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="1" checked style="width:auto; margin:0;"> {{ t.linia_1 }}</label>
                        <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="2" checked style="width:auto; margin:0;"> {{ t.linia_2 }}</label>
                        <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="3" checked style="width:auto; margin:0;"> {{ t.linia_3 }}</label>
                        <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="4" checked style="width:auto; margin:0;"> {{ t.linia_4 }}</label>
                        <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="plan" style="width:auto; margin:0;"> {{ t.plan_tab }}</label>
                        <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="progress" checked style="width:auto; margin:0;"> {{ t.prog_tab }}</label>
                        <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="wydajnosc" checked style="width:auto; margin:0;"> {{ t.wydajnosc_tab }}</label>
                        <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="raporty" checked style="width:auto; margin:0;"> {{ t.raporty_tab }}</label>
                        <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="db_art" style="width:auto; margin:0;"> {{ t.db_art_tab }}</label>
                        <label style="display:flex; align-items:center; gap:5px;"><input type="checkbox" name="dostep" value="admin" style="width:auto; margin:0;"> {{ t.admin_tab }}</label>
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
                            <input type="date" id="dataProdInput" name="data" value="{{ dzis }}" readonly style="background: #e9ecef;">
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
                        <input type="number" name="ilosc" min="1" placeholder="{{ t.ph_ilosc }}">
                    </div>
                </div>

                <div id="planInfoBox" class="plan-info-box"></div>

                <div class="row">
                    <div style="flex: 2;"><label>{{ t.nazwa_ind }}</label><input type="text" id="nazwaInput" name="nazwa" required readonly style="background: #e9ecef;"></div>
                    <div style="flex: 2;"><label>{{ t.odbiorca }}</label><input type="text" id="odbiorcaInput" name="odbiorca" required readonly style="background: #e9ecef;"></div>
                    <div style="flex: 1;"><label>{{ t.data_przyd }}</label><input type="date" name="data_przyd" min="{{ min_data }}" required></div>
                </div>

                <div class="row">
                    <div><label>{{ t.nr_partii }}</label><input type="text" name="partia_prod" value="{{ domyslna_partia }}" required></div>
                    <div><label>{{ t.ety_gora }}</label><input type="text" name="etykieta_gora" placeholder="{{ t.ph_etygora }}" required></div>
                    <div><label>{{ t.ety_dol }}</label><input type="text" name="etykieta_dol" placeholder="{{ t.ph_etydol }}" required></div>
                    <div><label>{{ t.operator }}</label><input type="text" name="operator" value="{{ pelna_nazwa }}" readonly style="background: #e9ecef; font-weight: bold;"></div>
                </div>

                <div style="background: #e8f4fd; padding: 15px; border-radius: 6px; margin-top: 15px; border: 1px solid #b8daff;">
                    <label style="margin-top:0; color: #004085; font-size: 14px;">🛡️ {{ t.bhp_tytul }}</label>
                    <div style="margin-top: 10px; display: flex; flex-direction: column; gap: 8px; font-size: 13px;">
                        <label style="font-weight: normal; display: flex; align-items: center; gap: 8px; cursor: pointer;"><input type="checkbox" name="potwierdzenie_kontroli" required style="width: auto; margin:0;"><span>{{ t.bhp_1 }}</span></label>
                        <label style="font-weight: normal; display: flex; align-items: center; gap: 8px; cursor: pointer;"><input type="checkbox" name="potwierdzenie_oczyszczenia" required style="width: auto; margin:0;"><span>{{ t.bhp_2 }}</span></label>
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
                    <div><label>{{ t.czas_start_lbl }}</label><input type="time" id="czasStart" name="czas_start" onchange="zapiszSzkic()" required></div>
                    <div><label>{{ t.czas_koniec_lbl }}</label><input type="time" id="czasKoniec" name="czas_koniec" onchange="zapiszSzkic()" required></div>
                </div>

                <script>
                    function ustawCzas(typ) {
                        const now = new Date();
                        const czasStr = String(now.getHours()).padStart(2, '0') + ':' + String(now.getMinutes()).padStart(2, '0');
                        if (typ === 'start') document.getElementById('czasStart').value = czasStr;
                        else document.getElementById('czasKoniec').value = czasStr;
                        zapiszSzkic();
                    }
                </script>

                <button type="submit">{{ t.dodaj_btn }}</button>
            </form>

            <hr style="margin: 25px 0;">

            {% if role == 'admin' %}
                <h3>📊 {{ t.wybierz_dzien_hist }}</h3>
                <form method="GET" action="/" style="display: flex; gap: 10px; align-items: flex-end; margin-bottom: 15px;">
                    <input type="hidden" name="linia" value="{{ linia }}">
                    <div style="flex: 1;"><label>{{ t.data_raportu_hist }}</label><input type="date" name="wybrana_data" value="{{ wybrana_data }}" onchange="this.form.submit()"></div>
                </form>
            {% endif %}

            <h3>📊 {{ t.realizacja_biez }} {{ t.linia_1.split()[0] }} {{ linia }} ({{ wybrana_data }})</h3>
            <table>
                <tr>
                    <th>{{ t.tabela_nr }}</th><th>{{ t.kol_kod }}</th><th>{{ t.kol_nazwa }}</th><th>{{ t.tabela_ilosc }}</th>
                    <th>{{ t.tabela_czas }}</th><th>{{ t.tabela_bhp }}</th><th>{{ t.tabela_szczegoly }}</th><th>{{ t.kol_akcja }}</th>
                </tr>
                {% for wpis in wpisy_dla_linii %}
                <tr>
                    <td>{{ loop.index }}</td><td>{{ wpis.kod }}</td><td>{{ wpis.nazwa }}</td><td><b>{{ wpis.ilosc }}</b> szt.</td>
                    <td>{{ wpis.czas_start }} - {{ wpis.czas_koniec }}</td><td>✅ BHP</td>
                    <td>{{ wpis.data_przyd }} | {{ wpis.partia_prod }} | {{ wpis.operator }}</td>
                    <td><a href="/usun/{{ wpis.id }}?data={{ wybrana_data }}" style="color: red; text-decoration: none;">{{ t.usun }}</a></td>
                </tr>
                {% else %}
                <tr><td colspan="8" style="text-align: center; color: #777;">{{ t.brak_wpisow }}</td></tr>
                {% endfor %}
            </table>

            <div class="summary-box">{{ t.lacznie }} {{ suma_sztuk }} szt.</div>
        {% else %}
            <div style="text-align: center; padding: 40px; color: #dc3545;"><h3>⚠️ {{ t.brak_dostepu }}</h3></div>
        {% endif %}
    {% endif %}
</div>
</body>
</html>
"""

# --- ZMIANA JĘZYKA, LOGOWANIE I PRZEŁĄCZANIE MODUŁÓW ---

@app.route('/set-lang/<lang_code>')
def set_lang(lang_code):
    if lang_code in T:
        session['lang'] = lang_code
    referer = request.headers.get('Referer')
    if referer:
        return redirect(referer)
    if session.get('modul') == 'produkcja':
        return redirect(url_for('produkcja_view'))
    return redirect(url_for('index'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    t, lang = get_lang_dict()
    wybrany_modul = request.form.get('modul', session.get('modul', 'koncowki'))
    if request.method == 'POST':
        username = (request.form.get('username') or '').strip()
        password = request.form.get('password') or ''
        modul = request.form.get('modul', 'koncowki')

        user_db = UzytkownikModel.query.filter_by(username=username, haslo=password).first()
        if not user_db:
            return render_template_string(LOGIN_TEMPLATE, t=t, lang=lang, wybrany_modul=modul, error="Nieprawidłowy login lub hasło!")

        dostep_lista = json.loads(user_db.dostep)
        if not czy_ma_dostep_do_modulu(dostep_lista, user_db.rola, modul):
            nazwa_m = "Moduł Produkcyjny" if modul == "produkcja" else "Końcówki Linii"
            return render_template_string(LOGIN_TEMPLATE, t=t, lang=lang, wybrany_modul=modul, error=f"Twoje konto nie ma uprawnień do wybranego modułu ({nazwa_m})!")

        zachowany_lang = session.get('lang', lang)
        session.clear()
        session['lang'] = zachowany_lang
        session['user'] = user_db.username
        session['rola'] = user_db.rola
        session['pelna_nazwa'] = user_db.pelna_nazwa
        session['dostep'] = dostep_lista
        session['modul'] = modul

        if modul == 'produkcja':
            return redirect(url_for('produkcja_view'))
        return redirect(url_for('index'))

    return render_template_string(LOGIN_TEMPLATE, t=t, lang=lang, wybrany_modul=wybrany_modul, error=None)

@app.route('/przelacz-modul/<modul>')
def przelacz_modul(modul):
    if 'user' not in session:
        return redirect(url_for('login'))
    if modul not in ['koncowki', 'produkcja']:
        return redirect(url_for('index'))

    dostep_lista = session.get('dostep', [])
    rola = session.get('rola')
    if not czy_ma_dostep_do_modulu(dostep_lista, rola, modul):
        return "Brak uprawnień do tego modułu!", 403

    session['modul'] = modul
    if modul == 'produkcja':
        return redirect(url_for('produkcja_view'))

    target_linia = request.args.get('linia')
    if target_linia:
        session['aktywna_linia'] = target_linia
        return redirect(url_for('index', linia=target_linia))
    return redirect(url_for('index'))

@app.route('/logout')
def logout():
    zachowany_lang = session.get('lang', 'pl')
    session.clear()
    session['lang'] = zachowany_lang
    return redirect(url_for('login'))

# --- WIDOK MODUŁU PRODUKCYJNEGO (LINIE 1-6 + PRACOWNICY/QR + ADMIN) ---

@app.route('/produkcja')
def produkcja_view():
    if 'user' not in session:
        return redirect(url_for('login'))

    dostep_lista = session.get('dostep', [])
    rola = session.get('rola')
    if not czy_ma_dostep_do_modulu(dostep_lista, rola, 'produkcja'):
        return "Brak uprawnień do Modułu Produkcyjnego!", 403

    t, lang = get_lang_dict()
    session['modul'] = 'produkcja'
    dzis = datetime.now().strftime('%Y-%m-%d')
    wybrana_data = request.args.get('wybrana_data', dzis) if rola == 'admin' else dzis

    pracownicy_rows = PracownikQrModel.query.order_by(PracownikQrModel.kolejnosc.asc(), PracownikQrModel.emp_id.asc()).all()
    employees_dict = {
        p.emp_id: {
            "emp_id": p.emp_id,
            "imie_nazwisko": p.imie_nazwisko,
            "stanowisko": p.stanowisko or "",
            "kolejnosc": p.kolejnosc
        }
        for p in pracownicy_rows
    }

    snapshot = build_monitor_snapshot(wybrana_data)
    uzytkownicy_baza_dict = {
        u.username: {"pelna_nazwa": u.pelna_nazwa, "rola": u.rola, "dostep": json.loads(u.dostep)}
        for u in UzytkownikModel.query.all()
    }
    can_access_koncowki = czy_ma_dostep_do_modulu(dostep_lista, rola, 'koncowki')

    return render_template_string(
        PRODUKCJA_MODULE_TEMPLATE,
        t=t,
        lang=lang,
        pelna_nazwa=session.get('pelna_nazwa'),
        role=rola,
        dostep=dostep_lista,
        can_access_koncowki=can_access_koncowki,
        employees_json=json.dumps(employees_dict),
        initial_snapshot_json=json.dumps(snapshot),
        uzytkownicy_baza=uzytkownicy_baza_dict
    )

@app.route('/api/produkcja/snapshot')
def api_produkcja_snapshot():
    if 'user' not in session:
        return jsonify({"error": "Brak sesji"}), 401
    dzis = datetime.now().strftime('%Y-%m-%d')
    wybrana_data = request.args.get('data', dzis) if session.get('rola') == 'admin' else dzis
    return jsonify(build_monitor_snapshot(wybrana_data))

@app.route('/api/produkcja/linia-akcja', methods=['POST'])
def api_produkcja_linia_akcja():
    if 'user' not in session:
        return jsonify({"error": "Brak sesji"}), 401

    login_uzytkownika = session.get('user') or '-'
    payload = request.get_json(silent=True) or {}
    line_id = int(payload.get('lineId', 0))
    action = payload.get('action')
    reason = (payload.get('reason') or '').strip()

    if line_id not in [1, 2, 3, 4, 5, 6]:
        return jsonify({"error": "Nieprawidłowy numer linii"}), 400

    stan = LiniaMonitorStanModel.query.get(line_id)
    if not stan:
        stan = LiniaMonitorStanModel(linia_id=line_id, status='IDLE', start_time_iso=None, reason='', started_by_login=None)
        db.session.add(stan)

    now_dt = datetime.now()
    now_iso = now_dt.isoformat()
    today_iso = now_dt.strftime('%Y-%m-%d')

    if action == 'start':
        if stan.status != 'IDLE' and stan.start_time_iso:
            prev_dt = parse_iso_dt(stan.start_time_iso)
            dur = max(0, int((now_dt - prev_dt).total_seconds())) if prev_dt else 0
            db.session.add(LiniaMonitorZdarzenieModel(
                linia_id=line_id,
                event_type='PRACA' if stan.status == 'RUNNING' else 'POSTÓJ',
                start_iso=stan.start_time_iso,
                end_iso=now_iso,
                duration_sec=dur,
                reason=stan.reason or '-',
                iso_date=stan.start_time_iso[:10],
                started_by_login=stan.started_by_login or login_uzytkownika,
                ended_by_login=login_uzytkownika
            ))
        stan.status = 'RUNNING'
        stan.start_time_iso = now_iso
        stan.reason = ''
        stan.started_by_login = login_uzytkownika

    elif action == 'stop':
        if not reason:
            return jsonify({"error": "Wybierz lub wpisz przyczynę postoju!"}), 400
        if stan.status != 'IDLE' and stan.start_time_iso:
            prev_dt = parse_iso_dt(stan.start_time_iso)
            dur = max(0, int((now_dt - prev_dt).total_seconds())) if prev_dt else 0
            db.session.add(LiniaMonitorZdarzenieModel(
                linia_id=line_id,
                event_type='PRACA' if stan.status == 'RUNNING' else 'POSTÓJ',
                start_iso=stan.start_time_iso,
                end_iso=now_iso,
                duration_sec=dur,
                reason=stan.reason or '-',
                iso_date=stan.start_time_iso[:10],
                started_by_login=stan.started_by_login or login_uzytkownika,
                ended_by_login=login_uzytkownika
            ))
        stan.status = 'STOPPED'
        stan.start_time_iso = now_iso
        stan.reason = reason
        stan.started_by_login = login_uzytkownika

    elif action == 'end':
        if stan.status != 'IDLE' and stan.start_time_iso:
            prev_dt = parse_iso_dt(stan.start_time_iso)
            dur = max(0, int((now_dt - prev_dt).total_seconds())) if prev_dt else 0
            db.session.add(LiniaMonitorZdarzenieModel(
                linia_id=line_id,
                event_type='PRACA' if stan.status == 'RUNNING' else 'POSTÓJ',
                start_iso=stan.start_time_iso,
                end_iso=now_iso,
                duration_sec=dur,
                reason=stan.reason or '-',
                iso_date=stan.start_time_iso[:10],
                started_by_login=stan.started_by_login or login_uzytkownika,
                ended_by_login=login_uzytkownika
            ))

        aktywne_na_linii = PracownikSesjaAktywnaModel.query.filter_by(line_id=line_id).all()
        for act in aktywne_na_linii:
            act_dt = parse_iso_dt(act.start_time_iso)
            dur_emp = max(0, int((now_dt - act_dt).total_seconds())) if act_dt else 0
            db.session.add(PracownikSesjaHistoriaModel(
                emp_id=act.emp_id,
                emp_name=act.emp_name,
                line_id=act.line_id,
                start_time_iso=act.start_time_iso,
                end_time_iso=now_iso,
                duration_sec=dur_emp,
                iso_date=act.iso_date or today_iso
            ))
            db.session.delete(act)

        stan.status = 'IDLE'
        stan.start_time_iso = None
        stan.reason = ''
        stan.started_by_login = None
    else:
        return jsonify({"error": "Nieznana akcja"}), 400

    db.session.commit()
    return jsonify({"ok": True, "snapshot": build_monitor_snapshot(today_iso)})

@app.route('/api/produkcja/skan-pracownika', methods=['POST'])
def api_produkcja_skan_pracownika():
    if 'user' not in session:
        return jsonify({"error": "Brak sesji"}), 401

    payload = request.get_json(silent=True) or {}
    line_id = int(payload.get('lineId', 0))
    emp_id = (payload.get('empId') or '').strip()
    new_name = (payload.get('newName') or '').strip()

    if line_id not in [1, 2, 3, 4, 5, 6] or not emp_id:
        return jsonify({"error": "Nieprawidłowe dane skanowania"}), 400

    prac = PracownikQrModel.query.get(emp_id)
    if not prac:
        if not new_name:
            return jsonify({"error": "Nieznany pracownik"}), 404
        max_kol = db.session.query(db.func.max(PracownikQrModel.kolejnosc)).scalar() or 0
        prac = PracownikQrModel(emp_id=emp_id, imie_nazwisko=new_name, stanowisko="", kolejnosc=max_kol + 1)
        db.session.add(prac)
        db.session.flush()

    now_dt = datetime.now()
    now_iso = now_dt.isoformat()
    today_iso = now_dt.strftime('%Y-%m-%d')

    same_line_act = PracownikSesjaAktywnaModel.query.filter_by(emp_id=emp_id, line_id=line_id).first()
    if same_line_act:
        act_dt = parse_iso_dt(same_line_act.start_time_iso)
        dur = max(0, int((now_dt - act_dt).total_seconds())) if act_dt else 0
        db.session.add(PracownikSesjaHistoriaModel(
            emp_id=same_line_act.emp_id,
            emp_name=same_line_act.emp_name,
            line_id=same_line_act.line_id,
            start_time_iso=same_line_act.start_time_iso,
            end_time_iso=now_iso,
            duration_sec=dur,
            iso_date=same_line_act.iso_date or today_iso
        ))
        db.session.delete(same_line_act)
    else:
        other_acts = PracownikSesjaAktywnaModel.query.filter_by(emp_id=emp_id).all()
        for act in other_acts:
            act_dt = parse_iso_dt(act.start_time_iso)
            dur = max(0, int((now_dt - act_dt).total_seconds())) if act_dt else 0
            db.session.add(PracownikSesjaHistoriaModel(
                emp_id=act.emp_id,
                emp_name=act.emp_name,
                line_id=act.line_id,
                start_time_iso=act.start_time_iso,
                end_time_iso=now_iso,
                duration_sec=dur,
                iso_date=act.iso_date or today_iso
            ))
            db.session.delete(act)

        db.session.add(PracownikSesjaAktywnaModel(
            emp_id=prac.emp_id,
            emp_name=prac.imie_nazwisko,
            line_id=line_id,
            start_time_iso=now_iso,
            iso_date=today_iso
        ))

    db.session.commit()

    pracownicy_rows = PracownikQrModel.query.order_by(PracownikQrModel.kolejnosc.asc(), PracownikQrModel.emp_id.asc()).all()
    employees_dict = {
        p.emp_id: {
            "emp_id": p.emp_id,
            "imie_nazwisko": p.imie_nazwisko,
            "stanowisko": p.stanowisko or "",
            "kolejnosc": p.kolejnosc
        }
        for p in pracownicy_rows
    }

    return jsonify({"ok": True, "snapshot": build_monitor_snapshot(today_iso), "employees": employees_dict})

@app.route('/api/produkcja/wymelduj-pracownika', methods=['POST'])
def api_produkcja_wymelduj_pracownika():
    if 'user' not in session:
        return jsonify({"error": "Brak sesji"}), 401

    payload = request.get_json(silent=True) or {}
    line_id = int(payload.get('lineId', 0))
    emp_id = (payload.get('empId') or '').strip()

    now_dt = datetime.now()
    now_iso = now_dt.isoformat()
    today_iso = now_dt.strftime('%Y-%m-%d')

    act = PracownikSesjaAktywnaModel.query.filter_by(emp_id=emp_id, line_id=line_id).first()
    if act:
        act_dt = parse_iso_dt(act.start_time_iso)
        dur = max(0, int((now_dt - act_dt).total_seconds())) if act_dt else 0
        db.session.add(PracownikSesjaHistoriaModel(
            emp_id=act.emp_id,
            emp_name=act.emp_name,
            line_id=act.line_id,
            start_time_iso=act.start_time_iso,
            end_time_iso=now_iso,
            duration_sec=dur,
            iso_date=act.iso_date or today_iso
        ))
        db.session.delete(act)
        db.session.commit()

    return jsonify({"ok": True, "snapshot": build_monitor_snapshot(today_iso)})

# --- WSPÓLNA BAZA PRACOWNIKÓW (API) ---

@app.route('/api/pracownicy/zapisz', methods=['POST'])
def api_pracownicy_zapisz():
    if 'user' not in session:
        return jsonify({"error": "Brak sesji"}), 401

    payload = request.get_json(silent=True) or {}
    emp_id = (payload.get('empId') or '').strip()
    imie_nazwisko = (payload.get('imieNazwisko') or '').strip()
    stanowisko = (payload.get('stanowisko') or '').strip()
    is_edit = bool(payload.get('isEdit'))

    if not emp_id or not imie_nazwisko:
        return jsonify({"error": "Wymagane ID oraz Imię i Nazwisko"}), 400

    prac = PracownikQrModel.query.get(emp_id)
    if prac:
        prac.imie_nazwisko = imie_nazwisko
        prac.stanowisko = stanowisko
    else:
        if is_edit:
            return jsonify({"error": "Nie znaleziono pracownika do edycji"}), 404
        max_kol = db.session.query(db.func.max(PracownikQrModel.kolejnosc)).scalar() or 0
        prac = PracownikQrModel(emp_id=emp_id, imie_nazwisko=imie_nazwisko, stanowisko=stanowisko, kolejnosc=max_kol + 1)
        db.session.add(prac)

    for act in PracownikSesjaAktywnaModel.query.filter_by(emp_id=emp_id).all():
        act.emp_name = imie_nazwisko

    db.session.commit()

    pracownicy_rows = PracownikQrModel.query.order_by(PracownikQrModel.kolejnosc.asc(), PracownikQrModel.emp_id.asc()).all()
    employees_dict = {
        p.emp_id: {
            "emp_id": p.emp_id,
            "imie_nazwisko": p.imie_nazwisko,
            "stanowisko": p.stanowisko or "",
            "kolejnosc": p.kolejnosc
        }
        for p in pracownicy_rows
    }
    return jsonify({"ok": True, "employees": employees_dict})

@app.route('/api/pracownicy/usun', methods=['POST'])
def api_pracownicy_usun():
    if 'user' not in session:
        return jsonify({"error": "Brak sesji"}), 401

    payload = request.get_json(silent=True) or {}
    emp_id = (payload.get('empId') or '').strip()
    prac = PracownikQrModel.query.get(emp_id)
    if prac:
        db.session.delete(prac)
        db.session.commit()

    pracownicy_rows = PracownikQrModel.query.order_by(PracownikQrModel.kolejnosc.asc(), PracownikQrModel.emp_id.asc()).all()
    employees_dict = {
        p.emp_id: {
            "emp_id": p.emp_id,
            "imie_nazwisko": p.imie_nazwisko,
            "stanowisko": p.stanowisko or "",
            "kolejnosc": p.kolejnosc
        }
        for p in pracownicy_rows
    }
    return jsonify({"ok": True, "employees": employees_dict})

# --- MODUŁ KOŃCÓWKI LINII / KOMPLETACJA / CENTRUM RAPORTÓW ---

@app.route('/')
def index():
    if 'user' not in session:
        return redirect(url_for('login'))

    dostep = session.get('dostep', [])
    role = session.get('rola')
    if not czy_ma_dostep_do_modulu(dostep, role, 'koncowki') and czy_ma_dostep_do_modulu(dostep, role, 'produkcja'):
        return redirect(url_for('produkcja_view'))

    session['modul'] = 'koncowki'
    t, lang = get_lang_dict()
    user = session.get('user')
    pelna_nazwa = session.get('pelna_nazwa')
    wypr_zmiana = session.get('wybrana_zmiana')
    can_access_produkcja = czy_ma_dostep_do_modulu(dostep, role, 'produkcja')

    if not wypr_zmiana:
        return render_template_string(
            KONCOWKI_TEMPLATE, t=t, lang=lang, wypr_zmiana=None, user=user, role=role,
            pelna_nazwa=pelna_nazwa, dostep=dostep, can_access_produkcja=can_access_produkcja,
            produkty_json="{}", plan_json="{}", zrealizowane_json="{}"
        )

    linia = request.args.get('linia', session.get('aktywna_linia', '1'))
    session['aktywna_linia'] = linia

    dzis = datetime.now().strftime('%Y-%m-%d')
    wybrana_data = request.args.get('wybrana_data', dzis) if role == 'admin' else dzis

    min_data = (datetime.now() + timedelta(days=4)).strftime('%Y-%m-%d')
    dof = datetime.now().timetuple().tm_yday
    domyslna_partia = f"{dof}/{wypr_zmiana}/{linia}"
    error_msg = session.pop('error_msg', None)

    produkty_baza = {p.kod: {"nazwa": p.nazwa, "odbiorca": p.odbiorca} for p in ArtykulModel.query.all()}

    plan_produkcyjny = {}
    for pl in PlanModel.query.all():
        plan_produkcyjny.setdefault(pl.data, {})[pl.kod] = pl.ilosc

    dostepne_daty_planu = sorted(plan_produkcyjny.keys(), reverse=True)
    wybrana_data_planu = request.args.get('wybrana_data_planu', dostepne_daty_planu[0] if dostepne_daty_planu else dzis)
    if wybrana_data_planu not in plan_produkcyjny and dostepne_daty_planu:
        wybrana_data_planu = dostepne_daty_planu[0]

    wpisy_baza_wszystkie = [{
        "id": w.id, "data": w.data, "linia": w.linia, "zmiana": w.zmiana, "kod": w.kod,
        "nazwa": w.nazwa, "odbiorca": w.odbiorca, "ilosc": w.ilosc,
        "data_przyd": w.data_przyd, "partia_prod": w.partia_prod,
        "etykieta_gora": w.etykieta_gora, "etykieta_dol": w.etykieta_dol,
        "czas_start": w.czas_start, "czas_koniec": w.czas_koniec, "operator": w.operator
    } for w in WpisProdukcyjnyModel.query.all()]

    dostepne_daty_wpisow = sorted(list(set([w['data'] for w in wpisy_baza_wszystkie] + [dzis])), reverse=True)
    dostepne_daty_wpisow_lub_planu = sorted(list(set(list(plan_produkcyjny.keys()) + [w['data'] for w in wpisy_baza_wszystkie] + [dzis])), reverse=True)

    wybrana_data_progress = request.args.get('wybrana_data_progress', dzis)
    if wybrana_data_progress not in dostepne_daty_wpisow_lub_planu and dostepne_daty_wpisow_lub_planu:
        wybrana_data_progress = dostepne_daty_wpisow_lub_planu[0]

    wybrana_data_wydajnosc = request.args.get('wybrana_data_wydajnosc', dzis)
    if wybrana_data_wydajnosc not in dostepne_daty_wpisow and dostepne_daty_wpisow:
        wybrana_data_wydajnosc = dostepne_daty_wpisow[0]

    zrealizowane_z_datami = {}
    for w in wpisy_baza_wszystkie:
        zrealizowane_z_datami.setdefault(w['data'], {})[w['kod']] = zrealizowane_z_datami.get(w['data'], {}).get(w['kod'], 0) + int(w['ilosc'])

    plan_wybrany_prog = plan_produkcyjny.get(wybrana_data_progress, {})
    suma_planu_calkowita_prog = sum(int(v) for v in plan_wybrany_prog.values())

    wykonane_dla_wybranej_daty_prog = {}
    for w in wpisy_baza_wszystkie:
        if w['data'] == wybrana_data_progress:
            wykonane_dla_wybranej_daty_prog[w['kod']] = wykonane_dla_wybranej_daty_prog.get(w['kod'], 0) + int(w['ilosc'])

    suma_wykonana_calkowita_prog = sum(int(v) for v in wykonane_dla_wybranej_daty_prog.values())
    suma_wkladu_plan_prog = sum(min(int(wykonane_dla_wybranej_daty_prog.get(k, 0)), int(p_val)) for k, p_val in plan_wybrany_prog.items())
    nadwyzka_calkowita_prog = sum(max(int(wykonane_dla_wybranej_daty_prog.get(k, 0)) - int(p_val), 0) for k, p_val in plan_wybrany_prog.items())
    pozostalo_calkowite_prog = max(suma_planu_calkowita_prog - suma_wkladu_plan_prog, 0)

    procent_podstawowy_prog = round(suma_wkladu_plan_prog / suma_planu_calkowita_prog * 100) if suma_planu_calkowita_prog > 0 else 0
    procent_calkowity_prog = round(suma_wykonana_calkowita_prog / suma_planu_calkowita_prog * 100) if suma_planu_calkowita_prog > 0 else 0

    wpisy_dla_linii = [w for w in wpisy_baza_wszystkie if w['data'] == wybrana_data and w['linia'] == linia and w['zmiana'] == wypr_zmiana]
    suma_sztuk = sum(int(w['ilosc']) for w in wpisy_dla_linii)

    grupy_wyd = {}
    for item in wpisy_baza_wszystkie:
        if item['data'] == wybrana_data_wydajnosc:
            key = f"L{str(item['linia'])}_Z{str(item['zmiana'])}"
            if key not in grupy_wyd:
                grupy_wyd[key] = {"suma_sztuk": 0, "przedzialy": []}
            grupy_wyd[key]["suma_sztuk"] += int(item['ilosc'])
            try:
                t1 = datetime.strptime(item['czas_start'], '%H:%M')
                t2 = datetime.strptime(item['czas_koniec'], '%H:%M')
                m1 = t1.hour * 60 + t1.minute
                m2 = t2.hour * 60 + t2.minute
                if m2 > m1:
                    grupy_wyd[key]["przedzialy"].append((m1, m2))
            except Exception:
                pass

    statystyki_wydajnosci_lista = []
    for l_nr in ['1', '2', '3', '4']:
        for z_nr in ['1', '2']:
            dane_g = grupy_wyd.get(f"L{l_nr}_Z{z_nr}", {"suma_sztuk": 0, "przedzialy": []})
            czas_minut = oblicz_czas_bez_nakladania(dane_g["przedzialy"])
            godziny = czas_minut / 60.0
            czas_godz = round(godziny, 2)
            sztuk_na_godzine = round(dane_g["suma_sztuk"] / godziny, 1) if godziny > 0 else 0
            statystyki_wydajnosci_lista.append({
                "linia": l_nr, "zmiana": z_nr, "suma_sztuk": dane_g["suma_sztuk"],
                "czas_godz": czas_godz, "sztuk_na_godzine": sztuk_na_godzine
            })

    uzytkownicy_baza_dict = {
        u.username: {"pelna_nazwa": u.pelna_nazwa, "rola": u.rola, "dostep": json.loads(u.dostep)}
        for u in UzytkownikModel.query.all()
    }

    return render_template_string(
        KONCOWKI_TEMPLATE,
        t=t, lang=lang, linia=linia, wypr_zmiana=wypr_zmiana, dzis=dzis,
        wybrana_data=wybrana_data, wybrana_data_planu=wybrana_data_planu,
        wybrana_data_progress=wybrana_data_progress, wybrana_data_wydajnosc=wybrana_data_wydajnosc,
        dostepne_daty_planu=dostepne_daty_planu, dostepne_daty_wpisow=dostepne_daty_wpisow,
        dostepne_daty_wpisow_lub_planu=dostepne_daty_wpisow_lub_planu, min_data=min_data,
        domyslna_partia=domyslna_partia, produkty_baza=produkty_baza,
        produkty_json=json.dumps(produkty_baza), plan_json=json.dumps(plan_produkcyjny),
        plan_produkcyjny=plan_produkcyjny, zrealizowane_json=json.dumps(zrealizowane_z_datami),
        plan_wybrany_prog=plan_wybrany_prog, wykonane_dla_wybranej_daty_prog=wykonane_dla_wybranej_daty_prog,
        suma_planu_calkowita_prog=suma_planu_calkowita_prog, suma_wykonana_calkowita_prog=suma_wykonana_calkowita_prog,
        suma_wkladu_plan_prog=suma_wkladu_plan_prog, nadwyzka_calkowita_prog=nadwyzka_calkowita_prog,
        pozostalo_calkowite_prog=pozostalo_calkowite_prog, procent_podstawowy_prog=procent_podstawowy_prog,
        procent_calkowity_prog=procent_calkowity_prog, wpisy_dla_linii=wpisy_dla_linii,
        suma_sztuk=suma_sztuk, statystyki_wydajnosci_lista=statystyki_wydajnosci_lista,
        user=user, role=role, pelna_nazwa=pelna_nazwa, dostep=dostep,
        can_access_produkcja=can_access_produkcja,
        uzytkownicy_baza=uzytkownicy_baza_dict, error_msg=error_msg
    )

@app.route('/wybierz-zmiane', methods=['POST'])
def wybierz_zmiane():
    session['wybrana_zmiana'] = request.form.get('zmiana')
    if session.get('modul') == 'produkcja':
        return redirect(url_for('produkcja_view'))
    return redirect(url_for('index'))

@app.route('/reset-zmiana')
def reset_zmiana():
    session.pop('wybrana_zmiana', None)
    if session.get('modul') == 'produkcja':
        return redirect(url_for('produkcja_view'))
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
    nowy_wpis = WpisProdukcyjnyModel(
        data=data_prod, linia=linia, zmiana=session.get('wybrana_zmiana'),
        kod=kod, nazwa=artykul.nazwa if artykul else "", odbiorca=artykul.odbiorca if artykul else "",
        ilosc=int(request.form.get('ilosc')), data_przyd=request.form.get('data_przyd'),
        partia_prod=request.form.get('partia_prod'), etykieta_gora=request.form.get('etykieta_gora'),
        etykieta_dol=request.form.get('etykieta_dol'), czas_start=czas_start, czas_koniec=czas_koniec,
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
    kod = request.form.get('kod', '').strip()
    nazwa = request.form.get('nazwa', '').strip()
    odbiorca = request.form.get('odbiorca', '').strip()
    if kod and nazwa:
        istniejacy = ArtykulModel.query.get(kod)
        if istniejacy:
            istniejacy.nazwa = nazwa
            istniejacy.odbiorca = odbiorca
        else:
            db.session.add(ArtykulModel(kod=kod, nazwa=nazwa, odbiorca=odbiorca))
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
        for w in surowy.split('\n'):
            czesci = w.split('\t')
            if len(czesci) >= 2:
                kod = czesci[0].strip()
                try:
                    ilosc = int(czesci[1].strip().replace(' ', ''))
                    istniejacy = PlanModel.query.filter_by(data=data_planu, kod=kod).first()
                    if istniejacy:
                        istniejacy.ilosc = ilosc
                    else:
                        db.session.add(PlanModel(data=data_planu, kod=kod, ilosc=ilosc))
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
    poz = PlanModel.query.filter_by(data=data_p, kod=kod).first()
    if poz:
        db.session.delete(poz)
        db.session.commit()
    return redirect(url_for('index', linia='plan', wybrana_data_planu=data_p))

@app.route('/admin/dodaj-uzytkownika', methods=['POST'])
def admin_dodaj_uzytkownika():
    if session.get('rola') != 'admin':
        return "Brak uprawnień!"
    username = request.form.get('username', '').strip()
    password = request.form.get('password', '').strip()
    pelna_nazwa = request.form.get('pelna_nazwa', '').strip()
    rola = request.form.get('rola')
    dostep = request.form.getlist('dostep')
    powrot_modul = request.form.get('powrot_modul')

    if username and password and pelna_nazwa:
        usr = UzytkownikModel.query.get(username)
        if usr:
            usr.haslo = password
            usr.pelna_nazwa = pelna_nazwa
            usr.rola = rola
            usr.dostep = json.dumps(dostep)
        else:
            db.session.add(UzytkownikModel(
                username=username, haslo=password, rola=rola,
                pelna_nazwa=pelna_nazwa, dostep=json.dumps(dostep)
            ))
        db.session.commit()

    if powrot_modul == 'produkcja':
        return redirect(url_for('produkcja_view'))
    return redirect(url_for('index', linia='admin'))

@app.route('/admin/usun-uzytkownika/<username>')
def admin_usun_uzytkownika(username):
    if session.get('rola') != 'admin':
        return "Brak uprawnień!"
    powrot_modul = request.args.get('powrot_modul')
    if username != 'admin':
        usr = UzytkownikModel.query.get(username)
        if usr:
            db.session.delete(usr)
            db.session.commit()
    if powrot_modul == 'produkcja':
        return redirect(url_for('produkcja_view'))
    return redirect(url_for('index', linia='admin'))

# --- EKSPORT RAPORTÓW W CENTRUM RAPORTÓW ---

def stworz_arkusz_linii(wb, sheet_title, wpisy_linii, wybrana_data, zmiana_label, l_nr, first_sheet_ref):
    if first_sheet_ref[0]:
        ws_linia = wb.active
        ws_linia.title = sheet_title
        first_sheet_ref[0] = False
    else:
        ws_linia = wb.create_sheet(title=sheet_title)

    ws_linia.views.sheetView[0].showGridLines = True
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_normal = Font(name="Calibri", size=10)
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")

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
    ws_linia.cell(row=5, column=1, value="Data produkcji:").font = font_bold
    ws_linia.cell(row=5, column=2, value=wybrana_data).font = font_normal
    ws_linia.cell(row=5, column=4, value="Dzień roku:").font = font_bold
    ws_linia.cell(row=5, column=5, value=d_z.timetuple().tm_yday).font = font_normal
    ws_linia.cell(row=5, column=6, value="Zmiana:").font = font_bold
    ws_linia.cell(row=5, column=7, value=zmiana_label).font = font_normal
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

    ws_linia['A3'] = f"ROZCHÓD WEWNĘTRZNY - OCENA PRODUKTU - WYRÓB GOTOWY (ZMIANA {zmiana_label})"
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
            str(item['kod']), item['nazwa'], item['odbiorca'],
            f"{item['data_przyd']} / {item['partia_prod']}", int(item['ilosc']),
            "+", "+", "+", item['etykieta_gora'], item['etykieta_dol'],
            item['czas_start'], item['czas_koniec'], " + ", item['operator']
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

@app.route('/generate', methods=['GET', 'POST'])
def generate():
    if 'user' not in session:
        return redirect(url_for('login'))

    dzis = datetime.now().strftime('%Y-%m-%d')
    wybrana_data = request.args.get('data', dzis)
    typ_raportu = request.args.get('typ_raportu', 'rozchod_osobne_arkusze')

    wpisy_dla_dnia = [{
        "linia": w.linia, "zmiana": w.zmiana, "kod": w.kod,
        "nazwa": w.nazwa, "odbiorca": w.odbiorca, "ilosc": w.ilosc,
        "data_przyd": w.data_przyd, "partia_prod": w.partia_prod,
        "etykieta_gora": w.etykieta_gora, "etykieta_dol": w.etykieta_dol,
        "czas_start": w.czas_start, "czas_koniec": w.czas_koniec, "operator": w.operator
    } for w in WpisProdukcyjnyModel.query.filter_by(data=wybrana_data).all()]

    wb = openpyxl.Workbook()
    font_title = Font(name="Calibri", size=14, bold=True, color="1F497D")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_normal = Font(name="Calibri", size=10)
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )

    if typ_raportu == 'postep_realizacja':
        ws_p = wb.active
        ws_p.title = "Postęp Realizacji"
        ws_p.views.sheetView[0].showGridLines = True
        ws_p['A1'] = f"RAPORT POSTĘPU I REALIZACJI PLANU - {wybrana_data}"
        ws_p['A1'].font = font_title

        plan_map = {p.kod: int(p.ilosc) for p in PlanModel.query.filter_by(data=wybrana_data).all()}
        wykonane_map = {}
        for item in wpisy_dla_dnia:
            wykonane_map[item['kod']] = wykonane_map.get(item['kod'], 0) + int(item['ilosc'])

        headers_p = ["Kod artykułu", "Nazwa produktu", "Odbiorca", "Plan", "Wykonano", "Pozostało", "Nadwyżka", "% Realizacji"]
        ws_p.row_dimensions[3].height = 25
        for c_idx, h_text in enumerate(headers_p, start=1):
            c = ws_p.cell(row=3, column=c_idx, value=h_text)
            c.font = font_header
            c.fill = fill_header
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin_border

        produkty_all = {p.kod: {"nazwa": p.nazwa, "odbiorca": p.odbiorca} for p in ArtykulModel.query.all()}
        r_p = 4
        for kod, plan_val in plan_map.items():
            wyk = wykonane_map.get(kod, 0)
            poz = max(plan_val - wyk, 0)
            nad = max(wyk - plan_val, 0)
            proc = round(wyk / plan_val * 100) if plan_val > 0 else 0
            p_info = produkty_all.get(kod, {})
            vals = [str(kod), p_info.get("nazwa", "Nieznany"), p_info.get("odbiorca", "-"), int(plan_val), int(wyk), int(poz), int(nad), f"{proc}%"]
            for c_idx, val in enumerate(vals, start=1):
                cell = ws_p.cell(row=r_p, column=c_idx, value=val)
                cell.font = font_normal
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center" if c_idx not in [2, 3] else "left", vertical="center")
            r_p += 1

        for col in ws_p.columns:
            max_l = max(len(str(cell.value or '')) for cell in col)
            ws_p.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = max(max_l + 4, 15)
        download_name = f"Raport_Postepu_{wybrana_data}.xlsx"

    elif typ_raportu == 'wydajnosc':
        ws_wyd = wb.active
        ws_wyd.title = "Raport Wydajności"
        ws_wyd.views.sheetView[0].showGridLines = True
        ws_wyd['A1'] = f"RAPORT WYDAJNOŚCI (SZT/H) - {wybrana_data}"
        ws_wyd['A1'].font = font_title

        stats = {}
        for item in wpisy_dla_dnia:
            key = f"L{item['linia']}_Z{item['zmiana']}"
            if key not in stats:
                stats[key] = {"sztuki": 0, "przedzialy": []}
            stats[key]["sztuki"] += int(item['ilosc'])
            try:
                t1 = datetime.strptime(item['czas_start'], '%H:%M')
                t2 = datetime.strptime(item['czas_koniec'], '%H:%M')
                m1 = t1.hour * 60 + t1.minute
                m2 = t2.hour * 60 + t2.minute
                if m2 > m1:
                    stats[key]["przedzialy"].append((m1, m2))
            except Exception:
                pass

        headers_w = ["Linia", "Zmiana", "Wyprodukowano sztuk", "Czas pracy (godziny)", "Wydajność (szt/h)"]
        ws_wyd.row_dimensions[3].height = 25
        for c_idx, h_text in enumerate(headers_w, start=1):
            c = ws_wyd.cell(row=3, column=c_idx, value=h_text)
            c.font = font_header
            c.fill = fill_header
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin_border

        r_w = 4
        for l_nr in ['1', '2', '3', '4']:
            for z_nr in ['1', '2']:
                s = stats.get(f"L{l_nr}_Z{z_nr}", {"sztuki": 0, "przedzialy": []})
                czas_minut = oblicz_czas_bez_nakladania(s["przedzialy"])
                godz = round(czas_minut / 60.0, 2)
                wyd = round(s["sztuki"] / godz, 1) if godz > 0 else 0
                vals_w = [f"Linia {l_nr}", f"Zmiana {z_nr}", int(s["sztuki"]), godz, f"{wyd} szt/h"]
                for c_idx, val in enumerate(vals_w, start=1):
                    cell = ws_wyd.cell(row=r_w, column=c_idx, value=val)
                    cell.font = font_bold if c_idx in [3, 5] else font_normal
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                r_w += 1

        for col in ws_wyd.columns:
            max_l = max(len(str(cell.value or '')) for cell in col)
            ws_wyd.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = max(max_l + 4, 18)
        download_name = f"Raport_Wydajnosci_{wybrana_data}.xlsx"

    elif typ_raportu == 'wydajnosc_rbh':
        ws_rbh = wb.active
        ws_rbh.title = "Wydajność szt na rbh"
        ws_rbh.views.sheetView[0].showGridLines = True
        ws_rbh['A1'] = f"RAPORT WYDAJNOŚCI NA ROBOCZOGODZINĘ (SZT./RBH) - {wybrana_data}"
        ws_rbh['A1'].font = font_title

        sztuki_per_line_shift = {(str(i), z): 0 for i in [1, 2, 3, 4, 5, 6] for z in ['1', '2']}
        for item in wpisy_dla_dnia:
            l_key = str(item['linia'])
            z_key = str(item['zmiana'])
            if (l_key, z_key) in sztuki_per_line_shift:
                sztuki_per_line_shift[(l_key, z_key)] += int(item['ilosc'] or 0)

        now_dt = datetime.now()
        emp_sec_per_line_shift = {(i, z): 0 for i in [1, 2, 3, 4, 5, 6] for z in ['1', '2']}
        next_day_iso = (datetime.strptime(wybrana_data, '%Y-%m-%d') + timedelta(days=1)).strftime('%Y-%m-%d')

        for eh in PracownikSesjaHistoriaModel.query.filter(PracownikSesjaHistoriaModel.iso_date.in_([wybrana_data, next_day_iso])).all():
            dzien_roboczy, zm = okresl_dzien_roboczy_i_zmiane(eh.start_time_iso)
            if dzien_roboczy == wybrana_data and (eh.line_id, zm) in emp_sec_per_line_shift:
                emp_sec_per_line_shift[(eh.line_id, zm)] += int(eh.duration_sec or 0)

        for act in PracownikSesjaAktywnaModel.query.filter(PracownikSesjaAktywnaModel.iso_date.in_([wybrana_data, next_day_iso])).all():
            dzien_roboczy, zm = okresl_dzien_roboczy_i_zmiane(act.start_time_iso)
            if dzien_roboczy == wybrana_data and (act.line_id, zm) in emp_sec_per_line_shift:
                dt_s = parse_iso_dt(act.start_time_iso)
                if dt_s:
                    emp_sec_per_line_shift[(act.line_id, zm)] += max(0, int((now_dt - dt_s).total_seconds()))

        headers_r = [
            "Linia",
            "Zmiana",
            "Wyprodukowano sztuk",
            "Suma czasów pracowników (HH:MM:SS)",
            "Suma czasów pracowników (rbh)",
            "Wydajność (szt. / rbh)"
        ]
        ws_rbh.row_dimensions[3].height = 25
        for c_idx, h_text in enumerate(headers_r, start=1):
            c = ws_rbh.cell(row=3, column=c_idx, value=h_text)
            c.font = font_header
            c.fill = fill_header
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin_border

        r_idx = 4
        for lid in [1, 2, 3, 4, 5, 6]:
            for zm in ['1', '2']:
                szt = sztuki_per_line_shift[(str(lid), zm)]
                sec_total = emp_sec_per_line_shift[(lid, zm)]
                rbh = round(sec_total / 3600.0, 2)
                wyd_rbh = round(szt / (sec_total / 3600.0), 1) if sec_total > 0 else 0.0

                row_vals = [
                    f"Linia {lid}",
                    f"Zmiana {zm}",
                    szt,
                    format_sec_hhmmss(sec_total),
                    rbh,
                    f"{wyd_rbh} szt./rbh"
                ]
                for c_idx, val in enumerate(row_vals, start=1):
                    cell = ws_rbh.cell(row=r_idx, column=c_idx, value=val)
                    cell.font = font_bold if c_idx in [3, 6] else font_normal
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                r_idx += 1

        for col in ws_rbh.columns:
            max_l = max(len(str(cell.value or '')) for cell in col)
            ws_rbh.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = max(max_l + 4, 20)
        download_name = f"Raport_Wydajnosci_RBH_{wybrana_data}.xlsx"

    elif typ_raportu == 'monitoring_linie_pracownicy':
        ws_m = wb.active
        ws_m.title = "Monitoring Linii i Pracowników"
        ws_m.views.sheetView[0].showGridLines = True
        ws_m['A1'] = f"RAPORT MONITORINGU LINII 1–6 I CZASU PRACOWNIKÓW - {wybrana_data}"
        ws_m['A1'].font = font_title

        next_day_iso = (datetime.strptime(wybrana_data, '%Y-%m-%d') + timedelta(days=1)).strftime('%Y-%m-%d')
        zdarzenia_all = LiniaMonitorZdarzenieModel.query.filter(
            LiniaMonitorZdarzenieModel.iso_date.in_([wybrana_data, next_day_iso])
        ).order_by(LiniaMonitorZdarzenieModel.start_iso.desc()).all()
        zdarzenia = [z for z in zdarzenia_all if okresl_dzien_roboczy_i_zmiane(z.start_iso)[0] == wybrana_data]

        sesje_hist_all = PracownikSesjaHistoriaModel.query.filter(
            PracownikSesjaHistoriaModel.iso_date.in_([wybrana_data, next_day_iso])
        ).order_by(PracownikSesjaHistoriaModel.start_time_iso.desc()).all()
        sesje_hist = [s for s in sesje_hist_all if okresl_dzien_roboczy_i_zmiane(s.start_time_iso)[0] == wybrana_data]

        prac_meta = {p.emp_id: p.stanowisko for p in PracownikQrModel.query.all()}

        line_shift_times = {(lid, zm): {"work_sec": 0, "stop_sec": 0} for lid in [1, 2, 3, 4, 5, 6] for zm in ['1', '2']}
        for z in zdarzenia:
            _, zm_z = okresl_dzien_roboczy_i_zmiane(z.start_iso)
            key = (z.linia_id, zm_z)
            if key in line_shift_times:
                if z.event_type == 'PRACA':
                    line_shift_times[key]["work_sec"] += int(z.duration_sec or 0)
                elif z.event_type == 'POSTÓJ':
                    line_shift_times[key]["stop_sec"] += int(z.duration_sec or 0)

        now_dt = datetime.now()
        for stan_l in LiniaMonitorStanModel.query.all():
            if stan_l.status in ['RUNNING', 'STOPPED'] and stan_l.start_time_iso:
                dz_s, zm_s = okresl_dzien_roboczy_i_zmiane(stan_l.start_time_iso)
                dt_s = parse_iso_dt(stan_l.start_time_iso)
                if dz_s == wybrana_data and dt_s and (stan_l.linia_id, zm_s) in line_shift_times:
                    elapsed_s = max(0, int((now_dt - dt_s).total_seconds()))
                    if stan_l.status == 'RUNNING':
                        line_shift_times[(stan_l.linia_id, zm_s)]["work_sec"] += elapsed_s
                    else:
                        line_shift_times[(stan_l.linia_id, zm_s)]["stop_sec"] += elapsed_s

        headers_m = [
            "Linia",
            "Typ zdarzenia",
            "Rozpoczęcie",
            "Zakończenie",
            "Czas trwania",
            "Przyczyna postoju",
            "Rozpoczął (login)",
            "Zakończył (login)"
        ]
        ws_m.row_dimensions[3].height = 24
        for c_idx, h_text in enumerate(headers_m, start=1):
            c = ws_m.cell(row=3, column=c_idx, value=h_text)
            c.font = font_header
            c.fill = fill_header
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin_border

        headers_sum = [
            "Linia",
            "Zmiana",
            "Łączny czas pracy (HH:MM:SS)",
            "Łączny czas postojów (HH:MM:SS)",
            "Praca / Postój (godz.)"
        ]
        for idx_s, h_s_text in enumerate(headers_sum, start=10):
            cs = ws_m.cell(row=3, column=idx_s, value=h_s_text)
            cs.font = font_header
            cs.fill = fill_header
            cs.alignment = Alignment(horizontal="center", vertical="center")
            cs.border = thin_border

        row_sum_idx = 4
        for lid in [1, 2, 3, 4, 5, 6]:
            for zm in ['1', '2']:
                w_sec = line_shift_times[(lid, zm)]["work_sec"]
                s_sec = line_shift_times[(lid, zm)]["stop_sec"]
                vals_sum = [
                    f"Linia {lid}",
                    f"Zmiana {zm}",
                    format_sec_hhmmss(w_sec),
                    format_sec_hhmmss(s_sec),
                    f"{round(w_sec / 3600.0, 2)}h / {round(s_sec / 3600.0, 2)}h"
                ]
                for offset_c, val_s in enumerate(vals_sum):
                    cell_s = ws_m.cell(row=row_sum_idx, column=10 + offset_c, value=val_s)
                    cell_s.font = font_bold if offset_c in [0, 2, 3] else font_normal
                    cell_s.border = thin_border
                    cell_s.alignment = Alignment(horizontal="center", vertical="center")
                row_sum_idx += 1

        rm = 4
        for z in zdarzenia:
            vals_m = [
                f"Linia {z.linia_id}",
                z.event_type,
                z.start_iso,
                z.end_iso,
                format_sec_hhmmss(z.duration_sec),
                z.reason or '-',
                z.started_by_login or '-',
                z.ended_by_login or '-'
            ]
            for c_idx, val in enumerate(vals_m, start=1):
                cell = ws_m.cell(row=rm, column=c_idx, value=val)
                cell.font = font_normal
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
            rm += 1

        ws_emp = wb.create_sheet(title="Czasy Pracowników")
        ws_emp.views.sheetView[0].showGridLines = True
        ws_emp['A1'] = f"CZASY PRACY PRACOWNIKÓW NA LINIACH - {wybrana_data}"
        ws_emp['A1'].font = font_title

        headers_e = ["ID Pracownika", "Pracownik", "Stanowisko", "Linia", "Rozpoczęcie", "Zakończenie", "Czas (HH:MM:SS)", "Czas (rbh)"]
        ws_emp.row_dimensions[3].height = 24
        for c_idx, h_text in enumerate(headers_e, start=1):
            c = ws_emp.cell(row=3, column=c_idx, value=h_text)
            c.font = font_header
            c.fill = fill_header
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin_border

        re_row = 4
        for s in sesje_hist:
            dur_s = int(s.duration_sec or 0)
            vals_e = [
                s.emp_id, s.emp_name, prac_meta.get(s.emp_id) or "-",
                f"Linia {s.line_id}", s.start_time_iso, s.end_time_iso,
                format_sec_hhmmss(dur_s), round(dur_s / 3600.0, 2)
            ]
            for c_idx, val in enumerate(vals_e, start=1):
                cell = ws_emp.cell(row=re_row, column=c_idx, value=val)
                cell.font = font_normal
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
            re_row += 1

        for sheet_curr in [ws_m, ws_emp]:
            for col in sheet_curr.columns:
                max_l = max(len(str(cell.value or '')) for cell in col)
                sheet_curr.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = max(max_l + 4, 18)

        download_name = f"Raport_Monitoring_Pracownicy_{wybrana_data}.xlsx"

    else:
        if not wpisy_dla_dnia:
            return f"Brak wpisów do wygenerowania dla daty {wybrana_data}!"

        linie_numery = ['1', '2', '3', '4']
        first_sheet_ref = [True]

        if typ_raportu == 'rozchod_zmiana_1':
            for l_nr in linie_numery:
                wpisy_z1 = [w for w in wpisy_dla_dnia if str(w['linia']) == l_nr and str(w['zmiana']) == '1']
                if wpisy_z1:
                    stworz_arkusz_linii(wb, f"Linia {l_nr} - Z1", wpisy_z1, wybrana_data, "1", l_nr, first_sheet_ref)
            download_name = f"Raport_Rozchod_Z1_{wybrana_data}.xlsx"

        elif typ_raportu == 'rozchod_zmiana_2':
            for l_nr in linie_numery:
                wpisy_z2 = [w for w in wpisy_dla_dnia if str(w['linia']) == l_nr and str(w['zmiana']) == '2']
                if wpisy_z2:
                    stworz_arkusz_linii(wb, f"Linia {l_nr} - Z2", wpisy_z2, wybrana_data, "2", l_nr, first_sheet_ref)
            download_name = f"Raport_Rozchod_Z2_{wybrana_data}.xlsx"

        else:
            for zm_nr in ['1', '2']:
                for l_nr in linie_numery:
                    wpisy_z = [w for w in wpisy_dla_dnia if str(w['linia']) == l_nr and str(w['zmiana']) == zm_nr]
                    if wpisy_z:
                        stworz_arkusz_linii(wb, f"Linia {l_nr} - Z{zm_nr}", wpisy_z, wybrana_data, zm_nr, l_nr, first_sheet_ref)
            download_name = f"Raport_Rozchod_Z1_Z2_{wybrana_data}.xlsx"

        if first_sheet_ref[0]:
            return f"Brak wpisów dla wybranego kryterium zmiany w dniu {wybrana_data}!"

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return send_file(
        bio,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == '__main__':
    app.run(port=5000, debug=True)
