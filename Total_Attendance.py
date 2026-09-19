import io
import json
import sqlite3
import traceback
import calendar
from collections import defaultdict
from datetime import date, datetime, timedelta
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font as XlFont, Border, Side, PatternFill, Color

# ============================================================
# GMS SALATIGA - DATA MINISTRY SCHEDULER
# Streamlit Web Version
# ============================================================

st.set_page_config(
    page_title="GMS Salatiga - Data Ministry Scheduler",
    page_icon="⛪",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ------------------------- CSS -------------------------------
st.markdown("""
<style>
.main .block-container {padding-top: 1rem; padding-bottom: 2rem; max-width: 1500px;}
h1, h2, h3 {font-weight: 700;}
.app-title {font-size: 2rem; font-weight: 800; margin-bottom: 0.1rem;}
.app-subtitle {font-size: 1rem; opacity: 0.75; margin-bottom: 1rem;}
.section-card {
    border: 2px solid #5bc0de; border-radius: 10px; padding: 1rem 1.1rem;
    margin-bottom: 1rem; background: rgba(91,192,222,0.04);
}
.section-card-gold {
    border: 2px solid #f0ad4e; border-radius: 10px; padding: 1rem 1.1rem;
    margin-bottom: 1rem; background: rgba(240,173,78,0.05);
}
.schedule-table {width:100%; border-collapse:collapse; font-size:15px;}
.schedule-table th, .schedule-table td {border:1px solid #bbb; padding:10px; text-align:center; vertical-align:middle;}
.schedule-table th {background:#ffe699; font-weight:800;}
.schedule-date-sat {background:#d9e2f3; font-weight:700;}
.schedule-date-sun {background:#d9ead3; font-weight:700;}
.schedule-midweek {background:#00ffff; font-weight:800; font-size:17px;}
.schedule-gladi {background:#ffd700; font-weight:800; font-size:17px;}
.schedule-unfilled {background:#ff0000; color:white; font-weight:800;}
.schedule-note {text-align:left !important;}
.touch-button button {min-height:50px; font-size:17px; font-weight:700;}
div[data-testid="stDataFrame"] {font-size: 15px;}
@media (max-width: 900px) {
    .main .block-container {padding-left:0.6rem; padding-right:0.6rem;}
    .app-title {font-size:1.55rem;}
    .schedule-table {font-size:12px;}
    .schedule-table th, .schedule-table td {padding:7px 5px;}
}
</style>
""", unsafe_allow_html=True)

# ============================================================
# DATABASE
# ============================================================
class DatabaseManager:
    def __init__(self, db_name="ministry_schedule.db"):
        self.db_name = db_name
        self.conn = sqlite3.connect(db_name, check_same_thread=False)
        self.create_tables()

    def create_tables(self):
        cursor = self.conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS setup_data (
            id INTEGER PRIMARY KEY,
            group_name TEXT,
            members TEXT,
            pics TEXT
        )''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS schedule_data (
            id INTEGER PRIMARY KEY,
            month TEXT,
            year INTEGER,
            num_weeks INTEGER,
            week_dates TEXT,
            schedule_json TEXT,
            special_services_json TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS absences (
            id INTEGER PRIMARY KEY,
            name TEXT,
            start_week INTEGER,
            duration INTEGER,
            replacements TEXT
        )''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS gladi_events (
            id INTEGER PRIMARY KEY,
            event_type TEXT,
            date TEXT,
            time TEXT,
            description TEXT
        )''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS trainees (
            id INTEGER PRIMARY KEY,
            name TEXT,
            group_name TEXT,
            trainings INTEGER,
            month TEXT,
            service TEXT,
            finished INTEGER
        )''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS service_reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            service TEXT NOT NULL,
            service_date TEXT NOT NULL,
            month TEXT NOT NULL,
            year INTEGER NOT NULL,
            report_json TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(service, service_date)
        )''')
        self.conn.commit()

    def save_setup(self, groups_data):
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM setup_data")
        for group_name, data in groups_data.items():
            cursor.execute(
                "INSERT INTO setup_data (group_name, members, pics) VALUES (?, ?, ?)",
                (group_name, ",".join(data['members']), ",".join(data['pics']))
            )
        self.conn.commit()

    def load_setup(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM setup_data")
        rows = cursor.fetchall()
        groups = {}
        for row in rows:
            groups[row[1]] = {
                'members': row[2].split(',') if row[2] else [],
                'pics': row[3].split(',') if row[3] else []
            }
        return groups

    def save_schedule(self, month, year, num_weeks, week_dates, schedule, special_services):
        cursor = self.conn.cursor()
        cursor.execute(
            "INSERT INTO schedule_data (month, year, num_weeks, week_dates, schedule_json, special_services_json) VALUES (?, ?, ?, ?, ?, ?)",
            (month, year, num_weeks, json.dumps(week_dates), json.dumps(schedule), json.dumps(special_services))
        )
        self.conn.commit()

    def load_schedule(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM schedule_data ORDER BY created_at DESC LIMIT 1")
        row = cursor.fetchone()
        if row:
            schedule_raw = json.loads(row[5])
            schedule = {
                int(week): data
                for week, data in schedule_raw.items()
            }
            return {
                'month': row[1], 'year': row[2], 'num_weeks': row[3],
                'week_dates': json.loads(row[4]),
                'schedule': schedule,
                'special_services': json.loads(row[6])
            }
        return None

    def save_absences(self, absences_list):
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM absences")
        for absence in absences_list:
            cursor.execute(
                "INSERT INTO absences (name, start_week, duration, replacements) VALUES (?, ?, ?, ?)",
                (absence['name'], (absence.get('weeks') or [0])[0], len(absence.get('weeks', [])),
                 json.dumps({'replacements': absence.get('replacements', []), 'dates': absence.get('dates', [])}))
            )
        self.conn.commit()

    def load_absences(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM absences")
        result = []
        for row in cursor.fetchall():
            try:
                decoded = json.loads(row[4])
                if isinstance(decoded, dict):
                    replacements = decoded.get('replacements', [])
                    dates = decoded.get('dates', [])
                else:
                    replacements, dates = decoded, []
            except Exception:
                replacements, dates = [row[4]], []
            result.append({
                'name': row[1],
                'weeks': list(range(row[2], row[2] + row[3])) if row[2] else [],
                'dates': dates, 'replacements': replacements
            })
        return result

    def save_gladi_events(self, events):
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM gladi_events")
        for event in events:
            cursor.execute(
                "INSERT INTO gladi_events (event_type, date, time, description) VALUES (?, ?, ?, ?)",
                (event['type'], event['date'], event['time'], event['description'])
            )
        self.conn.commit()

    def load_gladi_events(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM gladi_events")
        return [
            {'type': row[1], 'date': row[2], 'time': row[3], 'description': row[4]}
            for row in cursor.fetchall()
        ]

    def save_trainees(self, trainees):
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM trainees")
        for t in trainees:
            cursor.execute(
                "INSERT INTO trainees (name, group_name, trainings, month, service, finished) VALUES (?, ?, ?, ?, ?, ?)",
                (t['name'], t['group'], t['trainings'], t['month'], t['service'], 1 if t['finished'] else 0)
            )
        self.conn.commit()

    def load_trainees(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM trainees")
        return [
            {'name': row[1], 'group': row[2], 'trainings': row[3], 'month': row[4],
             'service': row[5], 'finished': bool(row[6])}
            for row in cursor.fetchall()
        ]

    # --------------------------------------------------------
    # Service Report persistence
    # --------------------------------------------------------
    def save_service_report(self, report):
        cursor = self.conn.cursor()
        service = report['service']
        service_date = report['service_date']
        month = report['month']
        year = int(report['year'])

        # ------------------------------------------------------------
        # Service Report bersifat ADDITIVE untuk data VOLUNTEER.
        # Jika laporan pada ibadah + tanggal yang sama sudah ada, data
        # volunteer lama TIDAK dihapus. Jumlah dari input terbaru
        # dianggap sebagai tambahan volunteer dan dijumlahkan ke data lama.
        # Departemen baru juga otomatis ditambahkan ke dictionary lama.
        # ------------------------------------------------------------
        cursor.execute(
            "SELECT report_json FROM service_reports WHERE service = ? AND service_date = ?",
            (service, service_date)
        )
        existing_row = cursor.fetchone()

        if existing_row:
            try:
                existing_report = json.loads(existing_row[0])
            except (TypeError, json.JSONDecodeError):
                existing_report = {}

            old_volunteers = existing_report.get('volunteers', {}) or {}
            new_volunteers = report.get('volunteers', {}) or {}

            # Pada revisi, angka yang terlihat di form adalah TOTAL TERKINI
            # per departemen, bukan angka tambahan. Contoh: data lama S-Pro=4
            # lalu user mengubah form menjadi 5, hasil akhir harus 5 (bertambah 1),
            # bukan 9. Departemen lama yang tidak disentuh tetap dipertahankan.
            merged_volunteers = dict(old_volunteers)
            for dept, new_count in new_volunteers.items():
                try:
                    current_count = int(new_count or 0)
                except (TypeError, ValueError):
                    current_count = 0
                merged_volunteers[dept] = current_count

            # Hapus departemen yang secara eksplisit dikoreksi menjadi 0, tetapi
            # jangan menghapus departemen lama yang tidak dikirim oleh form.
            merged_volunteers = {
                dept: int(count or 0) for dept, count in merged_volunteers.items()
                if int(count or 0) != 0
            }

            # Total volunteer dihitung dari data final semua departemen.
            merged_total_volunteers = sum(
                int(v or 0) for v in merged_volunteers.values()
            )

            # Revisi Service Report harus memperbarui DATA TERKINI secara penuh.
            # Data lama tetap menjadi dasar, tetapi field yang dikirim dari form
            # (attendance, gembala, pembicara, fulltimer, speaker_name, prayer
            # corner, dan field laporan lainnya) harus mengikuti input terbaru.
            # Untuk volunteer, gunakan jumlah TERKINI per departemen hasil form,
            # bukan menjumlahkan ulang angka lama dengan angka baru.
            merged_report = dict(existing_report)
            merged_report.update(report)
            merged_report['volunteers'] = merged_volunteers
            merged_report['total_volunteers'] = merged_total_volunteers

            # Total kehadiran keseluruhan selalu dihitung ulang dari DATA TERKINI:
            # Kehadiran Jemaat Saja + Gembala + Pembicara + Fulltimer + Total Volunteer.
            merged_report['total_attendance'] = _service_report_total_attendance(merged_report)

            # Metadata mengikuti laporan yang sedang disimpan.
            merged_report['service'] = service
            merged_report['service_date'] = service_date
            merged_report['month'] = month
            merged_report['year'] = year
            report = merged_report

        payload = json.dumps(report, ensure_ascii=False)
        cursor.execute("""
            INSERT INTO service_reports
                (service, service_date, month, year, report_json)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(service, service_date) DO UPDATE SET
                month=excluded.month,
                year=excluded.year,
                report_json=excluded.report_json
        """, (service, service_date, month, year, payload))
        self.conn.commit()

    def load_service_reports(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT report_json FROM service_reports ORDER BY service_date, service")
        result = []
        for row in cursor.fetchall():
            try:
                result.append(json.loads(row[0]))
            except Exception:
                continue
        return result

    def delete_service_report(self, service, service_date):
        cursor = self.conn.cursor()
        cursor.execute(
            "DELETE FROM service_reports WHERE service = ? AND service_date = ?",
            (service, service_date)
        )
        self.conn.commit()

    def close(self):
        self.conn.close()


# ============================================================
# CORE / APPLICATION STATE
# ============================================================
class MinistryScheduler:
    def __init__(self):
        self.db = DatabaseManager()
        self.services = ['Voltage', 'Teens', 'Youth', 'Umum 1', 'Umum 2', 'Umum 3']
        self.service_to_group = {
            'Voltage': 'VOLTAGE', 'Teens': 'AOG', 'Youth': 'AOG',
            'Umum 1': 'UMUM', 'Umum 2': 'UMUM', 'Umum 3': 'UMUM'
        }
        self.umum_support_pool = ["Siska", "Johanes", "Yessi", "Donny", "Isel"]
        self.groups = {
            'VOLTAGE': {'members': ["Johanes", "Siska", "Calista"], 'pics': ["Johanes", "Siska"]},
            'AOG': {'members': ["Yessi", "Isel", "Donny", "Relis", "Nuel"], 'pics': ["Yessi", "Isel", "Donny"]},
            'UMUM': {'members': ["Pipik", "Andree", "Hartono", "Risma", "Ayu", "Ferry"], 'pics': ["Pipik", "Andree", "Hartono", "Risma"]}
        }
        self.all_names = set()
        for g_data in self.groups.values():
            self.all_names.update(g_data['members'])
            self.all_names.update(g_data['pics'])
        self.absences = []
        self.special_services = []
        self.manual_overrides = []
        self.gladi_events = []
        self.trainees = []
        self.schedule_data = {}
        self.service_reports = []
        self.gladi_event_counter = 0

        self.base_loads = {
            "Johanes": 4, "Andree": 4, "Pipik": 5, "Yessi": 5,
            "Isel": 5, "Donny": 4, "Siska": 4, "Hartono": 4,
            "Relis": 5, "Ferry": 4, "Nuel": 4, "Calista": 4,
            "Ayu": 4, "Risma": 4,
        }
        self.max_loads_4_weeks = dict(self.base_loads)
        self.max_loads_5_weeks = dict(self.base_loads)

        self.teams = {
            'VOLTAGE': [
                {'name': 'TIM 1', 'pic': 'Johanes', 'member': 'Calista'},
                {'name': 'TIM 2', 'pic': 'Siska', 'member': 'Calista'},
                {'name': 'TIM 3', 'pic': 'Johanes', 'member': 'Siska'},
            ],
            'AOG': [
                {'name': 'TIM 1', 'pic': 'Yessi', 'member': 'Nuel'},
                {'name': 'TIM 2', 'pic': 'Isel', 'member': 'Relis'},
                {'name': 'TIM 3', 'pic': 'Donny', 'member': 'Nuel'},
                {'name': 'TIM 4', 'pic': 'Yessi', 'member': 'Relis'},
                {'name': 'TIM 5', 'pic': 'Isel', 'member': 'Donny'},
            ],
            'UMUM': [
                {'name': 'TIM 1', 'pic': 'Pipik', 'member': 'Ayu'},
                {'name': 'TIM 2', 'pic': 'Andree', 'member': 'Ferry'},
                {'name': 'TIM 3', 'pic': 'Hartono', 'member': 'Risma'},
                {'name': 'TIM 4', 'pic': 'Siska', 'member': 'Donny'},
                {'name': 'TIM 5', 'pic': 'Pipik', 'member': 'Johanes'},
                {'name': 'TIM 6', 'pic': 'Isel', 'member': 'Yessi'},
            ],
        }
        self.month_combo = "September"
        self.year_spin = 2026
        self.weeks_combo = 4
        self.week_date_inputs = []
        self.temp_additional_people = []

    # --------------------------------------------------------
    # Date logic
    # --------------------------------------------------------
    def get_indonesian_date(self, d):
        month_map = {
            1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April',
            5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus',
            9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'
        }
        day_map = {
            0: 'Senin', 1: 'Selasa', 2: 'Rabu', 3: 'Kamis',
            4: 'Jumat', 5: 'Sabtu', 6: 'Minggu'
        }
        return f"{day_map[d.weekday()]}, {d.day} {month_map[d.month]} {d.year}"

    def parse_indonesian_date(self, date_str):
        try:
            if ',' in date_str:
                date_str = date_str.split(',', 1)[1].strip()
            month_map = {
                'Januari': 1, 'Februari': 2, 'Maret': 3, 'April': 4,
                'Mei': 5, 'Juni': 6, 'Juli': 7, 'Agustus': 8,
                'September': 9, 'Oktober': 10, 'November': 11, 'Desember': 12
            }
            parts = date_str.split()
            if len(parts) >= 3:
                day = int(parts[0])
                month = month_map.get(parts[1], 1)
                year = int(parts[2])
                return date(year, month, day)
        except Exception:
            pass
        return None

    def get_week_number_from_date(self, date_obj):
        if not self.week_date_inputs:
            return None
        for i, (sat_date, sun_date) in enumerate(self.week_date_inputs):
            week_end = sat_date + timedelta(days=6)
            if sat_date <= date_obj <= week_end:
                return i + 1
        return None

    def is_person_on_absence(self, person_name, date_str):
        date_obj = self.parse_indonesian_date(date_str)
        if date_obj is None:
            return False
        week_num = self.get_week_number_from_date(date_obj)
        if week_num is None:
            week_num = (date_obj.day - 1) // 7 + 1
            if week_num > 5:
                week_num = 5
        for absence in self.absences:
            if person_name == absence['name'] and week_num in absence.get('weeks', []):
                return True
            if person_name == absence['name'] and date_str in absence.get('dates', []):
                return True
        return False

    def sort_special_services(self):
        self.special_services.sort(key=lambda x: self.parse_indonesian_date(x['date']) or date.max)

    def find_duplicate(self, d, event):
        for i, ss in enumerate(self.special_services):
            if ss['date'] == d and ss['event'].lower() == event.lower():
                return i
        return -1

    def find_gladi_duplicate(self, d, description, time, event_type):
        for i, ev in enumerate(self.gladi_events):
            if (ev['date'] == d and ev['description'].lower() == description.lower()
                    and ev['time'] == time and ev['type'] == event_type):
                return i
        return -1

    def find_gladi_same_date_events(self, d, description, event_type):
        matches = []
        for i, ev in enumerate(self.gladi_events):
            if (ev['date'] == d and ev['description'].lower() == description.lower()
                    and ev['type'] == event_type):
                matches.append((i, ev))
        return matches

    def find_gladi_any_type_on_date(self, d):
        matches = []
        for i, ev in enumerate(self.gladi_events):
            if ev['date'] == d:
                matches.append((i, ev))
        return matches

    def find_all_duplicates(self):
        seen = {}
        duplicates = []
        for i, ss in enumerate(self.special_services):
            key = (ss['date'].lower(), ss['event'].lower())
            if key not in seen:
                seen[key] = []
            seen[key].append((i, ss))
        for key, items in seen.items():
            if len(items) > 1:
                duplicates.append(items)
        return duplicates

    def find_all_gladi_duplicates(self):
        seen = {}
        duplicates = []
        for i, ev in enumerate(self.gladi_events):
            key = (ev['date'].lower(), ev['description'].lower(), ev['type'].lower())
            if key not in seen:
                seen[key] = []
            seen[key].append((i, ev))
        for key, items in seen.items():
            if len(items) > 1:
                duplicates.append(items)
        return duplicates

    def remove_gladi_duplicates_by_indices(self, indices_to_remove):
        for idx in sorted(indices_to_remove, reverse=True):
            if 0 <= idx < len(self.gladi_events):
                del self.gladi_events[idx]
        self.db.save_gladi_events(self.gladi_events)

    def remove_duplicates_by_indices(self, indices_to_remove):
        for idx in sorted(indices_to_remove, reverse=True):
            if 0 <= idx < len(self.special_services):
                del self.special_services[idx]
        self.sort_special_services()

    def update_combo_boxes(self):
        self.all_names = set()
        for g_data in self.groups.values():
            self.all_names.update(g_data['members'])
            self.all_names.update(g_data['pics'])

    # --------------------------------------------------------
    # Data operations
    # --------------------------------------------------------
    def _unfinished_trainee_names(self):
        return {t['name'] for t in self.trainees if not t.get('finished', False)}

    def sync_trainees_into_groups(self):
        for trainee in self.trainees:
            name = trainee['name'].strip().title()
            group = trainee['group']
            if group in self.groups and name not in self.groups[group]['members']:
                self.groups[group]['members'].append(name)
            self.all_names.add(name)

    def add_new_member(self, name, group, trainings, month, service, finished):
        name = name.strip().title()
        if not name:
            return False, "Please enter a name."
        if group not in self.groups:
            return False, "Invalid group selected."
        if not finished:
            for t in self.trainees:
                if (t['name'] == name and t['group'] == group
                        and t['month'] == month and t['service'] == service
                        and not t.get('finished', False)):
                    return False, f"{name} is already registered as a trainee for {service} in {month}."
        if name in self.all_names and name not in self.groups[group]['members'] and not any(t['name'] == name for t in self.trainees):
            return False, f"{name} already exists in another group."
        if finished:
            if name not in self.groups[group]['members']:
                self.groups[group]['members'].append(name)
            self.all_names.add(name)
            self.max_loads_4_weeks[name] = 4
            self.max_loads_5_weeks[name] = 5
            return True, f"{name} added to {group} group."
        if name not in self.groups[group]['members']:
            self.groups[group]['members'].append(name)
        self.trainees.append({
            'name': name, 'group': group, 'trainings': int(trainings),
            'month': month, 'service': service, 'finished': False
        })
        self.all_names.add(name)
        return True, f"{name} added as trainee to {group} group."

    def delete_trainees(self, indices):
        if not indices:
            return 0
        selected = []
        for idx in sorted(set(indices), reverse=True):
            if 0 <= idx < len(self.trainees):
                selected.append(self.trainees[idx])
                del self.trainees[idx]
        for trainee in selected:
            name = trainee['name']
            group = trainee['group']
            still_registered = any(
                t['name'] == name and t['group'] == group
                for t in self.trainees
            )
            if not still_registered and group in self.groups:
                self.groups[group]['members'] = [
                    m for m in self.groups[group]['members'] if m != name
                ]
                self.groups[group]['pics'] = [
                    p for p in self.groups[group]['pics'] if p != name
                ]
        self.update_combo_boxes()
        self.db.save_trainees(self.trainees)
        self.db.save_setup(self.groups)
        return len(selected)

    def delete_members(self, selections):
        if not selections:
            return 0
        deleted_count = 0
        for selection in selections:
            try:
                group_name, member_name = selection.split("|||", 1)
            except ValueError:
                continue
            member_name = member_name.strip().title()
            if group_name not in self.groups or not member_name:
                continue
            members = self.groups[group_name].get('members', [])
            if member_name in members:
                self.groups[group_name]['members'] = [
                    m for m in members if m != member_name
                ]
                self.groups[group_name]['pics'] = [
                    p for p in self.groups[group_name].get('pics', [])
                    if p != member_name
                ]
                deleted_count += 1
        self.all_names = set()
        for data in self.groups.values():
            self.all_names.update(data.get('members', []))
        for trainee in self.trainees:
            self.all_names.add(trainee['name'])
        self.update_combo_boxes()
        self.db.save_setup(self.groups)
        return deleted_count

    def save_setup(self, groups):
        for g_name, data in groups.items():
            mems = [n.strip().title() for n in data['members'] if n.strip()]
            pics = [n.strip().title() for n in data['pics'] if n.strip()]
            for trainee in self.trainees:
                if trainee['group'] == g_name and not trainee.get('finished', False):
                    if trainee['name'] not in mems:
                        mems.append(trainee['name'])
                    pics = [p for p in pics if p != trainee['name']]
            self.groups[g_name] = {'members': mems, 'pics': pics}
        self.sync_trainees_into_groups()
        self.update_combo_boxes()
        self.db.save_setup(self.groups)
        self.db.save_trainees(self.trainees)

    def add_absence(self, name, selected_weeks, replacement_mode, replacement, replacements_multiple):
        try:
            if not name:
                return False, "Please select an absent person."
            if not selected_weeks:
                return False, "Please select at least one week."
            for absence in self.absences:
                if absence['name'] == name:
                    overlap = [w for w in selected_weeks if w in absence['weeks']]
                    if overlap:
                        return False, f"{name} is already marked absent for Week(s): {overlap}. Please select different weeks or delete the existing record."
            replacements = []
            if len(selected_weeks) > 1 and replacement_mode == "Multiple replacement persons":
                replacements = list(replacements_multiple)
                if not replacements:
                    return False, "Please select at least one replacement."
                if name in replacements:
                    return False, "A person cannot be their own replacement!"
            else:
                if not replacement:
                    return False, "Please select a replacement."
                if replacement == name:
                    return False, "A person cannot be their own replacement!"
                replacements = [replacement]
            self.absences.append({'name': name, 'weeks': selected_weeks, 'replacements': replacements})
            self.db.save_absences(self.absences)
            return True, "Absence record added successfully."
        except Exception as e:
            return False, f"Error in Add Absence: {e}"

    def add_date_absence(self, name, selected_dates, replacement):
        if not name or not selected_dates:
            return False, "Pilih nama dan minimal satu tanggal izin."
        if not replacement or replacement == name:
            return False, "Pilih pengganti yang berbeda dari nama yang izin."
        try:
            date_strings = [
                self.get_indonesian_date(d)
                for d in selected_dates
                if isinstance(d, date)
            ]
            if not date_strings:
                return False, "Pilih minimal satu tanggal izin yang valid."
            for rec in self.absences:
                if rec.get('name') == name and set(date_strings).intersection(rec.get('dates', [])):
                    return False, f"{name} sudah memiliki izin pada salah satu tanggal tersebut."
            self.absences.append({'name': name, 'weeks': [], 'dates': date_strings, 'replacements': [replacement]})
            self.db.save_absences(self.absences)
            return True, "Izin tanggal tertentu berhasil ditambahkan."
        except Exception as e:
            return False, f"Gagal menyimpan izin tanggal: {e}"

    def add_gladi_event(self, event_type, d, time_text, desc, conflict_action=None):
        if not d or not time_text or not desc:
            return 'error', "Please fill all fields."
        same_date_events = self.find_gladi_any_type_on_date(d)
        if same_date_events:
            same_type_events = [ev for idx, ev in same_date_events if ev['type'] == event_type and ev['description'].lower() == desc.lower()]
            different_type_events = [ev for idx, ev in same_date_events if ev['type'] != event_type]
            if same_type_events:
                existing = same_type_events[0]
                if conflict_action == 'cancel' or conflict_action is None:
                    return 'duplicate', {'kind': 'same', 'existing': existing}
                if conflict_action == 'update':
                    dup_idx = self.gladi_events.index(existing)
                    self.gladi_events[dup_idx] = {
                        'id': self.gladi_event_counter, 'type': event_type,
                        'date': d, 'time': time_text, 'description': desc
                    }
                    self.gladi_event_counter += 1
                    self.db.save_gladi_events(self.gladi_events)
                    return 'success', "Existing Gladi event updated with new data!"
            elif different_type_events:
                existing = different_type_events[0]
                if conflict_action == 'replace':
                    dup_idx = self.gladi_events.index(existing)
                    self.gladi_events[dup_idx] = {
                        'id': self.gladi_event_counter, 'type': event_type,
                        'date': d, 'time': time_text, 'description': desc
                    }
                    self.gladi_event_counter += 1
                    self.db.save_gladi_events(self.gladi_events)
                    return 'success', f"Replaced {existing['type']} with {event_type} on {d}!"
                return 'conflict', {'kind': 'different', 'existing': existing,
                                    'new': {'type': event_type, 'date': d, 'time': time_text, 'description': desc}}
            else:
                if conflict_action == 'update':
                    pass
                elif conflict_action != 'proceed':
                    return 'date_warning', {'events': [ev for idx, ev in same_date_events],
                                            'new': {'type': event_type, 'date': d, 'time': time_text, 'description': desc}}
        self.gladi_event_counter += 1
        self.gladi_events.append({
            'id': self.gladi_event_counter, 'type': event_type,
            'date': d, 'time': time_text, 'description': desc
        })
        self.db.save_gladi_events(self.gladi_events)
        return 'success', f"{event_type} added successfully!"

    def add_special_service(self, d, event, pic, mem, additional_people, conflict_action=None):
        if not d or not event or not pic or not mem:
            return 'error', "Please fill all fields."
        if self.is_person_on_absence(pic, d):
            return 'error', f"{pic} has absence permission during the week of {d} and cannot serve as PIC!"
        if self.is_person_on_absence(mem, d):
            return 'error', f"{mem} has absence permission during the week of {d} and cannot serve as Member!"
        dup_idx = self.find_duplicate(d, event)
        filtered_additional = []
        warnings = []
        for person in additional_people:
            if person == pic:
                warnings.append(f"{person} is already assigned as PIC and cannot be added to additional people!")
            elif person == mem:
                warnings.append(f"{person} is already assigned as Member and cannot be added to additional people!")
            elif self.is_person_on_absence(person, d):
                warnings.append(f"{person} has absence permission during the week of {d} and cannot be added!")
            else:
                filtered_additional.append(person)
        additional_people = filtered_additional
        if additional_people and len(additional_people) != len(set(additional_people)):
            return 'error', "Duplicate names detected in additional people list!"
        if dup_idx >= 0:
            existing = self.special_services[dup_idx]
            if conflict_action == 'update':
                self.special_services[dup_idx] = {
                    'date': d, 'event': event, 'pic': pic,
                    'member': mem, 'additional_people': additional_people
                }
                self.sort_special_services()
                return 'success', "Existing event updated with new data!"
            return 'duplicate', {'existing': existing, 'warnings': warnings}
        self.special_services.append({
            'date': d, 'event': event, 'pic': pic,
            'member': mem, 'additional_people': additional_people
        })
        self.sort_special_services()
        msg = "Special service added!"
        if warnings:
            msg += " " + " ".join(warnings)
        return 'success', msg

    def add_override(self, week, svc, role, person, conflict_action=None):
        if not person:
            return 'error', "Please select a person."
        existing_idx = -1
        for idx, override in enumerate(self.manual_overrides):
            if override['week'] == week and override['service'] == svc and override['role'] == role:
                existing_idx = idx
                break
        if existing_idx >= 0:
            existing = self.manual_overrides[existing_idx]
            if conflict_action == 'update':
                self.manual_overrides[existing_idx]['person'] = person
                return 'success', f"Updated {svc} {role} for Week {week} to {person}!"
            return 'duplicate', existing
        for absence in self.absences:
            if person == absence['name'] and week in absence['weeks']:
                return 'error', f"{person} has absence permission during Week {week} and cannot be assigned!"
        self.manual_overrides.append({'week': week, 'service': svc, 'role': role, 'person': person})
        return 'success', "Manual assignment added."

    # --------------------------------------------------------
    # EXACT CORE SCHEDULING ALGORITHM FROM ORIGINAL APPLICATION
    # --------------------------------------------------------
    def generate_schedule(self):
        try:
            if not self.all_names:
                raise ValueError("Please complete Setup first.")
            num_weeks = 4 if self.weeks_combo == 4 else 5
            max_loads = self.max_loads_5_weeks if num_weeks == 5 else self.max_loads_4_weeks

            current_loads = defaultdict(int)
            for name in self.all_names:
                current_loads[name] = 0
            trainee_counts = {
                t['name']: 0 for t in self.trainees if not t.get('finished', False)
            }
            self.schedule_data = {
                w: {svc: {'pic': '', 'member': ''} for svc in self.services}
                for w in range(1, num_weeks + 1)
            }
            month_idx = [
                'January', 'February', 'March', 'April', 'May', 'June',
                'July', 'August', 'September', 'October', 'November', 'December'
            ].index(self.month_combo)
            calendar_rotation = self.year_spin * 12 + month_idx

            team_usage = {
                group: {idx: 0 for idx in range(len(self.teams.get(group, [])))}
                for group in self.teams
            }
            team_combo_history = {group: [] for group in GROUPS}

            def get_absence_map(week):
                absent_this_week = {}
                for abs_rec in self.absences:
                    if week in abs_rec.get('weeks', []):
                        week_index = abs_rec['weeks'].index(week)
                        repl_list = abs_rec.get('replacements', [])
                        if repl_list:
                            absent_this_week[abs_rec['name']] = repl_list[week_index % len(repl_list)]
                return absent_this_week

            def effective_person(person, absent_this_week):
                return absent_this_week.get(person, person)

            def team_actual_pair(team, absent_this_week):
                pic = effective_person(team.get('pic', ''), absent_this_week)
                mem = effective_person(team.get('member', ''), absent_this_week)
                return pic, mem

            def team_load_score(pic, mem, usage, preferred_order):
                pic_load = current_loads.get(pic, 0)
                mem_load = current_loads.get(mem, 0) if mem else 0
                pic_cap = max_loads.get(pic, num_weeks)
                mem_cap = max_loads.get(mem, num_weeks) if mem else num_weeks
                projected_pic = pic_load + 1
                projected_mem = mem_load + (1 if mem and mem != pic else 0)
                overload = max(0, projected_pic - pic_cap) + max(0, projected_mem - mem_cap)
                ratio_pic = projected_pic / max(pic_cap, 1)
                ratio_mem = projected_mem / max(mem_cap, 1) if mem else 0
                balance = max(ratio_pic, ratio_mem)
                total = projected_pic + projected_mem
                return (overload, balance, total, usage, preferred_order)

            for week in range(1, num_weeks + 1):
                absent_this_week = get_absence_map(week)
                used_team_this_week = {group: set() for group in GROUPS}
                weekly_team_plan = {}

                def build_week_team_plan(group_name, group_services):
                    auto_services = []
                    for s in group_services:
                        has_pic_override = any(
                            o['week'] == week and o['service'] == s and o['role'] == 'PIC'
                            for o in self.manual_overrides
                        )
                        has_mem_override = any(
                            o['week'] == week and o['service'] == s and o['role'] == 'Member'
                            for o in self.manual_overrides
                        )
                        if not has_pic_override and not has_mem_override:
                            auto_services.append(s)
                    if not auto_services:
                        return {}
                    valid_teams = []
                    for idx, team in enumerate(self.teams.get(group_name, [])):
                        raw_pic = str(team.get('pic', '')).strip()
                        raw_mem = str(team.get('member', '')).strip()
                        if not raw_pic or not raw_mem:
                            continue
                        pic, mem = team_actual_pair(team, absent_this_week)
                        if not pic or not mem or pic == mem:
                            continue
                        valid_teams.append((idx, team, pic, mem))
                    if not valid_teams:
                        return {}
                    plan_size = min(len(auto_services), len(valid_teams))
                    service_count = len(auto_services)
                    history = team_combo_history[group_name]
                    best = None

                    def evaluate_plan(chosen, remaining, pos):
                        nonlocal best
                        if pos == plan_size:
                            selected_indices = tuple(x[0] for x in chosen)
                            combo = frozenset(selected_indices)
                            repeated_combo = 1 if combo in history else 0
                            new_team_count = sum(
                                1 for idx in selected_indices
                                if team_usage[group_name].get(idx, 0) == 0
                            )
                            projected = defaultdict(int)
                            overload = 0
                            max_ratio = 0.0
                            total_ratio = 0.0
                            usage_total = 0
                            for idx, _, pic, mem in chosen:
                                projected[pic] += 1
                                if mem and mem != pic:
                                    projected[mem] += 1
                                usage_total += team_usage[group_name].get(idx, 0)
                            for person, add in projected.items():
                                cap = max_loads.get(person, num_weeks)
                                projected_load = current_loads.get(person, 0) + add
                                overload += max(0, projected_load - cap)
                                ratio = projected_load / max(cap, 1)
                                max_ratio = max(max_ratio, ratio)
                                total_ratio += ratio

                            team_count = max(len(self.teams.get(group_name, [])), 1)
                            rotation = tuple(
                                (idx - (calendar_rotation + week + service_pos)) % team_count
                                for service_pos, (idx, _, _, _) in enumerate(chosen)
                            )
                            score = (
                                repeated_combo,
                                -new_team_count,
                                rotation,
                                overload,
                                max_ratio,
                                total_ratio,
                                usage_total,
                                selected_indices,
                            )
                            if best is None or score < best[0]:
                                best = (score, dict(zip(auto_services[:plan_size], chosen)))
                            return
                        for item in remaining:
                            next_remaining = [r for r in remaining if r[0] != item[0]]
                            evaluate_plan(chosen + [item], next_remaining, pos + 1)

                    evaluate_plan([], valid_teams, 0)
                    if best is None:
                        return {}
                    selected_plan = {}
                    for svc_name, item in best[1].items():
                        selected_plan[svc_name] = item[0]
                    selected_combo = frozenset(selected_plan.values())
                    team_combo_history[group_name].append(selected_combo)
                    return selected_plan

                for svc_index, svc in enumerate(self.services):
                    group_name = self.service_to_group[svc]
                    group_data = self.groups[group_name]
                    unfinished_trainees = self._unfinished_trainee_names()
                    orig_pics = [
                        p for p in list(group_data.get('pics', []))
                        if p not in unfinished_trainees
                    ]
                    orig_mems = [
                        m for m in list(group_data.get('members', []))
                        if m not in unfinished_trainees
                    ]
                    if orig_pics:
                        _pic_rot = calendar_rotation % len(orig_pics)
                        orig_pics = orig_pics[_pic_rot:] + orig_pics[:_pic_rot]
                    if orig_mems:
                        _mem_rot = calendar_rotation % len(orig_mems)
                        orig_mems = orig_mems[_mem_rot:] + orig_mems[:_mem_rot]
                    override_pic = next(
                        (o['person'] for o in self.manual_overrides
                         if o['week'] == week and o['service'] == svc and o['role'] == 'PIC'),
                        None
                    )
                    override_mem = next(
                        (o['person'] for o in self.manual_overrides
                         if o['week'] == week and o['service'] == svc and o['role'] == 'Member'),
                        None
                    )

                    if override_pic or override_mem:
                        assigned_pic = override_pic or ''
                        assigned_mem = override_mem or ''
                        if override_pic and not override_mem:
                            compatible = []
                            for idx, team in enumerate(self.teams.get(group_name, [])):
                                if team.get('pic') == override_pic:
                                    pic, mem = team_actual_pair(team, absent_this_week)
                                    if pic == override_pic and mem and mem != override_pic:
                                        compatible.append((idx, team))
                            compatible.sort(
                                key=lambda x: team_load_score(
                                    x[1].get('pic', ''), x[1].get('member', ''),
                                    team_usage[group_name].get(x[0], 0), x[0]
                                )
                            )
                            if compatible:
                                assigned_mem = team_actual_pair(compatible[0][1], absent_this_week)[1]
                        elif override_mem and not override_pic:
                            compatible = []
                            for idx, team in enumerate(self.teams.get(group_name, [])):
                                if team.get('member') == override_mem:
                                    pic, mem = team_actual_pair(team, absent_this_week)
                                    if pic and mem == override_mem and pic != override_mem:
                                        compatible.append((idx, team))
                            compatible.sort(
                                key=lambda x: team_load_score(
                                    x[1].get('pic', ''), x[1].get('member', ''),
                                    team_usage[group_name].get(x[0], 0), x[0]
                                )
                            )
                            if compatible:
                                assigned_pic = team_actual_pair(compatible[0][1], absent_this_week)[0]
                        if not assigned_pic:
                            pic_candidates = []
                            for idx, p in enumerate(orig_pics):
                                actual = effective_person(p, absent_this_week)
                                if svc.startswith('Umum') and actual in self.umum_support_pool:
                                    continue
                                if actual not in {x[0] for x in pic_candidates}:
                                    pic_candidates.append((actual, idx))
                            pic_candidates.sort(
                                key=lambda x: (
                                    current_loads.get(x[0], 0) >= max_loads.get(x[0], num_weeks),
                                    current_loads.get(x[0], 0), x[1]
                                )
                            )
                            assigned_pic = pic_candidates[0][0] if pic_candidates else ''
                        if not assigned_mem:
                            mem_candidates = []
                            for idx, m in enumerate(orig_mems):
                                actual = effective_person(m, absent_this_week)
                                if actual != assigned_pic and actual not in {x[0] for x in mem_candidates}:
                                    mem_candidates.append((actual, idx))
                            mem_candidates.sort(
                                key=lambda x: (
                                    current_loads.get(x[0], 0) >= max_loads.get(x[0], num_weeks),
                                    current_loads.get(x[0], 0), x[1]
                                )
                            )
                            assigned_mem = mem_candidates[0][0] if mem_candidates else assigned_pic
                    else:
                        configured_teams = []
                        for idx, team in enumerate(self.teams.get(group_name, [])):
                            raw_pic = str(team.get('pic', '')).strip()
                            raw_mem = str(team.get('member', '')).strip()
                            if not raw_pic or not raw_mem:
                                continue
                            pic, mem = team_actual_pair(team, absent_this_week)
                            if not pic or not mem or pic == mem:
                                continue
                            configured_teams.append((idx, team, pic, mem))
                        available_team_count = len(configured_teams)
                        avoid_same_week = (
                            len(self.teams.get(group_name, [])) >=
                            sum(1 for s in self.services if self.service_to_group[s] == group_name)
                        )
                        if configured_teams:
                            if group_name not in weekly_team_plan:
                                group_services = [
                                    s for s in self.services
                                    if self.service_to_group[s] == group_name
                                ]
                                weekly_team_plan[group_name] = build_week_team_plan(
                                    group_name, group_services
                                )
                            planned_idx = weekly_team_plan[group_name].get(svc)
                            if planned_idx is not None:
                                planned = next(
                                    (x for x in configured_teams if x[0] == planned_idx),
                                    None
                                )
                                if planned is not None:
                                    selected_team_idx, _, assigned_pic, assigned_mem = planned
                                    team_usage[group_name][selected_team_idx] += 1
                                    used_team_this_week[group_name].add(selected_team_idx)
                                else:
                                    planned_idx = None
                            if planned_idx is None:
                                candidates = []
                                for idx, team, pic, mem in configured_teams:
                                    used_count = team_usage[group_name].get(idx, 0)
                                    used_this_week = idx in used_team_this_week[group_name]
                                    never_used = used_count == 0
                                    coverage_priority = 0 if never_used else 1
                                    same_week_priority = (1 if used_this_week and avoid_same_week else 0)
                                    load_score = team_load_score(
                                        pic, mem, used_count, idx
                                    )
                                    score = (
                                        coverage_priority,
                                        same_week_priority,
                                        (idx - (calendar_rotation + week + self.services.index(svc))) % max(len(self.teams.get(group_name, [])), 1),
                                        load_score[0],
                                        load_score[1],
                                        load_score[2],
                                        load_score[3],
                                    )
                                    candidates.append((score, idx, pic, mem))
                                candidates.sort(key=lambda x: x[0])
                                unused = [c for c in candidates if team_usage[group_name].get(c[1], 0) == 0]
                                if unused:
                                    chosen = unused[0]
                                else:
                                    chosen = candidates[0]
                                _, selected_team_idx, assigned_pic, assigned_mem = chosen
                                team_usage[group_name][selected_team_idx] += 1
                                used_team_this_week[group_name].add(selected_team_idx)
                        else:
                            pic_candidates = []
                            for idx, p in enumerate(orig_pics):
                                actual = effective_person(p, absent_this_week)
                                if svc.startswith('Umum') and actual in self.umum_support_pool:
                                    continue
                                if actual not in {x[0] for x in pic_candidates}:
                                    pic_candidates.append((actual, idx))
                            pic_candidates.sort(
                                key=lambda x: (
                                    current_loads.get(x[0], 0) >= max_loads.get(x[0], num_weeks),
                                    current_loads.get(x[0], 0), x[1]
                                )
                            )
                            assigned_pic = pic_candidates[0][0] if pic_candidates else ''
                            mem_candidates = []
                            for idx, m in enumerate(orig_mems):
                                actual = effective_person(m, absent_this_week)
                                if actual != assigned_pic and actual not in {x[0] for x in mem_candidates}:
                                    mem_candidates.append((actual, idx))
                            mem_candidates.sort(
                                key=lambda x: (
                                    current_loads.get(x[0], 0) >= max_loads.get(x[0], num_weeks),
                                    current_loads.get(x[0], 0), x[1]
                                )
                            )
                            assigned_mem = mem_candidates[0][0] if mem_candidates else assigned_pic

                    if not override_mem:
                        due_trainees = [
                            t for t in self.trainees
                            if (t.get('group') == group_name
                                and t.get('month') == self.month_combo
                                and t.get('service') == svc
                                and not t.get('finished', False)
                                and trainee_counts.get(t.get('name'), 0) < int(t.get('trainings', 0))
                                and t.get('name') != assigned_pic)
                        ]
                        if due_trainees:
                            due_trainees.sort(
                                key=lambda t: (
                                    trainee_counts.get(t['name'], 0),
                                    t['name'].casefold()
                                )
                            )
                            assigned_mem = due_trainees[0]['name']

                    if assigned_pic:
                        current_loads[assigned_pic] += 1
                    if assigned_mem and assigned_mem != assigned_pic:
                        current_loads[assigned_mem] += 1
                    if assigned_mem in trainee_counts:
                        trainee_counts[assigned_mem] += 1
                    if assigned_pic in trainee_counts:
                        trainee_counts[assigned_pic] += 1
                    self.schedule_data[week][svc] = {
                        'pic': assigned_pic,
                        'member': assigned_mem
                    }

            validation_errors = []
            group_slot_counts = {
                group: sum(
                    1 for w in range(1, num_weeks + 1)
                    for svc in self.services
                    if self.service_to_group[svc] == group
                )
                for group in GROUPS
            }
            for group in GROUPS:
                valid_teams = [
                    (idx, team) for idx, team in enumerate(self.teams.get(group, []))
                    if team.get('pic') and team.get('member')
                ]
                if len(valid_teams) <= group_slot_counts[group]:
                    missing = [
                        team.get('name', f'TIM {idx + 1}')
                        for idx, team in valid_teams
                        if team_usage.get(group, {}).get(idx, 0) == 0
                    ]
                    if missing:
                        validation_errors.append(
                            f"{group}: team belum terjadwal: {', '.join(missing)}"
                        )
            clean_automatic_run = not self.absences and not self.manual_overrides and not any(
                not t.get('finished', False) for t in self.trainees
            )
            if validation_errors and clean_automatic_run:
                raise ValueError(
                    "Automatic team coverage could not be satisfied. "
                    + " | ".join(validation_errors)
                )
            return True, None
        except Exception as e:
            return False, f"An error occurred:\n{str(e)}\n\nDetails:\n{traceback.format_exc()}"

    # --------------------------------------------------------
    # Excel export
    # --------------------------------------------------------
    def export_to_excel_bytes(self):
        if not self.schedule_data:
            raise ValueError("Please generate first.")
        if not self.week_date_inputs:
            raise ValueError("Please set week dates first.")
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Jadwal DM"
        header_font = XlFont(bold=True, size=16, color=Color("000000"))
        title_font = XlFont(bold=True, size=18)
        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        yellow_fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
        blue_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
        green_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
        cyan_fill = PatternFill(start_color="00FFFF", end_color="00FFFF", fill_type="solid")
        gold_fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
        ws.merge_cells('A1:G1')
        ws['A1'] = "JADWAL PELAYANAN DATA MINISTRY GMS SALATIGA"
        ws['A1'].font = title_font
        ws['A1'].alignment = center_align
        ws.merge_cells('A2:G2')
        ws['A2'] = f"BULAN {self.month_combo.upper()} {self.year_spin}"
        ws['A2'].font = header_font
        ws['A2'].alignment = center_align
        headers = [
            "IBADAH", "VOLTAGE\n(13.00)", "TEENS\n(16.00)", "YOUTH\n(18.30)",
            "UMUM 1\n(07.00)", "UMUM 2\n(09.30)", "UMUM 3\n(17.00)"
        ]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col, value=h)
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border
            cell.fill = yellow_fill
            
        current_row = 5
        
        # CRITICAL FIX: Use week_date_inputs as authoritative source for number of weeks
        num_weeks = len(self.week_date_inputs)
        
        for w in range(1, num_weeks + 1):
            # Skip if this week doesn't exist in schedule_data
            if w not in self.schedule_data:
                continue
                
            date_index = w - 1
            sat_date, sun_date = self.week_date_inputs[date_index]
                
            sat_date_str = self.get_indonesian_date(sat_date)
            sun_date_str = self.get_indonesian_date(sun_date)
            
            ws.cell(row=current_row, column=1, value="TANGGAL").border = thin_border
            ws.cell(row=current_row, column=1).fill = blue_fill
            ws.cell(row=current_row, column=1).font = XlFont(bold=True)
            for c in range(2, 5):
                ws.cell(row=current_row, column=c).border = thin_border
            ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=4)
            ws.cell(row=current_row, column=2, value=sat_date_str).border = thin_border
            ws.cell(row=current_row, column=2).fill = blue_fill
            for c in range(5, 8):
                ws.cell(row=current_row, column=c).border = thin_border
            ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=7)
            ws.cell(row=current_row, column=5, value=sun_date_str).border = thin_border
            ws.cell(row=current_row, column=5).fill = green_fill
            current_row += 1
            ws.cell(row=current_row, column=1, value="Volunteer").border = thin_border
            for col, svc in enumerate(self.services, 2):
                assignment = self.schedule_data[w][svc]
                pic = assignment.get('pic', '')
                mem = assignment.get('member', '')
                group_name = self.service_to_group.get(svc, '')
                matching_team = next(
                    (team for team in self.teams.get(group_name, [])
                     if team.get('pic', '') == pic and team.get('member', '') == mem),
                    None
                )
                if pic:
                    team_label = matching_team.get('name', '').strip() if matching_team else ''
                    people_lines = [pic] + ([mem] if mem else [])
                    val = "\n".join(([team_label] if team_label else []) + people_lines)
                else:
                    val = "UNFILLED"
                cell = ws.cell(row=current_row, column=col, value=val)
                cell.border = thin_border
                cell.alignment = center_align
                if not pic:
                    cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                    cell.font = XlFont(color="FFFFFF")
            current_row += 1
            ws.cell(row=current_row, column=1, value="Link Tally dan Seat Counter").border = thin_border
            for col in range(2, 8):
                ws.cell(row=current_row, column=col, value="").border = thin_border
            current_row += 2
            
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        ws.cell(row=current_row, column=1, value="IBADAH TENGAH MINGGU").font = header_font
        ws.cell(row=current_row, column=1).fill = cyan_fill
        ws.cell(row=current_row, column=1).alignment = center_align
        current_row += 1
        if self.special_services:
            for ss in self.special_services:
                ws.cell(row=current_row, column=1, value=ss['date']).border = thin_border
                ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=7)
                ws.cell(row=current_row, column=2, value=ss['event']).border = thin_border
                ws.cell(row=current_row, column=2).font = XlFont(bold=True)
                current_row += 1
                ws.cell(row=current_row, column=1, value="Volunteer").border = thin_border
                ws.cell(row=current_row, column=2, value=f"{ss['pic']} (PIC)").border = thin_border
                current_row += 1
                ws.cell(row=current_row, column=1, value="").border = thin_border
                ws.cell(row=current_row, column=2, value=ss['member']).border = thin_border
                current_row += 1
                additional = ss.get('additional_people', [])
                if additional:
                    ws.cell(row=current_row, column=1, value="").border = thin_border
                    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=7)
                    ws.cell(row=current_row, column=2, value=", ".join(additional)).border = thin_border
                    current_row += 1
                ws.cell(row=current_row, column=1, value="Link Tally dan Seat Counter").border = thin_border
                ws.cell(row=current_row, column=2, value="").border = thin_border
                current_row += 2
        if self.gladi_events:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
            ws.cell(row=current_row, column=1, value="JADWAL GLADI BERSIH & KOTOR").font = header_font
            ws.cell(row=current_row, column=1).fill = gold_fill
            ws.cell(row=current_row, column=1).alignment = center_align
            current_row += 1
            grouped_events = defaultdict(list)
            for event in self.gladi_events:
                key = (event['date'], event['description'], event['type'])
                grouped_events[key].append(event['time'])
            for (d, desc, event_type), times in grouped_events.items():
                ws.cell(row=current_row, column=1, value=d).border = thin_border
                ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=7)
                if len(times) > 1:
                    times_str = ", ".join(sorted(times))
                    event_text = f"{desc} ({event_type}) - Times: {times_str}"
                else:
                    event_text = f"{desc} ({event_type}) - Time: {times[0]}"
                ws.cell(row=current_row, column=2, value=event_text).border = thin_border
                ws.cell(row=current_row, column=2).font = XlFont(bold=True, size=14)
                ws.cell(row=current_row, column=2).alignment = left_align
                current_row += 2
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        ws.cell(row=current_row, column=1, value="Keterangan :").font = header_font
        current_row += 1
        notes = [
            "1. Yang belum punya seragam pakai kemeja & celana hitam, pakai ID Card, wanita harap memakai makeup dan pria rambut rapi, pakai parfum dan HT",
            "2. Datang 1 jam sebelum ibadah, ikut briefing dan berdoa bersama, mengingatkan semua pelayanan utk Absensi di GMS Church",
            "4. Tugas PIC adalah menggunakan Tally Counter & Report semua kehadiran di Group DM, merekap laporan dari Welcome Lounge, Prayer Corner, dan Baptisan",
            "5. Setelah ibadah standby di frontdesk",
            "6. Semua ibadah umum dan AOG memakai seragam lengkap",
            "7. Ibadah tengah minggu menggunakan kaos DM, celana hitam, dan sepatu hitam"
        ]
        for note in notes:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
            ws.cell(row=current_row, column=1, value=note).alignment = left_align
            current_row += 1
        ws.column_dimensions['A'].width = 20
        for col in ['B', 'C', 'D', 'E', 'F', 'G']:
            ws.column_dimensions[col].width = 25
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer.getvalue()

    # --------------------------------------------------------
    # Service Report
    # --------------------------------------------------------
    def save_service_report(self, report):
        # DatabaseManager melakukan merge/additive pada volunteer untuk
        # service + tanggal yang sama. Ambil kembali record hasil merge
        # supaya tampilan aplikasi memakai data final yang tersimpan.
        self.db.save_service_report(report)
        saved_reports = self.db.load_service_reports()
        self.service_reports = saved_reports
        self.service_reports.sort(key=lambda r: (r.get('service_date', ''), r.get('service', '')))

    def delete_service_report(self, service, service_date):
        self.db.delete_service_report(service, service_date)
        self.service_reports = [
            r for r in self.service_reports
            if not (r.get('service') == service and r.get('service_date') == service_date)
        ]

    def get_service_reports_for_month(self, month_name, year):
        return [
            r for r in self.service_reports
            if r.get('month') == month_name and int(r.get('year', 0)) == int(year)
        ]

    # --------------------------------------------------------
    # Save / recall
    # --------------------------------------------------------
    def save_current_state_to_db(self):
        try:
            self.db.save_setup(self.groups)
            self.db.save_trainees(self.trainees)
            self.db.save_absences(self.absences)
            self.db.save_gladi_events(self.gladi_events)
            week_dates = [
                (self.get_indonesian_date(sat), self.get_indonesian_date(sun))
                for sat, sun in self.week_date_inputs
            ]
            self.db.save_schedule(
                self.month_combo, self.year_spin,
                self.weeks_combo, week_dates,
                self.schedule_data, self.special_services
            )
            return True, None
        except Exception as e:
            return False, f"Could not save some data: {str(e)}"

    def recall_last_data(self):
        groups = self.db.load_setup()
        if groups:
            self.groups = groups
        self.update_combo_boxes()
        absences = self.db.load_absences()
        if absences:
            self.absences = absences
        gladi_events = self.db.load_gladi_events()
        if gladi_events:
            self.gladi_events = gladi_events
        trainees = self.db.load_trainees()
        if trainees:
            self.trainees = trainees
            self.sync_trainees_into_groups()
        self.service_reports = self.db.load_service_reports()
        schedule_data = self.db.load_schedule()
        if schedule_data:
            self.month_combo = schedule_data['month']
            self.year_spin = schedule_data['year']
            self.weeks_combo = schedule_data['num_weeks']
            self.special_services = schedule_data.get('special_services', [])
            self.sort_special_services()
            self.schedule_data = schedule_data.get('schedule', {})
            restored = []
            for pair in schedule_data.get('week_dates', []):
                if isinstance(pair, (list, tuple)) and len(pair) == 2:
                    sat = self.parse_indonesian_date(pair[0])
                    sun = self.parse_indonesian_date(pair[1])
                    if sat and sun:
                        restored.append((sat, sun))
            if restored:
                self.week_date_inputs = restored
        return True


# ============================================================
# SESSION STATE
# ============================================================
def init_state():
    if 'scheduler' not in st.session_state:
        st.session_state.scheduler = MinistryScheduler()
        # Service Report is loaded automatically because each report is persisted
        # in SQLite independently from the existing manual Recall workflow.
        st.session_state.scheduler.service_reports = st.session_state.scheduler.db.load_service_reports()
        st.session_state.active_tab = 0
        st.session_state.setup_members = {
            g: ", ".join(v['members']) for g, v in st.session_state.scheduler.groups.items()
        }
        st.session_state.setup_pics = {
            g: ", ".join(v['pics']) for g, v in st.session_state.scheduler.groups.items()
        }
        st.session_state.setup_widget_version = 0
        st.session_state.abs_selected_weeks = []
        st.session_state.abs_replacement_mode = "Same replacement person"
        st.session_state.abs_multiple_replacements = []
        st.session_state.gladi_pending = None
        st.session_state.ss_pending = None
        st.session_state.override_pending = None
        st.session_state.additional_people = []
        st.session_state.schedule_generated = False
        st.session_state.service_report_export_requested = False
        st.session_state.service_report_confirmed = False
        st.session_state.recall_done = False
        st.session_state.exit_requested = False
        st.session_state.logged_out = False
    return st.session_state.scheduler


scheduler = init_state()

# ============================================================
# EXIT / LOGOUT SCREEN
# ============================================================
if st.session_state.get("logged_out", False):
    st.markdown("""
<div style="
    max-width: 620px;
    margin: 90px auto 20px auto;
    padding: 42px 30px;
    text-align: center;
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 18px;
    box-shadow: 0 8px 30px rgba(0,0,0,0.08);
">
    <div style="font-size: 52px; margin-bottom: 12px;">✅</div>
    <h1 style="margin-bottom: 10px;">Anda telah keluar</h1>
    <p style="font-size: 17px; color: #6b7280; margin-bottom: 0;">
        Sesi aplikasi Data Ministry Scheduler telah dihentikan.
    </p>
    <p style="font-size: 15px; color: #9ca3af; margin-top: 8px;">
        Anda dapat menutup tab browser ini secara manual.
    </p>
</div>
""", unsafe_allow_html=True)
    if st.button("🔄 Kembali ke Aplikasi", type="primary", use_container_width=True):
        st.session_state.logged_out = False
        st.session_state.exit_requested = False
        st.rerun()
    st.stop()

# ============================================================
# HELPERS
# ============================================================
MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]
GROUPS = ["VOLTAGE", "AOG", "UMUM"]
SERVICES = scheduler.services
NAMES = sorted(scheduler.all_names)


def adjust_to_weekday(d, target_weekday):
    return d + timedelta(days=(target_weekday - d.weekday()) % 7)


def calculate_num_weeks(month_name, year):
    month = MONTHS.index(month_name) + 1
    saturdays = sum(
        1 for day_num in range(1, calendar.monthrange(year, month)[1] + 1)
        if date(year, month, day_num).weekday() == 5
    )
    if saturdays == 5:
        return 5
    return 4


def ensure_week_dates():
    num_weeks = scheduler.weeks_combo
    current = date.today()

    # Keep existing dates when the number of weeks is reduced (e.g. 5 -> 4).
    # This prevents stale Week 5 data from causing an index error.
    if scheduler.week_date_inputs:
        if len(scheduler.week_date_inputs) >= num_weeks:
            scheduler.week_date_inputs = scheduler.week_date_inputs[:num_weeks]
            return

        # If more weeks are needed, extend the existing list instead of
        # replacing the dates that the user has already selected.
        last_sat, last_sun = scheduler.week_date_inputs[-1]
        for _ in range(num_weeks - len(scheduler.week_date_inputs)):
            last_sat += timedelta(days=7)
            last_sun += timedelta(days=7)
            scheduler.week_date_inputs.append((last_sat, last_sun))
        return

    first_sat = adjust_to_weekday(current, 5)
    scheduler.week_date_inputs = [
        (first_sat + timedelta(days=7 * i), first_sat + timedelta(days=7 * i + 1))
        for i in range(num_weeks)
    ]


def rerun():
    st.rerun()


def show_schedule_html():
    if not scheduler.schedule_data:
        st.info("Belum ada jadwal. Klik Generate Schedule Preview terlebih dahulu.")
        return
    
    # CRITICAL FIX: Use week_date_inputs as authoritative source for number of weeks
    if not scheduler.week_date_inputs:
        st.info("Belum ada data tanggal minggu. Silakan atur di tab Regular Service.")
        return
    
    rows = []
    headers = [
        "IBADAH", "VOLTAGE (13.00)", "TEENS (16.00)", "YOUTH (18.30)",
        "UMUM 1 (07.00)", "UMUM 2 (09.30)", "UMUM 3 (17.00)"
    ]
    rows.append("<tr>" + "".join(f"<th>{h}</th>" for h in headers) + "</tr>")
    
    # Use only weeks that have both schedule data and corresponding dates.
    # This protects against stale schedule data when switching from 5 weeks to 4.
    schedule_weeks = [w for w in scheduler.schedule_data if isinstance(w, int)]
    num_weeks = min(len(scheduler.week_date_inputs), max(schedule_weeks, default=0))
    
    for w in range(1, num_weeks + 1):
        if w not in scheduler.schedule_data:
            continue
        
        date_index = w - 1
        if date_index >= len(scheduler.week_date_inputs):
            continue
        sat, sun = scheduler.week_date_inputs[date_index]
                
        rows.append(
            f"<tr><td><b>TANGGAL</b></td>"
            f"<td colspan='3' class='schedule-date-sat'>{scheduler.get_indonesian_date(sat)}</td>"
            f"<td colspan='3' class='schedule-date-sun'>{scheduler.get_indonesian_date(sun)}</td></tr>"
        )
        volunteer_cells = []
        for svc in SERVICES:
            data = scheduler.schedule_data[w][svc]
            pic = data.get('pic', '')
            mem = data.get('member', '')
            group_name = scheduler.service_to_group.get(svc, '')
            matching_team = next(
                (team for team in scheduler.teams.get(group_name, [])
                 if team.get('pic', '') == pic and team.get('member', '') == mem),
                None
            )
            if pic:
                team_name = matching_team.get('name', '').strip() if matching_team else ''
                lines = ([team_name] if team_name else []) + [pic] + ([mem] if mem else [])
                cell_text = "<br>".join(lines)
                volunteer_cells.append(f"<td>{cell_text}</td>")
            else:
                volunteer_cells.append("<td class='schedule-unfilled'>UNFILLED</td>")
        rows.append("<tr><td><b>Volunteer</b></td>" + "".join(volunteer_cells) + "</tr>")
        rows.append("<tr><td><b>Link Tally dan Seat Counter</b></td><td colspan='6'></td></tr>")
        rows.append("<tr><td colspan='7' style='height:8px;border:none'></td></tr>")
        
    rows.append("<tr><td colspan='7' class='schedule-midweek'>IBADAH TENGAH MINGGU</td></tr>")
    for ss in scheduler.special_services:
        rows.append(f"<tr><td>{ss['date']}</td><td colspan='6'><b>{ss['event']}</b></td></tr>")
        rows.append(f"<tr><td><b>Volunteer</b></td><td colspan='6'>{ss['pic']} (PIC)</td></tr>")
        rows.append(f"<tr><td></td><td colspan='6'>{ss['member']}</td></tr>")
        additional = ss.get('additional_people', [])
        if additional:
            rows.append(f"<tr><td></td><td colspan='6'>{', '.join(additional)}</td></tr>")
        rows.append("<tr><td><b>Link Tally dan Seat Counter</b></td><td colspan='6'></td></tr>")
        rows.append("<tr><td colspan='7' style='height:8px;border:none'></td></tr>")
        
    if scheduler.gladi_events:
        rows.append("<tr><td colspan='7' class='schedule-gladi'>JADWAL GLADI BERSIH & KOTOR</td></tr>")
        grouped = defaultdict(list)
        for event in scheduler.gladi_events:
            grouped[(event['date'], event['description'], event['type'])].append(event['time'])
        for (d, desc, event_type), times in grouped.items():
            times_str = ", ".join(sorted(times)) if len(times) > 1 else times[0]
            label = "Times" if len(times) > 1 else "Time"
            rows.append(f"<tr><td>{d}</td><td colspan='6'><b>{desc} ({event_type}) - {label}: {times_str}</b></td></tr>")
            rows.append("<tr><td colspan='7' style='height:8px;border:none'></td></tr>")
            
    notes = [
        "1. Yang belum punya seragam pakai kemeja & celana hitam, pakai ID Card, wanita harap memakai makeup dan pria rambut rapi, pakai parfum dan HT",
        "2. Datang 1 jam sebelum ibadah, ikut briefing dan berdoa bersama, mengingatkan semua pelayanan utk Absensi di GMS Church",
        "4. Tugas PIC adalah menggunakan Tally Counter & Report semua kehadiran di Group DM, merekap laporan dari Welcome Lounge, Prayer Corner, dan Baptisan",
        "5. Setelah ibadah standby di frontdesk",
        "6. Semua ibadah umum dan AOG memakai seragam lengkap",
        "7. Ibadah tengah minggu menggunakan kaos DM, celana hitam, dan sepatu hitam"
    ]
    rows.append("<tr><td colspan='7'><b>Keterangan :</b></td></tr>")
    for note in notes:
        rows.append(f"<tr><td colspan='7' class='schedule-note'>{note}</td></tr>")
        
    html = "<table class='schedule-table'>" + "".join(rows) + "</table>"
    st.markdown(html, unsafe_allow_html=True)


def absence_table():
    if not scheduler.absences:
        st.info("Belum ada data absence permission.")
        return
    import pandas as pd
    df = pd.DataFrame([
        {
            "Name": a['name'],
            "Month": scheduler.month_combo,
            "Year": scheduler.year_spin,
            "Weeks Absent": str(a['weeks']),
            "Duration": f"{len(a['weeks'])} Week{'s' if len(a['weeks']) != 1 else ''}",
            "Replacements": ", ".join(a['replacements'])
        }
        for a in scheduler.absences
    ])
    st.dataframe(df, use_container_width=True, hide_index=True)


def gladi_table():
    if not scheduler.gladi_events:
        st.info("Belum ada Gladi Event.")
        return
    import pandas as pd
    df = pd.DataFrame([
        {"Type": e['type'], "Date": e['date'], "Description": e['description'], "Time": e['time']}
        for e in scheduler.gladi_events
    ])
    st.dataframe(df, use_container_width=True, hide_index=True)


def special_table():
    if not scheduler.special_services:
        st.info("Belum ada Special Service.")
        return
    import pandas as pd
    df = pd.DataFrame([
        {
            "Type": "Special Service", "Date": s['date'], "Event": s['event'],
            "PIC": f"{s['pic']} (PIC)", "Member": s['member'],
            "Additional People": ", ".join(s.get('additional_people', []))
        }
        for s in scheduler.special_services
    ])
    st.dataframe(df, use_container_width=True, hide_index=True)


def override_table():
    if not scheduler.manual_overrides:
        st.info("Belum ada manual assignment.")
        return
    import pandas as pd
    df = pd.DataFrame([
        {"Week": f"Week {o['week']}", "Service": o['service'], "Role": o['role'], "Person": o['person']}
        for o in scheduler.manual_overrides
    ])
    st.dataframe(df, use_container_width=True, hide_index=True)


# ============================================================
# HEADER
# ============================================================
st.markdown('<div class="app-title">GMS SALATIGA - DATA MINISTRY SCHEDULER</div>', unsafe_allow_html=True)
st.markdown('<div class="app-subtitle">Web / Tablet Edition — responsive replacement for the PyQt5 desktop interface</div>', unsafe_allow_html=True)

# ============================================================
# SIDEBAR NAVIGATION
# ============================================================
tab_names = [
    "1. Setup Group PIC",
    "2. Absence Permissions",
    "3. Regular Service",
    "4. Special Service",
    "5. Gladi Events",
    "6. Manual Assignment",
    "7. Service Report",
    "8. Generate & Export"
]
with st.sidebar:
    st.markdown("## ⛪ Data Ministry")
    st.caption("GMS Salatiga")
    selected_tab = st.radio("MENU", tab_names, index=st.session_state.active_tab)
    st.session_state.active_tab = tab_names.index(selected_tab)
    st.divider()
    if st.button("↩️ Recall Last Data", use_container_width=True):
        try:
            scheduler.recall_last_data()
            st.session_state.setup_members = {g: ", ".join(v['members']) for g, v in scheduler.groups.items()}
            st.session_state.setup_pics = {g: ", ".join(v['pics']) for g, v in scheduler.groups.items()}
            st.session_state.setup_widget_version = st.session_state.get('setup_widget_version', 0) + 1
            st.session_state.schedule_generated = bool(scheduler.schedule_data)
            st.success("Last data recalled!")
        except Exception as e:
            st.error(f"Recall error: {e}")
    if st.button("💾 Save Current State", use_container_width=True):
        ok, msg = scheduler.save_current_state_to_db()
        if ok:
            st.success("Data saved successfully.")
        else:
            st.warning(msg)
    if st.button(" Exit / Keluar", use_container_width=True):
        st.session_state.exit_requested = True
    if st.session_state.get('exit_requested', False):
        st.warning("Apakah Anda yakin ingin keluar dari aplikasi?")
        exit_col1, exit_col2 = st.columns(2)
        with exit_col1:
            if st.button("❌ Batal", use_container_width=True):
                st.session_state.exit_requested = False
                st.rerun()
        with exit_col2:
            if st.button("✅ Ya, Keluar", type="primary", use_container_width=True):
                st.session_state.exit_requested = False
                st.session_state.logged_out = True
                st.rerun()
    st.caption("Gunakan browser tablet. Semua perubahan widget tersimpan pada session saat halaman aktif.")

# ============================================================
# SERVICE REPORT HELPERS
# ============================================================
# Daftar seluruh departemen volunteer yang dapat dipilih sebagai tambahan.
# Daftar ini bersifat global agar setiap jenis ibadah dapat menambahkan departemen
# yang sewaktu-waktu muncul tanpa mengubah konfigurasi utama ibadah.
ALL_SERVICE_REPORT_VOLUNTEER_DEPARTMENTS = [
    "PAW", "Pendoa", "Multimedia", "Usher", "S-Pro", "Sosmed",
    "Sound", "DM", "CM", "Kakak EK Voltage", "MUA", "FLC", "Venue",
    "WHL", "Hospitality", "MRI"
]

SERVICE_REPORT_CONFIG = {
    "Voltage": {
        "attendance_fields": [
            ("adult", "Jemaat orang tua/dewasa"),
            ("children", "Jemaat anak"),
            ("child_volunteer", "Volunteer anak"),
            ("new_souls", "Jiwa Baru"),
            ("altar_call", "Altarcall"),
            ("baptism", "Baptisan"),
        ],
        "volunteer_fields": [
            "PAW", "Pendoa", "Multimedia", "Usher", "S-Pro", "Sosmed", "Sound", "DM", "Kakak EK Voltage"
        ],
        "volunteer_notes": "Volunteer anak dapat diberi keterangan nama pada catatan volunteer.",
    },
    "Teens": {
        "attendance_fields": [
            ("attendance", "Kehadiran jemaat saja"),
            ("new_souls", "Jiwa Baru"),
            ("altar_call", "Altarcall"),
            ("baptism", "Baptisan"),
        ],
        "volunteer_fields": [
            "PAW", "Pendoa", "Multimedia", "Usher", "S-Pro", "Sosmed", "Sound", "DM", "MUA", "FLC", "Venue", "WHL", "Hospitality", "MRI"
        ],
    },
    "Youth": {
        "attendance_fields": [
            ("attendance", "Kehadiran jemaat saja"),
            ("new_souls", "Jiwa Baru"),
            ("altar_call", "Altarcall"),
            ("baptism", "Baptisan"),
        ],
        "volunteer_fields": [
            "PAW", "Pendoa", "Multimedia", "Usher", "S-Pro", "Sosmed", "Sound", "DM", "MUA", "FLC", "WHL"
        ],
    },
    "Umum 1": {
        "attendance_fields": [
            ("attendance", "Kehadiran jemaat"),
            ("new_souls", "Jiwa Baru"),
            ("altar_call", "Altarcall"),
            ("baptism", "Baptisan"),
        ],
        "volunteer_fields": [
            "PAW", "Pendoa", "Multimedia", "Usher", "S-Pro", "Sosmed", "Sound", "DM", "Venue", "FLC", "WHL", "MRI", "Hospitality", "MUA"
        ],
    },
    "Umum 2": {
        "attendance_fields": [
            ("attendance", "Kehadiran jemaat"),
            ("new_souls", "Jiwa Baru"),
            ("altar_call", "Altarcall"),
            ("baptism", "Baptisan"),
        ],
        "volunteer_fields": [
            "PAW", "Pendoa", "Multimedia", "Usher", "S-Pro", "Sosmed", "Sound", "DM", "Venue", "FLC", "WHL", "MRI", "Hospitality", "MUA"
        ],
    },
    "Umum 3": {
        "attendance_fields": [
            ("attendance", "Kehadiran jemaat"),
            ("new_souls", "Jiwa Baru"),
            ("altar_call", "Altarcall"),
            ("baptism", "Baptisan"),
        ],
        "volunteer_fields": [
            "PAW", "Pendoa", "Multimedia", "Usher", "S-Pro", "Sosmed", "Sound", "DM", "Venue", "FLC", "WHL", "MRI", "Hospitality", "MUA"
        ],
    },
}


def _service_report_number(label, key, value=0):
    return int(st.number_input(label, min_value=0, value=int(value or 0), step=1, key=key))


def _report_month_from_date(d):
    return MONTHS[d.month - 1]


def _format_report_date(iso_date):
    try:
        return scheduler.get_indonesian_date(date.fromisoformat(iso_date))
    except Exception:
        return iso_date


def _service_report_total_attendance(report):
    attendance = int(report.get('attendance_total', 0))
    pastor = int(report.get('pastor_count', 0))
    speaker = int(report.get('speaker_count', 0))
    fulltimer = int(report.get('fulltimer', 0))
    total_volunteers = int(report.get('total_volunteers', 0))
    # Total Kehadiran Keseluruhan =
    # Kehadiran Jemaat Saja + Gembala + Pembicara + Fulltimer + Total Volunteer.
    return attendance + pastor + speaker + fulltimer + total_volunteers


def _service_report_volunteer_total(volunteers):
    return sum(int(v or 0) for v in volunteers.values())


def _render_service_report_preview(report):
    st.markdown(f"### {report['service']} — {_format_report_date(report['service_date'])}")
    st.write(f"**Pembicara:** {report.get('speaker_name') or '-'}")
    st.write(
        f"**Kehadiran jemaat saja:** {report.get('attendance_total', 0)} | "
        f"**Jiwa Baru:** {report.get('new_souls', 0)} | "
        f"**Altarcall:** {report.get('altar_call', 0)} | "
        f"**Baptisan:** {report.get('baptism', 0)}"
    )
    st.write(
        f"**Gembala:** {report.get('pastor_count', 0)} | "
        f"**Pembicara:** {report.get('speaker_count', 0)} | "
        f"**Fulltimer:** {report.get('fulltimer', 0)} | "
        f"**Total Volunteers:** {report.get('total_volunteers', 0)} | "
        f"**Total Kehadiran Keseluruhan:** {_service_report_total_attendance(report)}"
    )
    if report.get('service') == 'Voltage':
        st.write(
            f"**Dewasa:** {report.get('adult', 0)} | **Anak:** {report.get('children', 0)} | "
            f"**Volunteer anak:** {report.get('child_volunteer', 0)}"
        )
    pc = report.get('prayer_corner', [])
    st.write(f"**Prayer Corner:** {sum(int(x.get('count', 0)) for x in pc)} orang")
    if pc:
        st.dataframe(
            [{"Nama": x.get('name', ''), "Jumlah": x.get('count', 0), "Gender": x.get('gender', '')} for x in pc],
            use_container_width=True,
            hide_index=True,
        )
    with st.expander("Rincian Volunteer", expanded=False):
        st.dataframe(
            [{"Volunteer": k, "Jumlah": v} for k, v in report.get('volunteers', {}).items()],
            use_container_width=True,
            hide_index=True,
        )


def _build_service_report_excel_bytes(reports, month_name, year):
    wb = Workbook()
    ws = wb.active
    ws.title = "Service Report"
    header_font = XlFont(bold=True, size=14)
    title_font = XlFont(bold=True, size=18)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    yellow = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    blue = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    green = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")

    ws.merge_cells('A1:D1')
    ws['A1'] = "SERVICE REPORT GMS SALATIGA"
    ws['A1'].font = title_font
    ws['A1'].alignment = center
    ws.merge_cells('A2:D2')
    ws['A2'] = f"REKAP SERVICE REPORT BULAN {month_name.upper()} {year}"
    ws['A2'].font = header_font
    ws['A2'].alignment = center
    current_row = 4

    ordered_services = ["Voltage", "Teens", "Youth", "Umum 1", "Umum 2", "Umum 3"]
    for service in ordered_services:
        service_reports = sorted(
            [r for r in reports if r.get('service') == service],
            key=lambda x: x.get('service_date', '')
        )
        if not service_reports:
            continue
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=4)
        c = ws.cell(current_row, 1, f"IBADAH {service.upper()}")
        c.font = header_font; c.fill = yellow; c.alignment = center
        current_row += 1
        for report in service_reports:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=4)
            c = ws.cell(current_row, 1, _format_report_date(report['service_date']))
            c.font = XlFont(bold=True); c.fill = blue; c.alignment = center
            current_row += 1
            rows = [
                ("Kehadiran Jemaat Saja", report.get('attendance_total', 0)),
            ]
            if service == 'Voltage':
                rows.extend([
                    ("Jemaat orang tua/dewasa", report.get('adult', 0)),
                    ("Jemaat anak", report.get('children', 0)),
                    ("Volunteer anak", report.get('child_volunteer', 0)),
                ])
            rows.extend([
                ("Jiwa Baru", report.get('new_souls', 0)),
                ("Altarcall", report.get('altar_call', 0)),
                ("Baptisan", report.get('baptism', 0)),
                ("Gembala", report.get('pastor_count', 0)),
                ("Pembicara", report.get('speaker_count', 0)),
                ("Fulltimer", report.get('fulltimer', 0)),
                ("Nama Pembicara", report.get('speaker_name') or '-'),
            ])
            for label, value in rows:
                ws.cell(current_row, 1, label).border = thin
                ws.cell(current_row, 1).alignment = left
                ws.cell(current_row, 2, value).border = thin
                ws.cell(current_row, 2).alignment = center
                ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=4)
                current_row += 1
            ws.cell(current_row, 1, "Volunteer").font = XlFont(bold=True); ws.cell(current_row, 1).fill = green
            ws.cell(current_row, 1).border = thin
            current_row += 1
            for name, value in report.get('volunteers', {}).items():
                ws.cell(current_row, 1, name).border = thin
                ws.cell(current_row, 2, int(value or 0)).border = thin
                ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=4)
                current_row += 1
            ws.cell(current_row, 1, "Total Volunteers").font = XlFont(bold=True); ws.cell(current_row, 1).border = thin
            ws.cell(current_row, 2, report.get('total_volunteers', 0)).font = XlFont(bold=True); ws.cell(current_row, 2).border = thin
            ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=4)
            current_row += 1
            ws.cell(current_row, 1, "Total Kehadiran Keseluruhan").font = XlFont(bold=True); ws.cell(current_row, 1).border = thin
            ws.cell(current_row, 2, _service_report_total_attendance(report)).font = XlFont(bold=True); ws.cell(current_row, 2).border = thin
            ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=4)
            current_row += 1
            ws.cell(current_row, 1, "Prayer Corner").font = XlFont(bold=True); ws.cell(current_row, 1).border = thin
            current_row += 1
            pc = report.get('prayer_corner', [])
            if pc:
                for person in pc:
                    ws.cell(current_row, 1, person.get('name', '')).border = thin
                    ws.cell(current_row, 2, int(person.get('count', 0))).border = thin
                    ws.cell(current_row, 3, person.get('gender', '')).border = thin
                    ws.cell(current_row, 4, '').border = thin
                    current_row += 1
            else:
                ws.cell(current_row, 1, '-').border = thin
                current_row += 1
            current_row += 1

    for col, width in {'A': 30, 'B': 20, 'C': 18, 'D': 18}.items():
        ws.column_dimensions[col].width = width

    # Monthly recap sheet.
    recap = wb.create_sheet("Rekap Bulanan")
    recap.merge_cells('A1:G1')
    recap['A1'] = f"REKAP TOTAL SERVICE REPORT - {month_name.upper()} {year}"
    recap['A1'].font = title_font; recap['A1'].alignment = center
    headers = ["IBADAH", "JUMLAH ENTRY", "KEHADIRAN JEMAAT SAJA", "GEMBALA", "PEMBICARA", "FULLTIMER", "TOTAL KEHADIRAN", "JIWA BARU", "ALTARCALL", "BAPTISAN", "PRAYER CORNER"]
    for col, h in enumerate(headers, 1):
        cell = recap.cell(3, col, h); cell.font = header_font; cell.fill = yellow; cell.border = thin; cell.alignment = center
    r = 4
    grand = {k: 0 for k in ['attendance_total','pastor_count','speaker_count','fulltimer','total_attendance','new_souls','altar_call','baptism','prayer_corner']}
    for service in ordered_services:
        service_reports = [x for x in reports if x.get('service') == service]
        vals = {
            'attendance_total': sum(int(x.get('attendance_total',0)) for x in service_reports),
            'pastor_count': sum(int(x.get('pastor_count',0)) for x in service_reports),
            'speaker_count': sum(int(x.get('speaker_count',0)) for x in service_reports),
            'fulltimer': sum(int(x.get('fulltimer',0)) for x in service_reports),
            'total_attendance': sum(_service_report_total_attendance(x) for x in service_reports),
            'new_souls': sum(int(x.get('new_souls',0)) for x in service_reports),
            'altar_call': sum(int(x.get('altar_call',0)) for x in service_reports),
            'baptism': sum(int(x.get('baptism',0)) for x in service_reports),
            'prayer_corner': sum(sum(int(p.get('count',0)) for p in x.get('prayer_corner',[])) for x in service_reports),
        }
        for k in grand: grand[k] += vals[k]
        row = [service, len(service_reports), vals['attendance_total'], vals['pastor_count'], vals['speaker_count'], vals['fulltimer'], vals['total_attendance'], vals['new_souls'], vals['altar_call'], vals['baptism'], vals['prayer_corner']]
        for col, value in enumerate(row,1):
            cell=recap.cell(r,col,value); cell.border=thin; cell.alignment=center
        r += 1
    row = ["TOTAL BULANAN", sum(1 for _ in reports), grand['attendance_total'], grand['pastor_count'], grand['speaker_count'], grand['fulltimer'], grand['total_attendance'], grand['new_souls'], grand['altar_call'], grand['baptism'], grand['prayer_corner']]
    for col,value in enumerate(row,1):
        cell=recap.cell(r,col,value); cell.border=thin; cell.font=XlFont(bold=True); cell.alignment=center
    for col,width in {'A':22,'B':15,'C':24,'D':12,'E':14,'F':12,'G':18,'H':15,'I':15,'J':15,'K':18}.items(): recap.column_dimensions[col].width=width
    return_bytes = io.BytesIO()
    wb.save(return_bytes)
    return_bytes.seek(0)
    return return_bytes.getvalue()


def _service_report_form():
    st.subheader("Input Service Report")
    services = list(SERVICE_REPORT_CONFIG.keys())
    selected_service = st.selectbox("Pilih Ibadah", services, key="sr_service")
    service_date = st.date_input("Tanggal Ibadah", value=date.today(), key="sr_date")
    cfg = SERVICE_REPORT_CONFIG[selected_service]

    # Jika laporan untuk ibadah + tanggal yang sama sudah tersimpan,
    # tampilkan jumlah volunteer TERAKHIR sebagai nilai awal form. Dengan cara
    # ini revisi S-Pro dari 4 menjadi 5 berarti total S-Pro menjadi 5, sehingga
    # hanya ada penambahan 1 orang. Departemen lama lainnya tetap muncul.
    existing_report = next(
        (r for r in scheduler.service_reports
         if r.get('service') == selected_service
         and r.get('service_date') == service_date.isoformat()),
        None
    )
    existing_volunteers = (existing_report or {}).get('volunteers', {}) or {}

    c1, c2, c3 = st.columns(3)
    with c1:
        pastor_count = _service_report_number("Gembala", "sr_pastor")
    with c2:
        speaker_count = _service_report_number("Pembicara", "sr_speaker_count")
    with c3:
        fulltimer = _service_report_number("Fulltimer", "sr_fulltimer")
    speaker_name = st.text_input("Gembala/Pembicara — Nama Pembicara", key="sr_speaker_name")

    st.markdown("#### Kehadiran")
    attendance_values = {}
    cols = st.columns(3)
    for i, (field, label) in enumerate(cfg['attendance_fields']):
        with cols[i % 3]:
            attendance_values[field] = _service_report_number(label, f"sr_{field}")

    if selected_service == 'Voltage':
        attendance_total = (
            attendance_values.get('adult', 0) +
            attendance_values.get('children', 0) +
            attendance_values.get('child_volunteer', 0)
        )
        st.info(f"Kehadiran jemaat saja otomatis: {attendance_total} orang")
    else:
        attendance_total = attendance_values.get('attendance', 0)

    st.markdown("#### Volunteer")
    st.caption("Catatan revisi: angka volunteer pada form adalah jumlah TERKINI per departemen. Jika S-Pro sebelumnya 4 lalu diubah menjadi 5, total S-Pro menjadi 5 (bertambah 1). Departemen lama lainnya tetap tersimpan.")
    volunteers = {}
    vol_cols = st.columns(3)
    for i, role in enumerate(cfg['volunteer_fields']):
        with vol_cols[i % 3]:
            volunteers[role] = _service_report_number(
                role,
                f"sr_vol_{role}",
                existing_volunteers.get(role, 0)
            )

    # Tambahan departemen bersifat opsional. Departemen tambahan yang sudah
    # pernah tersimpan ikut dipilih kembali agar tidak hilang saat revisi.
    # Departemen baru tetap dapat dipilih dari daftar global.
    existing_additional_departments = [
        dept for dept, count in existing_volunteers.items()
        if dept not in cfg['volunteer_fields'] and int(count or 0) != 0
    ]
    available_additional_departments = list(dict.fromkeys(
        existing_additional_departments + [
            dept for dept in ALL_SERVICE_REPORT_VOLUNTEER_DEPARTMENTS
            if dept not in cfg['volunteer_fields']
        ]
    ))
    additional_departments = st.multiselect(
        "Tambah Departemen Volunteer (opsional)",
        options=available_additional_departments,
        default=existing_additional_departments,
        help="Pilih departemen tambahan apabila pada ibadah ini ada volunteer dari departemen yang tidak tercantum pada daftar bawaan.",
        key="sr_additional_departments"
    )

    if additional_departments:
        st.caption("Masukkan jumlah volunteer untuk departemen tambahan yang dipilih.")
        additional_cols = st.columns(3)
        for i, role in enumerate(additional_departments):
            with additional_cols[i % 3]:
                volunteers[role] = _service_report_number(
                    role,
                    f"sr_vol_additional_{role}",
                    existing_volunteers.get(role, 0)
                )

    total_volunteers = _service_report_volunteer_total(volunteers)
    st.info(f"Total Volunteers otomatis: {total_volunteers} orang")

    st.markdown("#### Prayer Corner")
    prayer_count = st.number_input("Jumlah baris Prayer Corner", min_value=0, max_value=50, value=0, step=1, key="sr_prayer_rows")
    prayer_corner = []
    for i in range(int(prayer_count)):
        c1, c2, c3 = st.columns([2, 1, 1])
        with c1:
            name = st.text_input("Nama", key=f"sr_pc_name_{i}")
        with c2:
            count = _service_report_number("Jumlah didoakan", f"sr_pc_count_{i}")
        with c3:
            gender = st.selectbox("Gender", ["Cewek", "Cowok"], key=f"sr_pc_gender_{i}")
        if name.strip() or count:
            prayer_corner.append({'name': name.strip(), 'count': count, 'gender': gender})

    # Total kehadiran keseluruhan = jemaat + gembala + pembicara + fulltimer + volunteer.
    total_attendance = attendance_total + pastor_count + speaker_count + fulltimer + total_volunteers
    st.success(
        f"Total Kehadiran Keseluruhan otomatis: {total_attendance} orang "
        f"(Jemaat {attendance_total} + Gembala {pastor_count} + Pembicara {speaker_count} + Fulltimer {fulltimer} + Volunteer {total_volunteers})"
    )

    if st.button("💾 Simpan Service Report", type="primary", use_container_width=True):
        report = {
            'service': selected_service,
            'service_date': service_date.isoformat(),
            'month': _report_month_from_date(service_date),
            'year': service_date.year,
            **attendance_values,
            'attendance_total': attendance_total,
            'pastor_count': pastor_count,
            'speaker_count': speaker_count,
            'fulltimer': fulltimer,
            'speaker_name': speaker_name.strip(),
            'volunteers': volunteers,
            'total_volunteers': total_volunteers,
            'total_attendance': total_attendance,
            'prayer_corner': prayer_corner,
        }
        scheduler.save_service_report(report)
        st.session_state.service_report_confirmed = False
        st.success(f"Service Report {selected_service} tanggal {scheduler.get_indonesian_date(service_date)} berhasil disimpan ke database.")
        st.rerun()


def _service_report_preview_month():
    month_name = st.selectbox("Bulan Rekap", MONTHS, index=MONTHS.index(scheduler.month_combo), key="sr_month_preview")
    year = st.number_input("Tahun", min_value=2000, max_value=2100, value=int(scheduler.year_spin), step=1, key="sr_year_preview")
    reports = scheduler.get_service_reports_for_month(month_name, int(year))
    expected_weeks = calculate_num_weeks(month_name, int(year))
    st.info(f"Kalender bulan {month_name} {int(year)} memiliki {expected_weeks} minggu ibadah (berdasarkan jumlah hari Sabtu).")
    if not reports:
        st.warning("Belum ada Service Report untuk bulan tersebut.")
        return

    st.markdown("### Preview sebelum export")
    for report in reports:
        _render_service_report_preview(report)
    st.markdown("### Rekap Total Bulanan")
    summary = []
    for service in SERVICE_REPORT_CONFIG:
        items = [x for x in reports if x.get('service') == service]
        summary.append({
            'Ibadah': service,
            'Entry': len(items),
            'Kehadiran Jemaat Saja': sum(int(x.get('attendance_total',0)) for x in items),
            'Gembala': sum(int(x.get('pastor_count',0)) for x in items),
            'Pembicara': sum(int(x.get('speaker_count',0)) for x in items),
            'Fulltimer': sum(int(x.get('fulltimer',0)) for x in items),
            'Total Kehadiran': sum(_service_report_total_attendance(x) for x in items),
            'Jiwa Baru': sum(int(x.get('new_souls',0)) for x in items),
            'Altarcall': sum(int(x.get('altar_call',0)) for x in items),
            'Baptisan': sum(int(x.get('baptism',0)) for x in items),
            'Prayer Corner': sum(sum(int(p.get('count',0)) for p in x.get('prayer_corner',[])) for x in items),
        })
    import pandas as pd
    st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)
    if st.button("📊 Export Rekap Service Report ke Excel", type="primary", use_container_width=True):
        st.session_state.service_report_export_requested = True
    if st.session_state.get('service_report_export_requested', False):
        st.warning("Pastikan preview di atas sudah benar. Setelah dikonfirmasi, file Excel akan dibuat.")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("❌ Belum Benar / Kembali ke Input", key="sr_cancel_export", use_container_width=True):
                st.session_state.service_report_export_requested = False
                st.session_state.sr_mode = "📝 Input Report"
                rerun()
        with c2:
            if st.button("✅ Ya, Data Sudah Benar", key="sr_confirm_export", type="primary", use_container_width=True):
                try:
                    data = _build_service_report_excel_bytes(reports, month_name, int(year))
                    filename = f"Service_Report_GMS_Salatiga_{month_name}_{int(year)}.xlsx"
                    st.download_button(
                        "⬇️ Download Excel Service Report",
                        data=data,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="sr_download_excel",
                    )
                    st.session_state.service_report_export_requested = False
                    st.success("Service Report berhasil dibuat dan siap di-download.")
                except Exception as e:
                    st.error(f"Export Service Report error: {e}")



# ============================================================
# TAB 1 - SETUP GROUP PIC
# ============================================================
if st.session_state.active_tab == 0:
    st.header("1. Setup Group PIC")
    with st.container(border=True):
        st.subheader("New Member Registration & Training")
        c1, c2 = st.columns(2)
        with c1:
            new_name = st.text_input("Name", key="new_name")
            new_group = st.selectbox("Group", GROUPS, key="new_group")
            new_trainings = st.number_input("Number of training (in a month)", min_value=1, max_value=5, value=1, step=1, key="new_trainings")
        with c2:
            new_month = st.selectbox("Month", MONTHS, key="new_month")
            new_service = st.selectbox("Service (for training)", SERVICES, key="new_service")
            new_finished = st.checkbox("Finished Training? (Assign normally in schedule)", key="new_finished")
        if st.button(" Add New Member", type="secondary", use_container_width=True):
            ok, msg = scheduler.add_new_member(new_name, new_group, new_trainings, new_month, new_service, new_finished)
            if ok:
                scheduler.sync_trainees_into_groups()
                st.session_state.setup_members = {
                    g: ", ".join(v['members']) for g, v in scheduler.groups.items()
                }
                st.session_state.setup_pics = {
                    g: ", ".join(v['pics']) for g, v in scheduler.groups.items()
                }
                st.session_state.setup_widget_version = st.session_state.get('setup_widget_version', 0) + 1
                st.success(msg)
                st.rerun()
            else:
                st.warning(msg)

    st.subheader("Group Configuration")
    for group_name in GROUPS:
        with st.container(border=True):
            st.markdown(f"### {group_name} GROUP")
            col1, col2 = st.columns(2)
            with col1:
                st.session_state.setup_members[group_name] = st.text_area(
                    "Members (Comma separated)",
                    value=st.session_state.setup_members[group_name],
                    height=130,
                    key=f"members_{group_name}_{st.session_state.setup_widget_version}"
                )
            with col2:
                st.session_state.setup_pics[group_name] = st.text_area(
                    "PIC Ibadah (Can be PIC)",
                    value=st.session_state.setup_pics[group_name],
                    height=130,
                    key=f"pics_{group_name}_{st.session_state.setup_widget_version}"
                )
            
            members_from_textarea = [x.strip().title() for x in st.session_state.setup_members[group_name].replace(',', '\n').split('\n') if x.strip()]
            pics_from_textarea = [x.strip().title() for x in st.session_state.setup_pics[group_name].replace(',', '\n').split('\n') if x.strip()]
            scheduler.groups[group_name] = {'members': members_from_textarea, 'pics': pics_from_textarea}
            
            group_members = scheduler.groups.get(group_name, {}).get('members', [])
            trainee_names_in_group = {
                t['name'] for t in scheduler.trainees
                if t.get('group') == group_name and not t.get('finished', False)
            }
            active_group_members = [
                name for name in group_members if name not in trainee_names_in_group
            ]
            if active_group_members:
                st.markdown("**Manage Members**")
                selected_group_members = st.multiselect(
                    "Select member(s) to delete",
                    active_group_members,
                    key=f"delete_members_{group_name}_{st.session_state.setup_widget_version}"
                )
                if st.button(
                    "🗑️ Delete Selected Member",
                    key=f"delete_member_btn_{group_name}",
                    use_container_width=True
                ):
                    if not selected_group_members:
                        st.warning("Please select at least one member to delete.")
                    else:
                        st.session_state.member_delete_requested = True
                        st.session_state.member_delete_group = group_name
                if (st.session_state.get('member_delete_requested', False) and
                        st.session_state.get('member_delete_group') == group_name):
                    st.warning(
                        "⚠️ Member yang dipilih akan dihapus dari daftar member dan PIC "
                        "grup ini. Jadwal/history yang sudah tersimpan tidak akan dihapus. "
                        "Apakah Anda yakin?"
                    )
                    confirm_col1, confirm_col2 = st.columns(2)
                    with confirm_col1:
                        if st.button(
                            "❌ Batal",
                            key=f"cancel_delete_member_{group_name}",
                            use_container_width=True
                        ):
                            st.session_state.member_delete_requested = False
                            st.session_state.member_delete_group = None
                            st.rerun()
                    with confirm_col2:
                        if st.button(
                            "✅ Ya, Hapus Member",
                            key=f"confirm_delete_member_{group_name}",
                            type="primary",
                            use_container_width=True
                        ):
                            selections = [
                                f"{group_name}|||{name}"
                                for name in selected_group_members
                            ]
                            deleted_count = scheduler.delete_members(selections)
                            st.session_state.member_delete_requested = False
                            st.session_state.member_delete_group = None
                            st.session_state.setup_members = {
                                g: ", ".join(v['members'])
                                for g, v in scheduler.groups.items()
                            }
                            st.session_state.setup_pics = {
                                g: ", ".join(v['pics'])
                                for g, v in scheduler.groups.items()
                            }
                            st.session_state.setup_widget_version = st.session_state.get(
                                'setup_widget_version', 0
                            ) + 1
                            st.success(
                                f"Successfully deleted {deleted_count} active member(s) from {group_name} GROUP!"
                            )
                            st.rerun()

    st.subheader("Team Configuration & Preview")
    st.caption("Nama tim tidak boleh duplikat di dalam kelompok. Pilih PIC dan member dari daftar anggota yang sudah ada.")
    for group_name in GROUPS:
        with st.expander(f"{group_name} — Tim", expanded=True):
            current = scheduler.teams.get(group_name, [])
            count = st.number_input(f"Jumlah tim {group_name}", min_value=1, max_value=20, value=max(1, len(current)), step=1, key=f"team_count_{group_name}")
            revised = []
            people = sorted({
                str(name).strip()
                for group_data in scheduler.groups.values()
                for name in (group_data.get('members', []) + group_data.get('pics', []))
                if str(name).strip()
            }, key=str.casefold)
            if not people:
                people = ['']
            for idx in range(int(count)):
                old = current[idx] if idx < len(current) else {'name': f'TIM {idx + 1}', 'pic': '', 'member': ''}
                c1, c2, c3 = st.columns([1, 2, 2])
                with c1:
                    team_name = st.text_input(f"Nama tim #{idx + 1}", value=old.get('name', f'TIM {idx + 1}'), key=f"team_name_{group_name}_{idx}")
                with c2:
                    pic = st.selectbox(f"PIC tim #{idx + 1}", people, index=people.index(old['pic']) if old.get('pic') in people else 0, key=f"team_pic_{group_name}_{idx}")
                with c3:
                    member = st.selectbox(f"Member tim #{idx + 1}", people, index=people.index(old['member']) if old.get('member') in people else 0, key=f"team_mem_{group_name}_{idx}")
                revised.append({'name': team_name.strip(), 'pic': pic, 'member': member})
            names = [t['name'].casefold() for t in revised]
            if any(not n for n in names) or len(names) != len(set(names)):
                st.error(f"Nama tim {group_name} kosong atau duplikat. Ubah sebelum menyimpan.")
            else:
                scheduler.teams[group_name] = revised
            import pandas as pd
            st.dataframe(pd.DataFrame([{t['name']: f"{t['pic']} (PIC)" if row == 0 else t['member'] for t in revised} for row in range(2)]), use_container_width=True, hide_index=True)

    if st.button(" Save Setup & Map Services", type="primary", use_container_width=True):
        groups = {}
        for g in GROUPS:
            members = [x.strip().title() for x in st.session_state.setup_members[g].replace(',', '\n').split('\n') if x.strip()]
            pics = [x.strip().title() for x in st.session_state.setup_pics[g].replace(',', '\n').split('\n') if x.strip()]
            groups[g] = {'members': members, 'pics': pics}
        scheduler.save_setup(groups)
        st.success("Setup saved!")

    if scheduler.trainees:
        st.subheader("Registered Trainees")
        import pandas as pd
        st.dataframe(pd.DataFrame(scheduler.trainees), use_container_width=True, hide_index=True)
        trainee_labels = [
            f"{i + 1}. {t['name']} — {t['group']} — {t['month']} — {t['service']} — {t['trainings']} training"
            for i, t in enumerate(scheduler.trainees)
        ]
        selected_trainees = st.multiselect(
            "Select trainee record(s) to DELETE",
            trainee_labels,
            key="trainee_delete"
        )
        if st.button("🗑️ Delete Selected Trainee", use_container_width=True):
            if not selected_trainees:
                st.warning("Please select at least one trainee record to delete.")
            else:
                indices = [trainee_labels.index(x) for x in selected_trainees]
                deleted_count = scheduler.delete_trainees(indices)
                st.session_state.setup_members = {
                    g: ", ".join(v['members']) for g, v in scheduler.groups.items()
                }
                st.session_state.setup_pics = {
                    g: ", ".join(v['pics']) for g, v in scheduler.groups.items()
                }
                st.session_state.setup_widget_version = st.session_state.get(
                    'setup_widget_version', 0
                ) + 1
                st.success(f"Successfully deleted {deleted_count} trainee record(s)!")
                rerun()

# ============================================================
# TAB 2 - ABSENCE
# ============================================================
elif st.session_state.active_tab == 1:
    st.header("2. Absence Permissions")
    c1, c2 = st.columns([2, 1])
    with c1:
        abs_month = st.selectbox("Select Month & Year to determine weeks", MONTHS,
                                 index=MONTHS.index(scheduler.month_combo), key="abs_month")
    with c2:
        abs_year = st.number_input("Year", min_value=2020, max_value=2030, value=scheduler.year_spin, step=1, key="abs_year")
    num_abs_weeks = calculate_num_weeks(abs_month, abs_year)
    st.caption(f"Detected {num_abs_weeks} church weeks for {abs_month} {abs_year} (based on Saturdays, matching the original logic).")
    selected_weeks = st.multiselect("Select Week(s) Absent", list(range(1, num_abs_weeks + 1)), key="abs_weeks")
    st.write(f"**Duration: {len(selected_weeks)} Week{'s' if len(selected_weeks) != 1 else ''}**")
    names = sorted(scheduler.all_names)
    abs_name = st.selectbox("Absent Person", names if names else [""], key="abs_name")
    replacement_options = [n for n in names if n != abs_name]
    replacement = st.selectbox("Replacement Person (if 1 week or single replacement)", replacement_options if replacement_options else [""], key="abs_replacement")
    if len(selected_weeks) > 1:
        replacement_mode = st.radio(
            "Replacement method",
            ["Same replacement person", "Multiple replacement persons"],
            horizontal=False,
            key="abs_replacement_mode"
        )
        if replacement_mode == "Multiple replacement persons":
            multiple_replacements = st.multiselect(
                "Replacement Person(s)", replacement_options,
                key="abs_multiple_replacements"
            )
        else:
            multiple_replacements = []
    else:
        replacement_mode = "Same replacement person"
        multiple_replacements = []
    st.markdown("### Izin pada tanggal tertentu")
    date_abs_name = st.selectbox("Nama yang izin pada tanggal", names if names else [""], key="date_abs_name")
    selected_dates = st.date_input("Tanggal izin (bisa satu atau beberapa tanggal)", value=[], key="date_abs_dates")
    date_repl_options = [n for n in names if n != date_abs_name]
    date_replacement = st.selectbox("Pengganti untuk izin tanggal", date_repl_options if date_repl_options else [""], key="date_abs_replacement")
    if st.button("➕ Tambah Izin Tanggal", key="add_date_abs", use_container_width=True):
        try:
            if isinstance(selected_dates, date):
                normalized_dates = [selected_dates]
            elif isinstance(selected_dates, (list, tuple)):
                normalized_dates = [d for d in selected_dates if isinstance(d, date)]
            else:
                normalized_dates = []
            ok, msg = scheduler.add_date_absence(
                date_abs_name, normalized_dates, date_replacement
            )
            if ok:
                st.success(msg)
            else:
                st.warning(msg)
        except Exception as e:
            st.error(f"Gagal menambahkan izin tanggal: {e}")
    c1, c2 = st.columns(2)
    with c1:
        add_abs = st.button("➕ Add Absence", type="primary", use_container_width=True)
    with c2:
        delete_abs = st.button("🗑️ Delete Selected", use_container_width=True)
    if add_abs:
        scheduler.month_combo = abs_month
        scheduler.year_spin = int(abs_year)
        ok, msg = scheduler.add_absence(abs_name, selected_weeks, replacement_mode, replacement, multiple_replacements)
        if ok:
            st.success(msg)
        else:
            st.warning(msg)
    absence_table()
    if delete_abs:
        if not scheduler.absences:
            st.warning("Please select at least one absence record to delete.")
        else:
            st.session_state.abs_delete_mode = True
    if st.session_state.get('abs_delete_mode', False) and scheduler.absences:
        labels = [f"{i + 1}. {a['name']} - Week(s) {a['weeks']}" for i, a in enumerate(scheduler.absences)]
        to_delete = st.multiselect("Select absence record(s) to DELETE", labels, key="abs_delete_select")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Confirm Delete Absence", type="primary", use_container_width=True):
                indices = [labels.index(x) for x in to_delete]
                for idx in sorted(indices, reverse=True):
                    del scheduler.absences[idx]
                st.session_state.abs_delete_mode = False
                st.success(f"Successfully deleted {len(indices)} absence record(s)!")
                scheduler.db.save_absences(scheduler.absences)
                rerun()
        with c2:
            if st.button("Cancel Delete", use_container_width=True):
                st.session_state.abs_delete_mode = False
                rerun()

# ============================================================
# TAB 3 - REGULAR SERVICE
# ============================================================
elif st.session_state.active_tab == 2:
    st.header("3. Regular Service")
    st.info("Pengaturan bulan, tahun, jumlah minggu dan tanggal Sabtu/Minggu dipertahankan dari tab Regular Service pada aplikasi asli.")
    c1, c2 = st.columns(2)
    with c1:
        month_index = MONTHS.index(scheduler.month_combo) if scheduler.month_combo in MONTHS else 8
        scheduler.month_combo = st.selectbox("Month", MONTHS, index=month_index, key="regular_month")
    with c2:
        scheduler.year_spin = st.number_input("Year", min_value=2020, max_value=2030, value=int(scheduler.year_spin), step=1, key="regular_year")
    detected = calculate_num_weeks(scheduler.month_combo, int(scheduler.year_spin))
    st.caption(f"Original application detects 5 weeks when there are 5 Saturdays; otherwise 4 weeks. Detected: **{detected} weeks**.")
    week_choice = st.selectbox("Number of Weeks", [4, 5], index=0 if scheduler.weeks_combo == 4 else 1, format_func=lambda x: f"{x} Weeks", key="regular_weeks")
    scheduler.weeks_combo = week_choice
    ensure_week_dates()
    st.subheader("Select Dates for Each Week")
    new_week_dates = []
    for i in range(scheduler.weeks_combo):
        sat_default, sun_default = scheduler.week_date_inputs[i]
        c1, c2, c3 = st.columns([1, 3, 3])
        with c1:
            st.markdown(f"### Week {i + 1}")
        with c2:
            sat = st.date_input(f"Tanggal Sabtu — Week {i + 1}", value=sat_default, key=f"sat_{i}")
        with c3:
            sun = st.date_input(f"Tanggal Minggu — Week {i + 1}", value=sun_default, key=f"sun_{i}")
        if sat.weekday() != 5:
            st.warning(f"Week {i + 1}: field Sabtu must be a Saturday. It will be adjusted for schedule logic.")
            sat = adjust_to_weekday(sat, 5)
        if sun.weekday() != 6:
            st.warning(f"Week {i + 1}: field Minggu must be a Sunday. It will be adjusted for schedule logic.")
            sun = adjust_to_weekday(sun, 6)
        new_week_dates.append((sat, sun))
    scheduler.week_date_inputs = new_week_dates

# ============================================================
# TAB 4 - SPECIAL SERVICE
# ============================================================
elif st.session_state.active_tab == 3:
    st.header("4. Special Service")
    names = sorted(scheduler.all_names)
    c1, c2 = st.columns(2)
    with c1:
        ss_date = st.date_input("Date", value=date.today(), key="ss_date")
        ss_event = st.text_input("Event Name", placeholder="e.g., Leaders Community", key="ss_event")
        ss_pic = st.selectbox("PIC", names if names else [""], key="ss_pic")
    with c2:
        ss_mem = st.selectbox("Member", names if names else [""], key="ss_mem")
        additional_count = st.number_input("Number of additional people (Optional)", min_value=0, max_value=10, value=0, step=1, key="ss_additional_count")
        if additional_count:
            st.session_state.additional_people = st.multiselect(
                "Additional People",
                [n for n in names if n not in {ss_pic, ss_mem}],
                max_selections=int(additional_count),
                key="ss_additional_people"
            )
        else:
            st.session_state.additional_people = []
    if st.button("➕ Add Special Service", type="primary", use_container_width=True):
        d = scheduler.get_indonesian_date(ss_date)
        result, payload = scheduler.add_special_service(d, ss_event.strip(), ss_pic, ss_mem, st.session_state.additional_people)
        if result == 'success':
            st.success(payload)
            scheduler.db.save_schedule(scheduler.month_combo, scheduler.year_spin, scheduler.weeks_combo,
                                       [(scheduler.get_indonesian_date(a), scheduler.get_indonesian_date(b)) for a, b in scheduler.week_date_inputs],
                                       scheduler.schedule_data, scheduler.special_services)
        elif result == 'duplicate':
            st.session_state.ss_pending = {
                'date': d, 'event': ss_event.strip(), 'pic': ss_pic, 'mem': ss_mem,
                'additional': list(st.session_state.additional_people), 'existing': payload['existing']
            }
        else:
            st.warning(payload)
    if st.session_state.get('ss_pending'):
        p = st.session_state.ss_pending
        st.warning(
            f"A similar event already exists.\n\n"
            f"Date: {p['existing']['date']}\n\nEvent: {p['existing']['event']}\n\n"
            f"PIC: {p['existing']['pic']}\n\nMember: {p['existing']['member']}"
        )
        c1, c2 = st.columns(2)
        with c1:
            if st.button("❌ Cancel", key="ss_cancel", use_container_width=True):
                st.session_state.ss_pending = None
                rerun()
        with c2:
            if st.button("🔄 Update Existing", key="ss_update", type="primary", use_container_width=True):
                scheduler.add_special_service(p['date'], p['event'], p['pic'], p['mem'], p['additional'], conflict_action='update')
                st.session_state.ss_pending = None
                st.success("Existing event updated with new data!")
                rerun()
    special_table()
    if scheduler.special_services:
        st.subheader("Delete Special Service")
        labels = [f"{i + 1}. {s['date']} — {s['event']} — {s['pic']} / {s['member']}" for i, s in enumerate(scheduler.special_services)]
        selected_delete = st.multiselect("Select event(s) to delete", labels, key="ss_delete")
        if st.button("🗑️ Delete Selected Special Service", use_container_width=True):
            indices = [labels.index(x) for x in selected_delete]
            for idx in sorted(indices, reverse=True):
                del scheduler.special_services[idx]
            scheduler.sort_special_services()
            st.success(f"Successfully deleted {len(indices)} Special Service event(s)!")
            rerun()
    duplicates = scheduler.find_all_duplicates()
    if duplicates:
        st.subheader("Duplicate Cleanup")
        st.warning(f"{len(duplicates)} duplicate Special Service group(s) detected.")
        labels = []
        index_by_label = {}
        for group in duplicates:
            for idx, ss in group:
                label = f"[{idx + 1}] {ss['date']} — {ss['event']} — {ss['pic']} / {ss['member']}"
                labels.append(label)
                index_by_label[label] = idx
        selected_dup = st.multiselect("Select duplicate record(s) to DELETE", labels, key="ss_dup_delete")
        if st.button("🗑️ Delete Selected Duplicates", key="ss_delete_duplicates", use_container_width=True):
            indices = [index_by_label[x] for x in selected_dup]
            if indices:
                scheduler.remove_duplicates_by_indices(indices)
                st.success(f"Deleted {len(indices)} duplicate Special Service event(s).")
                rerun()

# ============================================================
# TAB 5 - GLADI EVENTS
# ============================================================
elif st.session_state.active_tab == 4:
    st.header("5. Gladi Events")
    with st.container(border=True):
        c1, c2 = st.columns(2)
        with c1:
            gladi_type = st.selectbox("Event Type", ["Gladi Kotor", "Gladi Bersih"], key="gladi_type")
            gladi_date = st.date_input("Date", value=date.today(), key="gladi_date")
        with c2:
            gladi_time = st.text_input("Time", placeholder="e.g., 18.00 WIB", key="gladi_time")
            gladi_desc = st.text_input("Description", placeholder="e.g., AOG Sound and Bound", key="gladi_desc")
        if st.button("➕ Add Gladi Event", type="primary", use_container_width=True):
            d = scheduler.get_indonesian_date(gladi_date)
            result, payload = scheduler.add_gladi_event(gladi_type, d, gladi_time.strip(), gladi_desc.strip())
            if result == 'success':
                st.success(payload)
            elif result in ('duplicate', 'conflict', 'date_warning'):
                st.session_state.gladi_pending = {
                    'result': result, 'payload': payload,
                    'type': gladi_type, 'date': d,
                    'time': gladi_time.strip(), 'desc': gladi_desc.strip()
                }
            else:
                st.warning(payload)
    if st.session_state.get('gladi_pending'):
        p = st.session_state.gladi_pending
        payload = p['payload']
        st.warning("⚠️ Gladi event conflict detected.")
        if p['result'] == 'duplicate':
            ev = payload['existing']
            st.write(f"**Existing:** {ev['date']} — {ev['type']} — {ev['description']} — {ev['time']}")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("❌ Cancel", key="gl_cancel1", use_container_width=True):
                    st.session_state.gladi_pending = None
                    rerun()
            with c2:
                if st.button("🔄 Update Existing", key="gl_update1", type="primary", use_container_width=True):
                    scheduler.add_gladi_event(p['type'], p['date'], p['time'], p['desc'], conflict_action='update')
                    st.session_state.gladi_pending = None
                    st.success("Existing Gladi event updated with new data!")
                    rerun()
        elif p['result'] == 'conflict':
            ev = payload['existing']
            st.write(f"**Existing event:** {ev['date']} — {ev['type']} — {ev['description']} — {ev['time']}")
            st.error("Gladi Bersih and Gladi Kotor cannot be on the same date!")
            c1, c2 = st.columns(2)
            with c1:
                if st.button(" Cancel (Keep Existing)", key="gl_cancel2", use_container_width=True):
                    st.session_state.gladi_pending = None
                    rerun()
            with c2:
                if st.button("🔄 Replace Existing with New", key="gl_replace", type="primary", use_container_width=True):
                    scheduler.add_gladi_event(p['type'], p['date'], p['time'], p['desc'], conflict_action='replace')
                    st.session_state.gladi_pending = None
                    st.success(f"Replaced {ev['type']} with {p['type']} on {p['date']}!")
                    rerun()
        else:
            st.write("**Existing event(s) on same date:**")
            for ev in payload['events']:
                st.write(f"- {ev['description']} ({ev['type']}) — {ev['time']}")
            st.write("What would you like to do?")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("❌ Cancel", key="gl_cancel3", use_container_width=True):
                    st.session_state.gladi_pending = None
                    rerun()
            with c2:
                if st.button("➕ Add Anyway", key="gl_proceed", type="primary", use_container_width=True):
                    scheduler.add_gladi_event(p['type'], p['date'], p['time'], p['desc'], conflict_action='proceed')
                    st.session_state.gladi_pending = None
                    st.success(f"{p['type']} added successfully!")
                    rerun()
    gladi_table()
    if scheduler.gladi_events:
        labels = [f"{i + 1}. {e['date']} — {e['description']} ({e['type']}) — {e['time']}" for i, e in enumerate(scheduler.gladi_events)]
        selected = st.multiselect("Select Gladi event(s) to delete", labels, key="gladi_delete")
        if st.button("🗑️ Delete Selected Gladi Event", use_container_width=True):
            indices = [labels.index(x) for x in selected]
            for idx in sorted(indices, reverse=True):
                del scheduler.gladi_events[idx]
            scheduler.db.save_gladi_events(scheduler.gladi_events)
            st.success(f"Successfully deleted {len(indices)} Gladi event(s)!")
            rerun()
    duplicates = scheduler.find_all_gladi_duplicates()
    if duplicates:
        st.subheader("Duplicate Cleanup")
        st.warning(f"{len(duplicates)} duplicate Gladi group(s) detected.")
        labels = []
        index_by_label = {}
        for group in duplicates:
            for idx, ev in group:
                label = f"[{idx + 1}] {ev['description']} ({ev['type']}) — {ev['date']} — {ev['time']}"
                labels.append(label)
                index_by_label[label] = idx
        selected_dup = st.multiselect("Select duplicate record(s) to DELETE", labels, key="gladi_dup_delete")
        if st.button("🗑️ Delete Selected Duplicates", use_container_width=True):
            indices = [index_by_label[x] for x in selected_dup]
            if indices:
                scheduler.remove_gladi_duplicates_by_indices(indices)
                st.success(f"Deleted {len(indices)} duplicate Gladi event(s).")
                rerun()

# ============================================================
# TAB 6 - MANUAL ASSIGNMENT
# ============================================================
elif st.session_state.active_tab == 5:
    st.header("6. Manual Assignment")
    names = sorted(scheduler.all_names)
    c1, c2 = st.columns(2)
    with c1:
        override_week = st.number_input("Week", min_value=1, max_value=5, value=1, step=1, key="override_week")
        override_svc = st.selectbox("Service", SERVICES, key="override_svc")
    with c2:
        override_role = st.selectbox("Assignment Type", ["PIC", "Member", "Team"], key="override_role")
    selected_group = scheduler.service_to_group.get(override_svc, "")
    configured_teams = scheduler.teams.get(selected_group, []) if selected_group else []
    team_options = [t.get('name', '').strip() for t in configured_teams if t.get('name', '').strip()]
    if override_role == "Team":
        override_team = st.selectbox(
            f"Team ({selected_group})",
            team_options if team_options else [""],
            key="override_team"
        )
        override_person = ""
    else:
        override_person = st.selectbox(
            "Person", names if names else [""], key="override_person"
        )
        override_team = ""
    if st.button("➕ Add Manual Assignment", type="primary", use_container_width=True):
        week = int(override_week)
        if override_role != "Team":
            result, payload = scheduler.add_override(week, override_svc, override_role, override_person)
            if result == 'success':
                st.success(payload)
            elif result == 'duplicate':
                st.session_state.override_pending = {
                    'week': week, 'svc': override_svc,
                    'role': override_role, 'person': override_person, 'existing': payload
                }
            else:
                st.warning(payload)
        else:
            selected_team = next(
                (t for t in configured_teams if t.get('name', '').strip() == override_team),
                None
            )
            if not selected_team:
                st.warning(f"Belum ada konfigurasi Team untuk group {selected_group}.")
            else:
                team_pic = selected_team.get('pic', '').strip()
                team_member = selected_team.get('member', '').strip()
                if not team_pic or not team_member:
                    st.warning(f"{override_team} belum memiliki PIC dan Member yang lengkap.")
                elif team_pic == team_member:
                    st.warning(f"{override_team} tidak dapat digunakan karena PIC dan Member sama.")
                else:
                    existing_pic = next(
                        (o for o in scheduler.manual_overrides
                         if o['week'] == week and o['service'] == override_svc and o['role'] == 'PIC'),
                        None
                    )
                    existing_mem = next(
                        (o for o in scheduler.manual_overrides
                         if o['week'] == week and o['service'] == override_svc and o['role'] == 'Member'),
                        None
                    )
                    st.session_state.override_pending = {
                        'week': week, 'svc': override_svc, 'role': 'Team',
                        'team': override_team, 'team_pic': team_pic, 'team_member': team_member,
                        'existing_pic': existing_pic, 'existing_mem': existing_mem
                    }
    if st.session_state.get('override_pending'):
        p = st.session_state.override_pending
        if p.get('role') == 'Team':
            st.warning(
                f"Set Team {p['team']} untuk Week {p['week']} — {p['svc']}? "
                f"PIC: {p['team_pic']} | Member: {p['team_member']}"
            )
            if p.get('existing_pic') or p.get('existing_mem'):
                st.info("Assignment PIC/Member pada Week dan Service tersebut sudah ada dan akan diperbarui menjadi Team yang dipilih.")
        else:
            st.warning(
                f"An assignment already exists for Week {p['week']}, {p['svc']}, {p['role']}. "
                f"Current person: {p['existing']['person']}. Update it with {p['person']}?"
            )
        c1, c2 = st.columns(2)
        with c1:
            if st.button("❌ No", key="ov_no", use_container_width=True):
                st.session_state.override_pending = None
                rerun()
        with c2:
            if st.button("🔄 Yes, Update", key="ov_yes", type="primary", use_container_width=True):
                if p.get('role') == 'Team':
                    pic_result, pic_payload = scheduler.add_override(
                        p['week'], p['svc'], 'PIC', p['team_pic'],
                        conflict_action='update'
                    )
                    if pic_result == 'error':
                        st.session_state.override_pending = None
                        st.warning(pic_payload)
                        rerun()
                    mem_result, mem_payload = scheduler.add_override(
                        p['week'], p['svc'], 'Member', p['team_member'],
                        conflict_action='update'
                    )
                    if mem_result == 'error':
                        st.session_state.override_pending = None
                        st.warning(mem_payload)
                        rerun()
                    st.session_state.override_pending = None
                    st.success(
                        f"Team {p['team']} berhasil ditetapkan untuk Week {p['week']} — {p['svc']} "
                        f"(PIC: {p['team_pic']}, Member: {p['team_member']})."
                    )
                    rerun()
                else:
                    scheduler.add_override(p['week'], p['svc'], p['role'], p['person'], conflict_action='update')
                    st.session_state.override_pending = None
                    st.success(f"Updated {p['svc']} {p['role']} for Week {p['week']} to {p['person']}!")
                    rerun()
    override_table()
    if scheduler.manual_overrides:
        labels = [f"{i + 1}. Week {o['week']} — {o['service']} — {o['role']} — {o['person']}" for i, o in enumerate(scheduler.manual_overrides)]
        selected = st.multiselect("Select manual assignment(s) to delete", labels, key="override_delete")
        if st.button("🗑️ Delete Selected", use_container_width=True):
            indices = [labels.index(x) for x in selected]
            for idx in sorted(indices, reverse=True):
                del scheduler.manual_overrides[idx]
            st.success(f"Successfully deleted {len(indices)} manual assignment(s)!")
            rerun()

# ============================================================
# TAB 7 - SERVICE REPORT
# ============================================================
elif st.session_state.active_tab == 6:
    st.header("7. Service Report")
    st.caption("Input laporan mingguan, simpan permanen ke SQLite, preview, lalu export rekap bulanan ke Excel.")
    sr_mode = st.radio(
        "Menu Service Report",
        ["📝 Input Report", "📊 Preview & Export"],
        horizontal=True,
        key="sr_mode"
    )
    if sr_mode == "📝 Input Report":
        _service_report_form()
    else:
        _service_report_preview_month()
    if scheduler.service_reports:
        st.divider()
        st.subheader("Data Service Report Tersimpan")
        import pandas as pd
        saved_rows = []
        for r in scheduler.service_reports:
            saved_rows.append({
                'Ibadah': r.get('service'),
                'Tanggal': _format_report_date(r.get('service_date','')),
                'Kehadiran Jemaat Saja': r.get('attendance_total',0),
                'Gembala': r.get('pastor_count',0),
                'Pembicara': r.get('speaker_count',0),
                'Fulltimer': r.get('fulltimer',0),
                'Total Kehadiran': _service_report_total_attendance(r),
                'Jiwa Baru': r.get('new_souls',0),
                'Altarcall': r.get('altar_call',0),
                'Baptisan': r.get('baptism',0),
                'Total Volunteer': r.get('total_volunteers',0),
                'Total Prayer': sum(int(p.get('count', 0)) for p in r.get('prayer_corner', [])),
            })
        st.dataframe(pd.DataFrame(saved_rows), use_container_width=True, hide_index=True)
        delete_options = [
            f"{r.get('service')} — {_format_report_date(r.get('service_date',''))}"
            for r in scheduler.service_reports
        ]
        selected_delete = st.multiselect("Pilih laporan untuk dihapus", delete_options, key="sr_delete_selection")
        if st.button("🗑️ Hapus Service Report Terpilih", use_container_width=True):
            # Ambil semua report yang dipilih SEBELUM melakukan penghapusan.
            # delete_service_report() juga menghapus item dari
            # scheduler.service_reports, sehingga mencari index dari list yang
            # sudah berubah dapat menyebabkan IndexError jika lebih dari satu
            # report dihapus sekaligus.
            selected_reports = []
            for label in selected_delete:
                idx = delete_options.index(label)
                if 0 <= idx < len(scheduler.service_reports):
                    selected_reports.append(scheduler.service_reports[idx])

            for report in selected_reports:
                scheduler.delete_service_report(
                    report.get('service'),
                    report.get('service_date')
                )

            st.success(f"{len(selected_reports)} laporan berhasil dihapus.")
            rerun()

# ============================================================
# TAB 8 - GENERATE & EXPORT
# ============================================================
elif st.session_state.active_tab == 7:
    st.header("8. Generate & Export")
    ensure_week_dates()
    if not scheduler.schedule_data:
        st.info("Please verify all data before generating the schedule.")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("⚙️ Generate Schedule Preview", type="primary", use_container_width=True):
            ok, msg = scheduler.generate_schedule()
            if ok:
                st.session_state.schedule_generated = True
                scheduler.save_current_state_to_db()
                st.success("Schedule generated successfully!")
            else:
                st.error(msg)
    with c2:
        if st.button("📊 Export to Excel", use_container_width=True):
            if not scheduler.schedule_data:
                st.warning("Please generate first.")
            else:
                st.session_state.export_requested = True
    if st.session_state.get('export_requested', False):
        st.warning("Are you sure you want to export the schedule? Please verify all data before proceeding.")
        c1, c2 = st.columns(2)
        with c1:
            if st.button(" Cancel Export", use_container_width=True):
                st.session_state.export_requested = False
                rerun()
        with c2:
            if st.button("✅ Yes, Export", type="primary", use_container_width=True):
                try:
                    data = scheduler.export_to_excel_bytes()
                    filename = f"Jadwal_DM_GMS_Salatiga_{scheduler.month_combo}_{scheduler.year_spin}.xlsx"
                    st.download_button(
                        "⬇️ Download Excel File",
                        data=data,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    scheduler.save_current_state_to_db()
                    st.success("Schedule exported and saved!")
                except Exception as e:
                    st.error(f"Export error: {e}")
    st.subheader("Schedule Preview")
    show_schedule_html()

# ============================================================
# END
# ============================================================