#!/usr/bin/env python3
"""
generate_bulletin.py
Génère le bulletin météo Bordeaux — HTML statique compatible e-mail Outlook.
Le fichier produit fonctionne en double mode :
  • Navigateur  → le JS écrase le grid avec le même contenu (rendu dynamique normal)
  • Outlook     → les cartes pré-rendues sont déjà dans le DOM (pas besoin de JS)

Données : API Open-Meteo (https://open-meteo.com/, CC-BY 4.0 — attribution dans
le pied du bulletin). Modèle "best_match" : AROME/ARPEGE (Météo-France) en début
de semaine, complété par un modèle global pour couvrir les 7 jours.
NB licence : l'offre gratuite Open-Meteo est réservée à un usage non commercial ;
si la conformité l'exige, basculer sur le portail officiel Météo-France
(portail-api.meteofrance.fr, licence Etalab) — seule fetch_openmeteo() change.
Garde-fou : si la semaine calendaire n'est pas couverte par la réponse API,
le script sort en erreur (exit 1) → le pipeline s'arrête, rien n'est envoyé.
"""

import json
import re
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime, date, timedelta
from pathlib import Path

# Console Windows (cp1252) : forcer la sortie en UTF-8 pour les messages d'etat
# (accents / fleches / emojis) sinon UnicodeEncodeError. N'affecte pas l'ecriture
# du HTML qui specifie deja encoding='utf-8'.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

# ────────────────────────────────────────────────────────────
# Récupération Open-Meteo (semaine calendaire, pas horaire, heure locale)
# ────────────────────────────────────────────────────────────

# Point de prévision : Bordeaux centre (agglomération bordelaise)
LAT, LON = 44.84, -0.58

API_URL = (
    "https://api.open-meteo.com/v1/forecast"
    f"?latitude={LAT}&longitude={LON}"
    "&hourly=temperature_2m,precipitation,weather_code,wind_gusts_10m,relative_humidity_2m"
    "&timezone=Europe%2FParis"
    "&start_date={start}&end_date={end}"
)

# Codes temps WMO (weather_code) → les 6 conditions du bulletin
WMO_TO_COND = {
    0: 'sunny', 1: 'sunny', 2: 'partly', 3: 'cloudy',
    45: 'cloudy', 48: 'cloudy',                                  # brouillard
    51: 'sprinkle', 53: 'sprinkle', 55: 'sprinkle',              # bruine
    56: 'sprinkle', 57: 'sprinkle',
    61: 'sprinkle', 63: 'rain', 65: 'rain', 66: 'rain', 67: 'rain',
    71: 'rain', 73: 'rain', 75: 'rain', 77: 'rain',              # neige (rare ici)
    80: 'sprinkle', 81: 'rain', 82: 'rain', 85: 'rain', 86: 'rain',
    95: 'storm', 96: 'storm', 99: 'storm',
}

# Sévérité croissante — choix de la condition de créneau + synthèse jour
SEV = ["sunny", "partly", "cloudy", "sprinkle", "rain", "storm"]

def fetch_openmeteo(start, end):
    url = API_URL.format(start=start.isoformat(), end=end.isoformat())
    req = urllib.request.Request(url, headers={
        "User-Agent": "RDI_PROJET_METEO/1.0 (bulletin interne)"})
    payload  = None
    last_err = None
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                payload = json.load(resp)
            break
        except (urllib.error.URLError, OSError, ValueError) as e:
            last_err = e
            if attempt < 2:
                time.sleep(10)
    if payload is None:
        sys.exit(f"[ERREUR] Open-Meteo inaccessible apres 3 tentatives : {last_err}")
    h = payload.get("hourly") or {}
    rows = []
    for i, ts in enumerate(h.get("time") or []):        # ts = "2026-06-10T08:00"
        if h["temperature_2m"][i] is None:              # trou de donnees : ligne ignoree
            continue
        rows.append({
            'date':     date.fromisoformat(ts[:10]),
            'hour':     int(ts[11:13]),
            'temp':     h["temperature_2m"][i],
            'wind_raf': h["wind_gusts_10m"][i] or 0,
            'precip':   h["precipitation"][i] or 0.0,
            'hum':      h["relative_humidity_2m"][i] if h["relative_humidity_2m"][i] is not None else 60,
            'code':     h["weather_code"][i],
        })
    if not rows:
        sys.exit("[ERREUR] Reponse Open-Meteo vide ou invalide — bulletin NON genere.")
    return rows

def cond_from(precip, hum):
    # Repli si weather_code absent : estimation grossière precip/humidité
    if precip > 2.0:  return "storm"
    if precip > 0.5:  return "rain"
    if precip > 0.0:  return "sprinkle"
    if hum < 45:      return "sunny"
    if hum < 70:      return "partly"
    return "cloudy"

# Créneaux du bulletin (Plan.md §6.1 : matin 08h30–12h30, aprem 13h30–17h30).
# Open-Meteo : la valeur horaire de précipitation = cumul de l'heure PRÉCÉDENTE,
# donc les heures 9→12 couvrent les cumuls 08h→12h, et 14→18 les cumuls 13h→18h.
MATIN_HOURS = (9, 10, 11, 12)
APREM_HOURS = (14, 15, 16, 17, 18)

def slot_agg(readings, target_hours):
    rows = [r for r in readings if r['hour'] in target_hours]
    if not rows:
        all_h = sorted(set(r['hour'] for r in readings))
        pivot = sum(target_hours) // len(target_hours)
        closest = min(all_h, key=lambda h: abs(h - pivot), default=None)
        rows = [r for r in readings if r['hour'] == closest] if closest else []
    if not rows:
        return {"precip": 0.0, "temp": 20, "cond": "cloudy", "peak": None}
    precip = round(sum(r['precip'] for r in rows), 1)
    temp   = round(max(r['temp'] for r in rows))
    # Condition du créneau : la plus défavorable des heures (codes WMO),
    # repli precip/humidité si le code manque.
    conds = [WMO_TO_COND[r['code']] for r in rows if r['code'] in WMO_TO_COND]
    if conds:
        cond = max(conds, key=SEV.index)
    else:
        hum  = round(sum(r['hum'] for r in rows) / len(rows))
        cond = cond_from(precip, hum)
    wettest = max(rows, key=lambda r: r['precip'])
    peak = f"{wettest['hour']}h" if wettest['precip'] > 0 else None
    return {"precip": precip, "temp": temp, "cond": cond, "peak": peak}

def day_agg(readings):
    return {
        "wind":  round(max((r['wind_raf'] for r in readings), default=0)),
        "matin": slot_agg(readings, MATIN_HOURS),
        "aprem": slot_agg(readings, APREM_HOURS),
    }

def wind_alert_label(max_raf):
    if max_raf >= 50: return "alerte"
    if max_raf >= 30: return "vigilance"
    return "normal"

def build_template():
    today  = date.today()
    monday = today - timedelta(days=today.weekday())
    sunday = monday + timedelta(days=6)
    # L'API sert aussi les jours passés de la semaine (analyse du modèle) :
    # les cartes estompées affichent des données réelles, plus des valeurs par défaut.
    rows   = fetch_openmeteo(monday, sunday)
    by_day = {}
    for r in rows:
        by_day.setdefault(r['date'], []).append(r)
    # Garde-fou fraîcheur : les 7 jours de la semaine doivent être couverts,
    # sinon on refuse de générer (exit 1 → le pipeline n'envoie rien).
    missing = [(monday + timedelta(days=i)).isoformat()
               for i in range(7) if monday + timedelta(days=i) not in by_day]
    if missing:
        sys.exit("[ERREUR] Donnees incompletes (jours manquants : "
                 + ", ".join(missing) + ") — bulletin NON genere.")
    return [day_agg(by_day[monday + timedelta(days=i)]) for i in range(7)]

def js_template_str(template):
    def slot(s):
        pk = f'"{s["peak"]}"' if s["peak"] else "null"
        return (f'{{ precip:{s["precip"]:.1f}, temp:{s["temp"]}, '
                f'cond:"{s["cond"]}", peak:{pk} }}')
    lines = [
        f'  {{ wind:{d["wind"]}, matin:{slot(d["matin"])}, aprem:{slot(d["aprem"])} }}'
        for d in template
    ]
    return "const TEMPLATE = [\n" + ",\n".join(lines) + "\n];"

# ────────────────────────────────────────────────────────────
# Icônes SVG (extrait de weather-icons.js inliné)
# ────────────────────────────────────────────────────────────

def _cloud(body, shadow):
    def shapes(f):
        return (f'<circle cx="22" cy="34" r="10" fill="{f}"/>'
                f'<circle cx="41" cy="33" r="13" fill="{f}"/>'
                f'<circle cx="31" cy="25" r="11" fill="{f}"/>'
                f'<rect x="12" y="33" width="42" height="15" rx="7.5" fill="{f}"/>')
    return (f'<g transform="translate(33,35) scale(1.09) translate(-33,-35)">{shapes(shadow)}</g>'
            f'<g>{shapes(body)}</g>')

def _drops(xs, y1, y2):
    lines = ''.join(f'<line x1="{x}" y1="{y1}" x2="{x-2.5}" y2="{y2}"/>' for x in xs)
    return f'<g stroke="#2f74c0" stroke-width="3.4" stroke-linecap="round">{lines}</g>'

_SUN_RAYS = (
    '<g stroke="#f6a623" stroke-width="3.6" stroke-linecap="round">'
    '<line x1="32" y1="3"  x2="32" y2="11"/><line x1="32" y1="37" x2="32" y2="45"/>'
    '<line x1="3"  y1="24" x2="11" y2="24"/><line x1="53" y1="24" x2="61" y2="24"/>'
    '<line x1="11" y1="3"  x2="16" y2="8"/> <line x1="48" y1="40" x2="53" y2="45"/>'
    '<line x1="53" y1="3"  x2="48" y2="8"/> <line x1="16" y1="40" x2="11" y2="45"/>'
    '</g>'
)
_SUN_DISC = (
    '<circle cx="32" cy="24" r="12" fill="#ffd23f"/>'
    '<circle cx="32" cy="24" r="12" fill="none" stroke="#f6a623" stroke-width="2"/>'
)

WEATHER_SVG = {
    'sunny': (
        '<svg viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">'
        '<g stroke="#f6a623" stroke-width="4" stroke-linecap="round">'
        '<line x1="32" y1="4"  x2="32" y2="13"/><line x1="32" y1="51" x2="32" y2="60"/>'
        '<line x1="4"  y1="32" x2="13" y2="32"/><line x1="51" y1="32" x2="60" y2="32"/>'
        '<line x1="12" y1="12" x2="18" y2="18"/><line x1="46" y1="46" x2="52" y2="52"/>'
        '<line x1="52" y1="12" x2="46" y2="18"/><line x1="18" y1="46" x2="12" y2="52"/>'
        '</g>'
        '<circle cx="32" cy="32" r="14" fill="#ffd23f"/>'
        '<circle cx="32" cy="32" r="14" fill="none" stroke="#f6a623" stroke-width="2.2"/>'
        '<circle cx="27" cy="27" r="4.5" fill="#ffe488"/></svg>'
    ),
    'partly':   f'<svg viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">{_SUN_RAYS}{_SUN_DISC}{_cloud("#ffffff","#c6d4e1")}</svg>',
    'cloudy':   f'<svg viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg"><g transform="translate(0,4)">{_cloud("#dde7f1","#aebfd0")}</g></svg>',
    'sprinkle': f'<svg viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg"><g transform="translate(0,-4)">{_cloud("#ffffff","#c6d4e1")}</g>{_drops([24,38],48,57)}</svg>',
    'rain':     f'<svg viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg"><g transform="translate(0,-5)">{_cloud("#dce7f2","#b2c4d6")}</g>{_drops([20,30,40,48],47,58)}</svg>',
    'storm':    (
        f'<svg viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">'
        f'<g transform="translate(0,-5)">{_cloud("#5d6f88","#46566b")}</g>'
        '<polygon points="35,38 23,55 31,55 27,64 44,46 35,46" '
        'fill="#ffd23f" stroke="#f6a623" stroke-width="1.4" stroke-linejoin="round"/></svg>'
    ),
}

UI_SVG = {
    'wind': (
        '<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" fill="none" '
        'stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
        '<path d="M3 8h11a2.5 2.5 0 1 0-2.5-2.5"/>'
        '<path d="M3 16h14a2.5 2.5 0 1 1-2.5 2.5"/>'
        '<path d="M3 12h7"/></svg>'
    ),
    'thermometer': (
        '<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" fill="none" '
        'stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
        '<path d="M14 14.8V5a2.5 2.5 0 0 0-5 0v9.8a4 4 0 1 0 5 0"/>'
        '<line x1="11.5" y1="17.5" x2="11.5" y2="9"/></svg>'
    ),
    'fire': (
        '<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">'
        '<path d="M12 2.5c2.4 3.4 5 5.2 5 9.2a5 5 0 0 1-10 0c0-2.1 1.1-3.7 2.5-5'
        'c.1 1.7.9 2.4 1.7 2c1-.5.3-2.3-.2-3.6C11.8 4 12 2.5 12 2.5z" fill="currentColor"/></svg>'
    ),
}

# Pré-rendu HTML statique des cartes (compatible Outlook)

COND_FR = {
    'sunny':    "Ensoleillé",
    'partly':   "Éclaircies",
    'cloudy':   "Couvert",
    'sprinkle': "Averses faibles",
    'rain':     "Pluie",
    'storm':    "Orages",
}

PRECIP_STYLE = {
    'p0': 'background:#e4f0fb;color:#1d4b78',
    'p1': 'background:#1f5ba8;color:#ffffff',
    'p2': 'background:#f4c20d;color:#473600',
    'p3': 'background:#ef8a64;color:#551b08',
    'p4': 'background:#b5271b;color:#ffffff',
}
WIND_COLOR = {
    'normal':    '#7c8a9a',
    'vigilance': '#e0930e',
    'alerte':    '#c0392b',
}

MON_FR = ["Janvier","Février","Mars","Avril","Mai","Juin",
          "Juillet","Août","Septembre","Octobre","Novembre","Décembre"]
DOW_FR = ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"]

def precip_cls(mm):
    if mm == 0:   return 'p0'
    if mm < 1.0:  return 'p1'
    if mm < 2.0:  return 'p2'
    if mm < 5.0:  return 'p3'
    return 'p4'

def slot_html(name, s):
    pcls   = precip_cls(s['precip'])
    pstyle = PRECIP_STYLE[pcls]
    fr_mm  = str(s['precip']).replace('.', ',')
    icon   = WEATHER_SVG.get(s['cond'], WEATHER_SVG['cloudy'])
    label  = COND_FR.get(s['cond'], "Couvert")
    return (
        f'<div class="slot {pcls}" style="{pstyle}">'
        f'<div class="slot-top"><span class="slot-name">{name}</span></div>'
        f'<div class="slot-main"><div class="slot-icon">{icon}</div><div>'
        f'<div class="slot-precip"><span class="mm">{fr_mm}<small>mm</small></span></div>'
        f'<div class="slot-cond">{label}</div>'
        f'</div></div></div>'
    )

def day_html(d, i):
    classes = ["day"]
    if i >= 5:       classes.append("is-weekend")
    if d['today']:   classes.append("is-today")
    if d['past']:    classes.append("is-past")
    cls = " ".join(classes)
    we_pill = '<span class="we-tag">we</span>' if i >= 5 else ""
    if d['today']:   pill = '<span class="today-pill">Auj.</span>'
    elif d['past']:  pill = '<span class="past-pill">Écoulé</span>'
    else:            pill = ""
    sev_m = SEV.index(d['matin']['cond']) if d['matin']['cond'] in SEV else 2
    sev_a = SEV.index(d['aprem']['cond']) if d['aprem']['cond'] in SEV else 2
    worst = d['matin']['cond'] if sev_m >= sev_a else d['aprem']['cond']
    tmax  = max(d['matin']['temp'], d['aprem']['temp'])
    hot   = tmax >= 30
    sum_icon = WEATHER_SVG.get(worst, WEATHER_SVG['cloudy'])
    sum_txt  = COND_FR.get(worst, "Couvert")
    therm_or_fire = UI_SVG['fire'] if hot else UI_SVG['thermometer']
    hot_cls = "hot" if hot else ""
    wlvl   = wind_alert_label(d['wind'])
    wcolor = WIND_COLOR[wlvl]
    return (
        f'<article class="{cls}">'
        f'<div class="day-head"><div class="day-id">'
        f'<span class="dow">{d["dow"]}{we_pill}{pill}</span>'
        f'<span class="date">{d["date"]}</span></div>'
        f'<div class="summary"><span class="cond-ic">{sum_icon}</span>'
        f'<span class="cond">{sum_txt}</span>'
        f'<span class="tmax {hot_cls}"><span class="ic">{therm_or_fire}</span>'
        f'<b>{tmax}°</b></span></div></div>'
        f'<div class="slots">'
        f'{slot_html("Matin", d["matin"])}'
        f'{slot_html("Après-midi", d["aprem"])}'
        f'</div>'
        f'<div class="day-foot"><div class="wind">'
        f'<span class="wi">{UI_SVG["wind"]}</span>'
        f'<span class="val">{d["wind"]} <span>km/h</span></span>'
        f'<span class="lvl" style="color:{wcolor};">'
        f'<span class="dot" style="background:{wcolor};"></span>{wlvl.capitalize()}'
        f'</span></div></div></article>'
    )

def build_week(template):
    today  = date.today()
    monday = today - timedelta(days=today.weekday())
    week = []
    for i, tpl in enumerate(template):
        d = monday + timedelta(days=i)
        week.append({**tpl, 'dow': DOW_FR[i],
                     'date': f'{d.day} {MON_FR[d.month-1]}',
                     'today': d == today, 'past': d < today})
    return week

def render_grid(week):
    return "\n".join(day_html(d, i) for i, d in enumerate(week))

def render_subtitle(today):
    monday = today - timedelta(days=today.weekday())
    sunday = monday + timedelta(days=6)
    days_left = 8 - today.isoweekday()
    s = "s" if days_left > 1 else ""
    week_label = f'semaine du {monday.day} au {sunday.day} {MON_FR[sunday.month-1]} {sunday.year}'
    return (f'Prévisions à <b>{days_left}\xa0jour{s}</b> · '
            f'<b>Agglomération Bordelaise</b> · <span class="mono">{week_label}</span>')

def render_stamp(now):
    py_to_fr = {0:"lun.", 1:"mar.", 2:"mer.", 3:"jeu.", 4:"ven.", 5:"sam.", 6:"dim."}
    g = py_to_fr[now.weekday()]
    return (f'<span class="live-dot"></span> Généré {g} '
            f'{now.day:02d}/{now.month:02d}/{now.year} · {now.strftime("%Hh%M")}'
            f' · Données <a href="https://open-meteo.com/" '
            f'style="color:inherit;">Open-Meteo.com</a> (CC-BY 4.0)')

# Chemins relatifs au script — structure autonome (Prod).
SCRIPT_DIR    = Path(__file__).resolve().parent
TEMPLATE_HTML = SCRIPT_DIR / "Bulletin Météo Bordeaux.html"
OUTPUT_DIR    = SCRIPT_DIR.parent / "Bulletins"

def generate():
    tpl  = build_template()
    week = build_week(tpl)
    now  = datetime.now()
    today = date.today()
    html = TEMPLATE_HTML.read_text(encoding='utf-8')
    html = re.sub(r'const TEMPLATE = \[[\s\S]*?\];', js_template_str(tpl), html)
    html = re.sub(
        r'(<div class="sub" id="subtitle">)[\s\S]*?(</div>)',
        lambda m: m.group(1) + render_subtitle(today) + m.group(2), html)
    html = re.sub(
        r'(<div class="stamp mono" id="stamp">)[\s\S]*?(</div>)',
        lambda m: m.group(1) + render_stamp(now) + m.group(2), html)
    html = html.replace(
        '<div class="grid" id="grid"></div>',
        f'<div class="grid" id="grid">{render_grid(week)}</div>')
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    outfile = OUTPUT_DIR / f"Meteo_{now.strftime('%Y%m%d-%H%M')}.html"
    outfile.write_text(html, encoding='utf-8')
    print(f"[OK] {outfile.name}")
    jours  = ["Lun","Mar","Mer","Jeu","Ven","Sam","Dim"]
    monday = today - timedelta(days=today.weekday())
    print("\n=== Prévisions semaine — source Open-Meteo (best_match) ===")
    for i, (label, d) in enumerate(zip(jours, tpl)):
        dm = monday + timedelta(days=i)
        marker = " ← aujourd'hui" if dm == today else (" (passé)" if dm < today else "")
        alert  = f" ⚡{wind_alert_label(d['wind'])}" if d['wind'] >= 30 else ""
        hm = "🔥" if d['matin']['temp'] >= 30 else ""
        ha = "🔥" if d['aprem']['temp'] >= 30 else ""
        print(f"  {label} {dm.strftime('%d/%m')}: "
              f"matin {d['matin']['temp']}°C{hm} {d['matin']['cond']:8s} "
              f"/ aprem {d['aprem']['temp']}°C{ha} {d['aprem']['cond']:8s} "
              f"| vent {d['wind']}km/h{alert}{marker}")
    print(f"\nFichier : {outfile}")
    return outfile

if __name__ == "__main__":
    generate()
