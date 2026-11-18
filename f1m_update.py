#!/usr/bin/env python3
"""
F1 Manager 2024 planner updater, version 1.0.0.

Reads a save from your SaveGames folder (read-only) and rebuilds the data tabs of
that career's workbook, "F1 Manager Planner - <team>.xlsx", next to this script.
When there is more than one save it lists them, with team, manager and race, and
asks which to use. Each career keeps its own workbook and its own Settings tab.

    python f1m_update.py            pick a save, update its workbook
    python f1m_update.py --open     the same, then open the workbook
    python f1m_update.py --watch    pick a save, then update every time that career saves
    python f1m_update.py --newest   skip the question and use the newest save
    python f1m_update.py --list     just list the saves

Needs Python 3.8+ and openpyxl  (pip install openpyxl).
The workbook must be closed in Excel while it is being updated.
The "Settings" tab and any tabs you add yourself are never touched.
"""
import argparse
import datetime as dt
import os
import re
import sqlite3
import struct
import sys
import tempfile
import time
import zlib

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.cell import MergedCell
    from openpyxl.chart import BarChart, DoughnutChart, RadarChart, Reference
    from openpyxl.chart.marker import DataPoint
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.chart.text import RichText, Text
    from openpyxl.chart.title import Title
    from openpyxl.drawing.line import LineProperties
    from openpyxl.drawing.text import CharacterProperties, Font as DrawingFont, Paragraph, ParagraphProperties, RegularTextRun
    from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
except ImportError:
    sys.exit("openpyxl is not installed. Run:  pip install openpyxl")

VERSION = "1.0.0"
HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_SAVE_DIR = os.path.join(os.environ.get("LOCALAPPDATA", ""), "F1Manager24", "Saved", "SaveGames")
DEFAULT_WORKBOOK = os.path.join(HERE, "F1 Manager Planner.xlsx")
NOT_CAREER_SAVES = {"customisation.sav", "player.sav", "scenario.sav"}

OWNED_SHEETS = ["Overview", "Action plan", "Board", "Car vs grid", "Part summary", "Part stats", "Designs", "Stock", "Spares", "Power unit", "Races",
                "Manufacturing", "Staff", "Facilities", "Scouting", "Pit crew", "Drivers", "Finance", "Womens team"]

# (label, default, note) - row order is fixed: formulas point at these cells.
SETTINGS = [
    ("Good rank: at or better than this needs no upgrade", 4, "Used on 'Car vs grid'"),
    ("Low mentality below", 40, "Used on 'Staff' and 'Overview'"),
    ("Worn part: condition below", 0.35, "Used on 'Stock' and 'Overview'"),
    ("Scouting: minimum opinion of your team", 45, "Candidates below this are left out"),
    ("Scouting: staff may rate this far below your current person", 8, "Drivers must rate above your weakest race driver"),
    ("Build work per day, normal, before Factory bonus", 200,
     "Inferred from other teams' builds in the save, not stored by the game. Change if the in-game estimate differs"),
    ("Build work per day, rushed, before Factory bonus", 300, "Normal x 1.5 (game speed multipliers 2 and 3)"),
    ("Rushed cost multiplier", 1.5, "From the save's speed table"),
    ("Emergency cost multiplier", 4, "From the save's speed table"),
    ("Minimum good spares to hold of each part", 1, "Used on 'Spares'. Raise it if you want more cover against crashes"),
    ("All-female team project: 1 = on, 0 = off", 0,
     "On: adds the 'Womens team' tab and puts its next steps into the Action plan"),
    ("Reserve: swap into a race seat when rated this many points higher", 3,
     "Used by the reserve driver plan on 'Action plan'. Ratings are averages of the driver's stats"),
    ("Reserve: give an FP1 outing every this many race weekends", 2,
     "2 = every other weekend. Sprint weekends are always skipped: they have only one practice session"),
    ("Cash to keep in the bank (safety buffer), $", 5000000,
     "The Action plan only spends what is above this, and never past the projected cost cap headroom"),
    ("Maximum number of actions on the plan", 25, "Actions that fit the money but fall beyond this are listed as waiting"),
    ("Graphics: 1 = game style (dark, charts), 0 = plain", 1, "Plain is quicker to print and easier on older versions of Excel"),
]
S_GOOD, S_MENT, S_WORN, S_OPIN, S_BELOW, S_NORMAL, S_RUSHED, S_RUSHX, S_EMERGX, S_MINSPARE, S_WOMEN, S_SWAP, S_FP1, S_BUFFER, S_MAXACT, S_GFX = (
    "Settings!$B$%d" % (i + 2) for i in range(len(SETTINGS)))

# car area a circuit asks for -> part stat that drives it (my reading of the game, not stored in the save)
AREAS = [
    ("Medium-speed cornering", "Medium-speed", 8),
    ("High-speed cornering", "High-speed", 9),
    ("Low-speed cornering", "Low-speed", 7),
    ("Top speed", "Top speed", 4),
    ("DRS", "DRS", 3),
    ("All cornering (airflow front)", None, 0),
    ("All cornering (airflow middle)", None, 13),
    ("Dirty air tolerance", None, 1),
    ("Brake cooling", None, 2),
    ("Engine cooling", None, 5),
]
CAR_STAT_LABEL = {0: "Top speed", 1: "Acceleration", 2: "Braking", 3: "Top speed", 4: "Low-speed", 5: "Medium-speed",
                  6: "High-speed", 7: "Low-speed", 8: "Medium-speed", 9: "High-speed", 13: "Engine cooling",
                  15: "Fuel efficiency", 16: "DRS"}
CAR_STAT_TO_PART_STAT = {"Top speed": 4, "Low-speed": 7, "Medium-speed": 8, "High-speed": 9, "DRS": 3}
PART_STAT_LABEL = {0: "Airflow front", 1: "Dirty air tolerance", 2: "Brake cooling", 3: "DRS delta", 4: "Drag reduction",
                   5: "Engine cooling", 6: "Fuel efficiency", 7: "Low-speed downforce", 8: "Medium-speed downforce",
                   9: "High-speed downforce", 10: "Power", 13: "Airflow middle", 15: "Durability"}
PART_LABEL = {3: "Body", 4: "Front wing", 5: "Rear wing", 6: "Sidepods", 7: "Floor", 8: "Suspension"}
ROLE_LABEL = {0: "Driver", 1: "Technical Chief", 2: "Race Engineer", 3: "Head of Aerodynamics", 4: "Sporting Director"}
MOOD = {0: "Optimistic", 1: "Positive", 2: "Neutral", 3: "Negative", 4: "Pessimistic"}

FONT = Font(name="Arial", size=10)
BOLD = Font(name="Arial", size=10, bold=True)
TITLE = Font(name="Arial", size=12, bold=True)
INPUT = Font(name="Arial", size=10, color="0000FF")
HEAD_FILL = PatternFill("solid", fgColor="E7E6E6")
MONEY = "$#,##0"
PCT = "0%"


# ---------------------------------------------------------------- save file

def read_header(path):
    """Team, manager, career and race from the save's header, without unpacking the database."""
    with open(path, "rb") as f:
        head = f.read(400000)

    def prop(name, kind):
        key = (struct.pack("<i", len(name) + 1) + name.encode() + b"\x00" +
               struct.pack("<i", len(kind) + 1) + kind.encode() + b"\x00")
        i = head.find(key)
        if i < 0:
            return None
        j = i + len(key) + 8 + 1          # value size, then a one-byte flag
        if kind == "Int8Property":
            return head[j]
        n = struct.unpack("<i", head[j:j + 4])[0]
        raw = head[j + 4:j + 4 - 2 * n].decode("utf-16-le", "ignore") if n < 0 else head[j + 4:j + 4 + n].decode("utf-8", "ignore")
        return raw.rstrip("\x00")

    team = prop("Team", "StrProperty") or ""
    m = re.match(r"\[STRING_LITERAL:Value=\|(.*)\|\]$", team)
    if m:
        team = m.group(1)
    elif team.startswith("["):
        team = words(team.strip("[]").split("_")[-1])
    return dict(team=team or "Unknown team", manager=" ".join(x for x in (prop("FirstName", "StrProperty"),
                                                                         prop("LastName", "StrProperty")) if x),
                career=prop("CareerID", "StrProperty") or path, race=prop("CurrentRace", "Int8Property"),
                races=prop("RacesInSeason", "Int8Property"))


def list_saves(save_dir):
    """Every career save in the folder, newest first, with what its header says."""
    out = []
    for name in os.listdir(save_dir):
        if name.lower().endswith(".sav") and name.lower() not in NOT_CAREER_SAVES:
            path = os.path.join(save_dir, name)
            try:
                info = read_header(path)
            except OSError:
                continue
            info.update(path=path, file=name, mtime=os.path.getmtime(path))
            out.append(info)
    if not out:
        raise RuntimeError("No career save (.sav) found in " + save_dir)
    return sorted(out, key=lambda x: -x["mtime"])


def describe(info):
    race = "race %s of %s" % (info["race"], info["races"]) if info["race"] and info["races"] else ""
    when = dt.datetime.fromtimestamp(info["mtime"]).strftime("%d %b %H:%M")
    return "%-24s %-20s %-16s %-16s saved %s" % (info["team"][:24], info["manager"][:20], race, info["file"][:16], when)


def choose_save(saves):
    """Ask which save to use. Enter takes the newest."""
    print("\nCareer saves found:\n")
    print("     %-24s %-20s %-16s %-16s %s" % ("Team", "Manager", "Next race", "File", "When"))
    for i, info in enumerate(saves, 1):
        print("  %d  %s" % (i, describe(info)))
    while True:
        answer = input("\nWhich save? Type its number, or press Enter for 1 (q to quit): ").strip().lower()
        if answer in ("", "1"):
            return saves[0]
        if answer == "q":
            sys.exit(0)
        if answer.isdigit() and 1 <= int(answer) <= len(saves):
            return saves[int(answer) - 1]
        print("Please type a number from 1 to %d." % len(saves))


def workbook_for(team):
    safe = re.sub(r'[<>:"/\\|?*]', "", team).strip() or "Unknown team"
    return os.path.join(HERE, "F1 Manager Planner - %s.xlsx" % safe)


def seed_settings(wb, legacy_path):
    """A new career workbook starts with the settings from the old single workbook, if there is one."""
    try:
        old = load_workbook(legacy_path)["Settings"]
    except Exception:
        return
    values = {old.cell(row=r, column=1).value: old.cell(row=r, column=2).value for r in range(2, old.max_row + 1)}
    ws = wb["Settings"]
    for r in range(2, len(SETTINGS) + 2):
        v = values.get(ws.cell(row=r, column=1).value)
        if isinstance(v, (int, float)):
            ws.cell(row=r, column=2).value = v


def unpack_save(save_path, out_dir):
    """The .sav is an Unreal header followed by a zlib stream holding three SQLite databases."""
    with open(save_path, "rb") as f:
        data = f.read()
    sep = b"\x00\x05\x00\x00\x00None\x00\x05\x00\x00\x00None\x00"
    i = data.find(sep)
    if i < 0:
        raise RuntimeError("This does not look like an F1 Manager 2024 career save: " + save_path)
    p = i + len(sep)
    _unknown, _packed, size_main, _size_b1, _size_b2 = struct.unpack("<5i", data[p:p + 20])
    raw = zlib.decompress(data[p + 20:])
    db_path = os.path.join(out_dir, "main.db")
    with open(db_path, "wb") as f:
        f.write(raw[:size_main])
    return db_path


def person_name(first, last):
    def clean(s):
        s = s or ""
        m = re.search(r"_([^_\]]+)\]$", s)
        if s.startswith("[") and m:
            s = re.sub(r"\d+$", "", m.group(1))
            s = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", s)
            s = re.sub(r"\b(Mc|Mac) (?=[A-Z])", r"\1", s)
        return s
    return (clean(first) + " " + clean(last)).strip()


def words(s):
    """'WeeklyDriverXPGain' -> 'Weekly Driver XP Gain'"""
    return re.sub(r"(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])", " ", s or "")


def median(values):
    v = sorted(values)
    n = len(v)
    if not n:
        return None
    return v[n // 2] if n % 2 else (v[n // 2 - 1] + v[n // 2]) / 2


# ---------------------------------------------------------------- extraction

def extract(db_path, settings):
    c = sqlite3.connect(db_path)
    q = lambda sql, *a: c.execute(sql, a).fetchall()
    D = {}

    team, first, last = q("select TeamID, FirstName, LastName from Player")[0]
    day, season = q("select Day, CurrentSeason from Player_State")[0]
    season_start = (dt.date(season, 1, 1) - dt.date(1899, 12, 30)).days
    f1_teams = [r[0] for r in q("select TeamID from Teams where Formula=1")]
    rivals = [t for t in f1_teams if t != team]
    D.update(team=team, day=day, season=season, n_teams=len(f1_teams))

    # ---- parts
    my_designs = q("""select DesignID, PartType, DesignNumber, DesignWork, DesignWorkMax, DayCreated, DayCompleted,
                      WindTunnelTime, CFD, ManufactureCount, BuildCost, BuildWorkMax, DesignCost
                      from Parts_Designs where TeamID=? and ValidFrom=? and PartType>=3 order by PartType, DesignNumber""",
                   team, season)
    stat_values = {}
    for did, st, val in q("""select v.DesignID, v.PartStat, v.Value from Parts_Designs_StatValues v
                             join Parts_Designs d on d.DesignID=v.DesignID where d.PartType>=3 and v.PartStat<>15"""):
        stat_values.setdefault(did, {})[st] = val
    loadout = {}
    for t, lo, pt, did in q("select TeamID, LoadoutID, PartType, DesignID from Parts_CarLoadout where PartType>=3"):
        loadout[(t, lo, pt)] = did
    items = q("""select i.DesignID, i.Condition, i.BuildWork, d.BuildWorkMax from Parts_Items i
                 join Parts_Designs d on d.DesignID=i.DesignID where d.TeamID=? and d.PartType>=3""", team)
    engineers_on = dict(q("""select p.DesignID, a.ActiveStaff from Parts_Projects p
                             join SubTeam_Assignments a on a.AssignmentID=p.AssignmentID
                             where a.TeamID=? and p.NewDesign=1""", team))
    building = {d[0] for d in my_designs} & {r[0] for r in q("select DesignID from Parts_Projects where NewDesign=0")}
    expertise = {(pt, st): e for pt, st, e in
                 q("select PartType, PartStat, Expertise from Parts_TeamExpertise where TeamID=?", team)}

    newest, newest_done = {}, {}
    for d in my_designs:
        newest[d[1]] = d
        if d[6] >= 0:
            newest_done[d[1]] = d
    design_no = {d[0]: d[2] for d in my_designs}

    def rival_best(t, pt, st):
        vals = [stat_values.get(loadout.get((t, lo, pt)), {}).get(st) for lo in (1, 2)]
        vals = [v for v in vals if v is not None]
        return max(vals) if vals else None

    part_stats, design_rows, stock_rows, summary_rows, flags, part_info = [], [], [], [], [], {}
    for pt, label in PART_LABEL.items():
        if pt not in newest:
            continue
        nd = newest[pt]
        car = {lo: loadout.get((team, lo, pt)) for lo in (1, 2)}
        for st in sorted(stat_values.get(nd[0], {})):
            riv = [v for v in (rival_best(t, pt, st) for t in rivals) if v is not None]
            mine = stat_values[nd[0]][st]
            part_stats.append([label, PART_STAT_LABEL.get(st, str(st)),
                               stat_values.get(car[1], {}).get(st), stat_values.get(car[2], {}).get(st), mine,
                               median(riv), max(riv) if riv else None, None,
                               1 + sum(1 for v in riv if v > mine), expertise.get((pt, st)), st])
        in_stock = lambda did: [i for i in items if i[0] == did and i[2] >= i[3]]
        done = newest_done.get(pt)
        note = []
        ready = len(in_stock(done[0])) if done else 0
        if done and not ready:
            note.append("Design #%d is finished but %s" % (done[2], "the first one is still being built"
                                                           if done[0] in building else "has never been built"))
        if done and ready:
            spare = ready - sum(1 for lo in (1, 2) if car[lo] == done[0])
            for lo in (1, 2):
                if car[lo] and design_no.get(car[lo], 0) < done[2]:
                    note.append("Car %d is on design #%d; %s" % (lo, design_no.get(car[lo], 0),
                                "a spare #%d is in stock" % done[2] if spare > 0 else "no spare #%d, build one" % done[2]))
        if nd[6] < 0:
            note.append("Design #%d is %d%% complete" % (nd[2], round(100 * nd[3] / nd[4])))
        part_info[pt] = dict(label=label, newest=nd, done=done, ready=ready, car=car,
                             building=bool(done and done[0] in building))
        summary_rows.append([label, nd[2], "In design" if nd[6] < 0 else "Finished",
                             design_no.get(car[1]), design_no.get(car[2]),
                             len(in_stock(done[0])) if done else 0, None, None, None, "; ".join(note)])
        flags += ["%s: %s" % (label, n) for n in note if "complete" not in n]

    for d in my_designs:
        did, pt = d[0], d[1]
        own = [i for i in items if i[0] == did]
        fitted = [str(lo) for lo in (1, 2) if loadout.get((team, lo, pt)) == did]
        design_rows.append([PART_LABEL[pt], d[2], "In design" if d[6] < 0 else "Finished",
                            min(1.0, d[3] / d[4]) if d[4] else None, d[5], d[6] if d[6] >= 0 else None,
                            d[7], d[8], engineers_on.get(did), d[9], len(own),
                            max((i[1] for i in own), default=None), " and ".join(fitted) or None])
        for i in sorted(own, key=lambda i: -i[1]):
            stock_rows.append([PART_LABEL[pt], d[2], i[1], "Being built" if i[2] < i[3] else "Ready",
                               "Yes" if fitted else "No", None])
    D.update(part_stats=part_stats, summary=summary_rows, designs=design_rows, stock=stock_rows)

    # ---- car areas
    areas = []
    for label, key, st in AREAS:
        mine = [r[4] for r in part_stats if r[10] == st]
        fitted = [max(v for v in (r[2], r[3]) if v is not None) for r in part_stats
                  if r[10] == st and (r[2] is not None or r[3] is not None)]
        totals = []
        for t in rivals:
            vals = [rival_best(t, pt, st) for pt in PART_LABEL if st in stat_values.get(newest.get(pt, [None])[0], {})]
            vals = [v for v in vals if v is not None]
            if vals:
                totals.append(sum(vals) / len(vals))
        if not mine:
            continue
        ours = sum(mine) / len(mine)
        areas.append([label, PART_STAT_LABEL[st], sum(fitted) / len(fitted) if fitted else None, ours,
                      median(totals), None, 1 + sum(1 for v in totals if v > ours), key])
    D["areas"] = areas
    rank_by_key = {a[7]: a[6] for a in areas if a[7]}

    # ---- races
    races = []
    for rid, rday, name in q("""select r.RaceID, r.Day, t.Name from Races r join Races_Tracks t on t.TrackID=r.TrackID
                                where r.SeasonID=? and r.State=0 order by r.Day""", season):
        req = q("""select s.CarStat, s.IsCrucial from Parts_RecommendedTrackStats s join Races r on r.TrackID=s.TrackID
                   where r.RaceID=?""", rid)
        crucial = [CAR_STAT_LABEL.get(cs, str(cs)) for cs, cr in req if cr]
        also = [CAR_STAT_LABEL.get(cs, str(cs)) for cs, cr in req if not cr]
        ranks = [rank_by_key[k] for k in crucial if k in rank_by_key]
        races.append([words(name), rday - day, rday, ", ".join(crucial) or None,
                      ", ".join(also) or None, max(ranks) if ranks else None])
    D["races"] = races

    # ---- overview numbers
    balance = q("select Balance from Finance_TeamBalance where TeamID=?", team)[0][0]
    res = q("select max(WindTunnelBlocksRemaining), max(CfdBlocksRemaining) from Parts_Resources")[0]
    period_end = q("select min(Day) from Seasons_Deadlines where SeasonID=? and DeadlineType=5 and Day>=?", season, day)[0][0]
    eng_total = (q("select TotalStaff from SubTeam_Ownership where TeamID=? and SubTeamType=0", team) or [[0]])[0][0]
    eng_used = q("select coalesce(sum(ActiveStaff),0) from SubTeam_Assignments where TeamID=? and SubTeamType=0", team)[0][0]
    scouts_total = (q("select TotalStaff from SubTeam_Ownership where TeamID=? and SubTeamType=1", team) or [[0]])[0][0]
    scouts_used = q("select coalesce(sum(ActiveStaff),0) from SubTeam_Assignments where TeamID=? and SubTeamType=1", team)[0][0]

    def building_effect(btype, effect):
        r = q("""select e.Value, h.DegradationValue from Buildings_HQ h join Buildings_Effects e on e.BuildingID=h.BuildingID
                 where h.TeamID=? and h.BuildingType=? and e.EffectID=?""", team, btype, effect)
        return r[0] if r else (None, None)
    design_slots = building_effect(2, 8)[0] or 0
    build_slots = building_effect(3, 9)[0] or 0
    factory_bonus = building_effect(3, 4)[0] or 0
    designs_running = sum(1 for d in my_designs if d[6] < 0)
    builds_running = q("""select count(*) from Parts_Projects p join Parts_Designs d on d.DesignID=p.DesignID
                          where p.NewDesign=0 and d.TeamID=?""", team)[0][0]
    standing = q("select Position, Points from Races_TeamStandings where SeasonID=? and TeamID=? and RaceFormula=1", season, team)
    D["overview"] = [
        ("In-game date", day, "date"), ("Season", season, None),
        ("Championship position", standing[0][0] if standing else None, None),
        ("Points", standing[0][1] if standing else None, None),
        ("Balance", balance, MONEY),
        ("Next race", races[0][0] if races else None, None),
        ("Days to next race", races[0][1] if races else None, None),
        ("Wind tunnel hours left this period", res[0], None),
        ("CFD hours left this period", res[1], None),
        ("Days until testing hours reset", (period_end - day + 1) if period_end else None, None),
        ("Engineers idle", eng_total - eng_used, None), ("Engineers owned", eng_total, None),
        ("Design slots free", design_slots - designs_running, None),
        ("Manufacturing slots free", build_slots - builds_running, None),
        ("Scouts idle", scouts_total - scouts_used, None),
    ]
    cap = (q("select CurrentValue from Regulations_Enum_Changes where Name='SpendingCap'") or [[None]])[0][0]
    cap_spent = -q("""select coalesce(sum(Value),0) from Finance_Transactions
                      where TeamID=? and Day>=? and AffectsCostCap=1 and Value<0""", team, season_start)[0][0]
    season_days = (dt.date(season + 1, 1, 1) - dt.date(season, 1, 1)).days
    projected = cap_spent / max(1, day - season_start + 1) * season_days
    if cap:
        D["overview"] += [("Cost cap", cap, MONEY), ("Spent against the cap this season", cap_spent, MONEY),
                          ("Projected at season end, at this rate", projected, MONEY),
                          ("Cap headroom at season end", cap - projected, MONEY)]
        if projected > 0.95 * cap:
            flags.append("Cost cap: projected spend of $%.1fM is within 5%% of the $%.1fM cap" % (projected / 1e6, cap / 1e6))
    D["cap"] = dict(cap=cap, spent=cap_spent, projected=projected, balance=balance)
    D["factory_bonus"] = factory_bonus
    if eng_total - eng_used > 0:
        flags.append("%d of %d engineers are not assigned to any project" % (eng_total - eng_used, eng_total))
    if (res[0] or res[1]) and period_end:
        flags.append("%d wind tunnel and %d CFD hours unused; they reset in %d days" % (res[0], res[1], period_end - day + 1))
    if scouts_total - scouts_used > 0:
        flags.append("%d scouts idle" % (scouts_total - scouts_used))

    # ---- spares: how many good units to hold beyond the two on the cars
    races_done = q("select count(*) from Races where SeasonID=? and State=2", season)[0][0]
    rate_normal = settings[5] * (1 + factory_bonus)
    part_of = {d[0]: d[1] for d in my_designs}
    spares, spare_plan = [], {}
    for pt, info in part_info.items():
        done = info["done"]
        if not done:
            continue
        work, cost = q("select BaseManufactureWork, BaseManufactureCost from Parts_Enum_Type where Value=?", pt)[0]
        ready_all = [i for i in items if part_of.get(i[0]) == pt and i[2] >= i[3]]
        ready_new = [i for i in ready_all if i[0] == done[0]]
        good = sum(1 for i in ready_new if i[1] >= settings[2])
        lost = max(0, sum(d[9] for d in my_designs if d[1] == pt) - len(ready_all))
        build_days = int(-(-work // rate_normal))
        first = races[0][1] if races else 0
        cover = max(1, sum(1 for r in races if r[1] <= first + build_days + 2))
        target = max(settings[9], int(-(-(lost / races_done * cover) // 1)) if races_done else 0)
        on_new = sum(1 for lo in (1, 2) if info["car"][lo] == done[0])
        note = []
        if info["newest"][6] < 0:
            note.append("Design #%d is %d%% done: build spares of that once it lands" % (
                info["newest"][2], round(100 * info["newest"][3] / info["newest"][4])))
        if len(ready_new) > good:
            note.append("%d unit(s) of #%d are worn" % (len(ready_new) - good, done[2]))
        spares.append([info["label"], done[2], on_new, len(ready_new), good, len(ready_all) - len(ready_new), lost,
                       None, cover, None, None, None, cost, None, "; ".join(note) or None])
        spare_plan[pt] = dict(target=target, good=good, on_new=on_new, lost=lost, cover=cover, cost=cost,
                              build_days=build_days, to_build=max(0, 2 + target - good))
    D.update(spares=spares, races_done=races_done)

    # ---- power unit: engine, ERS and gearbox wear against the season allocation, and when a grid penalty is due
    PU = {0: ("Engine", "EngineLimit", "ExceededEngineLimit"), 1: ("ERS", "ErsLimit", "ExceededERSLimit"),
          2: ("Gearbox", "GearboxLimit", "ExceededGearboxLimit")}
    reg = dict(q("select Name, CurrentValue from Regulations_Enum_Changes"))
    pens = {n: (g, m) for n, g, m in q("select Name, GridPenalty, FirstOffenceModifier from Races_Enum_PenaltyType")}
    thresh = {v: (mi, ma, cost) for v, mi, ma, cost in q(
        "select Value, DamageMinorThreshold, DamageMajorThreshold, BaseManufactureCost from Parts_Enum_Type where Value<=2")}
    units = q("""select i.ItemID, d.PartType, i.ManufactureNumber, i.Condition, i.AssociatedCar, i.BuildWork, d.BuildWorkMax
                 from Parts_Items i join Parts_Designs d on d.DesignID=i.DesignID
                 where d.TeamID=? and d.PartType<=2 and d.ValidFrom=?""", team, season)
    fitted_pu = {(lo, pt): item for lo, pt, item in q("select LoadoutID, PartType, ItemID from Parts_CarLoadout where TeamID=? and PartType<=2", team)}
    wtype = dict(q("select Day, WeekendType from Races where SeasonID=? and State=0", season))
    weights = [1.25 if wtype.get(r[2], 0) else 1.0 for r in races]          # a sprint weekend runs more laps
    avg_grid = q("select avg(StartingPos) from Races_Results where Season=? and TeamID=?", season, team)[0][0]
    hard_to_pass = ("Monaco", "Hungaroring", "Marina Bay", "Zandvoort", "Imola")
    pu_summary, pu_units, pu_new, pu_bundles = [], [], [], []
    plan_cells = {}

    def simulate(ready, used, limit, wear, major, pen_key, forced=None):
        """Run the healthiest unit each weekend. A new unit comes in at race `forced`, or when even the healthiest
        would end a weekend in major damage. Returns the cell text per race and the new units taken."""
        cond = {u[2]: u[3] for u in ready}
        total, cells, taken = used, [], []
        for i, w in enumerate(weights):
            pick = max(cond, key=cond.get) if cond else None
            if i != forced and pick is not None and cond[pick] - wear * w >= major:
                name = "#%d" % pick if isinstance(pick, int) else "new unit"
                cells.append("%s: %d%% to %d%%" % (name, round(cond[pick] * 100), round((cond[pick] - wear * w) * 100)))
                cond[pick] -= wear * w
            else:
                total += 1
                g, mod = pens.get(pen_key, (0, 1))
                places = 0 if total <= limit else g * (mod if total == limit + 1 else 1)
                cells.append("NEW: %s" % ("%d-place grid penalty" % places if places else "no penalty"))
                cond["new%d" % total] = 1 - wear * w
                taken.append((i, places))
        return cells, taken

    for car in (1, 2):
        comp = {}
        for pt, (label, limit_key, pen_key) in PU.items():
            minor, major, cost = thresh[pt]
            limit = reg.get(limit_key)
            pool = [u for u in units if u[1] == pt and u[4] == car]
            ready = [u for u in pool if u[5] >= u[6]]
            wear = sum(1 - u[3] for u in ready) / races_done if races_done else 0.12
            fit = next((u for u in ready if u[0] == fitted_pu.get((car, pt))), None)
            best = max(ready, key=lambda u: u[3]) if ready else None
            for u in sorted(ready, key=lambda u: u[2]):
                state = "Good" if u[3] >= minor else "Minor damage" if u[3] >= major else "Major damage"
                pu_units.append(["Car %d" % car, label, "#%d" % u[2], u[3], state, "Yes" if fit and u[0] == fit[0] else None])
            cells, taken = simulate(ready, len(pool), limit, wear, major, pen_key)
            comp[pt] = dict(label=label, limit=limit, pool=pool, ready=ready, wear=wear, fit=fit, best=best, major=major,
                            cost=cost, pen_key=pen_key, cells=cells, taken=taken)
            if fit and fit[3] < major:
                flags.append("Car %d %s on the car is in major damage (%d%%)" % (car, label.lower() if label != "ERS" else "ERS", fit[3] * 100))
        # one trip to the back of the grid: take every new part this car needs in the same weekend,
        # no later than the first one is due, and at a track where passing is possible if there is one
        due = [cp["taken"][0][0] for cp in comp.values() if cp["taken"]]
        bundle_at = None
        if due:
            first_due = min(due)
            easy = [i for i in range(first_due + 1) if not any(h.lower() in races[i][0].lower() for h in hard_to_pass)]
            bundle_at = easy[-1] if easy else first_due     # as late as possible, but not at a track where passing is hard
            for cp in comp.values():
                if cp["taken"]:
                    cp["cells"], cp["taken"] = simulate(cp["ready"], len(cp["pool"]), cp["limit"], cp["wear"], cp["major"],
                                                      cp["pen_key"], forced=bundle_at)
            parts = [cp for cp in comp.values() if cp["taken"]]
            places = sum(cp["taken"][0][1] for cp in parts)
            later = [(cp["label"], races[i][0]) for cp in parts for i, _p in cp["taken"][1:]]
            pu_bundles.append(dict(car=car, race=races[bundle_at][0], days=races[bundle_at][1], places=places,
                                   parts=[cp["label"] for cp in parts], cost=sum(cp["cost"] for cp in parts), later=later,
                                   first_due=races[first_due][0]))
        for pt, cp in comp.items():
            for i, cell in enumerate(cp["cells"]):
                plan_cells[(i, car, pt)] = cell
            fit, best = cp["fit"], cp["best"]
            need = cp["wear"] * sum(weights)
            usable = sum(max(0, u[3] - cp["major"]) for u in cp["ready"])
            if cp["taken"]:
                verdict = "New unit at %s, taken with the car's other new parts so the penalty is served once" % races[cp["taken"][0][0]][0]
            elif fit and best and best[3] - fit[3] >= 0.05:
                verdict = "Fit #%d now (%d%%, against %d%% on the car): it is already in your allocation, so no penalty" % (
                    best[2], best[3] * 100, fit[3] * 100)
            else:
                verdict = "Enough left: rotate the units you have, healthiest first"
            pu_summary.append(["Car %d" % car, cp["label"], "%d of %d" % (len(cp["pool"]), cp["limit"]), "#%d" % fit[2] if fit else None,
                               fit[3] if fit else None, "#%d" % best[2] if best else None, best[3] if best else None,
                               cp["wear"], need, usable, verdict])
    pu_plan = [[r[0] + (" (sprint)" if wtype.get(r[2], 0) else ""), r[1]] +
               [plan_cells.get((i, car, pt)) for car in (1, 2) for pt in (0, 1, 2)] for i, r in enumerate(races)]
    D["pu"] = dict(summary=pu_summary, plan=pu_plan, units=pu_units, bundles=pu_bundles, avg_grid=avg_grid,
                   limits="Engine %s, ERS %s, gearbox %s per car" % (reg.get("EngineLimit"), reg.get("ErsLimit"), reg.get("GearboxLimit")),
                   penalties="First engine or ERS over the limit: %d places, each one after: %d. Gearbox: %d places each" % (
                       pens["ExceededEngineLimit"][0] * pens["ExceededEngineLimit"][1], pens["ExceededEngineLimit"][0],
                       pens["ExceededGearboxLimit"][0] * pens["ExceededGearboxLimit"][1]),
                   damage="Engine: minor damage below %d%%, major below %d%%. ERS and gearbox: %d%% and %d%%" % (
                       thresh[0][0] * 100, thresh[0][1] * 100, thresh[1][0] * 100, thresh[1][1] * 100),
                   hard=[r[0] for r in races if any(h.lower() in r[0].lower() for h in hard_to_pass)])

    # ---- manufacturing
    D["manufacturing"] = [[PART_LABEL[v], work, cost] for v, work, cost in
                          q("select Value, BaseManufactureWork, BaseManufactureCost from Parts_Enum_Type where Value>=3 order by Value")]

    # ---- staff
    staff, ratings = [], dict(q("select StaffID, avg(Val) from Staff_PerformanceStats group by StaffID"))
    my_staff = q("""select k.StaffID, b.FirstName, b.LastName, g.StaffType, k.PosInTeam, k.EndSeason, k.Salary, s.Mentality,
                    s.MentalityOpinion from Staff_Contracts k join Staff_BasicData b on b.StaffID=k.StaffID
                    join Staff_GameData g on g.StaffID=k.StaffID left join Staff_State s on s.StaffID=k.StaffID
                    where k.TeamID=? and k.ContractType=0 order by g.StaffType, k.PosInTeam""", team)
    all_teams = dict(q("select TeamID, TeamName from Teams"))
    for sid, fn, ln, stype, pos, ends, sal, ment, mood in my_staff:
        op = dict(q("select Status, Value from Staff_Mentality_Statuses where StaffID=?", sid))
        # what is already agreed or under way beyond the current contract
        nxt = None
        future = q("select TeamID, EndSeason, Salary from Staff_Contracts where StaffID=? and ContractType=3", sid)
        talks = q("select 1 from Staff_Contracts where StaffID=? and ContractType=1 and TeamID=?", sid, team)
        if future and future[0][0] == team:
            nxt = "Renewed to %d" % future[0][1]
        elif future:
            nxt = "Leaving for %s" % all_teams.get(future[0][0], "another team")
        elif talks:
            nxt = "In talks"
        role = ROLE_LABEL.get(stype, "Staff")
        if stype == 0:
            role = "Reserve driver" if pos >= 3 else "Driver %d" % pos
        elif stype == 2:
            role = "Race Engineer %d" % pos
        staff.append([person_name(fn, ln), role, ment, MOOD.get(mood), ends, sal, ratings.get(sid),
                      op.get(0), op.get(2), op.get(3), nxt])
        if nxt and nxt.startswith("Leaving"):
            flags.append("%s (%s): %s at the end of the season" % (person_name(fn, ln), role, nxt[0].lower() + nxt[1:]))
        elif ends == season and not nxt:
            flags.append("%s (%s): contract ends this season, no renewal agreed" % (person_name(fn, ln), role))
        if ment is not None and ment < settings[1]:
            flags.append("%s (%s): mentality %d" % (person_name(fn, ln), role, ment))
    for sid, fn, ln, stype, pos, ends, sal in q(
            """select k.StaffID, b.FirstName, b.LastName, g.StaffType, k.PosInTeam, k.EndSeason, k.Salary
               from Staff_Contracts k join Staff_BasicData b on b.StaffID=k.StaffID join Staff_GameData g on g.StaffID=k.StaffID
               where k.TeamID=?1 and k.ContractType=3 and k.StaffID not in
               (select StaffID from Staff_Contracts where TeamID=?1 and ContractType=0)""", team):
        role = ROLE_LABEL.get(stype, "Staff")
        if stype == 0:
            role = "Reserve driver" if pos >= 3 else "Driver %d" % pos
        ment = q("select Mentality from Staff_State where StaffID=?", sid)
        for r in staff:   # whoever holds that seat now and has nothing agreed is being replaced, not renewed
            if r[1] == role and r[4] == season and not r[10]:
                r[10] = "Replaced by %s" % person_name(fn, ln)
                flags[:] = [f for f in flags if not f.startswith("%s (%s): contract ends" % (r[0], r[1]))]
        staff.append([person_name(fn, ln), role + " from next season", ment[0][0] if ment else None, None, ends, sal,
                      ratings.get(sid), None, None, None, "Signed, joins next season"])

    # seats with nobody in them right now
    def seat_label(stype, pos):
        if stype == 0:
            return "Reserve driver" if pos >= 3 else "Driver %d" % pos
        return "Race Engineer %d" % pos if stype == 2 else ROLE_LABEL[stype]
    filled_now = {(stype, min(pos, 3) if stype == 0 else pos) for _sid, _fn, _ln, stype, pos, *_ in my_staff}
    filled_next = {seat_label(stype, pos) for stype, pos in q(
        """select g.StaffType, k.PosInTeam from Staff_Contracts k join Staff_GameData g on g.StaffID=k.StaffID
           where k.TeamID=? and k.ContractType=3""", team)}
    vacancies = []
    for seat in [(0, 1), (0, 2), (0, 3), (1, 1), (2, 1), (2, 2), (3, 1), (4, 1)]:
        if seat in filled_now:
            continue
        label = seat_label(*seat)
        covered = label in filled_next
        staff.append(["(vacant)", label, None, None, None, None, None, None, None, None,
                      "Filled from next season" if covered else "Nobody signed"])
        flags.append("No %s: the seat is empty %s" % (label.lower(), "until next season" if covered else "now and for next season"))
        vacancies.append((seat[0], label, covered))
    D["staff"] = staff

    # ---- facilities
    fac = []
    for btype, bname, lvl, cond, refurb, upkeep in q(
            """select h.BuildingType, t.Name, b.UpgradeLevel, h.DegradationValue, b.RefurbishCost, b.UpkeepCost
               from Buildings_HQ h join Buildings b on b.BuildingID=h.BuildingID join Building_Enum_Types t on t.Type=h.BuildingType
               where h.TeamID=? order by h.BuildingType""", team):
        grid = q("""select avg(b.UpgradeLevel) from Buildings_HQ h join Buildings b on b.BuildingID=h.BuildingID
                    where h.BuildingType=? and h.TeamID<>?""", btype, team)[0][0]
        nxt = q("select BuildingID, ConstructionCost from Buildings where Type=? and UpgradeLevel=?", btype, lvl + 1)
        gives = None
        if nxt:
            cur = dict(q("""select n.Name, e.Value from Buildings_Effects e join Building_Enum_Effects n on n.Effect=e.EffectID
                            join Buildings_HQ h on h.BuildingID=e.BuildingID where h.TeamID=? and h.BuildingType=?""", team, btype))
            new = q("""select n.Name, e.Value from Buildings_Effects e join Building_Enum_Effects n on n.Effect=e.EffectID
                       where e.BuildingID=?""", nxt[0][0])
            gives = "; ".join("%s %s to %s" % (words(n), "%g" % (cur.get(n) or 0), "%g" % v)
                              for n, v in new if v is not None and v != cur.get(n))
        fac.append([bname, lvl, cond if lvl > 0 else None, None, grid, refurb or None, nxt[0][1] if nxt else None,
                    upkeep or None, gives, None])
        if lvl > 0 and cond <= 0.5:
            flags.append("%s is at %d%% condition" % (bname, round(cond * 100)))
    D["facilities"] = fac

    # ---- scouting
    weakest = {}
    for sid, fn, ln, stype, pos, *_ in my_staff:
        if stype == 0 and pos >= 3:
            continue
        r = ratings.get(sid)
        if r is not None and (stype not in weakest or r < weakest[stype][1]):
            weakest[stype] = (person_name(fn, ln), r)
    scouted = {r[0] for r in q("select StaffID from Scouting where TeamID=?", team)}
    team_name = dict(q("select TeamID, TeamName from Teams"))
    licence = dict(q("select StaffID, HasSuperLicense from Staff_DriverData"))
    signed_with_us = {r[0] for r in q("select StaffID from Staff_Contracts where TeamID=? and ContractType=3", team)}
    cands = []
    for sid, fn, ln, stype, dob, opin, thr in q(
            """select g.StaffID, b.FirstName, b.LastName, g.StaffType, b.DOB_ISO, p.TeamOpinionRating, p.ContractAcceptanceThreshold
               from Staff_GameData g join Staff_BasicData b on b.StaffID=g.StaffID
               join Staff_ContractPatience p on p.StaffID=g.StaffID and p.TeamID=? where g.Retired=0""", team):
        if stype not in weakest or opin is None or opin < settings[3] or sid not in ratings:
            continue
        k = q("select TeamID, EndSeason, Salary from Staff_Contracts where StaffID=? and ContractType=0 order by PosInTeam", sid)
        if (k and k[0][0] == team) or sid in signed_with_us:
            continue
        floor = weakest[stype][1] if stype == 0 else weakest[stype][1] - settings[4]
        if ratings[sid] < floor or (stype == 0 and ratings[sid] <= floor):
            continue
        age = season - int(dob[:4]) if dob else None
        cands.append([ROLE_LABEL[stype], person_name(fn, ln), ratings[sid], weakest[stype][1], None, opin, thr, age,
                      team_name.get(k[0][0], "Free agent") if k else "Free agent", k[0][1] if k else None,
                      k[0][2] if k else None, "Yes" if sid in scouted else "No", weakest[stype][0], stype,
                      ("Yes" if licence.get(sid) else "No") if stype == 0 else None,
                      next(("Signed with %s" % team_name.get(f[0]) for f in
                            q("select TeamID from Staff_Contracts where StaffID=? and ContractType=3", sid) if f[0] != team), None)])
    cands.sort(key=lambda r: (r[13], -r[2]))
    D["scouting"] = [r[:13] + r[14:16] for r in cands]

    # ---- pit crew
    crew = []
    for name, val in q("""select e.Name, p.Val from Staff_PitCrew_PerformanceStats p
                          join Staff_Enum_PerformanceStatTypes e on e.Value=p.StatID where p.TeamID=?""", team):
        others = [r[0] for r in q("""select p.Val from Staff_PitCrew_PerformanceStats p join Staff_Enum_PerformanceStatTypes e
                                     on e.Value=p.StatID where e.Name=? and p.TeamID<>? and p.TeamID in (%s)"""
                                  % ",".join(map(str, f1_teams)), name, team)]
        crew.append([words(name), val, sum(others) / len(others) if others else None,
                     max(others) if others else None, None,
                     1 + sum(1 for v in others if (v < val if name == "Fatigue" else v > val))])
    stages = []
    for name, mine, mine_err, oth, oth_err in q(
            """select s.Name, avg(case when t.TeamID=?1 then t.Duration end), avg(case when t.TeamID=?1 then t.IncidentDelay end),
               avg(case when t.TeamID<>?1 then t.Duration end), avg(case when t.TeamID<>?1 then t.IncidentDelay end)
               from Races_PitStopTimings t join Staff_Enum_PitStopStage s on s.Value=t.PitStopStage
               where t.SeasonID=?2 and t.PitStopStage<=6 and t.TeamID in (%s) group by t.PitStopStage order by t.PitStopStage"""
            % ",".join(map(str, f1_teams)), team, season):
        stages.append([words(name), mine, mine_err, oth, oth_err, None])
    crew.sort(key=lambda r: r[0] == "Fatigue")   # fatigue is on a different scale: keep it last
    D.update(crew=crew, stages=stages)

    # ---- drivers
    drivers = [(sid, person_name(fn, ln)) for sid, fn, ln, stype, *_ in my_staff if stype == 0]
    stat_names = [r[0] for r in q("""select distinct e.Name from Staff_PerformanceStats p join Staff_Enum_PerformanceStatTypes e
                                     on e.Value=p.StatID where p.StaffID=? order by e.Name""", drivers[0][0])] if drivers else []
    drv = []
    for n in stat_names:
        row = [n]
        for sid, _ in drivers:
            v = q("""select p.Val from Staff_PerformanceStats p join Staff_Enum_PerformanceStatTypes e on e.Value=p.StatID
                     where p.StaffID=? and e.Name=?""", sid, n)
            row.append(v[0][0] if v else None)
        drv.append(row)
    res_rows = []
    for sid, name in drivers:
        r = q("""select count(*), avg(FinishingPos), min(FinishingPos), avg(StartingPos), sum(DNF), sum(Points)
                 from Races_Results where Season=? and DriverID=?""", season, sid)[0]
        res_rows.append([name] + list(r))
    D.update(driver_names=[n for _, n in drivers], driver_stats=drv, driver_results=res_rows)

    # ---- finance
    D["finance"] = [[words(n), v] for n, v in q(
        """select e.Name, sum(f.Value) from Finance_Transactions f join Finance_Enum_TransactionType e
           on e.TransactionType=f.TransactionType where f.TeamID=? and f.Day>=? group by 1 having sum(f.Value)<>0 order by 2""",
        team, season_start)]
    # ---- all-female team project (optional)
    women, women_steps = [], []
    if settings[10]:
        seat_holder = {}
        for r in staff:
            base = r[1].replace(" from next season", "")
            base = "Race Engineer" if base.startswith("Race Engineer") else "Driver" if base in ("Driver 1", "Driver 2") else base
            seat_holder.setdefault(base, []).append(r)
        formula_of = dict(q("select TeamID, Formula from Teams"))
        standings = {sid: (f, pos) for sid, f, pos in q(
            "select DriverID, RaceFormula, Position from Races_DriverStandings where SeasonID=?", season)}
        for sid, fn, ln, stype in q("""select g.StaffID, b.FirstName, b.LastName, g.StaffType from Staff_GameData g
                                       join Staff_BasicData b on b.StaffID=g.StaffID
                                       where b.Gender=1 and g.Retired=0 and g.StaffType<=4"""):
            name, role = person_name(fn, ln), ROLE_LABEL[stype]
            rating = ratings.get(sid)
            cur = q("select TeamID, PosInTeam, EndSeason from Staff_Contracts where StaffID=? and ContractType=0", sid)
            fut = q("select TeamID, EndSeason from Staff_Contracts where StaffID=? and ContractType=3", sid)
            opin = q("select TeamOpinionRating from Staff_ContractPatience where StaffID=? and TeamID=?", sid, team)
            opin = opin[0][0] if opin else None
            drv = q("select Improvability, HasSuperLicense, HasRacedEnoughToJoinF1 from Staff_DriverData where StaffID=?", sid)
            f1 = [k for k in cur if formula_of.get(k[0]) == 1]
            main = f1[0] if f1 else cur[0] if cur else None
            mine = any(k[0] == team for k in cur) or any(k[0] == team for k in fut)
            if not main:
                attached = "Free agent"
            elif stype == 0 and formula_of.get(main[0]) == 1:
                attached = "%s %s" % (all_teams.get(main[0]), "race seat" if main[1] <= 2 else "reserve" if main[1] == 3 else "affiliate")
            else:
                attached = all_teams.get(main[0])
            ordinal = lambda n: "%d%s" % (n, "th" if 10 <= n % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th"))
            racing = "F%d, %s in the standings" % (standings[sid][0], ordinal(standings[sid][1])) if sid in standings else None
            elsewhere = [k for k in fut if k[0] != team]
            holders = seat_holder.get(role, [])
            current_best = max((h[6] for h in holders if h[6] is not None), default=None)
            opening = next((h for h in holders if h[0] == "(vacant)" or (h[10] or "").startswith("Leaving")), None)
            reserve_open = any(h[0] == "(vacant)" for h in seat_holder.get("Reserve driver", []))
            available = not main or main[2] == season
            if mine:
                step, order = "On your team", 0
            elif elsewhere:
                step, order = "Signed with %s to %d: revisit then" % (all_teams.get(elsewhere[0][0]), elsewhere[0][1]), 5
            elif opin is None or opin < settings[3]:
                step, order = "Not interested yet (opinion %s): revisit as the team improves" % opin, 6
            elif stype == 0 and drv and drv[0][1] and reserve_open:
                step, order = "Has a super licence: sign as reserve driver now", 1
            elif stype == 0:
                step = ("Sign as an affiliate now and develop: no super licence yet" if not main else
                        "Approach as an affiliate now (contract ends this season): no super licence yet" if available else
                        "Shortlist as an affiliate (contract runs to %d): no super licence yet" % main[2])
                order = 2 if available else 4
            elif opening and available:
                step, order = "Sign for the %s seat: %s" % (role, "it is vacant" if opening[0] == "(vacant)"
                                                              else opening[0] + " is leaving"), 1
            else:
                until = min((h[4] for h in holders if h[4]), default=None)
                step = "Shortlist: no %s seat opens before %s" % (role, until + 1 if until else "next season")
                order = 3 if available else 4
            potential = (rating or 0) + (drv[0][0] / 4 if drv else 0)
            women.append([role, name, rating, current_best, None, drv[0][0] if drv else None,
                          ("Yes" if drv[0][1] else "No") if drv else None, ("Yes" if drv[0][2] else "No") if drv else None,
                          racing, opin, attached, main[2] if main else None, step, stype, order, potential])
        women.sort(key=lambda r: (r[13], r[14], -r[15]))
        women_steps = [r for r in women if r[14] in (1, 2)]
    D["women"] = [r[:13] for r in women]
    D["women_on_team"] = sum(1 for r in women if r[14] == 0)

    # ---- board: this season's target, how confidence moves, and where to finish so next year's target stays reachable
    D["board"] = None
    tgt = q("select TargetPos from Board_SeasonObjectives where TeamID=? and SeasonID=?", team, season)
    if tgt:
        target = tgt[0][0]
        pos_pts = {p: a for p, a, _b in q("select Pos, PtsFromConstructor, PtsFromBestDriver from Board_TeamRating_ChampionshipPosPts")}
        last_pos = max(pos_pts)
        ptsf = lambda p: pos_pts.get(min(p, last_pos), 0)      # places past the table score as the last one
        mult = dict(q("select NumSeasonsAgo, Multiplier from Board_TeamRating_RelevancyMultipliers"))
        past = dict(q("select SeasonID, Position from Races_TeamStandings where TeamID=? and RaceFormula=1 and SeasonID<?",
                      team, season))
        cols = "PtsFromConstructorResults, PtsFromDriverResults, PtsFromSeasonsEntered, PtsFromChampionshipsWon, PtsFromNewTeamHype"
        mine_r = (q("select %s from Board_TeamRating where TeamID=? and SeasonID=?" % cols, team, season) or [(0, 0, 0, 0, 0)])[0]
        rating_now = sum(mine_r)
        rivals_r = {t: sum(r) for t, *r in q("select TeamID, %s from Board_TeamRating where SeasonID=? and TeamID<>?" % cols,
                                              season, team) if t in f1_teams}
        bands = q("select Name, MinTeamRating, MaxTeamRating, ExpectedRank from Board_Enum_ObjectiveTypes order by MinTeamRating")
        band_of = lambda r: next(("%s (expects about P%d)" % (words(n), er) for n, lo, hi, er in bands if lo <= r <= hi), None)
        pay = dict(q("select SeasonExpectation, Budget from Board_Payments"))
        conf = (q("select Confidence from Board_Confidence where Season=?", season) or [[None]])[0][0]
        levels = q("select Name, UpperBound, ReviewBonus from Board_Enum_ConfidenceLevels order by UpperBound")
        level = next((words(n) for n, ub, _b in levels if conf is not None and conf <= ub), None)
        lt = q("""select t.Name, o.TargetEndYear from Board_Objectives o join Board_Enum_ObjectiveTypes t on t.Value=o.Type
                  where o.TeamID=? and o.State=0 order by o.StartYear desc""", team)
        now_pos = standing[0][0] if standing else None
        finish = []
        for p_ in range(1, len(f1_teams) + 1):
            cons_next = ptsf(p_) * mult.get(1, 0) + sum(ptsf(past[season - k]) * mult.get(k + 1, 0)
                                                        for k in (1, 2, 3) if season - k in past)
            r_next = cons_next + sum(mine_r[1:])
            rank = 1 + sum(1 for v in rivals_r.values() if v > r_next)
            finish.append([p_, "Met" if p_ <= target else "Failed", round(r_next), band_of(r_next), rank, pay.get(rank),
                           "Aim here" if p_ == target else None])
        race_rules = [["%s than expected by %g to %g places" % ("Worse" if hi <= 0 else "Better", min(abs(lo), abs(hi)),
                                                                      max(abs(lo), abs(hi))), mod, words(n)] for lo, hi, mod, n in q(
            """select m.MinPosDelta, m.MaxPosDelta, m.ConfidenceMod, e.Name from Board_Confidence_RaceMods m
               join Board_Enum_BoardPerformance e on e.Value=m.BoardPerformance order by m.MinPosDelta""")]
        season_rules = [[words(n), s_mod, l_mod] for n, s_mod, l_mod in q(
            "select Name, ConfidenceModFromSeasonObj, ConfidenceModFromLongTermObj from Board_Enum_ObjectiveStates where Value>0")]

        # benchmark: what the teams finishing around the target actually employ
        table_season = season if q("select 1 from Races where SeasonID=? and State=2", season) else season - 1
        order = [t for t, in q("""select TeamID from Races_TeamStandings where SeasonID=? and RaceFormula=1 and TeamID<>?
                                  order by Position""", table_season, team)]
        near = order[max(0, target - 2):target + 1]
        roles = [("Race drivers", 0, "k.PosInTeam<=2"), ("Technical Chief", 1, "1"), ("Race Engineers", 2, "1"),
                 ("Head of Aerodynamics", 3, "1"), ("Sporting Director", 4, "1")]
        opinion = dict(q("select StaffID, TeamOpinionRating from Staff_ContractPatience where TeamID=?", team))
        future_elsewhere = {sid for sid, in q("select StaffID from Staff_Contracts where ContractType=3 and TeamID<>?", team)}
        bench = []
        for label, stype, extra in roles:
            people = q("""select k.StaffID, k.Salary from Staff_Contracts k join Staff_GameData g on g.StaffID=k.StaffID
                          where k.ContractType=0 and g.StaffType=? and %s and k.TeamID in (%s)""" % (extra, ",".join(map(str, near)) or "0"), stype)
            vals = [ratings[sid] for sid, _ in people if sid in ratings]
            sals = [sal for _, sal in people if sal]
            mine = [ratings.get(sid) for sid, fn, ln, st, pos, *_ in my_staff if st == stype and (stype != 0 or pos <= 2)]
            mine = [v for v in mine if v is not None]
            b = sum(vals) / len(vals) if vals else None
            # good value: available, willing, licensed if a driver, and rated at the benchmark rather than far above it
            options = []
            if b is not None:
                for sid, fn, ln in q("""select g.StaffID, b.FirstName, b.LastName from Staff_GameData g
                                        join Staff_BasicData b on b.StaffID=g.StaffID where g.Retired=0 and g.StaffType=?""", stype):
                    r = ratings.get(sid)
                    if r is None or r < b - 2 or (opinion.get(sid) or -99) < settings[3] or sid in future_elsewhere:
                        continue
                    if stype == 0 and not licence.get(sid):
                        continue
                    k = q("select TeamID, EndSeason, Salary from Staff_Contracts where StaffID=? and ContractType=0", sid)
                    if any(t == team for t, *_ in k) or any(e and e > season for _t, e, _s in k):
                        continue   # ours already, or tied up beyond this season
                    options.append((r, person_name(fn, ln), k[0][2] if k else None, team_name.get(k[0][0]) if k else "Free agent"))
                options.sort(key=lambda o: o[0])
            bench.append([label, b, sum(mine) / len(mine) if mine else None, None, sum(sals) / len(sals) if sals else None,
                          "; ".join("%s (%.1f, %s%s)" % (o[1], o[0], o[3], ", paid $%.1fM now" % (o[2] / 1e6) if o[2] else "")
                                    for o in options[:3]) or "Nobody suitable is free"])
        D["board"] = dict(target=target, now_pos=now_pos, races_left=len(races), conf=conf, level=level,
                          long_term="%s, by %d" % (words(lt[0][0]), lt[0][1]) if lt else None,
                          rating_now=rating_now, rating_parts=mine_r, band_now=band_of(rating_now), pay_now=pay.get(target),
                          finish=finish, race_rules=race_rules, season_rules=season_rules, bench=bench,
                          near=", ".join(team_name.get(t, "?") for t in near), table_season=table_season)
        if now_pos and races:
            if now_pos > target:
                flags.append("Board: you are P%d against a target of P%d with %d races left" % (now_pos, target, len(races)))
            elif now_pos <= target - 2:
                flags.append("Board: you are P%d, %d places ahead of the P%d target. Finishing higher raises next season's target"
                             % (now_pos, target - now_pos, target))

    # ---- action plan: up to 10 actions ranked by impact per dollar (rule-based)
    n_r = max(len(races), 1)
    weight = {0: 1.0, 13: 1.0, 1: 0.3, 2: 0.2, 5: 0.2}
    for key, st in CAR_STAT_TO_PART_STAT.items():
        weight[st] = sum((2 if key in (r[3] or "") else 0) + (1 if key in (r[4] or "") else 0) for r in races) / n_r

    def weighted(pt, fn):
        """weighted average over the part's stats of fn(stat); weights follow the remaining calendar"""
        sts = [st for st in stat_values.get(part_info[pt]["newest"][0], {})]
        tot = sum(weight.get(st, 0.2) for st in sts)
        return sum(weight.get(st, 0.2) * fn(st) for st in sts) / tot if tot else 0

    def impact_from(points):
        return 5 if points >= 60 else 4 if points >= 40 else 3 if points >= 20 else 2 if points >= 8 else 1

    def race_for(ready_day):
        for r in races:
            if r[1] - 2 >= ready_day:
                return r[0]
        return "Next season"

    rate_n = settings[5] * (1 + factory_bonus)
    rate_r = settings[6] * (1 + factory_bonus)
    type_info = {v: (work, cost, dwork, dcost) for v, work, cost, dwork, dcost in q(
        "select Value, BaseManufactureWork, BaseManufactureCost, BaseDesignWork, BaseDesignCost from Parts_Enum_Type")}
    eng_speed = dict(q("select EngineerNum, SpeedIncrease from Parts_EngineerData"))
    next_race_days = races[0][1] if races else 999
    actions = []

    def add(action, when, takes, in_time_for, cost, impact, why, counts_to_cap=True):
        actions.append(dict(action=action, when=when, takes=takes, race=in_time_for, cost=cost, impact=impact, why=why,
                            cap=counts_to_cap))

    def build_plan(pt, start_day):
        """choose normal or rushed: rush only if that is what gets the part to the next race it can reach"""
        work, cost = type_info[pt][0], type_info[pt][1]
        days_n, days_r = -(-work // rate_n), -(-work // rate_r)
        if race_for(start_day + days_r) != race_for(start_day + days_n):
            return "rushed", int(days_r), cost * settings[7]
        return "normal", int(days_n), cost

    for pt, info in part_info.items():
        done, nd, car, label = info["done"], info["newest"], info["car"], info["label"]
        gain_car = lambda lo, d: weighted(pt, lambda st: stat_values[d[0]].get(st, 0) - stat_values.get(car[lo], {}).get(st, 0))
        if done and not info["ready"] and not info["building"]:
            g = sum(gain_car(lo, done) for lo in (1, 2)) / 2
            speed, days, cost = build_plan(pt, 0)
            add("Build 2 x %s #%d (%s) and fit them" % (label, done[2], speed), "Now", days, race_for(days), 2 * cost,
                impact_from(g), "Finished design never built. Worth about +%d weighted stat points per car over what is fitted" % g)
        elif done and info["ready"]:
            spare = info["ready"] - sum(1 for lo in (1, 2) if car[lo] == done[0])
            for lo in (1, 2):
                if car[lo] and design_no.get(car[lo], 0) < done[2]:
                    g = gain_car(lo, done)
                    if g <= 0:
                        continue
                    if spare > 0:
                        spare -= 1
                        add("Fit the spare %s #%d to car %d" % (label, done[2], lo), "Now", 0, race_for(0), 0,
                            impact_from(g), "Already in stock. About +%d weighted stat points on that car" % g)
                    elif nd[6] >= 0 or nd[3] < 0.8 * nd[4]:  # not worth it if a newer design is nearly finished
                        speed, days, cost = build_plan(pt, 0)
                        add("Build 1 x %s #%d (%s) for car %d" % (label, done[2], speed, lo), "Now", days, race_for(days),
                            cost, impact_from(g), "Car %d is on design #%d. About +%d weighted stat points" % (lo, design_no[car[lo]], g))
        if nd[6] < 0:
            n_eng = engineers_on.get(nd[0]) or 1
            left = max(0, nd[4] - nd[3]) / (200 * (1 + eng_speed.get(n_eng, 0)))
            left = int(-(-left // 1))
            g = sum(gain_car(lo, nd) for lo in (1, 2)) / 2
            speed, days, cost = build_plan(pt, left)
            add("Build 2 x %s #%d (%s) as soon as the design completes" % (label, nd[2], speed),
                "In about %d days" % left, days, race_for(left + days), 2 * cost, impact_from(g),
                "Design is %d%% done. About +%d weighted stat points per car" % (round(100 * nd[3] / nd[4]), g))

    for pt, sp in spare_plan.items():
        info = part_info[pt]
        if info["newest"][6] < 0 and info["newest"][3] >= 0.8 * info["newest"][4]:
            continue   # a newer design is about to land; spares of the old one would be wasted
        unfitted = max(0, sp["good"] - sp["on_new"])
        for_cars = max(0, (2 - sp["on_new"]) - unfitted) if info["ready"] else 2
        extra = int(sp["to_build"] - for_cars)
        if extra > 0:
            add("Build %d x %s #%d as spare%s" % (extra, info["label"], info["done"][2], "s" if extra > 1 else ""),
                "Now", sp["build_days"] * extra, race_for(sp["build_days"]), extra * sp["cost"], 2,
                "You have lost %d %s units this season (%.1f a race). Target is %d good spare%s beyond the two on the cars"
                % (sp["lost"], info["label"].lower(), sp["lost"] / max(1, races_done), sp["target"], "s" if sp["target"] > 1 else ""))

    nxt_race = "%s (in %d days)" % (races[0][0], races[0][1]) if races else "the next race"
    for r in pu_summary:
        if r[10].startswith("Fit #"):
            car, label = int(r[0][-1]), r[1]
            major = thresh[[k for k, v in PU.items() if v[0] == label][0]][1]
            add("Fit %s %s to car %d" % (label.lower() if label != "ERS" else "ERS", r[5], car), "Before %s" % nxt_race,
                0, races[0][0] if races else None, 0, 4 if r[4] < major else 2,
                "Healthiest unit in your allocation (%d%% against %d%% on the car), so no grid penalty" % (r[6] * 100, r[4] * 100))
    for b in pu_bundles:
        names = [x.lower() if x != "ERS" else "ERS" for x in b["parts"]]
        what = names[0] if len(names) == 1 else ", ".join(names[:-1]) + " and " + names[-1]
        tip = ""
        if b["places"] and avg_grid and avg_grid >= 15:
            tip = " You have averaged P%.0f on the grid this season, so it mostly just moves the car to the back." % avg_grid
        if b["race"] != b["first_due"]:
            tip += " The first part is due at %s, where overtaking is hard, so it is brought forward." % b["first_due"]
        add("Car %d: new %s%s at %s%s" % (b["car"], what, " together" if len(names) > 1 else "", b["race"],
                                                  " (%d-place grid penalty)" % b["places"] if b["places"] else ""),
            "That race weekend, in %d days" % b["days"], None, b["race"], b["cost"], 3,
            "The units in the allocation would otherwise fall into major damage. Taking them in one weekend serves the penalty once.%s "
            "Cost shown is the game's base price; check the price in game" % tip)

    free_slots = int(design_slots - designs_running)
    idle_eng = eng_total - eng_used
    deficits = sorted(((weighted(pt, lambda st, pt=pt: (median([v for v in (rival_best(t, pt, st) for t in rivals) if v is not None]) or 0)
                                 - stat_values[i["newest"][0]].get(st, 0)), pt)
                       for pt, i in part_info.items() if i["newest"][6] >= 0), reverse=True)
    hours_note = ""
    if res[0] or res[1]:
        hours_note = " Put the %d wind tunnel and %d CFD hours on it: they reset in %d days." % (
            res[0], res[1], (period_end - day + 1) if period_end else 0)
    for n, (deficit, pt) in enumerate(deficits[:max(0, min(free_slots, 2))]):
        if deficit <= 0:
            continue
        eng = max(1, min(6, (idle_eng or 2) // max(1, min(free_slots, 2))))
        days = int(-(-type_info[pt][2] // (200 * (1 + eng_speed.get(eng, 0)))))
        sts = sorted(stat_values[part_info[pt]["newest"][0]], key=lambda st: -weight.get(st, 0.2))[:2]
        build_days = int(-(-type_info[pt][0] // rate_n))
        add("Start a new %s design, focused on %s" % (PART_LABEL[pt].lower(), " and ".join(PART_STAT_LABEL[st].lower() for st in sts)),
            "Now", days, race_for(days + build_days), type_info[pt][3], impact_from(deficit * 0.6),
            "About %d weighted points behind the grid median, with %d engineers.%s" % (deficit, eng, hours_note if n == 0 else ""))
    if idle_eng > 0 and designs_running and free_slots <= 0:
        add("Assign the %d idle engineers to the designs in progress" % idle_eng, "Now", 0, None, 0, 3,
            "Each extra engineer adds 10% design speed, up to 6 per project")

    team_ment = [r[2] for r in staff if r[2] is not None]
    low_ment = team_ment and sum(team_ment) / len(team_ment) < 50
    for r in fac:
        bname, lvl, cond, grid, refurb, upcost, gives = r[0], r[1], r[2], r[4], r[5], r[6], r[8] or ""
        dev = bname in ("Design Centre", "Factory", "Wind Tunnel", "CFD Simulator", "Suspension Simulator", "Car Part Test Centre")
        eff = None if cond is None else 1 if cond > 0.75 else 0.8 if cond > 0.5 else 0.6 if cond > 0.25 else 0.4
        m = re.search(r"Weekly Income (\d+) to (\d+)", gives)
        if eff is not None and eff < 1 and refurb and not (upcost and grid and lvl < grid - 0.5 and upcost <= 5 * refurb):
            add("Refurbish the %s" % bname, "Now", None, None, refurb, 4 if dev else 2 if eff <= 0.6 else 1,
                "At %d%% condition it gives only %d%% of its effect" % (round(cond * 100), eff * 100), counts_to_cap=False)
        elif m and upcost and (int(m.group(2)) - int(m.group(1))) > 0 and upcost / (int(m.group(2)) - int(m.group(1))) <= 12:
            add("Upgrade the %s" % bname, "Now", None, None, upcost, 2,
                "Pays for itself in about %d weeks of extra income" % round(upcost / (int(m.group(2)) - int(m.group(1)))),
                counts_to_cap=False)
        elif low_ment and "Mentality" in gives and upcost and (lvl == 0 or (grid and lvl < grid - 0.5)):
            add("%s the %s" % ("Build" if lvl == 0 else "Upgrade", bname), "Now", None, None, upcost, 2,
                "Team mentality averages %d. Next level gives: %s" % (sum(team_ment) / len(team_ment), gives.lower()),
                counts_to_cap=False)
        elif bname == "Design Centre" and upcost and grid and lvl <= grid - 1 and balance >= 1.3 * upcost:
            add("Upgrade the Design Centre", "Once parts are ordered", None, "Next season", upcost, 4,
                "A level below the grid average (%.1f). Next level gives: %s" % (grid, gives.lower()), counts_to_cap=False)

    contract_end = q("select min(Day) from Seasons_Deadlines where SeasonID=? and DeadlineType=6", season)[0][0]
    for r in staff:
        if r[10] and r[10].startswith("Leaving"):
            add("Line up a replacement for %s (%s)" % (r[0], r[1]), "Before day %d" % (contract_end - day) if contract_end
                else "Before the season ends", None, None, None, 3, "%s at the end of the season. Candidates are on the Scouting tab, filtered by the Role column" % r[10])
    seen_roles = set()
    for r in women_steps:
        key = (r[0], r[14])
        if key in seen_roles:
            continue
        seen_roles.add(key)
        if r[0] == "Driver" and r[14] == 2:
            best = [w for w in women_steps if w[0] == "Driver" and w[14] == 2][:2]
            add("Women's team: sign %s as affiliate driver%s" % (" and ".join(w[1] for w in best), "s" if len(best) > 1 else ""),
                "Now", None, None, None, 2,
                "Best available prospects by rating and improvability: %s. They need a super licence before a race seat. "
                "Cost depends on the contract" % "; ".join("%s (%.1f, improvability %s)" % (w[1], w[2] or 0, w[5]) for w in best))
        else:
            add("Women's team: sign %s (%s)" % (r[1], r[0].lower()),
                "Now", None, None, None, 2, "%s. Rated %.1f%s. Cost depends on the contract" % (
                    r[12], r[2] or 0, " against %.1f for your best in that role" % r[3] if r[3] else ""))
    for stype, label, covered in vacancies:
        if label == "Reserve driver":
            continue   # handled by the reserve driver plan below
        pool = [cnd for cnd in cands if cnd[13] == stype and cnd[8] == "Free agent" and not cnd[15]]
        names = ", ".join("%s (%.1f, age %s)" % (cnd[1], cnd[2], cnd[7]) for cnd in pool[:3])
        add("Sign a %s%s" % (label.lower(), " for the rest of this season" if covered else ""), "Now", None, None, None,
            2 if label == "Reserve driver" else 4,
            "The seat is empty. Free agents open to you: %s. Cost depends on the contract" % (names or "none listed on the Scouting tab"))
    expiring = [r for r in staff if r[4] == season and not r[10]]
    if expiring:
        add("Open renewal talks with " + ", ".join(r[0] for r in expiring),
            "Before day %d, when contracts expire" % (contract_end - day) if contract_end else "Before the season ends",
            None, None, None, 4, "Replacements open to you are mostly weaker (see Scouting). Cost depends on the terms you agree")

    # ---- reserve driver plan: when to run the reserve, who stands down, and whether to swap seats
    weekend_type = dict(q("select Day, WeekendType from Races where SeasonID=? and State=0", season))
    weekend_name = dict(q("select Type, Name from Races_Enum_WeekendType"))
    next_of = {r[0]: r[10] for r in staff}
    drivers_now = [dict(sid=sid, name=person_name(fn, ln), pos=pos, rating=ratings.get(sid), ment=ment,
                        nxt=next_of.get(person_name(fn, ln)))
                   for sid, fn, ln, stype, pos, ends, sal, ment, mood in my_staff if stype == 0]
    racers = [d for d in drivers_now if d["pos"] <= 2 and d["rating"] is not None]
    reserve_now = next((d for d in drivers_now if d["pos"] >= 3), None)
    swap_margin, every = settings[11], max(1, int(settings[12]))
    boost = q("select avg(Value), min(Value), max(Value) from Staff_Mentality_Events where Event=13")[0]
    boost_text = ("worth +%d to +%d mentality for about two weeks in this save" % (boost[1], boost[2])
                  if boost[0] is not None else "a mentality boost for the reserve")

    def on_way_out(d):
        return bool(d.get("nxt")) and d["nxt"].startswith(("Leaving", "Replaced"))

    def stands_down(pair):   # who gives up FP1: someone on the way out first, then the lower-rated driver
        return sorted(pair, key=lambda d: (not on_way_out(d), d["rating"] or 0))[0]

    lines, plan_rows, pick, swap = [], [], None, None
    weakest = min(racers, key=lambda d: d["rating"]) if racers else None
    reserve = reserve_now
    if reserve_now is None:
        pool = [cnd for cnd in cands if cnd[13] == 0 and cnd[14] == "Yes" and not cnd[15] and cnd[8] == "Free agent"]
        if pool:
            pick = pool[0]
            reserve = dict(sid=None, name=pick[1], pos=3, rating=pick[2], ment=None, nxt=None)
            lines.append(("Reserve now", "None signed. Best licensed free agent with no deal elsewhere for next season: "
                          "%s (rated %.1f, age %s, opinion %d of your team). The plan below assumes you sign them."
                          % (pick[1], pick[2], pick[7], pick[5])))
        else:
            lines.append(("Reserve now", "None signed, and no licensed free agent rated above your weakest race driver is free "
                          "for next season. See the Scouting tab for others"))
    else:
        lines.append(("Reserve now", "%s (rated %.1f%s)" % (reserve_now["name"], reserve_now["rating"] or 0,
                                                            ", mentality %d" % reserve_now["ment"] if reserve_now["ment"] is not None else "")))
    if reserve and weakest and reserve["rating"] is not None:
        gap = reserve["rating"] - weakest["rating"]
        if gap >= swap_margin:
            swap = (reserve, weakest)
            lines.append(("Race seat", "Swap. %s rates %.1f against %.1f for %s (%+.1f; your threshold is +%g). Put %s in %s's seat "
                          "before the next race, and %s becomes the reserve.%s"
                          % (reserve["name"], reserve["rating"], weakest["rating"], weakest["name"], gap, swap_margin,
                             reserve["name"], weakest["name"], weakest["name"],
                             " %s is on the way out anyway (%s), so there is little to lose." % (weakest["name"], weakest["nxt"][0].lower() + weakest["nxt"][1:])
                             if on_way_out(weakest) else "")))
        else:
            lines.append(("Race seat", "Keep the race drivers. %s rates %+.1f against your weakest race driver, %s (threshold +%g)."
                          % (reserve["name"], gap, weakest["name"], swap_margin)))
    pair = ([swap[0]] + [d for d in racers if d is not swap[1]]) if swap else racers
    fp1 = swap[1] if swap else reserve
    if fp1 and pair:
        lines.append(("FP1 outings", "%s every %d race weekend%s, never on sprint weekends. %s gives up the session: %s."
                      % (fp1["name"], every, "s" if every > 1 else "", stands_down(pair)["name"],
                         "on the way out at the end of the season" if on_way_out(stands_down(pair)) else "the lower-rated race driver")))
        last = q("""select max(r.Day) from Races_PracticeResults p join Races r on r.RaceID=p.RaceID
                    where p.TeamID=? and p.SeasonID=? and p.DriverID=?""", team, season, fp1["sid"])[0][0] if fp1["sid"] and not swap else None
        since = q("select count(*) from Races where SeasonID=? and State=2 and Day>?", season, last)[0][0] if last else None
        gap = 0 if swap else (since if since is not None else 99)   # after a swap, let the new race driver settle in first
        for name, days, rday, *_ in races:
            wt = weekend_type.get(rday, 0)
            label = "%s (%s)" % (name, words(weekend_name.get(wt, "Normal")).lower())
            if wt != 0:
                plan_rows.append([label, days, "No", "Sprint weekend: only one practice session, so both race drivers need it"])
                gap += 1
            elif gap >= every - 1:
                sd = stands_down(pair)
                plan_rows.append([label, days, "Yes: %s in %s's car" % (fp1["name"], sd["name"]),
                                  "%s keeps race-ready, and the outing is %s. %s sits out: %s"
                                  % (fp1["name"], boost_text, sd["name"],
                                     "on the way out at the end of the season" if on_way_out(sd) else "the lower-rated race driver")])
                gap = 0
            else:
                plan_rows.append([label, days, "No", "Not due: outings every %d weekends" % every])
                gap += 1
    if not plan_rows:
        plan_rows = [[r[0], r[1], "No", "No reserve driver"] for r in races]

    nxt_race = "%s (in %d days)" % (races[0][0], races[0][1]) if races else "the next race"
    if reserve_now and swap:
        add("Swap reserve %s into %s's seat" % (reserve_now["name"], weakest["name"]), "Before %s" % nxt_race, None,
            races[0][0] if races else None, 0, 3, lines[1][1])
    if pick:
        add("Sign %s as reserve driver%s" % (pick[1], ", then put them in %s's seat" % weakest["name"] if swap else ""),
            "Now", None, races[0][0] if swap and races else None, None, 3 if swap else 2,
            "Seat is empty. Licensed, free now and for next season, rated %.1f. %s Cost depends on the contract"
            % (pick[2], lines[1][1] if len(lines) > 1 else ""))
    if reserve_now:
        first = next((r for r in plan_rows if r[2].startswith("Yes")), None)
        if first:
            add("Run %s in FP1 at %s, %s" % (fp1["name"], first[0].split(" (")[0], first[2][5:].replace(fp1["name"] + " in ", "in ")),
                "That race weekend, in %d days" % first[1], None, first[0].split(" (")[0], 0, 1, first[3])
    D["reserve_plan"] = dict(lines=lines, rows=plan_rows)

    if scouts_total - scouts_used > 0:
        todo = ["%s (%s)" % (cnd[1], cnd[0].lower()) for cnd in cands
                if cnd[11] == "No" and cnd[8] == "Free agent" and not cnd[15]][:int(scouts_total - scouts_used)]
        if todo:
            add("Send the idle scouts to " + " and ".join(todo), "Now", 6, None, 0, 1, "Free; a detailed scout takes about 6 days")
    crew_rank = {r[0]: r[5] for r in crew}
    weak = [k for k in ("Speed", "Car Setup") if crew_rank.get(k, 0) >= D["n_teams"] - 1]
    if weak and crew_rank.get("Wheel Gun", 99) <= 3:
        add("Switch pit crew training from stage drills to gym and car setup practice", "Now", None, None, 0, 2,
            "Stage skills are top three on the grid but %s rank%s near last" % (" and ".join(weak), "s" if len(weak) == 1 else ""))

    for a in actions:
        a["score"] = a["impact"] / (1 + (a["cost"] or 0) / 1e6)
    actions.sort(key=lambda a: (-a["score"], -(a["impact"]), a["takes"] or 0))
    # Take actions in rank order while the money lasts: cash above the safety buffer, and the cost cap headroom
    # (facility work is paid in cash but does not count against the cap).
    cash_left = balance - settings[13]
    cap_left = (cap - projected) if cap else float("inf")
    chosen, waiting = [], []
    for a in actions:
        cost = a["cost"] or 0
        cap_cost = cost if a["cap"] else 0
        if len(chosen) >= int(settings[14]):
            a["reason"] = "Beyond the maximum number of actions"
            waiting.append(a)
        elif cost > cash_left:
            a["reason"] = "Needs $%.1fM; $%.1fM left above the buffer" % (cost / 1e6, max(0, cash_left) / 1e6)
            waiting.append(a)
        elif cap_cost > cap_left:
            a["reason"] = "Would take projected spending past the cost cap"
            waiting.append(a)
        else:
            chosen.append(a)
            cash_left -= cost
            cap_left -= cap_cost
    D["actions"], D["waiting"] = chosen, waiting
    D["budget"] = dict(cash=balance, buffer=settings[13], cap_headroom=(cap - projected) if cap else None)
    D["flags"] = flags
    c.close()
    return D


# ---------------------------------------------------------------- workbook

def style_table(ws, head_row, n_cols, first_col=1):
    for col in range(first_col, first_col + n_cols):
        cell = ws.cell(row=head_row, column=col)
        cell.font, cell.fill = BOLD, HEAD_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="top")


def put_table(ws, top, headers, rows, formats=None, widths=None):
    for j, h in enumerate(headers, 1):
        ws.cell(row=top, column=j, value=h)
    style_table(ws, top, len(headers))
    for i, row in enumerate(rows, top + 1):
        for j, v in enumerate(row, 1):
            cell = ws.cell(row=i, column=j, value=v)
            cell.font = FONT
            if formats and formats.get(j):
                cell.number_format = formats[j]
    for j, w in enumerate(widths or [], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    return top + len(rows)


def read_settings(wb):
    values = [s[1] for s in SETTINGS]
    if wb is not None and "Settings" in wb.sheetnames:
        ws = wb["Settings"]
        for i in range(len(SETTINGS)):
            v = ws.cell(row=i + 2, column=2).value
            if isinstance(v, (int, float)):
                values[i] = v
    return values


SETTINGS_NOTE = "Blue values are yours to change. This tab and any tabs you add are never overwritten by the updater."


def ensure_settings(wb):
    if "Settings" in wb.sheetnames:
        ws = wb["Settings"]   # an older workbook: add any setting introduced since, leave the rest alone
        for i, (label, value, note) in enumerate(SETTINGS, 2):
            if not isinstance(ws.cell(row=i, column=2).value, (int, float)):
                for col, v in enumerate((label, value, note), 1):
                    ws.cell(row=i, column=col, value=v).font = INPUT if col == 2 else FONT
        for row in range(len(SETTINGS) + 2, ws.max_row + 1):   # move the footnote below the settings
            if ws.cell(row=row, column=1).value == SETTINGS_NOTE:
                ws.cell(row=row, column=1).value = None
        ws.cell(row=len(SETTINGS) + 3, column=1, value=SETTINGS_NOTE).font = FONT
        return
    ws = wb.create_sheet("Settings")
    ws.append(["Setting", "Value", "Note"])
    style_table(ws, 1, 3)
    for label, value, note in SETTINGS:
        ws.append([label, value, note])
    for row in ws.iter_rows(min_row=2, max_row=len(SETTINGS) + 1):
        for cell in row:
            cell.font = FONT
        row[1].font = INPUT
        if isinstance(row[1].value, (int, float)) and row[1].value >= 100000:
            row[1].number_format = MONEY
    ws.cell(row=len(SETTINGS) + 3, column=1, value=SETTINGS_NOTE).font = FONT
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width, ws.column_dimensions["C"].width = 58, 10, 90


def build(wb, D, save_path):
    for name in OWNED_SHEETS:
        if name in wb.sheetnames:
            del wb[name]
    sheets = {name: wb.create_sheet(name, i) for i, name in enumerate(OWNED_SHEETS)}
    n_races = max(len(D["races"]), 1)
    rank_head = "Rank of %d" % D["n_teams"]

    # Overview
    ws = sheets["Overview"]
    ws["A1"], ws["A1"].font = "F1 Manager 2024 planner: %s" % D.get("display_team", ""), TITLE
    ws["A2"], ws["A2"].font = "Planner version %s" % VERSION, FONT
    info = [("Save file", os.path.basename(save_path), None),
            ("Save written", dt.datetime.fromtimestamp(os.path.getmtime(save_path)), "yyyy-mm-dd hh:mm"),
            ("Workbook updated", dt.datetime.now().replace(microsecond=0), "yyyy-mm-dd hh:mm")] + D["overview"]
    for i, (label, value, fmt) in enumerate(info, 3):
        ws.cell(row=i, column=1, value=label).font = BOLD
        cell = ws.cell(row=i, column=2, value=value)
        cell.font, cell.alignment = FONT, Alignment(horizontal="left")
        if fmt:
            cell.number_format = "d mmm yyyy" if fmt == "date" else fmt
    ws["A13"].comment = Comment("Taken as the only non-zero row of the save's testing-hours table. "
                                "Check against the game the first time.", "updater")
    top = len(info) + 5
    ws.cell(row=top, column=1, value="Things to look at").font = TITLE
    for i, text in enumerate(D["flags"] or ["Nothing flagged"], top + 1):
        ws.cell(row=i, column=1, value=text).font = FONT
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 44, 26

    # Action plan
    ws = sheets["Action plan"]
    ws["A1"], ws["A1"].font = "Action plan: ranked by impact per dollar, as far as the money goes", TITLE
    B = D["budget"]
    summary = [("Cash in the bank", B["cash"]), ("Keep as a safety buffer", "=" + S_BUFFER),
               ("Available to spend", "=C2-C3"),
               ("Cost cap headroom at season end (projected)", B["cap_headroom"] if B["cap_headroom"] is not None else "n/a")]
    for i, (label, value) in enumerate(summary, 2):
        ws.cell(row=i, column=2, value=label).font = BOLD
        cell = ws.cell(row=i, column=3, value=value)
        cell.font, cell.number_format, cell.alignment = FONT, MONEY, Alignment(horizontal="left")
    HEAD = 7
    acts = D["actions"]
    first, last = HEAD + 1, HEAD + len(acts)
    rows = [['=RANK(I{r},$I${f}:$I${l})'.format(r=i, f=first, l=last), a["action"], a["when"], a["takes"], a["race"],
             a["cost"], "=SUM($F${f}:F{r})".format(f=first, r=i), a["impact"], "=H{r}/(1+N(F{r})/1000000)".format(r=i),
             a["why"]] for i, a in enumerate(acts, first)]
    put_table(ws, HEAD, ["Rank", "Action", "When to start", "Takes (days)", "On the car in time for", "Cost",
                         "Spent so far", "Impact (1-5)", "Value score", "Why"], rows, {6: MONEY, 7: MONEY, 9: "0.00"},
              [6, 52, 28, 8, 20, 12, 12, 8, 8, 90])
    for row in ws.iter_rows(min_row=first, max_row=last):
        row[7].font = INPUT
        for cell in (row[1], row[2], row[9]):
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row=last + 1, column=2, value="Total cost of the listed actions").font = BOLD
    cell = ws.cell(row=last + 1, column=6, value="=SUM(F%d:F%d)" % (first, last))
    cell.font, cell.number_format = BOLD, MONEY
    ws.cell(row=last + 2, column=2, value="Left to spend after these").font = BOLD
    cell = ws.cell(row=last + 2, column=6, value="=C4-F%d" % (last + 1))
    cell.font, cell.number_format = BOLD, MONEY
    notes = ["Actions are taken in rank order while the money lasts: cash above the buffer, and the projected cost cap headroom. "
             "Facility work is paid in cash but does not count against the cap.",
             "Value score = impact divided by (1 + cost in $ millions), so free and cheap actions rank first.",
             "Impact is set by rules: for parts it comes from stat points gained, weighted by what the remaining circuits need. "
             "It is in blue because you can overtype it; the rank follows.",
             "A blank cost depends on contract terms. Days and races assume the build and design rates on the Settings tab.",
             "This tab is rebuilt from the save on every update."]
    for i, text in enumerate(notes, last + 4):
        ws.cell(row=i, column=2, value=text).font = FONT
    ws.freeze_panes = "A%d" % first
    last = last + 4 + len(notes)   # end of the plan block

    wait = D["waiting"]
    if wait:
        top = last + 1
        ws.cell(row=top, column=2, value="Waiting for money").font = TITLE
        for col, head in zip((2, 3, 6, 10), ("Action", "Why it is waiting", "Cost", "What it would do")):
            ws.cell(row=top + 1, column=col, value=head)
        style_table(ws, top + 1, 9, first_col=2)
        for i, a in enumerate(wait, top + 2):
            for col, v in zip((2, 3, 6, 10), (a["action"], a["reason"], a["cost"], a["why"])):
                cell = ws.cell(row=i, column=col, value=v)
                cell.font, cell.alignment = FONT, Alignment(wrap_text=True, vertical="top")
            ws.cell(row=i, column=6).number_format = MONEY
            ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=5)
        ws.merge_cells(start_row=top + 1, start_column=3, end_row=top + 1, end_column=5)
        last = top + 1 + len(wait)
    notes = []

    top = last + 2
    ws.cell(row=top, column=2, value="Reserve driver plan").font = TITLE
    row = top + 1
    for label, text in D["reserve_plan"]["lines"]:
        ws.cell(row=row, column=2, value=label).font = BOLD
        cell = ws.cell(row=row, column=3, value=text)
        cell.font, cell.alignment = FONT, Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=10)
        ws.row_dimensions[row].height = 15 * max(1, -(-len(text) // 150))
        row += 1
    row += 1
    for col, head in zip((2, 3, 4, 5), ("Race", "Reserve in FP1", "Days away", "Why")):
        ws.cell(row=row, column=col, value=head)
    style_table(ws, row, 4, first_col=2)
    ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=10)
    for name, days, verdict, why in D["reserve_plan"]["rows"]:
        row += 1
        for col, v in zip((2, 3, 4, 5), (name, verdict, days, why)):
            cell = ws.cell(row=row, column=col, value=v)
            cell.font, cell.alignment = FONT, Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=10)
        ws.row_dimensions[row].height = 30
    row += 2
    for text in ("The race driver who sits out FP1 can take a 'missed practice' mentality hit; this save holds no example of its size.",
                 "Thresholds for swapping and how often to run the reserve are on the Settings tab."):
        ws.cell(row=row, column=2, value=text).font = FONT
        row += 1

    # Board
    ws = sheets["Board"]
    Bd = D["board"]
    if not Bd:
        ws["A1"] = "The save holds no board objective for this season yet."
    else:
        ws["A1"], ws["A1"].font = "Board objective and how to meet it without raising the bar", TITLE
        info = [("Board target this season", "P%d" % Bd["target"]),
                ("Where you are now", "P%d" % Bd["now_pos"] if Bd["now_pos"] else "No races yet"),
                ("Races left", Bd["races_left"]),
                ("Board confidence", "%s (%s)" % (Bd["conf"], Bd["level"]) if Bd["conf"] is not None else "n/a"),
                ("Long-term objective", Bd["long_term"] or "None open"),
                ("Team rating now", Bd["rating_now"]), ("Rating band", Bd["band_now"]),
                ("Board payment for a P%d target" % Bd["target"], Bd["pay_now"])]
        for i, (label, value) in enumerate(info, 3):
            ws.cell(row=i, column=1, value=label).font = BOLD
            cell = ws.cell(row=i, column=2, value=value)
            cell.font, cell.alignment = FONT, Alignment(horizontal="left")
        ws.cell(row=10, column=2).number_format = MONEY
        parts = ["constructor results %d" % Bd["rating_parts"][0], "driver results %d" % Bd["rating_parts"][1],
                 "seasons entered %d" % Bd["rating_parts"][2], "championships %d" % Bd["rating_parts"][3],
                 "new team hype %d" % Bd["rating_parts"][4]]
        ws.cell(row=11, column=1, value="Rating is made of").font = BOLD
        ws.cell(row=11, column=2, value=", ".join(parts)).font = FONT

        top = 14
        ws.cell(row=top - 1, column=1, value="Where to finish this season, and what it does to next season").font = TITLE
        end = put_table(ws, top, ["Finish", "This season's objective", "Team rating next season (estimate)", "Rating band",
                                  "Likely target next season (rank among team ratings)", "Board payment for that target", "Advice"],
                        Bd["finish"], {6: MONEY})
        for row in ws.iter_rows(min_row=top + 1, max_row=end):
            if row[6].value:
                for cell in row:
                    cell.font = BOLD
        notes = ["The rating counts last season's finish twice over, so it is what mostly sets next season's target.",
                 "Finish on the target, or one place better. Every place above that pushes next season's rating up.",
                 "Estimate assumes driver results, seasons, championships and new team hype stay as they are; hype may fade.",
                 "Rivals' ratings are this season's; theirs will move too."]
        for i, text in enumerate(notes, end + 1):
            ws.cell(row=i, column=1, value=text).font = FONT

        top = end + len(notes) + 3
        ws.cell(row=top - 1, column=1, value="How the board reacts").font = TITLE
        end = put_table(ws, top, ["Each race: result against what the board expected", "Confidence change", "Board mood"],
                        Bd["race_rules"])
        top = end + 2
        end = put_table(ws, top, ["End of season: objective", "Confidence (season objective)", "Confidence (long-term objective)"],
                        Bd["season_rules"])
        ws.cell(row=end + 1, column=1, value="Confidence at 0 or below means the board has lost confidence in you.").font = FONT

        top = end + 4
        ws.cell(row=top - 1, column=1, value="Staff needed for the target: what the teams around P%d employ (%s, %d table)"
                                             % (Bd["target"], Bd["near"], Bd["table_season"])).font = TITLE
        rows = [r[:3] + ["=IF(OR(B{r}=\"\",C{r}=\"\"),\"\",C{r}-B{r})".format(r=i), r[4], r[5]]
                for i, r in enumerate(Bd["bench"], top + 1)]
        end = put_table(ws, top, ["Role", "Benchmark rating", "Yours", "Difference", "Their average salary",
                                  "Good value: available and rated at the benchmark, not far above it"], rows,
                        {2: "0.0", 3: "0.0", 4: "0.0", 5: MONEY})
        for row in ws.iter_rows(min_row=top + 1, max_row=end):
            row[5].alignment = Alignment(wrap_text=True, vertical="top")
        steps = ["Setting up a new team:",
                 "1. If you get a choice of board expectation, take the lowest you can live with. A higher target pays more "
                 "($2.5M a year per place, from the save) but every place is harder to hold.",
                 "2. Hire to the benchmark above, not the best on the market: a star costs far more than the target needs.",
                 "3. Aim to finish on the target, one place better at most. A big over-achievement in year one sets a big target for year two.",
                 "4. Once the target is safe late in the season, move design work to next year's car instead of chasing extra places.",
                 "Ratings are plain averages of each person's stats. The benchmark uses the teams placed either side of the target."]
        for i, text in enumerate(steps, end + 2):
            ws.cell(row=i, column=1, value=text).font = BOLD if i == end + 2 else FONT
        for col, w in zip("ABCDEFG", (44, 22, 18, 30, 16, 16, 70)):
            ws.column_dimensions[col].width = w

    # Car vs grid
    ws = sheets["Car vs grid"]
    rows = []
    for i, a in enumerate(D["areas"], 2):
        key = a[7]
        count = lambda col: 'COUNTIF(Races!$%s$2:$%s$%d,"*%s*")' % (col, col, n_races + 1, key)
        rows.append([a[0], a[1], a[2], a[3], a[4], "=D%d-E%d" % (i, i), a[6],
                     "=%s+%s" % (count("D"), count("E")) if key else None, "=" + count("D") if key else None,
                     '=IF(G{r}<={good},"Already good: no upgrade needed",IF(I{r}="","Not a circuit requirement",'
                     'IF(I{r}>=3,"Upgrade: high priority",IF(H{r}>0,"Upgrade","Low priority"))))'.format(r=i, good=S_GOOD)])
    end = put_table(ws, 1, ["Area", "Part stat", "On the cars now (avg)", "Newest designs (avg)", "Grid median", "Gap",
                            rank_head, "Races left needing it", "Crucial at", "Verdict"],
                    rows, {3: "0", 4: "0", 5: "0", 6: "0"}, [30, 24, 14, 14, 12, 8, 9, 12, 10, 36])
    ws.cell(row=end + 2, column=1, value="Which part stat drives which area is my reading of the game; the save does not state it. "
                                         "Grid values are the best part each rival has fitted.").font = FONT
    ws.freeze_panes = "A2"

    # Part stats
    ws = sheets["Part stats"]
    rows = [r[:7] + ["=E%d-F%d" % (i, i)] + r[8:10] for i, r in enumerate(D["part_stats"], 2)]
    n_ps = len(rows) + 1
    put_table(ws, 1, ["Part", "Stat", "Car 1", "Car 2", "Newest design", "Grid median", "Grid best", "Gap to median",
                      rank_head, "Team expertise"], rows, {j: "0" for j in (3, 4, 5, 6, 7, 8, 10)},
              [14, 26, 9, 9, 10, 10, 10, 10, 9, 11])
    ws.freeze_panes = "A2"

    # Part summary
    ws = sheets["Part summary"]
    n = len(D["summary"])
    rows = []
    for i, r in enumerate(D["summary"], 2):
        rng = lambda col: "'Part stats'!$%s$2:$%s$%d" % (col, col, n_ps)
        rows.append(r[:6] + ["=AVERAGEIFS(%s,%s,A%d)" % (rng("H"), rng("A"), i),
                             "=AVERAGEIFS(%s,%s,A%d)" % (rng("I"), rng("A"), i),
                             "=RANK(G%d,$G$2:$G$%d,1)" % (i, n + 1), r[9]])
    put_table(ws, 1, ["Part", "Newest design", "Status", "Car 1 design", "Car 2 design", "Newest finished in stock",
                      "Average gap to median", "Average rank", "Design priority (1 = furthest behind)", "Notes"],
              rows, {7: "0", 8: "0.0"}, [14, 9, 11, 9, 9, 11, 12, 11, 16, 80])
    ws.freeze_panes = "A2"

    # Designs
    ws = sheets["Designs"]
    put_table(ws, 1, ["Part", "Design", "Status", "Design progress", "Started", "Finished", "Wind tunnel hours", "CFD hours",
                      "Engineers now", "Built so far", "In stock", "Best condition", "On car"], D["designs"],
              {4: PCT, 5: "d mmm yyyy", 6: "d mmm yyyy", 12: PCT}, [14, 8, 11, 10, 13, 13, 10, 8, 10, 8, 8, 10, 9])
    ws.freeze_panes = "A2"

    # Stock
    ws = sheets["Stock"]
    rows = [r[:5] + ['=IF(AND(D{r}="Ready",C{r}<{worn}),"Worn","")'.format(r=i, worn=S_WORN)]
            for i, r in enumerate(D["stock"], 2)]
    put_table(ws, 1, ["Part", "Design", "Condition", "State", "This design is on a car", "Flag"], rows, {3: PCT},
              [14, 8, 10, 12, 14, 10])
    ws.freeze_panes = "A2"

    # Spares
    ws = sheets["Spares"]
    ws["A1"], ws["A1"].font = "Races completed this season", BOLD
    ws["B1"] = D["races_done"]
    rows = []
    for i, r in enumerate(D["spares"], 4):
        rows.append(r[:7] + ["=IF($B$1>0,G{r}/$B$1,0)".format(r=i), r[8],
                             "=MAX({m},ROUNDUP(H{r}*I{r},0))".format(r=i, m=S_MINSPARE), "=E{r}-2".format(r=i),
                             "=MAX(0,2+J{r}-E{r})".format(r=i), r[12], "=L{r}*M{r}".format(r=i), r[14]])
    end = put_table(ws, 3, ["Part", "Newest finished design", "Cars running it", "Units ready", "Of which in good condition",
                            "Older designs in stock (fallback)", "Units lost this season", "Lost per race",
                            "Races to cover before a rebuild arrives", "Target spares", "Good spares now",
                            "To build", "Unit cost", "Cost", "Note"], rows, {8: "0.0", 13: MONEY, 14: MONEY},
                    [14, 10, 9, 8, 11, 13, 10, 8, 14, 8, 8, 8, 11, 12, 60])
    if rows:
        ws.cell(row=end + 1, column=1, value="Total").font = BOLD
        for col, fmt in (("L", "0"), ("N", MONEY)):
            cell = ws["%s%d" % (col, end + 1)]
            cell.value, cell.font, cell.number_format = "=SUM(%s4:%s%d)" % (col, col, end), BOLD, fmt
    for i, text in enumerate([
            "'To build' covers the two cars plus the target spares, counting only good units of the newest finished design.",
            "'Units lost' is everything built this season that is no longer in stock: worn out, crashed or scrapped.",
            "'Good condition' means at or above the worn threshold on the Settings tab; the minimum spares figure is there too.",
            "Older designs still work as emergency cover but are slower, so they are not counted as spares."], end + 3):
        ws.cell(row=i, column=1, value=text).font = FONT
    ws.freeze_panes = "A4"

    # Power unit
    ws = sheets["Power unit"]
    P = D["pu"]
    ws["A1"], ws["A1"].font = "Power unit: engine, ERS and gearbox wear, and when to take a grid penalty", TITLE
    info = [("Allocation for the season", P["limits"]), ("Grid penalties", P["penalties"]), ("Damage", P["damage"]),
            ("Your average grid position", "P%.1f" % P["avg_grid"] if P["avg_grid"] else "n/a")]
    for i, (label, value) in enumerate(info, 3):
        ws.cell(row=i, column=1, value=label).font = BOLD
        ws.cell(row=i, column=2, value=value).font = FONT
    rec = []
    for b_ in P["bundles"]:
        names = [x.lower() if x != "ERS" else "ERS" for x in b_["parts"]]
        what = names[0] if len(names) == 1 else ", ".join(names[:-1]) + " and " + names[-1]
        rec.append("Car %d: take a new %s%s at %s (in %d days)%s." % (
            b_["car"], what, " together" if len(names) > 1 else "", b_["race"], b_["days"],
            ", a %d-place grid penalty" % b_["places"] if b_["places"] else ", no penalty"))
    for r in P["summary"]:
        if r[4] is not None and r[6] is not None and r[6] - r[4] >= 0.05 and r[10].startswith("New unit"):
            rec.append("Until then, or if you decide against new parts: fit %s %s to %s (%d%%, against %d%% on the car)." % (
                r[1].lower() if r[1] != "ERS" else "ERS", r[5], r[0].lower(), r[6] * 100, r[4] * 100))
    if not rec:
        rec = ["No new parts needed: rotate the units you have, healthiest first."]
    ws.cell(row=8, column=1, value="Recommendation").font = BOLD
    for i, text in enumerate(rec):
        cell = ws.cell(row=8 + i, column=2, value=text)
        cell.font = Font(name="Arial", size=10, bold=True)
    top = 11 + len(rec)
    ws.cell(row=top - 1, column=1, value="Each car").font = TITLE
    end = put_table(ws, top, ["Car", "Component", "Used of allocation", "On the car", "Its condition", "Healthiest unit",
                              "Its condition", "Wear per race (estimate)", "Needed to finish the season",
                              "Usable before major damage", "Verdict"], P["summary"],
                    {5: PCT, 7: PCT, 8: PCT, 9: PCT, 10: PCT}, [24, 22, 12, 10, 11, 11, 11, 11, 12, 12, 80])
    sum_rows = (top + 1, end)
    for row in ws.iter_rows(min_row=top + 1, max_row=end):
        row[10].alignment = Alignment(wrap_text=True, vertical="top")
    top = end + 3
    ws.cell(row=top - 1, column=1, value="Race by race: which unit to run").font = TITLE
    end = put_table(ws, top, ["Race", "Days away", "Car 1 engine", "Car 1 ERS", "Car 1 gearbox", "Car 2 engine", "Car 2 ERS",
                              "Car 2 gearbox"], P["plan"])
    for row in ws.iter_rows(min_row=top + 1, max_row=end):
        for cell in row[2:8]:
            if cell.value and str(cell.value).startswith("NEW"):
                cell.font = Font(name="Arial", size=10, bold=True, color="C00000")
    for col in "CDEFGH":
        ws.column_dimensions[col].width = max(ws.column_dimensions[col].width or 0, 27)
    ws.column_dimensions["A"].width = 32
    top = end + 3
    ws.cell(row=top - 1, column=1, value="Every unit").font = TITLE
    end = put_table(ws, top, ["Car", "Component", "Unit", "Condition", "State", "On the car now"], P["units"], {4: PCT})
    unit_rows = (top + 1, end)
    notes = ["Wear per race is this season's average for that car and component: everything used up so far, divided by races run.",
             "Sprint weekends are counted as 1.25 races. Swapping between units already in your allocation never costs a penalty.",
             "The plan runs the healthiest unit each weekend and calls for a new one only when even that would end the weekend in major damage.",
             "If you need more than one new part, taking them in the same weekend means one trip to the back of the grid instead of two.",
             "Allocations reset next season, so a new unit now only has to last the races that are left."]
    for i, text in enumerate(notes, end + 2):
        ws.cell(row=i, column=1, value=text).font = FONT
    D["pu_rows"] = dict(summary=sum_rows, units=unit_rows)

    # Races
    ws = sheets["Races"]
    put_table(ws, 1, ["Race", "Days away", "Date", "Crucial", "Also recommended", "Our worst rank in a crucial area"],
              D["races"], {3: "d mmm yyyy"}, [26, 9, 13, 28, 30, 16])
    ws.freeze_panes = "A2"

    # Manufacturing
    ws = sheets["Manufacturing"]
    for i, (label, value) in enumerate([("Factory build speed bonus", D["factory_bonus"]),
                                        ("Build work per day, normal", "=%s*(1+B1)" % S_NORMAL),
                                        ("Build work per day, rushed", "=%s*(1+B1)" % S_RUSHED)], 1):
        ws.cell(row=i, column=1, value=label).font = BOLD
        ws.cell(row=i, column=2, value=value).font = FONT
    ws["B1"].number_format = "0.0%"
    rows = [[p, work, "=ROUNDUP(B{r}/$B$2,0)".format(r=i), "=ROUNDUP(B{r}/$B$3,0)".format(r=i), cost,
             "=E{r}*{x}".format(r=i, x=S_RUSHX), "=E{r}*{x}".format(r=i, x=S_EMERGX), "=F{r}-E{r}".format(r=i)]
            for i, (p, work, cost) in enumerate(D["manufacturing"], 6)]
    put_table(ws, 5, ["Part", "Build work", "Normal days", "Rushed days", "Normal cost", "Rushed cost",
                      "Emergency cost (instant)", "Extra to rush"], rows, {5: MONEY, 6: MONEY, 7: MONEY, 8: MONEY},
              [30, 10, 9, 9, 12, 12, 14, 12])
    ws.cell(row=len(rows) + 8, column=1, value="Days are per part. Build rates are on the Settings tab.").font = FONT

    # Staff
    ws = sheets["Staff"]
    rows = [r[:11] + ['=TRIM(IF(AND(E{r}=Overview!$B$7,K{r}=""),"Contract ends this season, nothing agreed. ","")'
                      '&IF(LEFT(K{r},7)="Leaving","Needs a replacement. ","")&IF(AND(C{r}<>"",C{r}<{m}),"Low mentality",""))'
                      .format(r=i, m=S_MENT)] for i, r in enumerate(D["staff"], 2)]
    put_table(ws, 1, ["Name", "Role", "Mentality (0-100)", "Mood", "Contract ends", "Salary", "Rating (avg of stats)",
                      "Car opinion", "Driver opinion", "HQ opinion", "Next season", "Flag"], rows, {6: MONEY, 7: "0.0"},
              [24, 28, 10, 11, 9, 12, 10, 8, 8, 8, 30, 44])
    ws.freeze_panes = "A2"

    # Facilities
    ws = sheets["Facilities"]
    rows = []
    for i, r in enumerate(D["facilities"], 2):
        rows.append(r[:3] + ['=IF(C{r}="","",IF(C{r}>0.75,1,IF(C{r}>0.5,0.8,IF(C{r}>0.25,0.6,0.4))))'.format(r=i)] + r[4:9] +
                    ['=IF(B{r}=0,"Not built",IF(D{r}<1,"Refurbish: running at "&TEXT(D{r},"0%"),'
                     'IF(B{r}<E{r}-0.5,"Below grid average: consider upgrade","Leave alone")))'.format(r=i)])
    put_table(ws, 1, ["Building", "Level", "Condition", "Effect now", "Grid average level", "Refurbish cost", "Upgrade cost",
                      "Upkeep per month", "What the next level gives", "Suggested action"], rows,
              {3: PCT, 4: PCT, 5: "0.0", 6: MONEY, 7: MONEY, 8: MONEY}, [24, 7, 10, 9, 10, 13, 13, 12, 70, 38])
    ws["D1"].comment = Comment("Condition tiers from the save: above 75% full effect, 50-75% 80%, 25-50% 60%, below 40%.", "updater")
    ws.freeze_panes = "A2"

    # Scouting
    ws = sheets["Scouting"]
    rows = [r[:4] + ["=C%d-D%d" % (i, i)] + r[5:] for i, r in enumerate(D["scouting"], 2)]
    put_table(ws, 1, ["Role", "Name", "Rating (avg of stats)", "Your current", "Difference", "Opinion of your team",
                      "Acceptance threshold", "Age", "Current team", "Contract ends", "Current salary", "Scouted",
                      "Compared with", "Super licence", "Next season"], rows, {3: "0.0", 4: "0.0", 5: "0.0", 7: "0.00", 11: MONEY},
              [22, 24, 10, 9, 10, 10, 11, 6, 26, 9, 12, 8, 24, 9, 30])
    ws.freeze_panes = "A2"

    # Pit crew
    ws = sheets["Pit crew"]
    rows = [r[:4] + ["=B%d-C%d" % (i, i)] + r[5:] for i, r in enumerate(D["crew"], 2)]
    end = put_table(ws, 1, ["Stat", "Yours", "Grid average", "Grid best", "Gap to average", rank_head], rows,
                    {2: "0.0", 3: "0.0", 4: "0.0", 5: "0.0"}, [34, 10, 12, 10, 12, 10])
    ws.cell(row=end + 1, column=1, value="Fatigue: lower is better, and it is ranked that way. Every other stat: higher is better.").font = FONT
    top = end + 4
    ws.cell(row=top - 1, column=1, value="Average pit stop this season, seconds per stage").font = TITLE
    rows = [r[:5] + ["=(B{r}+C{r})-(D{r}+E{r})".format(r=i)] for i, r in enumerate(D["stages"], top + 1)]
    end = put_table(ws, top, ["Stage", "Yours", "Your error delay", "Rivals", "Rivals' error delay", "Time lost to rivals"],
                    rows, {j: "0.00" for j in (2, 3, 4, 5, 6)})
    if rows:
        ws.cell(row=end + 1, column=1, value="Total").font = BOLD
        for col in "BCDEF":
            cell = ws["%s%d" % (col, end + 1)]
            cell.value, cell.font, cell.number_format = "=SUM(%s%d:%s%d)" % (col, top + 1, col, end), BOLD, "0.00"

    # Drivers
    ws = sheets["Drivers"]
    names = D["driver_names"]
    end = put_table(ws, 1, ["Stat"] + names, D["driver_stats"], None, [16] + [20] * len(names) + [10, 10])
    if D["driver_stats"]:
        ws.cell(row=end + 1, column=1, value="Average").font = BOLD
        for j in range(2, len(names) + 2):
            col = get_column_letter(j)
            cell = ws["%s%d" % (col, end + 1)]
            cell.value, cell.font, cell.number_format = "=AVERAGE(%s2:%s%d)" % (col, col, end), BOLD, "0.0"
    top = end + 4
    ws.cell(row=top - 1, column=1, value="This season").font = TITLE
    put_table(ws, top, ["Driver", "Races", "Average finish", "Best finish", "Average grid", "DNFs", "Points"],
              D["driver_results"], {3: "0.0", 5: "0.0"})

    # Finance
    ws = sheets["Finance"]
    end = put_table(ws, 1, ["This season by type", "Amount"], D["finance"], {2: MONEY + ";-" + MONEY}, [34, 16])
    for label, formula in (("Income", '=SUMIF(B2:B%d,">0")' % end), ("Spending", '=SUMIF(B2:B%d,"<0")' % end),
                           ("Net", "=SUM(B2:B%d)" % end), ("Balance now", "=Overview!$B$10")):
        end += 1
        ws.cell(row=end + 1, column=1, value=label).font = BOLD
        cell = ws.cell(row=end + 1, column=2, value=formula)
        cell.font, cell.number_format = BOLD, MONEY + ";-" + MONEY

    # Womens team
    ws = sheets["Womens team"]
    if D["women"]:
        ws["A1"], ws["A1"].font = "All-female team project", TITLE
        ws["A2"] = "Seats held by women now or signed for next season"
        ws["E2"] = D["women_on_team"]
        ws["A3"] = "Women drivers holding a super licence"
        ws["E3"] = '=COUNTIFS($A$6:$A$%d,"Driver",$G$6:$G$%d,"Yes")' % (len(D["women"]) + 5, len(D["women"]) + 5)
        rows = [r[:4] + ["=IF(OR(C{r}=\"\",D{r}=\"\"),\"\",C{r}-D{r})".format(r=i)] + r[5:] for i, r in enumerate(D["women"], 6)]
        end = put_table(ws, 5, ["Role", "Name", "Rating (avg of stats)", "Your best in that role", "Difference",
                                "Improvability", "Super licence", "Raced enough for F1", "Racing this season",
                                "Opinion of your team", "Attached to", "Contract ends", "Suggested step"], rows,
                        {3: "0.0", 4: "0.0", 5: "0.0"}, [20, 24, 10, 10, 10, 12, 8, 10, 24, 10, 34, 9, 70])
        for i, text in enumerate([
                "Drivers need a super licence for a race seat, so the route is: sign as affiliates, let them race F2 and F3, promote when licensed.",
                "Improvability is the game's own 0-100 figure for how far a driver can still grow.",
                "Turn this tab off with 'All-female team project' on the Settings tab."], end + 2):
            ws.cell(row=i, column=1, value=text).font = FONT
        ws.freeze_panes = "A6"
    else:
        ws["A1"] = "The all-female team project is switched off. Set 'All-female team project' on the Settings tab to 1 to turn it on."

    for ws in sheets.values():
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None and cell.font.name != "Arial":
                    cell.font = FONT
    wb.active = 0


# ---------------------------------------------------------------- game-style graphics

BG, PANEL, PANEL2, LINE = "15151E", "1F1F2B", "282836", "3A3A4E"
RED, WHITE, TEXT, MUTED = "E10600", "FFFFFF", "E6E6EE", "9A9AAF"
GOLD, GREEN, AMBER, CYAN = "FFC800", "00C46A", "FF9F1C", "00B8D9"
SCALE_BAD, SCALE_MID, SCALE_GOOD = "B3141E", "2C2C3A", "0A8A55"
TAB_COLOURS = {"Overview": RED, "Action plan": RED, "Board": RED, "Power unit": RED, "Car vs grid": CYAN, "Part summary": CYAN,
               "Part stats": CYAN, "Designs": CYAN, "Stock": CYAN, "Spares": CYAN, "Races": CYAN, "Manufacturing": CYAN,
               "Staff": GOLD, "Scouting": GOLD, "Pit crew": GOLD, "Drivers": GOLD, "Womens team": GOLD,
               "Facilities": GREEN, "Finance": GREEN}


def _fill(colour):
    return PatternFill("solid", fgColor=colour)


def _skin_sheet(ws):
    """Dark pit-wall look: dark background, red header bars, banded table rows, white text, gold for editable cells."""
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = TAB_COLOURS.get(ws.title, LINE)
    last_col = max(ws.max_column, 12) + 10
    for col in range(1, last_col + 1):
        ws.column_dimensions[get_column_letter(col)].fill = _fill(BG)   # empty cells beyond the data
    headers = {}
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            rgb = cell.fill.fgColor.rgb if cell.fill and cell.fill.fill_type == "solid" else None
            if rgb and str(rgb).endswith("E7E6E6"):
                headers.setdefault(cell.row, []).append(cell.column)
    banded = {}
    for r, cols in headers.items():
        c1, c2 = min(cols), max(cols)
        rr = r + 1
        while rr <= ws.max_row and any(ws.cell(row=rr, column=c).value not in (None, "") for c in range(c1, c2 + 1)):
            for c in range(c1, c2 + 1):
                banded[(rr, c)] = PANEL if (rr - r) % 2 else PANEL2
            rr += 1
    thin = Side(style="thin", color=LINE)
    for row in ws.iter_rows(min_row=1, max_row=max(ws.max_row, 1) + 5, max_col=last_col):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            f = cell.font
            colour = f.color.rgb if f is not None and f.color is not None and isinstance(f.color.rgb, str) else ""
            if cell.row in headers and cell.column in headers[cell.row]:
                cell.fill = _fill(RED)
                cell.font = Font(name="Arial", size=10, bold=True, color=WHITE)
                cell.alignment = Alignment(wrap_text=True, vertical="center")
                continue
            cell.fill = _fill(banded.get((cell.row, cell.column), BG))
            if (cell.row, cell.column) in banded:
                cell.border = Border(bottom=thin)
            if f is not None and f.sz and f.sz >= 12 and f.b:
                cell.font = Font(name="Arial Black", size=13, color=WHITE)
                cell.border = Border(bottom=Side(style="medium", color=RED))
            elif colour.endswith("0000FF"):
                cell.font = Font(name="Arial", size=10, bold=True, color=GOLD)
            else:
                cell.font = Font(name="Arial", size=f.sz or 10 if f is not None else 10, bold=bool(f is not None and f.b),
                                 italic=bool(f is not None and f.i), color=TEXT)
    for r, cols in headers.items():
        lines = 1
        for c in cols:
            text = str(ws.cell(row=r, column=c).value or "")
            width = ws.column_dimensions[get_column_letter(c)].width or 10
            lines = max(lines, -(-len(text) // max(1, int(width * 1.1))))
        ws.row_dimensions[r].height = max(ws.row_dimensions[r].height or 0, 14 * lines + 8)


def _rich(colour=TEXT, size=900, bold=False):
    cp = CharacterProperties(solidFill=colour, sz=size, b=bold, latin=DrawingFont(typeface="Arial"))
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def _dark_chart(ch, title, width=16, height=8):
    cp = CharacterProperties(solidFill=WHITE, sz=1200, b=True, latin=DrawingFont(typeface="Arial"))
    ch.title = Title(tx=Text(rich=RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=title)])])),
                     overlay=False)
    ch.graphical_properties = GraphicalProperties(solidFill=PANEL, ln=LineProperties(solidFill=LINE))
    ch.plot_area.graphicalProperties = GraphicalProperties(solidFill=PANEL, ln=LineProperties(noFill=True))
    for ax in (getattr(ch, "x_axis", None), getattr(ch, "y_axis", None)):
        if ax is None:
            continue
        ax.delete = False
        ax.txPr = _rich()
        ax.tickLblPos = "low"
        ax.spPr = GraphicalProperties(ln=LineProperties(solidFill=LINE))
        if getattr(ax, "majorGridlines", None) is not None:
            ax.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=LINE))
    if ch.legend is not None:
        ch.legend.txPr = _rich()
        ch.legend.position = "b"
    ch.width, ch.height = width, height


def _colour_points(series, colours):
    for i, colour in enumerate(colours):
        pt = DataPoint(idx=i)
        pt.graphicalProperties = GraphicalProperties(solidFill=colour, ln=LineProperties(noFill=True))
        series.dPt.append(pt)


def _scale(ws, rng, low_good=False, mid=None):
    lo, hi = (SCALE_GOOD, SCALE_BAD) if low_good else (SCALE_BAD, SCALE_GOOD)
    if mid is None:
        ws.conditional_formatting.add(rng, ColorScaleRule(start_type="min", start_color=lo, mid_type="percentile",
                                                          mid_value=50, mid_color=SCALE_MID, end_type="max", end_color=hi))
    else:
        ws.conditional_formatting.add(rng, ColorScaleRule(start_type="min", start_color=lo, mid_type="num", mid_value=mid,
                                                          mid_color=SCALE_MID, end_type="max", end_color=hi))


def _bar(ws, rng):
    ws.conditional_formatting.add(rng, DataBarRule(start_type="min", end_type="max", color="0A6E85", showValue=True))


def _tile(ws, top, col, label, formula, fmt=None, accent=RED):
    """A dashboard tile: small caps label over a big number, two columns wide."""
    c1, c2 = get_column_letter(col), get_column_letter(col + 1)
    ws.merge_cells("%s%d:%s%d" % (c1, top, c2, top))
    ws.merge_cells("%s%d:%s%d" % (c1, top + 1, c2, top + 2))
    lab = ws["%s%d" % (c1, top)]
    lab.value, lab.font = label.upper(), Font(name="Arial", size=8, bold=True, color=MUTED)
    lab.alignment = Alignment(horizontal="center", vertical="bottom")
    val = ws["%s%d" % (c1, top + 1)]
    val.value, val.font = formula, Font(name="Arial Black", size=18, color=WHITE)
    val.alignment = Alignment(horizontal="center", vertical="center")
    if fmt:
        val.number_format = fmt
    for r in range(top, top + 3):
        for c in (col, col + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = _fill(PANEL)
    for c in (col, col + 1):
        ws.cell(row=top, column=c).border = Border(top=Side(style="thick", color=accent))


def apply_game_style(wb, D):
    names = [n for n in OWNED_SHEETS if n in wb.sheetnames]
    for n in names:
        _skin_sheet(wb[n])

    # Overview: dashboard tiles and the car's gap to the grid
    ws = wb["Overview"]
    look = lambda label: 'IFERROR(INDEX($B:$B,MATCH("%s",$A:$A,0)),"-")' % label
    for col in range(4, 13):
        ws.column_dimensions[get_column_letter(col)].width = 11
    for r in (3, 4, 5, 7, 8, 9):
        ws.row_dimensions[r].height = 20
    _tile(ws, 3, 4, "Championship", '="P"&%s' % look("Championship position"))
    _tile(ws, 3, 6, "Points", "=" + look("Points"))
    _tile(ws, 3, 8, "Balance", '=IFERROR(%s/1000000,"-")' % look("Balance"), '"$"0.0"M"', GREEN)
    _tile(ws, 3, 10, "Board target", "=IFERROR(IF(Board!B3=\"\",\"-\",Board!B3),\"-\")", None, GOLD)
    _tile(ws, 7, 4, "Next race in", '=%s&" days"' % look("Days to next race"), None, CYAN)
    _tile(ws, 7, 6, "Board confidence", '=IFERROR(VALUE(LEFT(Board!B6,FIND(" ",Board!B6)-1)),"-")', None, GOLD)
    _tile(ws, 7, 8, "Cap headroom", '=IFERROR(%s/1000000,"-")' % look("Cap headroom at season end"), '"$"0.0"M"', GREEN)
    _tile(ws, 7, 10, "Idle engineers", "=" + look("Engineers idle"), None, CYAN)
    flag_top = next((r for r in range(1, ws.max_row + 1) if ws.cell(row=r, column=1).value == "Things to look at"), None)
    if flag_top:
        for r in range(flag_top + 1, ws.max_row + 1):
            cell = ws.cell(row=r, column=1)
            if cell.value and not str(cell.value).startswith("▸"):
                cell.value = "▸  " + str(cell.value)
                cell.font = Font(name="Arial", size=10, color=AMBER if cell.value != "▸  Nothing flagged" else GREEN)
    cvg = wb["Car vs grid"]
    n = len(D["areas"])
    if n:
        ch = BarChart()
        ch.type, ch.grouping = "bar", "clustered"
        ch.add_data(Reference(cvg, min_col=6, min_row=1, max_row=n + 1), titles_from_data=True)
        ch.set_categories(Reference(cvg, min_col=1, min_row=2, max_row=n + 1))
        _dark_chart(ch, "Car: gap to the grid median", 17, 8.5)
        ch.legend = None
        _colour_points(ch.series[0], [GREEN if (a[3] or 0) - (a[4] or 0) >= 0 else RED for a in D["areas"]])
        ch.y_axis.title = None
        ws.add_chart(ch, "D12")

    # Car vs grid: spider chart and gap bars
    if n:
        _scale(cvg, "F2:F%d" % (n + 1), mid=0)
        _scale(cvg, "G2:G%d" % (n + 1), low_good=True)
        radar = RadarChart()
        radar.type = "marker"
        radar.add_data(Reference(cvg, min_col=4, max_col=5, min_row=1, max_row=n + 1), titles_from_data=True)
        radar.set_categories(Reference(cvg, min_col=1, min_row=2, max_row=n + 1))
        _dark_chart(radar, "Your newest parts against the grid median", 16, 10)
        for s_, colour in zip(radar.series, (RED, CYAN)):
            s_.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=colour, w=28575))
            s_.marker.symbol = "circle"
            s_.marker.graphicalProperties = GraphicalProperties(solidFill=colour, ln=LineProperties(solidFill=colour))
        cvg.add_chart(radar, "A%d" % (n + 5))
        bar = BarChart()
        bar.type = "bar"
        bar.add_data(Reference(cvg, min_col=6, min_row=1, max_row=n + 1), titles_from_data=True)
        bar.set_categories(Reference(cvg, min_col=1, min_row=2, max_row=n + 1))
        _dark_chart(bar, "Gap to the grid median", 16, 10)
        bar.legend = None
        _colour_points(bar.series[0], [GREEN if (a[3] or 0) - (a[4] or 0) >= 0 else RED for a in D["areas"]])
        cvg.add_chart(bar, "E%d" % (n + 5))

    # Part summary: which part is furthest behind
    ps = wb["Part summary"]
    m = len(D["summary"])
    if m:
        _scale(ps, "G2:G%d" % (m + 1), mid=0)
        _scale(ps, "H2:H%d" % (m + 1), low_good=True)
        bar = BarChart()
        bar.add_data(Reference(ps, min_col=7, min_row=1, max_row=m + 1), titles_from_data=True)
        bar.set_categories(Reference(ps, min_col=1, min_row=2, max_row=m + 1))
        _dark_chart(bar, "Average gap to the grid median, by part", 16, 8)
        bar.legend = None
        bar.series[0].graphicalProperties = GraphicalProperties(solidFill=CYAN, ln=LineProperties(noFill=True))
        ps.add_chart(bar, "A%d" % (m + 4))
    st = wb["Part stats"]
    k = len(D["part_stats"])
    if k:
        _scale(st, "H2:H%d" % (k + 1), mid=0)
        _scale(st, "I2:I%d" % (k + 1), low_good=True)

    # Designs, stock, races, scouting, facilities, staff, women's team: colour where it helps the eye
    if D["designs"]:
        _bar(wb["Designs"], "D2:D%d" % (len(D["designs"]) + 1))
        _scale(wb["Designs"], "L2:L%d" % (len(D["designs"]) + 1))
    if D["stock"]:
        _scale(wb["Stock"], "C2:C%d" % (len(D["stock"]) + 1))
    if D["races"]:
        _bar(wb["Races"], "B2:B%d" % (len(D["races"]) + 1))
        _scale(wb["Races"], "F2:F%d" % (len(D["races"]) + 1), low_good=True)
    if D["scouting"]:
        _scale(wb["Scouting"], "E2:E%d" % (len(D["scouting"]) + 1), mid=0)
        _scale(wb["Scouting"], "F2:F%d" % (len(D["scouting"]) + 1))
    if D["facilities"]:
        _scale(wb["Facilities"], "C2:C%d" % (len(D["facilities"]) + 1))
        _scale(wb["Facilities"], "D2:D%d" % (len(D["facilities"]) + 1))
    sf = wb["Staff"]
    rows = len(D["staff"])
    if rows:
        _scale(sf, "C2:C%d" % (rows + 1))
        for col in "HIJ":
            _scale(sf, "%s2:%s%d" % (col, col, rows + 1), mid=0)
        people = [i for i, r in enumerate(D["staff"]) if r[2] is not None and "next season" not in r[1]]
        if people:
            bar = BarChart()
            bar.type = "bar"
            last = max(people) + 2
            bar.add_data(Reference(sf, min_col=3, min_row=1, max_row=last), titles_from_data=True)
            bar.set_categories(Reference(sf, min_col=1, min_row=2, max_row=last))
            _dark_chart(bar, "Mentality (0-100)", 16, 8)
            bar.legend = None
            bar.y_axis.scaling.min, bar.y_axis.scaling.max = 0, 100
            _colour_points(bar.series[0], [(GREEN if (r[2] or 0) >= 60 else AMBER if (r[2] or 0) >= 40 else RED)
                                           for r in D["staff"][:last - 1]])
            sf.add_chart(bar, "A%d" % (rows + 4))
    if D.get("women") and "Womens team" in wb.sheetnames:
        _scale(wb["Womens team"], "E6:E%d" % (len(D["women"]) + 5), mid=0)

    # Power unit: condition colours
    if "Power unit" in wb.sheetnames and D.get("pu_rows"):
        pw = wb["Power unit"]
        a, b = D["pu_rows"]["summary"]
        if b >= a:
            for col in "EG":
                _scale(pw, "%s%d:%s%d" % (col, a, col, b))
        a, b = D["pu_rows"]["units"]
        if b >= a:
            _scale(pw, "D%d:D%d" % (a, b))
        for row in pw.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("NEW"):
                    cell.font = Font(name="Arial", size=10, bold=True, color=AMBER)

    # Spares: what to build stands out
    sp = wb["Spares"]
    if D["spares"]:
        _scale(sp, "L4:L%d" % (len(D["spares"]) + 3), low_good=True)

    # Pit crew: yours against the grid
    pc = wb["Pit crew"]
    c = len(D["crew"])
    if c:
        _scale(pc, "E2:E%d" % (c + 1), mid=0)
        _scale(pc, "F2:F%d" % (c + 1), low_good=True)
        bar = BarChart()
        shown = c - (1 if D["crew"][-1][0] == "Fatigue" else 0)
        bar.add_data(Reference(pc, min_col=2, max_col=3, min_row=1, max_row=shown + 1), titles_from_data=True)
        bar.set_categories(Reference(pc, min_col=1, min_row=2, max_row=shown + 1))
        _dark_chart(bar, "Pit crew: yours against the grid average (fatigue left out)", 18, 9)
        for s_, colour in zip(bar.series, (RED, MUTED)):
            s_.graphicalProperties = GraphicalProperties(solidFill=colour, ln=LineProperties(noFill=True))
        pc.add_chart(bar, "H2")

    # Drivers: spider chart of their stats
    dr = wb["Drivers"]
    names = D["driver_names"]
    if names and D["driver_stats"]:
        rows = len(D["driver_stats"])
        radar = RadarChart()
        radar.type = "marker"
        radar.add_data(Reference(dr, min_col=2, max_col=1 + len(names), min_row=1, max_row=rows + 1), titles_from_data=True)
        radar.set_categories(Reference(dr, min_col=1, min_row=2, max_row=rows + 1))
        _dark_chart(radar, "Driver stats", 15, 11)
        for s_, colour in zip(radar.series, (RED, CYAN, GOLD, GREEN)):
            s_.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=colour, w=28575))
            s_.marker.symbol = "circle"
            s_.marker.graphicalProperties = GraphicalProperties(solidFill=colour, ln=LineProperties(solidFill=colour))
        dr.add_chart(radar, "%s2" % get_column_letter(len(names) + 3))

    # Finance: where the money comes from and goes
    fi = wb["Finance"]
    f = len(D["finance"])
    if f:
        bar = BarChart()
        bar.type = "bar"
        bar.add_data(Reference(fi, min_col=2, min_row=1, max_row=f + 1), titles_from_data=True)
        bar.set_categories(Reference(fi, min_col=1, min_row=2, max_row=f + 1))
        _dark_chart(bar, "This season's money, by type", 17, max(8, 0.55 * f))
        bar.legend = None
        _colour_points(bar.series[0], [GREEN if v >= 0 else RED for _n, v in D["finance"]])
        fi.add_chart(bar, "D2")

    # Board: what each finishing place does to next season's rating
    Bd = D.get("board")
    if Bd:
        bd = wb["Board"]
        bd.merge_cells("B11:C11")
        bd["B11"].alignment = Alignment(wrap_text=True, vertical="top")
        bd.row_dimensions[11].height = 45
        nf = len(Bd["finish"])
        bar = BarChart()
        bar.add_data(Reference(bd, min_col=3, min_row=14, max_row=14 + nf), titles_from_data=True)
        bar.set_categories(Reference(bd, min_col=1, min_row=15, max_row=14 + nf))
        _dark_chart(bar, "Next season's team rating by where you finish (gold = target)", 18, 6.5)
        bar.legend = None
        _colour_points(bar.series[0], [GOLD if r[0] == Bd["target"] else GREEN if r[1] == "Met" else RED for r in Bd["finish"]])
        bd.add_chart(bar, "D2")
        diff_top = next((r for r in range(1, bd.max_row + 1) if bd.cell(row=r, column=1).value == "Role"), None)
        if diff_top:
            _scale(bd, "D%d:D%d" % (diff_top + 1, diff_top + len(Bd["bench"])), mid=0)

    # Action plan: spending gauge and value bars
    ap = wb["Action plan"]
    acts = D["actions"]
    if acts:
        first, last = 8, 7 + len(acts)
        _bar(ap, "I%d:I%d" % (first, last))
        ap.conditional_formatting.add("H%d:H%d" % (first, last), ColorScaleRule(
            start_type="num", start_value=1, start_color=SCALE_MID, end_type="num", end_value=5, end_color="8C6D00"))
        for r in range(1, 7):
            ap.row_dimensions[r].height = 22
        dn = DoughnutChart()
        dn.add_data(Reference(ap, min_col=6, min_row=last + 1, max_row=last + 2))
        dn.set_categories(Reference(ap, min_col=2, min_row=last + 1, max_row=last + 2))
        _dark_chart(dn, "Budget: planned against left", 11, 4.6)
        dn.holeSize = 60
        _colour_points(dn.series[0], [RED, GREEN])
        dn.legend.position = "r"
        ap.add_chart(dn, "E1")


def update(save_path, workbook, team=""):
    new_book = not os.path.exists(workbook)
    wb = None if new_book else load_workbook(workbook)
    settings = read_settings(wb)
    with tempfile.TemporaryDirectory() as tmp:
        D = extract(unpack_save(save_path, tmp), settings)
    D["display_team"] = team
    if wb is None:
        wb = Workbook()
        del wb[wb.sheetnames[0]]
    ensure_settings(wb)
    if new_book and os.path.exists(DEFAULT_WORKBOOK) and os.path.abspath(workbook) != os.path.abspath(DEFAULT_WORKBOOK):
        seed_settings(wb, DEFAULT_WORKBOOK)
    build(wb, D, save_path)
    if settings[15]:
        apply_game_style(wb, D)
    # Write to a temporary file first, then swap it in. If the update is interrupted
    # (window closed, PC shut down) the existing workbook is left intact, never half-written.
    tmp = workbook + ".tmp"
    try:
        wb.save(tmp)
        os.replace(tmp, workbook)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    print("Updated %s from %s (%d flags)" % (os.path.basename(workbook), os.path.basename(save_path), len(D["flags"])))
    return save_path


def main():
    ap = argparse.ArgumentParser(description="Update an F1 Manager 2024 planner workbook from a save.")
    ap.add_argument("--version", action="version", version="F1 Manager 2024 Planner %s" % VERSION)
    ap.add_argument("--save-dir", default=DEFAULT_SAVE_DIR, help="folder holding the .sav files")
    ap.add_argument("--save", help="use this exact save file")
    ap.add_argument("--newest", action="store_true", help="use the newest save without asking")
    ap.add_argument("--list", action="store_true", help="list the saves and stop")
    ap.add_argument("--workbook", help="workbook to write (default: one per career, named after the team)")
    ap.add_argument("--open", action="store_true", help="open the workbook afterwards (Windows)")
    ap.add_argument("--watch", action="store_true", help="keep running and update whenever the chosen career saves")
    args = ap.parse_args()

    if args.save:
        chosen = read_header(args.save)
        chosen.update(path=args.save, file=os.path.basename(args.save), mtime=os.path.getmtime(args.save))
    else:
        saves = list_saves(args.save_dir)
        if args.list:
            for i, info in enumerate(saves, 1):
                print("  %d  %s" % (i, describe(info)))
            return
        interactive = sys.stdin is not None and sys.stdin.isatty()
        chosen = saves[0] if (args.newest or len(saves) == 1 or not interactive) else choose_save(saves)
    workbook = args.workbook or workbook_for(chosen["team"])
    print("\nUsing %s (%s) -> %s" % (chosen["file"], chosen["team"], os.path.basename(workbook)))

    def run_once(path):
        try:
            update(path, workbook, chosen["team"])
            return True
        except PermissionError:
            print("The workbook is open in Excel. Close it and it will be updated.")
            return False

    if not args.watch:
        ok = run_once(chosen["path"])
        if ok and args.open and hasattr(os, "startfile"):
            os.startfile(workbook)
        sys.exit(0 if ok else 1)

    print("Watching for saves of this career (%s) - press Ctrl+C to stop." % chosen["team"])
    done_for = None
    while True:
        try:
            if args.save:
                path = args.save
            else:   # follow the career, whichever of its save files is newest (autosave or a manual save)
                mine = [x for x in list_saves(args.save_dir) if x["career"] == chosen["career"]]
                path = mine[0]["path"] if mine else chosen["path"]
            stamp = (path, os.path.getmtime(path))
            if stamp != done_for:
                time.sleep(3)  # let the game finish writing
                if run_once(path):
                    done_for = (path, os.path.getmtime(path))
        except KeyboardInterrupt:
            break
        except Exception as e:  # keep watching through a half-written save
            print("Could not read the save yet: %s" % e)
        time.sleep(5)


if __name__ == "__main__":
    main()
