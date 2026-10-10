#!/usr/bin/env python3
"""Импорт «Универсальной таблицы» в Supabase (запускается у вас на компьютере).

  pip install pandas openpyxl requests
  python import_students.py Универсальная_таблица.xlsx --dry-run      # только отчёт, ничего не отправляется
  SUPABASE_URL=https://XXXX.supabase.co SUPABASE_SERVICE_KEY=... python import_students.py Универсальная_таблица.xlsx

Ключ service_role (Settings → API) никуда не коммитьте и на сайт не кладите.
Скрипт рассчитан на ПУСТУЮ базу (после schema.sql + patch_v2.sql); если дети уже есть, остановится (--force отменяет проверку).
"""
import os, re, sys, uuid, collections
import pandas as pd

DEFAULT_START = "2025-09-01"      # если у ребёнка нет даты договора
BATCH = 500

def clean(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NaT: return None
    s = str(v).strip()
    return s or None

def norm_name(s): return re.sub(r"\s+", " ", s.replace("ё", "е").replace("Ё", "Е")).strip().lower()

def norm_group(s):
    s = clean(s) or "без названия"
    s = re.sub(r"\s*лет\s*$", "", s, flags=re.I)
    return re.sub(r"\s*[-–—]\s*", "-", s).strip()

def price_type(group):
    if re.match(r"^(1[5-9]|[2-9]\d)\+", group) or "взросл" in group.lower(): return "Подростки и взрослые"
    m = re.search(r"(\d+)-(\d+)", group)
    if not m: return None
    mid = (int(m[1]) + int(m[2])) / 2
    if mid <= 6.5: return "Группа 3–6 лет"
    if mid <= 9.5: return "Группа 7–9 лет"
    if mid < 15: return "Группа 10–14 лет"
    return "Подростки и взрослые"

def to_date(v):
    d = pd.to_datetime(v, errors="coerce", dayfirst=True)
    return None if pd.isna(d) else d.date().isoformat()

def contract_date(c):
    m = re.search(r"от\s+(\d{2})\.(\d{2})\.(\d{4})", c or "")
    return f"{m[3]}-{m[2]}-{m[1]}" if m else None

def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    dry, force = "--dry-run" in sys.argv, "--force" in sys.argv
    if not args: sys.exit(__doc__)
    df = pd.read_excel(args[0], sheet_name=0)
    df.columns = ["n", "contract", "name", "dob", "age", "schedule", "parent", "phone", "trial", "notes", "teacher", "group"]

    rows, skipped = [], collections.Counter()
    for r in df.to_dict("records"):
        name = clean(r["name"])
        if not name: skipped["без ФИО"] += 1; continue
        if "резерв" in name.lower(): skipped["строки «резерв»"] += 1; continue
        if not clean(r["teacher"]): skipped["без преподавателя"] += 1; continue
        rows.append(dict(name=name, key=norm_name(name), contract=clean(r["contract"]), dob=to_date(r["dob"]) if clean(r["dob"]) else None,
                         parent=clean(r["parent"]), phone=clean(r["phone"]), trial=to_date(r["trial"]) if clean(r["trial"]) else None,
                         notes=clean(r["notes"]), teacher=clean(r["teacher"]), group=norm_group(r["group"]), schedule=clean(r["schedule"])))

    # --- дети: одно имя = один ребёнок, если даты рождения не противоречат
    by_name = collections.defaultdict(list)
    for r in rows: by_name[r["key"]].append(r)
    children, child_of = [], {}
    for key, rs in by_name.items():
        dobs = sorted({r["dob"] for r in rs if r["dob"]})
        clusters = {d: [] for d in dobs} or {None: []}
        first = next(iter(clusters))
        for r in rs: clusters[r["dob"] if r["dob"] in clusters else first].append(r)
        for d, grp in clusters.items():
            pick = lambda f: next((x[f] for x in grp if x[f]), None)
            c = dict(id=str(uuid.uuid4()), full_name=grp[0]["name"], birth_date=d, parent_name=pick("parent"), phone=pick("phone"),
                     contract=pick("contract"), trial_date=pick("trial"), notes=pick("notes"))
            children.append(c)
            for x in grp: child_of[id(x)] = c["id"]

    # --- группы (преподаватель + название)
    groups = {}
    sched = collections.defaultdict(collections.Counter)
    for r in rows:
        groups.setdefault((r["teacher"], r["group"]), None)
        if r["schedule"]: sched[(r["teacher"], r["group"])][r["schedule"]] += 1

    # --- записи в группы (ребёнок + группа, берём самую раннюю дату договора)
    enr = {}
    for r in rows:
        k = (child_of[id(r)], (r["teacher"], r["group"]))
        start = contract_date(r["contract"]) or DEFAULT_START
        enr[k] = min(enr.get(k, start), start)

    no_price = sorted({g for (_, g) in groups if not price_type(g)})
    print(f"Строк в файле: {len(df)}; загружается: {len(rows)}; пропущено: {dict(skipped) or 0}")
    print(f"Преподавателей: {len({t for t, _ in groups})}; групп: {len(groups)}; детей: {len(children)} (дубли склеены); записей в группы: {len(enr)}")
    print("Без телефона:", sum(1 for c in children if not c['phone']), "| без даты рождения:", sum(1 for c in children if not c['birth_date']), "| e-mail в таблице нет, его заполните отдельно")
    print("Группы, для которых не удалось определить тип цены (задайте вручную на сайте/в базе):", no_price or "нет")
    if dry: return

    url, key = os.environ.get("SUPABASE_URL"), os.environ.get("SUPABASE_SERVICE_KEY")
    if not url or not key: sys.exit("Задайте SUPABASE_URL и SUPABASE_SERVICE_KEY")
    import requests
    H = {"apikey": key, "Authorization": f"Bearer {key}", "Content-Type": "application/json"}
    api = lambda t: f"{url.rstrip('/')}/rest/v1/{t}"
    def get(t, q=""):
        r = requests.get(api(t) + "?" + q, headers=H); r.raise_for_status(); return r.json()
    def post(t, data):
        for i in range(0, len(data), BATCH):
            r = requests.post(api(t), headers={**H, "Prefer": "return=minimal"}, json=data[i:i + BATCH])
            if not r.ok: sys.exit(f"Ошибка при записи в {t}: {r.status_code} {r.text}")
    if get("children", "select=id&limit=1") and not force: sys.exit("В базе уже есть дети. Остановлено (--force, если так и надо).")

    teachers = {t["name"]: t["id"] for t in get("teachers", "select=id,name")}
    new_t = [{"name": n} for n in sorted({t for t, _ in groups}) if n not in teachers]
    if new_t: post("teachers", new_t); teachers = {t["name"]: t["id"] for t in get("teachers", "select=id,name")}
    ptypes = {p["name"]: p["id"] for p in get("price_types", "select=id,name")}
    gid = {}
    glist = []
    for (t, g) in groups:
        gid[(t, g)] = str(uuid.uuid4())
        s = sched[(t, g)].most_common(1)
        glist.append(dict(id=gid[(t, g)], teacher_id=teachers[t], name=g, schedule=s[0][0] if s else None, price_type_id=ptypes.get(price_type(g) or "")))
    post("groups", glist)
    post("children", children)
    post("enrollments", [dict(child_id=c, group_id=gid[g], start_date=s) for (c, g), s in enr.items()])
    print("Готово: данные загружены.")

if __name__ == "__main__":
    main()
