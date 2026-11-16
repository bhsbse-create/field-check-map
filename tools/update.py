"""현장 체크 지도 데이터 갱신 스크립트.

회사 엑셀(Tenant Profile)에서 한 담당자 몫만 뽑아 암호화한 뒤 data.enc.json 을 새로 만든다.
엑셀 원본과 평문 데이터는 절대 저장소에 올리지 않는다 (.gitignore 참고).

사용 예 (저장소 폴더에서):
    python tools/update.py --excel "C:/어딘가/2026_.xlsx" --password 2189

옵션:
    --name        담당자 이름 (기본: 최서윤)
    --password    사이트 비밀번호 (새 비번으로 바꿀 때는 --old-password 에 예전 비번)
    --old-password  예전 비번. 기존 좌표를 재사용하려고 기존 data.enc.json 을 여는 데 씀 (기본: --password 와 같음)
    --kakao-key   카카오 REST API 키. 있으면 새 건물 좌표를 주소로 정확히 찾음 (없으면 무료 OSM 검색, 대략 위치)
    --regeocode   기존 좌표를 버리고 전부 다시 찾기 (카카오 키와 같이 쓰면 전부 정확해짐)
    --out         출력 파일 (기본: data.enc.json)

필요한 것: pip install openpyxl cryptography
"""
import argparse, base64, collections, json, math, os, re, sys, time, urllib.parse, urllib.request

import openpyxl
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ITER = 250000
OFFICE = dict(name="SC제일은행 본점빌딩 (회사)", addr="서울 종로구 종로 47", lat=37.57101, lon=126.98239)
SUB_CENTER = {"CBD": (37.5665, 126.9780), "GBD": (37.4979, 127.0276), "YBD": (37.5251, 126.9255),
              "Pangyo": (37.4130, 127.0990), "Yongsan": (37.5326, 126.9905), "Magok": (37.5605, 126.8255),
              "Seongsu": (37.5446, 127.0557), "DMC": (37.5794, 126.8895)}


# ---------- 암호화 ----------
def derive(pw, salt, it=ITER):
    return PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=it).derive(pw.encode())


def encrypt(obj, pw):
    raw = json.dumps(obj, ensure_ascii=False, default=str).encode("utf-8")
    salt, iv = os.urandom(16), os.urandom(12)
    ct = AESGCM(derive(pw, salt)).encrypt(iv, raw, None)
    return dict(v=1, iter=ITER, salt=base64.b64encode(salt).decode(), iv=base64.b64encode(iv).decode(),
                ct=base64.b64encode(ct).decode())


def decrypt(enc, pw):
    key = derive(pw, base64.b64decode(enc["salt"]), enc["iter"])
    pt = AESGCM(key).decrypt(base64.b64decode(enc["iv"]), base64.b64decode(enc["ct"]), None)
    return json.loads(pt.decode("utf-8"))


# ---------- 엑셀 읽기 ----------
def norm(h):
    return re.sub(r"\s+", "", str(h or "")).lower()


def find_header(rows, must):
    for i, r in enumerate(rows[:10]):
        hs = [norm(x) for x in r]
        if all(any(m in h for h in hs) for m in must):
            return i, hs
    raise SystemExit(f"헤더를 못 찾았어요: {must}")


def col(hs, *keys):
    for k in keys:
        for i, h in enumerate(hs):
            if k in h:
                return i
    return None


def read_excel(path, name):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sh_list = next(ws for ws in wb.worksheets if ws.title.strip() == "리스트")
    sh_floor = next(ws for ws in wb.worksheets if "층별" in ws.title)

    rows = list(sh_list.iter_rows(values_only=True))
    hi, hs = find_header(rows, ["id", "property", "address", "담당자"])
    c = dict(id=col(hs, "id"), sub=col(hs, "submarket"), name=col(hs, "property"), addr=col(hs, "address"),
             who=col(hs, "담당자"), gfa=col(hs, "연면적"), appr=col(hs, "사용승인"), kind=col(hs, "일반/집합"))
    bl = []
    for r in rows[hi + 1:]:
        if r[c["who"]] and name in str(r[c["who"]]) and r[c["id"]] is not None:
            bl.append(dict(id=r[c["id"]], sub=r[c["sub"]], name=str(r[c["name"]]).strip(), addr=str(r[c["addr"]] or "").strip(),
                           gfa=r[c["gfa"]] if c["gfa"] is not None else None,
                           approved=str(r[c["appr"]]) if c["appr"] is not None and r[c["appr"]] else None,
                           kind=r[c["kind"]] if c["kind"] is not None else None))
    ids = {b["id"] for b in bl}

    rows = list(sh_floor.iter_rows(values_only=True))
    hi, hs = find_header(rows, ["id", "buildingname", "tenantname", "floor"])
    f = dict(id=col(hs, "id"), bn=col(hs, "buildingname"), t=col(hs, "tenantname"), p=col(hs, "purpose"),
             o=col(hs, "office/retail"), fl=col(hs, "floor"), a=col(hs, "area(sqm)"), ra=col(hs, "rentablearea(sqm)", "rentable"),
             kor=col(hs, "한국어분류"), ind=col(hs, "industry(kor)"), fd=col(hs, "foreign/domestic"), tl=col(hs, "tenant/landlord"))
    floors = collections.defaultdict(list)
    for r in rows[hi + 1:]:
        if r[f["id"]] in ids:
            g = lambda k: r[f[k]] if f[k] is not None else None
            floors[r[f["id"]]].append(dict(f=g("fl"), t=g("t"), p=g("p"), o=g("o"), a=g("a"), ra=g("ra"),
                                           ind=g("ind") or g("kor"), fd=g("fd"), tl=g("tl"), bn=g("bn")))
    return bl, floors


def floor_key(s):
    s = str(s or "").strip()
    m = re.match(r"^(지하|지|B)\s*(\d+)", s, re.I)
    if m:
        return -int(m.group(2))
    m = re.match(r"^(RF|R/T|옥탑|PH)\s*(\d*)", s, re.I)
    if m:
        return 900 + int(m.group(2) or 0)
    m = re.search(r"(\d+)", s)
    return int(m.group(1)) if m else 500


# ---------- 좌표 ----------
def http_json(url, headers=None):
    req = urllib.request.Request(url, headers=headers or {"User-Agent": "field-check-map/1.0"})
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.load(r)


def kakao(addr, key):
    try:
        d = http_json("https://dapi.kakao.com/v2/local/search/address.json?" + urllib.parse.urlencode({"query": addr}),
                      {"Authorization": "KakaoAK " + key})
        if d.get("documents"):
            x = d["documents"][0]
            return float(x["y"]), float(x["x"]), "exact"
    except Exception as e:
        print("  카카오 오류:", e)
    return None


def nominatim(q, viewbox=None):
    p = {"q": q, "format": "json", "limit": 1, "countrycodes": "kr"}
    if viewbox:
        p.update(viewbox=viewbox, bounded=1)
    try:
        r = http_json("https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode(p))
    except Exception:
        r = []
    time.sleep(1.1)
    return r


def osm(b):
    a = b["addr"].replace("서울시", "서울특별시")
    d = nominatim(a)
    if not d:
        return None
    lat, lon = float(d[0]["lat"]), float(d[0]["lon"])
    nm = re.sub(r"\(.*?\)", " ", b["name"]).strip()
    toks = [t for t in nm.split() if len(t) >= 2]
    vb = f"{lon - 0.02},{lat + 0.02},{lon + 0.02},{lat - 0.02}"
    if toks:
        r = nominatim(" ".join(toks), vb)
        if r and toks[0][:3].lower() in r[0]["display_name"].split(",")[0].replace(" ", "").lower():
            return float(r[0]["lat"]), float(r[0]["lon"]), "exact"
    return lat, lon, "approx"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True)
    ap.add_argument("--password", required=True)
    ap.add_argument("--old-password")
    ap.add_argument("--name", default="최서윤")
    ap.add_argument("--kakao-key")
    ap.add_argument("--regeocode", action="store_true")
    ap.add_argument("--out", default=os.path.join(ROOT, "data.enc.json"))
    a = ap.parse_args()

    bl, floors = read_excel(a.excel, a.name)
    print(f"{a.name} 담당 건물 {len(bl)}개, 층 데이터 {sum(len(v) for v in floors.values())}줄")
    if not bl:
        raise SystemExit("담당 건물이 0개예요. 이름이나 엑셀을 확인하세요.")

    old = {}
    cur = os.path.join(ROOT, "data.enc.json")
    if os.path.exists(cur) and not a.regeocode:
        try:
            prev = decrypt(json.load(open(cur)), a.old_password or a.password)
            old = {b["id"]: b for b in prev["buildings"]}
            print(f"기존 좌표 {len(old)}개 재사용")
        except Exception:
            print("기존 data.enc.json 을 못 열었어요 (비번이 다르면 --old-password). 좌표를 새로 찾아요.")

    out = []
    for b in bl:
        o = old.get(b["id"])
        if o and o.get("lat") and o.get("addr") == b["addr"]:
            lat, lon, prec = o["lat"], o["lon"], o["prec"]
        else:
            g = kakao(b["addr"], a.kakao_key) if a.kakao_key else None
            g = g or osm(b)
            if not g:
                c = SUB_CENTER.get(b["sub"], (37.5665, 126.978))
                g = (c[0], c[1], "approx")
            lat, lon, prec = g
            print(f"  좌표 {'정확' if prec == 'exact' else '대략'}: {b['name']}")
        if "SC제일은행" in b["name"] and "본점" in b["name"]:
            lat, lon, prec = OFFICE["lat"], OFFICE["lon"], "exact"
        fl = sorted(floors.get(b["id"], []), key=lambda x: (-floor_key(x["f"]), str(x["t"] or "")))
        out.append(dict(b, lat=lat, lon=lon, prec=prec, floors=fl,
                        n_tenant=len({x["t"] for x in fl if x["t"]})))

    seen = collections.Counter()
    for b in out:
        if b["prec"] != "exact" and not (old.get(b["id"]) and old[b["id"]].get("lat") == b["lat"]):
            k = (round(b["lat"], 5), round(b["lon"], 5))
            n = seen[k]
            seen[k] += 1
            if n:
                b["lat"] += 0.0006 * math.sqrt(n) * math.sin(n * 2.399)
                b["lon"] += 0.0006 * math.sqrt(n) * math.cos(n * 2.399) * 1.25

    payload = dict(updated=time.strftime("%Y-%m-%d"), office=OFFICE, buildings=out)
    json.dump(encrypt(payload, a.password), open(a.out, "w"))
    c = collections.Counter(b["prec"] for b in out)
    print(f"완료 → {a.out}  (정확 {c['exact']}개 / 대략 {c['approx']}개)")
    print("이제 git add data.enc.json → commit → push 하면 몇 분 뒤 사이트에 반영돼요.")


if __name__ == "__main__":
    main()
