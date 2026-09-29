"""Synthetic SAP extracts for development, training and regression testing.

Entirely fabricated: 'Demo Steel Ltd', its vendors, customers and projects do not
exist. The extracts imitate the layout of real SAP exports (ALV spreadsheet exports
with title lines, an 'unconverted' | separated MB51 list, a tab separated FBL3N,
trailing-minus amounts with Indian digit grouping, leading zeros, subtotal lines)
so that the ingestion layer is exercised the way real data will exercise it.

Background ('noise') transactions are generated so that they do NOT trigger any test.
On top of them, every test has planted exceptions and planted near-misses; PLANTED
lists what each test must find, and tests/test_library.py checks exactly that.
"""

from __future__ import annotations

import datetime as dt
import random
from pathlib import Path

from openpyxl import Workbook

D = dt.date.fromisoformat
PERIOD_FROM, PERIOD_TO = D("2025-04-01"), D("2026-03-31")
HOLIDAYS = [(D("2025-08-15"), "Independence Day"), (D("2025-10-02"), "Gandhi Jayanti"),
            (D("2026-01-26"), "Republic Day")]
HOLIDAY_DATES = {d for d, _ in HOLIDAYS}

# What each test must flag in the synthetic data (and nothing else).
# Keys are test IDs; values are (result column, expected set of values).
PLANTED = {
    "PRC-SPLIT-01": ("po_number", {"4500009001", "4500009002", "4500009003"}),
    "PRC-NOPR-01": ("exception_id", {"4500009010/10"}),
    "PRC-POAFTGR-01": ("exception_id", {"4500009020"}),
    "PRC-OPEN-01": ("exception_id", {"4500009030/10"}),
    "PRC-PRICEVAR-01": ("exception_id", {"4500009043/10"}),
    "PRC-NEWVEN-01": ("exception_id", {"4500009050", "4500009051"}),
    "PRC-GRIR-01": ("exception_id", {"4500009060/10"}),
    "PRC-VENCON-01": ("exception_id", {"MG-REFR"}),
    "PRC-LATE-01": ("exception_id", {"4500009070/10"}),
    "INV-NEG-01": ("material", {"10000090"}),
    "INV-NONMOV-01": ("exception_id", {"10000091|1000"}),
    "INV-MANADJ-01": ("exception_id", {"4900000001/1", "4900000002/1"}),
    "INV-GRNI-01": ("exception_id", {"4500009090/10"}),
    "SAL-PRICE-01": ("exception_id", {"9000009001/10"}),
    "SAL-CN-01": ("exception_id", {"9000009101", "9000009102"}),
    "SAL-AR-01": ("doc_number", {"1800009001"}),
    "SAL-BLOCKED-01": ("exception_id", {"9000009201"}),
    "SAL-RETURN-01": ("exception_id", {"7100000001/10"}),
    "SAL-RETURN-02": ("exception_id", {"7100000003/10", "7100000004/10", "7100000005/10", "7100000006/10",
                                       "7100000007/10"}),
    "PRJ-COST-01": ("exception_id", {"P-1001", "P-1003"}),
    "PRJ-TIME-01": ("exception_id", {"P-1001", "P-1004"}),
    "PRJ-POSTCLOSE-01": ("exception_id", {"P-1005"}),
    "PRJ-IDLE-01": ("exception_id", {"P-1006"}),
    "PRJ-CWIP-01": ("exception_id", {"400001-0", "400003-0"}),
    "PRJ-TECOCWIP-01": ("exception_id", {"400004-0"}),
    "FIN-DUPINV-01": ("doc_number", {"5100009001", "5100009002"}),
    "FIN-DUPINV-02": ("second_document", {"5100009006"}),
    "FIN-JELATE-01": ("doc_number", {"100009001"}),
    "FIN-MANJE-01": ("doc_number", {"100009010", "100009040"}),
    "FIN-SENSGL-01": ("doc_number", {"100009020"}),
    "FIN-REV-01": ("doc_number", {"100009030"}),
    "FIN-ROUND-01": ("doc_number", {"100009010", "100009040"}),
    "FIN-HOLIDAY-01": ("doc_number", {"100009050", "100009051"}),
    "FIN-ADV-01": ("doc_number", {"1700009001"}),
    "MSE-PAY45-01": ("doc_number", {"5100009101", "5100009102"}),
    "MSE-SHARE-01": ("exception_id", {"MSE-SHARE"}),
}


# --------------------------------------------------------------------------- formatting

def sap_date(d):
    return d.strftime("%d.%m.%Y") if d else ""


def indian(amount: float) -> str:
    """1,23,45,678.90 with SAP's trailing minus for negatives."""
    neg = amount < 0
    whole, frac = f"{abs(amount):.2f}".split(".")
    if len(whole) > 3:
        head, tail = whole[:-3], whole[-3:]
        groups = []
        while len(head) > 2:
            groups.insert(0, head[-2:])
            head = head[:-2]
        if head:
            groups.insert(0, head)
        whole = ",".join(groups + [tail])
    return f"{whole}.{frac}" + ("-" if neg else "")


def matnr(n) -> str:
    return str(n).zfill(18)


def lifnr(n) -> str:
    return str(n).zfill(10)


def weekday_on_or_after(d: dt.date) -> dt.date:
    while d.weekday() == 6 or d in HOLIDAY_DATES:
        d += dt.timedelta(days=1)
    return d


def write_xlsx(path: Path, header: list, rows: list, title_lines=()):
    wb = Workbook()
    ws = wb.active
    for t in title_lines:
        ws.append([t])
    if title_lines:
        ws.append([])
    ws.append(header)
    for r in rows:
        ws.append(r)
    wb.save(path)


# --------------------------------------------------------------------------- generator

def generate(target: Path, seed: int = 7) -> Path:
    rng = random.Random(seed)
    target = Path(target)
    ext = target / "extracts"
    crit = target / "criteria"
    ext.mkdir(parents=True, exist_ok=True)
    crit.mkdir(parents=True, exist_ok=True)

    def rand_date(lo: dt.date, hi: dt.date) -> dt.date:
        return lo + dt.timedelta(days=rng.randint(0, (hi - lo).days))

    # ---------------- vendors
    vendors = []  # (lifnr, name, created, posting_block, purch_block, deletion, msme)
    for v in range(100001, 100021):
        vendors.append((v, f"Vendor {v} Pvt Ltd", rand_date(D("2012-01-01"), D("2018-12-31")), "", "", "", ""))
    for v in range(100021, 100031):
        vendors.append((v, f"MSE Supplier {v}", D("2016-06-01"), "", "", "", "1" if v <= 100025 else "2"))
    vendors.append((100031, "Medium Engg Works", D("2016-06-01"), "", "", "", "3"))
    vendors.append((100035, "Newly Registered Traders", D("2025-09-01"), "", "", "", ""))
    vendors.append((100036, "Blocked Supplies Co", D("2014-03-01"), "X", "X", "", ""))
    vendors.append((100037, "Recent Vendor Ltd", D("2025-01-01"), "", "", "", ""))

    # ---------------- purchase orders, goods receipts, invoices (noise)
    po_rows, md_rows, grir_rows, vi_rows = [], [], [], []
    materials = {10000001 + i: (f"MG-0{i % 5 + 1}", rng.uniform(5000, 50000)) for i in range(30)}
    po_no, md_no, inv_no, pay_no = 4500000001, 5000000001, 5100000001, 1500000001

    def add_po(po, item, vendor, material, mgroup, plant, pgroup, po_date, qty, price, *, created=None,
               pr="auto", doc_type="NB", delivery=None, open_qty=0.0, open_value=0.0, deletion="",
               completed="", text=None):
        po_rows.append([
            str(po), item, doc_type, str(vendor), text or f"Material {material}", matnr(material), mgroup,
            plant, pgroup, dt.datetime.combine(po_date, dt.time()), sap_date(created or po_date),
            qty, "EA", round(price, 2), 1, round(qty * price, 2), "INR",
            ("" if pr is None else (str(10000000 + po % 100000) if pr == "auto" else pr)),
            sap_date(delivery or po_date + dt.timedelta(days=30)), open_qty, open_value, deletion, completed,
        ])

    def add_md(movement, material, plant, sloc, posting, qty, amount, po=None, po_item=None, vendor=None,
               doc=None, user="STORE01"):
        nonlocal md_no
        number = doc or md_no
        if doc is None:
            md_no += 1
        md_rows.append([str(number), "0001", movement, sap_date(posting), matnr(material), plant, sloc,
                        f"{abs(qty):.3f}" + ("-" if qty < 0 else ""), "EA", indian(amount),
                        str(po) if po else "", str(po_item) if po_item else "", lifnr(vendor) if vendor else "",
                        user, sap_date(posting)])

    def add_invoice(vendor, amount, posting, doc_date, reference, cleared_on=None, doc=None, doc_type="RE"):
        nonlocal inv_no, pay_no
        number = doc or inv_no
        if doc is None:
            inv_no += 1
        clearing_doc = ""
        if cleared_on:
            clearing_doc = str(pay_no)
            pay_no += 1
            vi_rows.append(["1000", lifnr(vendor), clearing_doc, "2025", "KZ", sap_date(cleared_on), sap_date(cleared_on),
                            "", indian(amount), clearing_doc, sap_date(cleared_on), "", "Payment"])
        vi_rows.append(["1000", lifnr(vendor), str(number), "2025", doc_type, sap_date(posting), sap_date(doc_date),
                        reference, indian(-amount), clearing_doc, sap_date(cleared_on) if cleared_on else "", "", ""])

    noise_vendors = list(range(100001, 100021)) * 12 + list(range(100021, 100031)) * 2
    rng.shuffle(noise_vendors)
    for i, vendor in enumerate(noise_vendors):
        material = 10000001 + rng.randrange(30)
        mgroup, base = materials[material]
        price = base * rng.uniform(0.95, 1.05)
        qty = int(rng.uniform(1_200_000, 6_000_000) / price) + 1
        po_date = rand_date(PERIOD_FROM, D("2026-02-28"))
        plant = rng.choice(["1000", "2000"])
        add_po(po_no, 10, vendor, material, mgroup, plant, rng.choice(["P01", "P02", "P03"]), po_date, qty, price)
        gr_date = po_date + dt.timedelta(days=rng.randint(5, 25))
        value = round(qty * price, 2)
        add_md("101", material, plant, "0001", gr_date, qty, value, po=po_no, po_item=10, vendor=vendor)
        grir_rows.append([str(po_no), 10, lifnr(vendor), matnr(material), plant, qty, qty, value, value])
        posting = gr_date + dt.timedelta(days=2)
        cleared = posting + dt.timedelta(days=rng.randint(10, 30))
        add_invoice(vendor, value, posting, gr_date, f"INV/{vendor % 1000}/{i + 1:04d}",
                    cleared_on=cleared if cleared <= PERIOD_TO else None)
        po_no += 1

    # ---------------- planted procurement scenarios
    P = add_po
    # split POs: 3 POs within 18 days, each < 10 lakh, together 10.5 lakh
    P(4500009001, 10, 100005, 10000050, "MG-SPL", "1000", "P01", D("2025-06-02"), 40, 10000)
    P(4500009002, 10, 100005, 10000050, "MG-SPL", "1000", "P01", D("2025-06-10"), 35, 10000)
    P(4500009003, 10, 100005, 10000050, "MG-SPL", "1000", "P01", D("2025-06-20"), 30, 10000)
    # near-miss: together only 8 lakh
    P(4500009004, 10, 100006, 10000060, "MG-SPL", "1000", "P01", D("2025-07-01"), 40, 10000)
    P(4500009005, 10, 100006, 10000060, "MG-SPL", "1000", "P01", D("2025-07-15"), 40, 10000)
    # near-miss: 45 days apart
    P(4500009006, 10, 100007, 10000070, "MG-SPL", "1000", "P02", D("2025-08-01"), 40, 10000)
    P(4500009007, 10, 100007, 10000070, "MG-SPL", "1000", "P02", D("2025-09-15"), 40, 10000)
    P(4500009008, 10, 100007, 10000070, "MG-SPL", "1000", "P02", D("2025-10-30"), 40, 10000)
    # no PR; near-misses: excluded doc type, below minimum value
    P(4500009010, 10, 100008, 10000100, "MG-NPR", "2000", "P03", D("2025-07-07"), 100, 20000, pr=None)
    P(4500009011, 10, 100008, 10000101, "MG-NPR", "2000", "P03", D("2025-07-08"), 100, 15000, pr=None, doc_type="ZRC")
    P(4500009012, 10, 100008, 10000102, "MG-NPR", "2000", "P03", D("2025-07-09"), 10, 5000, pr=None)
    # PO created after GR (document date back-dated to before the GR)
    P(4500009020, 10, 100009, 10000103, "MG-GR", "1000", "P01", D("2025-08-01"), 50, 10000,
      created=D("2025-08-20"), delivery=D("2025-08-30"))
    add_md("101", 10000103, "1000", "0001", D("2025-08-05"), 50, 500000, po=4500009020, po_item=10, vendor=100009)
    # open POs
    P(4500009030, 10, 100015, 10000104, "MG-OPN", "1000", "P02", D("2025-06-01"), 80, 10000,
      delivery=D("2025-07-15"), open_qty=80, open_value=800000)
    P(4500009031, 10, 100015, 10000105, "MG-OPN", "1000", "P02", D("2025-11-01"), 60, 10000,
      delivery=D("2025-12-15"), open_qty=60, open_value=600000)
    P(4500009032, 10, 100015, 10000106, "MG-OPN", "1000", "P02", D("2025-04-10"), 30, 10000,
      delivery=D("2025-05-10"), open_qty=30, open_value=300000, completed="X")
    # price variation: median 1025, one PO at 1600
    for po, vendor, price in [(4500009040, 100016, 1000), (4500009041, 100017, 1050),
                              (4500009042, 100018, 980), (4500009043, 100019, 1600)]:
        P(po, 10, vendor, 10000080, "MG-PRV", "2000", "P03", D("2025-09-01") + dt.timedelta(days=po % 10 * 20),
          100, price)
    # new / blocked vendors; near-miss: vendor created 9 months before the PO
    P(4500009050, 10, 100035, 10000107, "MG-NV", "1000", "P01", D("2025-09-10"), 60, 10000)
    P(4500009051, 10, 100036, 10000108, "MG-NV", "1000", "P01", D("2025-10-05"), 60, 10000)
    P(4500009052, 10, 100037, 10000109, "MG-NV", "1000", "P01", D("2025-10-01"), 60, 10000)
    # GR/IR: invoiced more than received
    grir_rows.append(["4500009060", 10, lifnr(100015), matnr(10000110), "1000", 100, 120, 1000000, 1200000])
    # delayed delivery; near-miss within the 30-day grace
    P(4500009070, 10, 100016, 10000111, "MG-LT", "1000", "P02", D("2025-06-01"), 70, 10000, delivery=D("2025-07-01"))
    add_md("101", 10000111, "1000", "0001", D("2025-09-15"), 70, 700000, po=4500009070, po_item=10, vendor=100016)
    P(4500009071, 10, 100016, 10000112, "MG-LT", "1000", "P02", D("2025-06-01"), 70, 10000, delivery=D("2025-07-01"))
    add_md("101", 10000112, "1000", "0001", D("2025-07-20"), 70, 700000, po=4500009071, po_item=10, vendor=100016)
    # vendor concentration: one vendor holds 90% of MG-REFR
    for po, vendor, qty in [(4500009080, 100010, 300), (4500009081, 100010, 300),
                            (4500009082, 100010, 300), (4500009083, 100011, 100)]:
        P(po, 10, vendor, 10000120, "MG-REFR", "1000", "P01", D("2025-05-05") + dt.timedelta(days=po % 10 * 60),
          qty, 10000)
    # GR not invoiced: old (flag) and recent (near-miss)
    P(4500009090, 10, 100017, 10000121, "MG-GRN", "1000", "P01", D("2025-06-01"), 50, 8000, delivery=D("2025-06-30"))
    add_md("101", 10000121, "1000", "0001", D("2025-06-15"), 50, 400000, po=4500009090, po_item=10, vendor=100017)
    grir_rows.append(["4500009090", 10, lifnr(100017), matnr(10000121), "1000", 50, 0, 400000, 0])
    P(4500009091, 10, 100017, 10000122, "MG-GRN", "1000", "P01", D("2026-02-01"), 20, 8000, delivery=D("2026-03-01"))
    add_md("101", 10000122, "1000", "0001", D("2026-02-20"), 20, 160000, po=4500009091, po_item=10, vendor=100017)
    grir_rows.append(["4500009091", 10, lifnr(100017), matnr(10000122), "1000", 20, 0, 160000, 0])

    # ---------------- stock and issues
    stock_rows = []
    for material in range(10000001, 10000031):
        stock_rows.append([matnr(material), f"Material {material}", "1000", "0001", 100, "EA", 500000.0])
        add_md("261", material, "1000", "0001", rand_date(D("2026-02-01"), D("2026-03-25")), -5, -25000)
    stock_rows.append([matnr(10000090), "Negative stock item", "1000", "0002", -5, "EA", -25000.0])
    stock_rows.append([matnr(10000091), "Idle spare (no movement)", "1000", "0001", 50, "EA", 2500000.0])
    stock_rows.append([matnr(10000092), "Low value idle item", "1000", "0001", 5, "EA", 50000.0])
    stock_rows.append([matnr(10000094), "Recently issued item", "1000", "0001", 20, "EA", 800000.0])
    add_md("261", 10000094, "1000", "0001", D("2025-04-05"), -2, -80000)
    add_md("701", 10000093, "1000", "0001", D("2025-11-20"), -10, -350000, doc=4900000001, user="STORE02")
    add_md("551", 10000093, "1000", "0001", D("2025-12-05"), -4, -200000, doc=4900000002, user="STORE02")
    add_md("701", 10000093, "1000", "0001", D("2026-01-10"), -1, -20000, doc=4900000003, user="STORE02")

    # ---------------- planted vendor-ledger scenarios
    add_invoice(100012, 236000, D("2025-07-10"), D("2025-07-08"), "INV/778", D("2025-08-01"), doc=5100009001, doc_type="KR")
    add_invoice(100012, 236000, D("2025-08-05"), D("2025-07-08"), "INV-778", D("2025-08-30"), doc=5100009002, doc_type="KR")
    add_invoice(100013, 100000, D("2025-09-01"), D("2025-08-30"), "B-55", D("2025-09-20"), doc=5100009003, doc_type="KR")
    add_invoice(100013, 120000, D("2025-09-03"), D("2025-08-30"), "B-55", D("2025-09-22"), doc=5100009004, doc_type="KR")
    add_invoice(100014, 118000, D("2025-11-04"), D("2025-11-03"), "A-100", D("2025-11-25"), doc=5100009005, doc_type="KR")
    add_invoice(100014, 118000, D("2025-11-06"), D("2025-11-05"), "B-771", D("2025-11-28"), doc=5100009006, doc_type="KR")
    add_invoice(100014, 118000, D("2026-01-12"), D("2026-01-10"), "C-900", D("2026-02-02"), doc=5100009007, doc_type="KR")
    # MSE payments: micro paid in 106 days, small unpaid 80 days, medium (not covered), micro paid in 39 days
    add_invoice(100021, 300000, D("2025-07-01"), D("2025-06-28"), "MSE/11", D("2025-10-15"), doc=5100009101, doc_type="KR")
    add_invoice(100026, 200000, D("2026-01-10"), D("2026-01-08"), "MSE/12", None, doc=5100009102)
    add_invoice(100031, 400000, D("2025-06-01"), D("2025-05-30"), "MED/1", D("2025-10-01"), doc=5100009103, doc_type="KR")
    add_invoice(100022, 150000, D("2025-09-01"), D("2025-08-30"), "MSE/13", D("2025-10-10"), doc=5100009104, doc_type="KR")
    # advances (special G/L A): old open, recent open, old cleared
    vi_rows.append(["1000", lifnr(100015), "1700009001", "2025", "KA", "20.04.2025", "20.04.2025", "ADV-1",
                    indian(500000), "", "", "A", "Mobilisation advance"])
    vi_rows.append(["1000", lifnr(100015), "1700009002", "2025", "KA", "15.01.2026", "15.01.2026", "ADV-2",
                    indian(300000), "", "", "A", "Advance"])
    vi_rows.append(["1000", lifnr(100015), "1700009003", "2025", "KA", "01.05.2025", "01.05.2025", "ADV-3",
                    indian(200000), "1500999001", "01.08.2025", "A", "Advance adjusted"])

    # ---------------- G/L documents
    gl_rows = []
    gl_no = 100000001

    def add_gl(doc, doc_type, tcode, posting, entry, amount, dr_gl, cr_gl, user="FIUSER1", reversed_with="",
               header=""):
        period = (posting.month - 4) % 12 + 1
        for line, (gl, amt) in enumerate([(dr_gl, amount), (cr_gl, -amount)], 1):
            gl_rows.append(["1000", gl, str(doc), str(line).zfill(3), "2025", period, doc_type, sap_date(posting),
                            sap_date(posting), sap_date(entry), indian(amt), user, tcode, "",
                            reversed_with, header])

    auto_types = [("RE", "MIRO"), ("KR", "FB60"), ("WE", "MIGO"), ("KZ", "F110"), ("DZ", "F-28"), ("RV", "VF01")]
    accounts = ["400100", "400200", "210000", "110000", "300000"]
    for _ in range(200):
        doc_type, tcode = rng.choice(auto_types)
        posting = weekday_on_or_after(rand_date(PERIOD_FROM, D("2026-03-25")))
        entry = weekday_on_or_after(posting + dt.timedelta(days=rng.randint(0, 2)))
        dr, cr = rng.sample(accounts, 2)
        add_gl(gl_no, doc_type, tcode, posting, entry, round(rng.uniform(10000, 5000000), 2) + 0.37, dr, cr,
               user=rng.choice(["MMUSER1", "APUSER1", "ARUSER1"]))
        gl_no += 1
    for _ in range(20):
        posting = weekday_on_or_after(rand_date(PERIOD_FROM, D("2026-03-25")))
        add_gl(gl_no, "SA", "FB50", posting, posting, round(rng.uniform(10000, 400000), 2) + 0.11,
               "400200", "110000")
        gl_no += 1
    add_gl(100009001, "SA", "FB50", D("2025-06-30"), D("2025-08-20"), 300000.55, "400100", "210000")
    add_gl(100009002, "SA", "FB50", D("2026-03-31"), D("2026-04-25"), 200000.10, "400100", "210000")
    add_gl(100009010, "SA", "FB50", D("2025-09-15"), D("2025-09-15"), 5000000.00, "400200", "210000",
           header="Provision")
    add_gl(100009011, "SA", "FB50", D("2025-09-16"), D("2025-09-16"), 200000.00, "400200", "210000")
    add_gl(100009020, "SA", "FB50", D("2025-10-20"), D("2025-10-20"), 750000.25, "450000", "110000",
           header="Write back")
    add_gl(100009030, "SA", "FB50", D("2025-12-10"), D("2025-12-10"), 600000.40, "400100", "210000",
           reversed_with="100009031")
    add_gl(100009031, "AB", "FB08", D("2025-12-20"), D("2025-12-20"), -600000.40, "400100", "210000",
           reversed_with="100009030")
    add_gl(100009040, "SA", "FB01", D("2026-02-10"), D("2026-02-10"), 2500000.00, "400100", "210000")
    add_gl(100009050, "SA", "FB50", D("2025-10-10"), D("2025-10-12"), 150000.60, "400100", "210000")
    add_gl(100009051, "SA", "FB50", D("2025-10-01"), D("2025-10-02"), 80000.30, "400100", "210000")

    # ---------------- sales
    customers = [[str(c), f"Customer {c}", "01.04.2015", "", "", "", ""] for c in range(300001, 300011)]
    customers[8] = ["300009", "Blocked Customer", "01.04.2015", "X", "01", "", ""]
    billing, fbl5n = [], []
    prices = {20000001: 50000, 20000002: 42000}
    inv_dates, inv_info = {}, {}
    fixed_dates = {5: "2025-05-10", 6: "2025-06-01", 7: "2025-06-01", 8: "2025-07-01", 9: "2025-08-01",
                   11: "2025-09-01", 12: "2025-10-01"}
    for i in range(1, 61):
        doc = 9000000000 + i
        material = rng.choice(list(prices))
        date = rand_date(PERIOD_FROM, PERIOD_TO)
        if i in fixed_dates:
            date = D(fixed_dates[i])
        inv_dates[doc] = date
        qty = rng.randint(50, 200)
        inv_info[doc] = (material, date, qty)
        price = prices[material] * rng.uniform(1.0, 1.05)
        value = round(qty * price, 2)
        customer = 300001 + rng.randrange(8)
        billing.append([str(doc), "000010", "F2", sap_date(date), str(customer), matnr(material), qty, "TO",
                        value, "", "", "1000"])
        due = date + dt.timedelta(days=30)
        cleared = date + dt.timedelta(days=rng.randint(10, 28))
        fbl5n.append(["1000", str(customer), str(1800000000 + i), "RV", sap_date(date), sap_date(due),
                      indian(value), "1400000001" if cleared <= PERIOD_TO else "",
                      sap_date(cleared) if cleared <= PERIOD_TO else "", ""])
    billing.append(["9000009001", "000010", "F2", "12.08.2025", "300002", matnr(20000001), 100, "TO", 4600000.0, "", "", "1000"])
    billing.append(["9000009002", "000010", "F2", "13.08.2025", "300002", matnr(20000001), 100, "TO", 5000000.0, "", "", "1000"])
    billing.append(["9000009101", "000010", "G2", "10.09.2025", "300003", matnr(20000001), 16, "TO", 800000.0, "9000000010", "", "1000"])
    billing.append(["9000009102", "000010", "G2", sap_date(D("2025-05-10") + dt.timedelta(days=150)), "300004",
                    matnr(20000001), 2, "TO", 100000.0, "9000000005", "", "1000"])
    billing.append(["9000009103", "000010", "G2", sap_date(D("2025-06-01") + dt.timedelta(days=10)), "300005",
                    matnr(20000002), 1, "TO", 50000.0, "9000000006", "", "1000"])
    billing.append(["9000009201", "000010", "F2", "05.11.2025", "300009", matnr(20000002), 10, "TO", 420000.0, "", "", "1000"])
    fbl5n.append(["1000", "300001", "1800009001", "RV", "15.09.2025", "15.10.2025", indian(1200000), "", "", ""])
    fbl5n.append(["1000", "300004", "1800009002", "RV", "16.01.2026", "15.02.2026", indian(900000), "", "", ""])
    for row in billing:  # CMO-style region and branch sales office
        customer = int(row[4])
        row += ["ER" if customer % 2 else "WR", f"BSO-{customer % 4 + 1:02d}"]

    # ---------------- sales returns / quality complaints (ZSDR084 style)
    def ret(no, invoice, qty, request, receipt, reason, qc="Accepted", approval="EW/2025/123", cn="CN-1",
            customer="300001", delivery=""):
        material = inv_info[invoice][0] if invoice else 20000001
        return [str(no), "10", customer, str(invoice) if invoice else "", delivery, matnr(material), qty, "TO",
                round(qty * prices[material], 2), sap_date(D(request)), sap_date(D(receipt)) if receipt else "",
                reason, qc, approval, cn, "BSO-01"]

    returns = [
        # late return: 120 days after invoice 9000000007 (01.06.2025)
        ret(7100000001, 9000000007, 10, "2025-09-29", "2025-10-10", "Quality - lamination"),
        # near-miss: clean return within 20 days
        ret(7100000002, 9000000008, 5, "2025-07-21", "2025-07-30", "Quality - surface defect"),
        # no link to invoice or delivery
        ret(7100000003, None, 8, "2025-11-10", "2025-11-20", "Quality - dimension"),
        # returned more than invoiced (invoice 9000000009 has at most 200)
        ret(7100000004, 9000000009, 500, "2025-08-20", "2025-08-28", "Quality - bend"),
        # risk reason: wrong specification
        ret(7100000005, 9000000011, 6, "2025-09-15", "2025-09-25", "Wrong specification supplied"),
        # QC decision not recorded
        ret(7100000006, 9000000012, 4, "2025-10-15", "2025-10-25", "Quality - rust", qc=""),
        # credit note issued, no receipt of material recorded
        ret(7100000007, 9000000011, 3, "2025-09-20", None, "Quality - edge crack", cn="CN-7"),
    ]

    # ---------------- projects
    wbs = [
        ["P-1001", "P-1001", "Blast furnace relining", "REL", "01.04.2022", "31.12.2024", "", "", "1000"],
        ["P-1002", "P-1002", "Coke oven battery upgrade", "REL", "01.04.2024", "31.12.2026", "", "", "1000"],
        ["P-1003", "P-1003", "Sinter plant bag filter", "REL", "01.04.2025", "31.03.2027", "", "", "2000"],
        ["P-1004", "P-1004", "Rail mill reheating furnace", "TECO", "01.04.2023", "31.03.2025", "30.06.2025", "30.06.2025", "1000"],
        ["P-1005", "P-1005", "Oxygen plant compressor", "TECO", "01.04.2024", "30.06.2025", "31.05.2025", "31.05.2025", "2000"],
        ["P-1006", "P-1006", "Raw material handling conveyor", "REL", "01.04.2024", "30.09.2026", "", "", "2000"],
        ["P-1007", "P-1007", "Township water line", "CLSD", "01.04.2018", "31.03.2020", "31.03.2020", "31.03.2020", "1000"],
    ]
    crore = 10_000_000
    budget = [["P-1001", 100 * crore, 125 * crore, 2 * crore], ["P-1002", 50 * crore, 52 * crore, 0],
              ["P-1003", 0, 3 * crore, 0], ["P-1004", 20 * crore, 19 * crore, 0],
              ["P-1005", 10 * crore, 9 * crore, 0], ["P-1006", 8 * crore, 5 * crore, 0],
              ["P-1007", 5 * crore, 5 * crore, 0]]
    cji3 = []
    co_doc = 7000000001

    def add_cost(wbs_id, date, amount, text="Contractor bill"):
        nonlocal co_doc
        cji3.append([f"WBS {wbs_id}", str(co_doc), sap_date(date), "5100100", round(amount, 2), text])
        co_doc += 1

    for m in range(12):
        month = D("2025-04-15") + dt.timedelta(days=30 * m)
        for wbs_id in ("P-1001", "P-1002", "P-1003"):
            add_cost(wbs_id, month, rng.uniform(500000, 3000000))
    for m in range(3):
        add_cost("P-1004", D("2025-04-10") + dt.timedelta(days=30 * m), 1500000)
        add_cost("P-1006", D("2025-04-10") + dt.timedelta(days=15 * m), 2000000)
    add_cost("P-1005", D("2025-04-20"), 1000000)
    cji3.append(["WBS P-1005", "7000009001", "10.09.2025", "5100100", 400000.0, "Additional civil work"])
    assets = [
        ["400001", "0", "4000", "CWIP: coal handling plant", "01.05.2019", 30 * crore, 30 * crore, ""],
        ["400002", "0", "4000", "CWIP: new weighbridge", "01.01.2025", 5 * crore, 5 * crore, ""],
        ["400003", "0", "4000", "CWIP: effluent treatment", "01.02.2023", 2 * crore, 2 * crore, ""],
        ["400004", "0", "4000", "CWIP: oxygen compressor", "01.10.2024", 1.5 * crore, 1.5 * crore, "P-1005"],
        ["400005", "0", "4000", "CWIP: BF relining", "01.06.2025", 10 * crore, 10 * crore, "P-1001"],
        ["100001", "0", "1000", "Office building", "01.04.2010", 40 * crore, 22 * crore, ""],
    ]

    # ---------------- write the extracts
    write_xlsx(ext / "ME2N.xlsx",
               ["Purchasing Document", "Item", "Purch. Doc. Type", "Supplier/Supplying Plant", "Short Text",
                "Material", "Material Group", "Plant", "Purchasing Group", "Document Date", "Created On",
                "Order Quantity", "Order Unit", "Net Price", "Price Unit", "Net Order Value", "Currency",
                "Purchase Requisition", "Delivery Date", "Still to be delivered (qty)",
                "Still to be delivered (value)", "Deletion Indicator", "Delivery Completed"],
               po_rows, title_lines=["Purchasing Documents per Supplier", "Demo Steel Ltd (synthetic data)"])

    with open(ext / "MB51.txt", "w", encoding="utf-8") as fh:
        fh.write("Material Document List\n\n")
        head = ["Material Document", "Item", "Movement Type", "Posting Date", "Material", "Plant",
                "Storage Location", "Quantity", "Unit of Entry", "Amount in LC", "Purchase Order",
                "Purchase Order Item", "Supplier", "User Name", "Entry Date"]
        rule = "-" * 180
        fh.write(rule + "\n|" + "|".join(head) + "|\n" + rule + "\n")
        for r in md_rows:
            fh.write("|" + "|".join(str(c) for c in r) + "|\n")
        fh.write(rule + "\n")

    write_xlsx(ext / "MB52.xlsx", ["Material", "Material Description", "Plant", "Storage Location",
                                   "Unrestricted", "Base Unit of Measure", "Value Unrestricted"],
               stock_rows, title_lines=["Display Warehouse Stocks of Material"])
    write_xlsx(ext / "MB5S.xlsx", ["Purchasing Document", "Item", "Supplier", "Material", "Plant",
                                   "Quantity received", "Quantity invoiced", "Value Received", "Value Invoiced"],
               grir_rows)
    write_xlsx(ext / "LFA1.xlsx", ["LIFNR", "NAME1", "ERDAT", "SPERR", "SPERM", "LOEVM", "J_1ISSIST"],
               [[lifnr(v), n, int(c.strftime("%Y%m%d")), pb, pu, de, ms] for v, n, c, pb, pu, de, ms in vendors])

    fbl1n_rows = list(vi_rows)
    fbl1n_rows.insert(len(fbl1n_rows) // 2, ["*", "", "", "", "", "", "", "", indian(-12345678.9), "", "", "", ""])
    fbl1n_rows.append(["", "", "", "", "", "", "", "", indian(-98765432.1), "", "", "", ""])
    write_xlsx(ext / "FBL1N.xlsx", ["Company Code", "Account", "Document Number", "Fiscal Year", "Document Type",
                                    "Posting Date", "Document Date", "Reference", "Amount in local currency",
                                    "Clearing Document", "Clearing Date", "Special G/L ind.", "Text"],
               fbl1n_rows, title_lines=["Vendor Line Item Display", "Demo Steel Ltd (synthetic data)"])

    with open(ext / "FBL3N.tsv", "w", encoding="utf-8") as fh:
        fh.write("\t".join(["Company Code", "G/L Account", "Document Number", "Line item", "Fiscal Year",
                            "Posting period", "Document Type", "Posting Date", "Document Date", "Entry Date",
                            "Amount in local currency", "User Name", "Transaction Code", "Text", "Reversed with",
                            "Doc.Header Text"]) + "\n")
        for r in gl_rows:
            fh.write("\t".join(str(c) for c in r) + "\n")

    write_xlsx(ext / "FBL5N.xlsx", ["Company Code", "Customer", "Document Number", "Document Type", "Posting Date",
                                    "Net due date", "Amount in local currency", "Clearing Document",
                                    "Clearing Date", "Special G/L ind."], fbl5n)
    write_xlsx(ext / "KNA1.xlsx", ["KUNNR", "NAME1", "ERDAT", "SPERR", "AUFSD", "FAKSD", "LOEVM"], customers)
    write_xlsx(ext / "VF05N.xlsx", ["Billing Document", "Item", "Billing Type", "Billing Date", "Sold-To Party",
                                    "Material", "Billed Quantity", "Sales Unit", "Net Value", "Reference Document",
                                    "Cancelled", "Plant", "Region", "Sales Office"], billing)
    write_xlsx(ext / "ZSDR084.xlsx", ["Complaint No", "Return Item", "Customer", "Invoice No", "Delivery No",
                                      "Material", "Return Qty", "UoM", "Return Value", "Complaint Date",
                                      "Receipt Date", "Reason", "QC Status", "Approval Ref", "Credit Note", "Branch"],
               returns, title_lines=["Quality Complaint / Return Report (synthetic)"])
    write_xlsx(ext / "CN43N.xlsx", ["WBS Element", "Project Definition", "Description", "System Status",
                                    "Basic start date", "Basic finish date", "Actual finish", "TECO date", "Plant"],
               wbs)
    write_xlsx(ext / "S_ALR_87013558.xlsx", ["WBS Element", "Budget", "Actual", "Commitment"], budget)
    write_xlsx(ext / "CJI3.xlsx", ["Object", "Document Number", "Posting Date", "Cost Element", "Val/COArea Crcy",
                                   "Name"], cji3)
    write_xlsx(ext / "AR01.xlsx", ["Asset", "Subnumber", "Asset Class", "Asset description", "Capitalized on",
                                   "Acquis.val.", "Book val.", "WBS Element"], assets)
    write_xlsx(crit / "approved_prices.xlsx", ["Material", "Unit", "Valid From", "Valid To", "Approved Price",
                                               "Authority"],
               [[matnr(m), "TO", "01.04.2025", "31.03.2026", p, "Price circular 7/2025 (synthetic)"]
                for m, p in prices.items()])
    write_xlsx(crit / "holidays.xlsx", ["Date", "Description"], [[sap_date(d), n] for d, n in HOLIDAYS])

    sources = {
        "po_items": ("extracts/ME2N.xlsx", "ME2N (selection WE101)", len(po_rows), sum(r[15] for r in po_rows)),
        "material_docs": ("extracts/MB51.txt", "MB51", len(md_rows), sum(_num(r[9]) for r in md_rows)),
        "stock": ("extracts/MB52.xlsx", "MB52 as on 31.03.2026", len(stock_rows), sum(r[6] for r in stock_rows)),
        "grir": ("extracts/MB5S.xlsx", "MB5S", len(grir_rows), sum(r[7] for r in grir_rows)),
        "vendor_master": ("extracts/LFA1.xlsx", "SE16N LFA1", len(vendors), None),
        "vendor_items": ("extracts/FBL1N.xlsx", "FBL1N all items", len(vi_rows), sum(_num(r[8]) for r in vi_rows)),
        "gl_items": ("extracts/FBL3N.tsv", "FBL3N", len(gl_rows), None),
        "customer_items": ("extracts/FBL5N.xlsx", "FBL5N all items", len(fbl5n), sum(_num(r[6]) for r in fbl5n)),
        "customer_master": ("extracts/KNA1.xlsx", "SE16N KNA1", len(customers), None),
        "billing_items": ("extracts/VF05N.xlsx", "VF05N", len(billing), sum(r[8] for r in billing)),
        "sales_returns": ("extracts/ZSDR084.xlsx", "ZSDR084", len(returns), sum(r[6] for r in returns)),
        "approved_prices": ("criteria/approved_prices.xlsx", "Price circular keyed in by audit", None, None),
        "wbs_master": ("extracts/CN43N.xlsx", "CN43N", len(wbs), None),
        "wbs_budget": ("extracts/S_ALR_87013558.xlsx", "S_ALR_87013558", len(budget), sum(r[2] for r in budget)),
        "wbs_costs": ("extracts/CJI3.xlsx", "CJI3", len(cji3), sum(r[4] for r in cji3)),
        "assets": ("extracts/AR01.xlsx", "AR01 as on 31.03.2026", len(assets), sum(r[6] for r in assets)),
        "holidays": ("criteria/holidays.xlsx", "Holiday circular keyed in by audit", None, None),
    }
    profile = target / "demo_profile.toml"
    profile.write_text(demo_profile_text(sources), encoding="utf-8")
    return profile


DEMO_HEADER = """\
# Demo run profile for the synthetic 'Demo Steel Ltd' extracts.
# The demo lists every topic that has a library test (more than the sixteen a real plan lists)
# plus two on-site topics, so that every test runs. The SAP control figures (control_rows,
# control_total) are the counts and totals the generator wrote, standing in for the figures an
# auditor would copy from the SAP report.

[audit]
company = "Demo Steel Ltd (synthetic)"
unit = "Plant 1000 and 2000"
tier = "A"
lead_team = "Demo team"
period_from = 2025-04-01
period_to = 2026-03-31
cutoff_date = 2026-03-31

[sap]
decimal_notation = "1,234,567.89"
date_format = "DD.MM.YYYY"
"""

DEMO_FOOTER = """
[annexure_ii]
topics = [
  "T1A01", "T1A02", "T1A03", "T1A04", "T1A06", "T1A07", "T1A08", "T1A10", "T1B01", "T1B03",
  "T2A01", "T2A02", "T2A05", "T2A07",
  "T3A01", "T3A04", "T3A05", "T3A07", "T3A08",
  "T4A02", "T4A03", "T4A04", "T4A05", "T4A07", "T4A08",
  "T7A01", "T7A02", "T7A04", "T7A05", "T7A08", "T7B01", "T7B02", "T7B05",
  "T9A01", "TCS01", "TCS02",
]

[annexure_ii.topic.T1A01]
reason_codes = "C, D"
recommend = "Y"
[annexure_ii.topic.T1A07]
reason_codes = "C"
recommend = "N"

[params."PRC-NOPR-01"]
exclude_doc_types = ["ZRC"]

[params."FIN-SENSGL-01"]
sensitive_gls = ["450000"]

[params."SAL-RETURN-01"]
max_return_days = 60             # demo value; a real audit takes it from the approved return policy
"""


def self_test() -> tuple[bool, list[str]]:
    """Runs every library test on freshly generated synthetic data and checks that each finds
    exactly its planted exceptions (no more, no fewer). Used by the window's Self-test button."""
    import tempfile

    from .engine import EXCEPTIONS, run
    from .loader import load_library
    from .profile import load_profile

    lib = load_library()
    with tempfile.TemporaryDirectory(prefix="auditwb_selftest_") as tmp:
        result = run(load_profile(generate(Path(tmp))), lib, progress=lambda *_: None)
    lines, failed = [], 0
    for test_id in lib.tests:
        if test_id not in PLANTED:
            failed += 1
            lines.append(f"  FAIL  {test_id}: no planted cases defined")
            continue
        r = result.result(test_id)
        column, expected = PLANTED[test_id]
        found = {str(row[r.columns.index(column)]) for row in r.rows} if r.executed else set()
        ok = r.status == EXCEPTIONS and found == expected
        failed += not ok
        detail = f"found {len(found)} of {len(expected)} planted" if ok else (
            f"status {r.status}; missed {sorted(expected - found)}; unexpected {sorted(found - expected)} {r.message}")
        lines.append(f"  {'PASS' if ok else 'FAIL'}  {test_id}: {detail}")
    if failed:
        lines.append(f"Self-test FAILED for {failed} rule(s). Do not use this copy; report the lines above.")
    else:
        lines.append(f"Self-test passed: all {len(lib.tests)} rules found exactly their planted exceptions "
                     f"and none of the near-misses.")
    return not failed, lines


def _num(value) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    return -float(s[:-1].replace(",", "")) if s.endswith("-") else float(s.replace(",", ""))


def demo_profile_text(sources: dict) -> str:
    """sources: dataset -> (file, sap report, rows or None, control total or None)."""
    parts = [DEMO_HEADER]
    for ds, (file, report, rows, total) in sources.items():
        parts.append(f"[datasets.{ds}]\nfiles = [\"{file}\"]\nsap_report = \"{report}\"")
        if rows is not None:
            parts.append(f"control_rows = {rows}")
        if total is not None:
            parts.append(f"control_total = {round(total, 2)}")
    parts.append(DEMO_FOOTER)
    return "\n".join(parts)
