"""
Damodaran Yatirim Sistemi — Engine
Kaynak: The Corporate Life Cycle (2024) + Investment Valuation

Moduller:
1. yasam_dongusu  — 6 asama tespiti
2. hikaye_sayilar — 3P testi + narativ
3. icsel_deger   — basitleştirilmiş DCF
4. fiyat_deger   — PD vs icsel deger
5. risk_analizi  — 3 katman
6. karar         — AL/TUT/SAT
"""

import re
from typing import Optional

# ── KOLON SABiTLERi ──────────────────────────────────────────────────────────
C_SEKTOR    = 'Hisse Sektör'
C_PAZAR     = 'Hisse Pazar Adı'
C_FIILI     = 'Fiili Dolaşımdaki Pay'
C_HALKA     = 'Halka Açıklık Oranı'
C_BETA      = 'Beta'
C_CARI      = 'Cari Oran'
C_ROA       = 'Aktif Karlılık (ROA) (%)'
C_FAVOK     = 'FAVÖK (Yıllık)'
C_FAVOK_MRJ = 'Favök Marjı (Yıllık)'
C_ROE       = 'Özsermaye Karlılığı (ROE) Yıllık (%)'
C_FV_FAVOK  = 'Firma Değeri / FAVÖK'
C_FV_NS     = 'Firma Değeri / Net Satış'
C_FK        = 'Fiyat Kazanç'
C_KAPANIS   = 'Hisse Kapanış'
C_PD        = 'Piyasa Değeri'
C_PD_EFK    = 'Piyasa Değeri / Esas Faaliyet Karı'
C_PDDD      = 'Piyasa Değeri / Defter Değeri'
C_AKTIF_BUY = 'Aktif Büyüme (%)'
C_FAVOK_BUY = 'FAVÖK Büyüme (%) (Yıllık)'
C_NS_BUY    = 'Net Satışlar Büyümesi (%) (Yıllık)'
C_OZK_BUY   = 'Özsermaye Büyümesi (%)'
C_PEG       = 'Peg Oranı'
C_ROIC      = 'Roic'
C_ROCE      = 'ROCE Oranı (%)'
C_FAVOK_FIN = 'FAVÖK / Net Finansman Gider'
C_NETBORC_F = 'Net Borç / FAVÖK (Yıllık) (%)'
C_NET_FIN_G = 'Net Finansman Giderleri (Yıllık)'
C_PIOTROSKI = 'Piotroski F Skor'
C_TOBIN_Q   = 'Tobin Q Oranı'
C_BODE      = 'Toplam Borç / Özsermaye'
C_BRUT_KAR  = 'Brüt Kar/Zarar'
C_DONEN     = 'Dönen Varlıklar'
C_DURAN     = 'Duran Varlıklar'
C_EFK       = 'Esas Faaliyet Karı /Zararı Net (Yıllık)'
C_FIN_GID   = 'Finansman Giderleri'
C_NAKIT     = 'İşletme Faaliyetlerinden Nakit Akışları'
C_YATIRIM_NK= 'Yatırım Faaliyetlerinden Kaynaklanan Nakit Akışlar'
C_KVB       = 'Kısa Vadeli Borçlar'
C_NAKIT_BEN = 'Nakit ve Nakit Benzerleri'
C_NK        = 'Net Dönem Karı / Zararı (Yıllık)'
C_NET_BORC  = 'Net Borç'
C_MDV       = 'Maddi Duran Varlıklar'
C_OZKAYNAK  = 'Özkaynaklar'
C_UVB       = 'Uzun Vadeli Borçlar'
C_NS        = 'Net Satışlar'
C_MARJ      = 'Favök Marjı (Yıllık)'  # EFK marj kolonu — FAVOK marji ile proxy

# Türkiye parametreleri
RISKSIZ_FAIZ  = 0.40   # %40 — TCMB politika faizi baz
ULKE_RISK_PRM = 0.05   # %5 — Türkiye ülke risk primi
PIYASA_PRM    = 0.06   # %6 — Equity risk premium
SURDURULEBILIR_BUYUME = 0.20  # %20 — uzun vade nominal büyüme (enflasyon dahil)


def safe_float(v) -> Optional[float]:
    if v is None or v == '': return None
    try:
        if isinstance(v, (int, float)): return float(v)
        return float(str(v).replace(',', '.').replace('%', '').strip())
    except: return None


def hesapla_pd(row) -> Optional[float]:
    v = safe_float(row.get(C_PD, ''))
    if v and v > 0: return v  # Fastweb PD zaten tam TL cinsinden
    return None


def fmt_milyon(v) -> str:
    if v is None: return '-'
    v = float(v)
    if abs(v) >= 1e12: return f"{v/1e12:.1f}T"
    if abs(v) >= 1e9:  return f"{v/1e9:.1f}Mr"
    if abs(v) >= 1e6:  return f"{v/1e6:.0f}M"
    return f"{v:.0f}"


def donem_from_filename(fn: str) -> Optional[str]:
    m = re.search(r'(\d{4})(\d{2})', fn)
    if m: return m.group(1) + m.group(2)
    return None


def fix_xlsx_styles(raw: bytes) -> bytes:
    import zipfile, io
    MINIMAL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>'
    try:
        zin = zipfile.ZipFile(io.BytesIO(raw))
        names = zin.namelist()
        if 'xl/styles.xml' not in names: return raw
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                zout.writestr(name, MINIMAL if name == 'xl/styles.xml' else zin.read(name))
        return buf.getvalue()
    except: return raw


def read_excel_bytes(raw: bytes) -> Optional[dict]:
    import pandas as pd, io
    for attempt in range(2):
        try:
            xf = pd.read_excel(io.BytesIO(raw), header=0)
            if 'Kod' not in xf.columns: continue
            result = {}
            for _, row in xf.iterrows():
                kod = str(row.get('Kod', '')).strip()
                if not kod or kod == 'nan': continue
                result[kod] = {c: row[c] for c in xf.columns if c != 'Kod'}
            return result if result else None
        except:
            if attempt == 0: raw = fix_xlsx_styles(raw)
    return None


# ── MODÜL 1: YASAM DONGUSU ───────────────────────────────────────────────────
def yasam_dongusu(quarters: dict, donems: list, kod: str) -> dict:
    """
    Damodaran 6 Asama
    Kriterler: Gelir buyumesi + Faaliyet marji + Yeniden yatirim
    """
    def seri(col):
        return [safe_float(quarters[d].get(kod, {}).get(col, '')) for d in donems]

    efk_s  = [v for v in seri(C_EFK)  if v is not None]
    ns_s   = [v for v in seri(C_NS)   if v is not None]
    marj_s = [v for v in seri(C_MARJ) if v is not None]
    nkt_s  = [v for v in seri(C_NAKIT) if v is not None]
    dur_s  = [v for v in seri(C_DURAN) if v is not None]
    roic_s = [v for v in seri(C_ROIC) if v is not None]

    # Tek dönem varsa direkt kolonları kullan
    son_row = quarters[donems[-1]].get(kod, {})
    ns_buy_direkt  = safe_float(son_row.get(C_NS_BUY, ''))
    marj_direkt    = safe_float(son_row.get(C_FAVOK_MRJ, ''))

    if len(ns_s) < 4 or len(efk_s) < 4:
        # Tek dönem — direkt kolonlarla minimal analiz
        ns_buy  = ns_buy_direkt
        son_marj = marj_direkt
        son_efk  = efk_s[-1] if efk_s else None
        efk_poz  = 1.0 if (son_efk and son_efk > 0) else 0.0
        marj_trend = 0
        yy_buy = None
        son_roic = safe_float(son_row.get(C_ROIC, ''))
        # Asama tespiti (tek dönem)
        if not son_efk or son_efk <= 0:
            asama, label, emoji, renk = 1, "Baslangic", "🌱", "#94A3B8"
            metrik, aciklama = "EV/Pazar Potansiyeli", "EFK negatif."
        elif ns_buy and ns_buy >= 50 and (not son_marj or son_marj < 15):
            asama, label, emoji, renk = 2, "Genc Buyume", "🧒", "#38BDF8"
            metrik, aciklama = "EV/Satis", f"Gelir hizla buyuyor ({ns_buy:.0f}%)."
        elif ns_buy and ns_buy >= 20:
            asama, label, emoji, renk = 3, "Yuksek Buyume", "🚀", "#4ADE80"
            metrik, aciklama = "PEG", f"Guclu buyume ({ns_buy:.0f}%)."
        elif ns_buy and ns_buy < -5:
            asama, label, emoji, renk = 6, "Dusus", "📉", "#F87171"
            metrik, aciklama = "PD/DD", "Gelir geriliyor."
        elif (not ns_buy or ns_buy < 10) and son_marj and son_marj >= 10:
            asama, label, emoji, renk = 5, "Olgun/Stabil", "🏛️", "#FCD34D"
            metrik, aciklama = "F/K, FV/FAVÖK", "Buyume dusuk, marj yuksek."
        else:
            asama, label, emoji, renk = 4, "Olgun Buyume", "💪", "#A78BFA"
            metrik, aciklama = "PEG", "Olgun buyume evresi."
        return {
            "asama": asama, "label": label, "emoji": emoji, "renk": renk,
            "metrik": metrik, "aciklama": aciklama,
            "ns_buy": round(ns_buy, 1) if ns_buy else None,
            "marj_son": round(son_marj, 1) if son_marj else None,
            "marj_trend": 0, "yy_buy": None,
            "efk_poz": round(efk_poz*100, 0),
            "son_roic": round(son_roic, 1) if son_roic else None,
        }

    son_efk  = efk_s[-1]
    son_marj = marj_s[-1] if marj_s else None
    son_nkt  = nkt_s[-1]  if nkt_s  else None
    son_roic = roic_s[-1] if roic_s else None

    # NS buyume (2Y)
    ns_buy = None
    if len(ns_s) >= 8 and ns_s[-8] and ns_s[-8] > 0:
        ns_buy = (ns_s[-1]/ns_s[-8] - 1)*100
    elif len(ns_s) >= 4 and ns_s[-4] and ns_s[-4] > 0:
        ns_buy = (ns_s[-1]/ns_s[-4] - 1)*100

    # Marj trendi
    marj_son = sum(marj_s[-4:])/4 if len(marj_s)>=4 else son_marj
    marj_onc = sum(marj_s[-8:-4])/4 if len(marj_s)>=8 else marj_son
    marj_trend = (marj_son - marj_onc) if (marj_son and marj_onc) else 0

    # Yeniden yatirim proxy (duran varlik buyumesi)
    yy_buy = None
    if len(dur_s) >= 4 and dur_s[-4] and dur_s[-4] > 0:
        yy_buy = (dur_s[-1]/dur_s[-4] - 1)*100

    # EFK istikrari
    efk_poz = sum(1 for x in efk_s[-8:] if x > 0) / min(len(efk_s), 8)
    nkt_poz = son_nkt is not None and son_nkt > 0

    # Asama tespiti
    if efk_poz < 0.5 or son_efk <= 0:
        asama, label, emoji, renk = 1, "Baslangic", "🌱", "#94A3B8"
        metrik = "EV/Pazar Potansiyeli"
        aciklama = "EFK istikrarsiz. Nakit yakiliyor. Deger hikayeye gore sekillenir."

    elif ns_buy and ns_buy >= 50 and (not son_marj or son_marj < 10) and marj_trend >= 0:
        asama, label, emoji, renk = 2, "Genc Buyume", "🧒", "#38BDF8"
        metrik = "EV/Satis"
        aciklama = f"Gelir hizla buyuyor ({ns_buy:.0f}%). Marj dusuk ama yukseliyor. Bar Mitzvah oncesi."

    elif ns_buy and ns_buy >= 20 and efk_poz >= 0.7 and marj_trend >= 0:
        asama, label, emoji, renk = 3, "Yuksek Buyume", "🚀", "#4ADE80"
        metrik = "EV/Satis, PEG"
        aciklama = f"Guclu buyume ({ns_buy:.0f}%). EFK istikrari yuksek, marj yukseliyor."

    elif (ns_buy and ns_buy < -5) or (marj_trend and marj_trend < -5 and son_marj and son_marj < 5):
        asama, label, emoji, renk = 6, "Dusus", "📉", "#F87171"
        metrik = "PD/DD"
        aciklama = "Gelir ve marj geriliyor. Pazar daralması veya rekabet baskisi."

    elif (not ns_buy or ns_buy < 10) and son_marj and son_marj >= 10 and nkt_poz:
        asama, label, emoji, renk = 5, "Olgun/Stabil", "🏛️", "#FCD34D"
        metrik = "F/K, FV/FAVÖK"
        aciklama = "Buyume yavasladi, marj yuksek ve stabil. Fazla nakit uretiyor."

    else:
        asama, label, emoji, renk = 4, "Olgun Buyume", "💪", "#A78BFA"
        metrik = "PEG"
        aciklama = f"Buyume olgunlasiyor{f' ({ns_buy:.0f}%)' if ns_buy else ''}. Karlilik stabil."

    return {
        "asama": asama, "label": label, "emoji": emoji, "renk": renk,
        "metrik": metrik, "aciklama": aciklama,
        "ns_buy": round(ns_buy, 1) if ns_buy else None,
        "marj_son": round(marj_son, 1) if marj_son else None,
        "marj_trend": round(marj_trend, 1),
        "yy_buy": round(yy_buy, 1) if yy_buy else None,
        "efk_poz": round(efk_poz*100, 0),
        "son_roic": round(son_roic, 1) if son_roic else None,
    }


# ── MODÜL 2: 3P TESTi ───────────────────────────────────────────────────────
def uc_p_testi(row: dict, yd: dict) -> dict:
    """
    Damodaran 3P Testi: Possible + Plausible + Probable
    Her biri 0-33 puan
    """
    puan = {"possible": 0, "plausible": 0, "probable": 0, "toplam": 0, "detay": []}

    ns_buy  = yd.get("ns_buy")
    marj    = yd.get("marj_son")
    asama   = yd.get("asama")
    efk_poz = yd.get("efk_poz", 0)
    roic    = safe_float(row.get(C_ROIC, ''))
    piotroski = safe_float(row.get(C_PIOTROSKI, ''))
    bode    = safe_float(row.get(C_BODE, ''))
    favok_fin = safe_float(row.get(C_FAVOK_FIN, ''))

    # POSSIBLE — Is modeli calisabilir mi?
    p1 = 0
    if asama in [2, 3, 4]: p1 += 15  # Aktif buyume asamasi
    if ns_buy and ns_buy > 10: p1 += 10  # Gelir buyuyor
    if marj and marj > 0: p1 += 8  # Pozitif marj
    p1 = min(p1, 33)
    puan["possible"] = p1
    puan["detay"].append(f"Possible: {p1}/33 — Is modeli {'calisir' if p1>20 else 'belirsiz'}")

    # PLAUSIBLE — Buyume surdurulebilir mi?
    p2 = 0
    if roic and roic > 10: p2 += 12  # ROIC > sermaye maliyeti
    if efk_poz >= 75: p2 += 10  # EFK tutarli
    if marj and marj > 5: p2 += 6  # Pozitif marj
    if bode and bode < 2: p2 += 5  # Dusuk borclanma
    p2 = min(p2, 33)
    puan["plausible"] = p2
    puan["detay"].append(f"Plausible: {p2}/33 — Surdurulebilirlik {'yuksek' if p2>20 else 'dusuk'}")

    # PROBABLE — Piyasa bunu hala fiyatlamadi mi?
    p3 = 0
    if piotroski and piotroski >= 7: p3 += 10  # Guclu finansal saglik
    if favok_fin and favok_fin > 2: p3 += 8  # Fin giderleri karsilanabilir
    if asama in [2, 3]: p3 += 10  # Erken asama = piyasa henuz fiyatlamadi
    if ns_buy and ns_buy > 30: p3 += 5  # Guclu momentum
    p3 = min(p3, 33)
    puan["probable"] = p3
    puan["detay"].append(f"Probable: {p3}/33 — Piyasa {'henuz fiyatlamadi' if p3>20 else 'fiyatlamis olabilir'}")

    puan["toplam"] = p1 + p2 + p3
    return puan


# ── MODÜL 3: iCSEL DEGER ────────────────────────────────────────────────────
def icsel_deger_hesapla(row: dict, quarters: dict, donems: list, kod: str, yd: dict) -> dict:
    """
    Damodaran Basitlestirilmis DCF
    5 Girdi: Gelir buyumesi + Hedef marj + Yeniden yatirim + Risk + Buyume suresi
    """
    efk_son = safe_float(row.get(C_EFK, ''))
    ns_son  = safe_float(row.get(C_NS, ''))
    marj    = yd.get("marj_son")
    ns_buy  = yd.get("ns_buy")
    roic    = safe_float(row.get(C_ROIC, ''))
    beta    = safe_float(row.get(C_BETA, ''))
    bode    = safe_float(row.get(C_BODE, ''))
    asama   = yd.get("asama")

    if not efk_son or efk_son <= 0 or not ns_son:
        return {"deger": None, "aciklama": "EFK negatif — DCF uygulanamaz"}

    # 1. BUYUME ORANI (asama bazli)
    if asama == 2:   buyume = min((ns_buy or 30)/100, 0.50)
    elif asama == 3: buyume = min((ns_buy or 20)/100, 0.40)
    elif asama == 4: buyume = min((ns_buy or 10)/100, 0.20)
    elif asama == 5: buyume = 0.08
    elif asama == 6: buyume = 0.02
    else:            buyume = 0.15

    # 2. iSKONTO ORANI (WACC proxy)
    # Beta yoksa asama bazli tahmin
    if not beta:
        beta_tahmin = {1:1.5, 2:1.3, 3:1.1, 4:0.9, 5:0.8, 6:1.2}.get(asama, 1.0)
    else:
        beta_tahmin = beta

    # Borç oranı
    borc_orani = min((bode or 0.5) / ((bode or 0.5) + 1), 0.5)
    ozk_orani  = 1 - borc_orani

    # WACC = Özkaynak maliyeti × özk oranı + Borç maliyeti × borç oranı × (1-vergi)
    ozk_maliyeti  = RISKSIZ_FAIZ + beta_tahmin * PIYASA_PRM + ULKE_RISK_PRM
    borc_maliyeti = RISKSIZ_FAIZ + 0.03  # spread
    wacc = ozk_maliyeti * ozk_orani + borc_maliyeti * borc_orani * 0.75

    # 3. BUYUME SURESi (asama bazli)
    sureler = {1: 10, 2: 8, 3: 7, 4: 5, 5: 3, 6: 2}
    sure = sureler.get(asama, 5)

    # 4. YENiDEN YATIRIM ORANI
    # ROIC varsa: YY oranı = büyüme / ROIC
    if roic and roic > 0:
        yy_orani = min(buyume / (roic/100), 0.8)
    else:
        yy_orani = 0.4  # varsayilan

    # 5. SERBEST NAKiT AKISI (FCF)
    fcf_baslangic = efk_son * (1 - yy_orani) * 0.75  # vergi sonrasi

    # DCF Hesabi (2 asama)
    pv_toplam = 0
    efk_mevcut = efk_son
    for yil in range(1, sure + 1):
        efk_mevcut = efk_mevcut * (1 + buyume)
        fcf = efk_mevcut * (1 - yy_orani) * 0.75
        pv_toplam += fcf / ((1 + wacc) ** yil)

    # Terminal deger (Gordon modeli)
    terminal_efk = efk_mevcut * (1 + SURDURULEBILIR_BUYUME)
    terminal_yy  = min(SURDURULEBILIR_BUYUME / max(wacc, 0.01), 0.6)
    terminal_fcf = terminal_efk * (1 - terminal_yy) * 0.75
    terminal_deger = terminal_fcf / max(wacc - SURDURULEBILIR_BUYUME, 0.01)
    pv_terminal = terminal_deger / ((1 + wacc) ** sure)

    if pv_toplam != pv_toplam or pv_terminal != pv_terminal:
        return {"deger": None, "aciklama": "Hesaplama hatasi (NaN)"}
    icsel_deger_toplam = pv_toplam + pv_terminal

    return {
        "deger": icsel_deger_toplam,
        "wacc": round(wacc*100, 1),
        "buyume": round(buyume*100, 1),
        "sure": sure,
        "yy_orani": round(yy_orani*100, 1),
        "pv_buyume": round(pv_toplam) if pv_toplam == pv_toplam else 0,
        "pv_terminal": round(pv_terminal) if pv_terminal == pv_terminal else 0,
        "aciklama": f"WACC:{wacc*100:.0f}% Buyume:{buyume*100:.0f}% Sure:{sure}Y"
    }


# ── MODÜL 4: FiYAT vs DEGER ─────────────────────────────────────────────────
def fiyat_deger_analizi(row: dict, icsel: dict) -> dict:
    pd_val = hesapla_pd(row)
    id_val = icsel.get("deger")

    if not pd_val or not id_val or id_val <= 0:
        return {"guvenlik_marji": None, "karar_sinyal": None}

    guvenlik_marji = (id_val - pd_val) / id_val * 100

    if guvenlik_marji >= 40:
        karar_sinyal = "GUCLU AL"
        renk = "#4ADE80"
    elif guvenlik_marji >= 20:
        karar_sinyal = "AL"
        renk = "#86EFAC"
    elif guvenlik_marji >= -10:
        karar_sinyal = "TUT/IZLE"
        renk = "#FCD34D"
    elif guvenlik_marji >= -30:
        karar_sinyal = "DIKKATLI"
        renk = "#FB923C"
    else:
        karar_sinyal = "PAHALI"
        renk = "#F87171"

    return {
        "pd_val": pd_val,
        "id_val": id_val,
        "guvenlik_marji": round(guvenlik_marji, 1),
        "karar_sinyal": karar_sinyal,
        "renk": renk,
        "pd_fmt": fmt_milyon(pd_val),
        "id_fmt": fmt_milyon(id_val),
    }


# ── MODÜL 5: RiSK ANALiZi ───────────────────────────────────────────────────
def risk_analizi(row: dict, yd: dict) -> dict:
    """
    Damodaran 3 Katman Risk:
    1. Is riski (Piotroski + borç + cari oran)
    2. Buyume riski (EFK istikrari + nakit)
    3. Piyasa riski (beta + halka açıklık)
    """
    piotroski  = safe_float(row.get(C_PIOTROSKI, ''))
    bode       = safe_float(row.get(C_BODE, ''))
    cari       = safe_float(row.get(C_CARI, ''))
    beta       = safe_float(row.get(C_BETA, ''))
    favok_fin  = safe_float(row.get(C_FAVOK_FIN, ''))
    halka      = safe_float(row.get(C_HALKA, ''))
    netborc_f  = safe_float(row.get(C_NETBORC_F, ''))

    # Is riski (0-100, dusuk = iyi)
    is_risk = 50
    if piotroski:
        is_risk -= (piotroski - 4.5) * 5  # 9=min risk, 0=max risk
    if bode:
        is_risk += min(bode * 5, 25)
    if cari:
        is_risk -= min((cari - 1) * 10, 20)
    is_risk = max(0, min(100, is_risk))

    # Buyume riski
    buy_risk = 50
    efk_poz = yd.get("efk_poz", 50)
    buy_risk -= (efk_poz - 50) * 0.5
    if favok_fin:
        buy_risk -= min((favok_fin - 2) * 5, 20)
    if netborc_f:
        buy_risk += min(netborc_f / 10, 25)
    buy_risk = max(0, min(100, buy_risk))

    # Piyasa riski
    piyasa_risk = 50
    if beta:
        piyasa_risk = beta * 50
    if halka:
        piyasa_risk += (20 - min(halka, 20)) * 0.5  # dusuk halka aciklik = risk
    piyasa_risk = max(0, min(100, piyasa_risk))

    toplam_risk = (is_risk + buy_risk + piyasa_risk) / 3

    risk_seviye = "DUSUK" if toplam_risk < 35 else "ORTA" if toplam_risk < 60 else "YUKSEK"
    risk_renk   = "#4ADE80" if toplam_risk < 35 else "#FCD34D" if toplam_risk < 60 else "#F87171"

    return {
        "is_risk": round(is_risk, 0),
        "buy_risk": round(buy_risk, 0),
        "piyasa_risk": round(piyasa_risk, 0),
        "toplam_risk": round(toplam_risk, 0),
        "seviye": risk_seviye,
        "renk": risk_renk,
    }


# ── MODÜL 6: NIHAI KARAR ────────────────────────────────────────────────────
def nihai_karar(yd: dict, uc_p: dict, fiyat: dict, risk: dict) -> dict:
    """
    Damodaran Karar Çerçevesi:
    Guvenlik marji + Risk seviyesi + 3P puani + Yasam dongusu kombinasyonu
    """
    gm        = fiyat.get("guvenlik_marji")
    risk_puan = risk.get("toplam_risk", 50)
    uc_p_puan = uc_p.get("toplam", 0)
    asama     = yd.get("asama")
    karar_sinyal = fiyat.get("karar_sinyal")

    if gm is None:
        return {"karar": "VERi YOK", "puan": 0, "renk": "#475569", "aciklama": "Icsel deger hesaplanamadi."}

    # Temel puan
    puan = 50

    # Guvenlik marji etkisi
    puan += gm * 0.5

    # Risk duzeltmesi (yuksek risk puani dusuruyor)
    puan -= (risk_puan - 50) * 0.3

    # 3P etkisi
    puan += (uc_p_puan - 50) * 0.2

    # Asama bonusu (erken asama = potansiyel)
    asama_bonus = {2: 10, 3: 8, 4: 3, 5: -2, 6: -10, 1: -5}
    puan += asama_bonus.get(asama, 0)

    puan = max(0, min(100, puan))

    if puan >= 75:
        karar, renk = "GUCLU AL 🟢", "#4ADE80"
    elif puan >= 60:
        karar, renk = "AL 🟢", "#86EFAC"
    elif puan >= 45:
        karar, renk = "TUT / iZLE 🟡", "#FCD34D"
    elif puan >= 30:
        karar, renk = "DiKKATLi 🟠", "#FB923C"
    else:
        karar, renk = "PAHALI / KACIN 🔴", "#F87171"

    aciklama = (
        f"Guvenlik Marji: {gm:.0f}% | "
        f"Risk: {risk.get('seviye','?')} | "
        f"3P: {uc_p_puan}/99 | "
        f"Asama: {yd.get('label','?')}"
    )

    return {"karar": karar, "puan": round(puan, 0), "renk": renk, "aciklama": aciklama}


# ── TAM ANALiZ ───────────────────────────────────────────────────────────────
def tam_analiz(kod: str, quarters: dict, donems: list) -> dict:
    """Tek hisse için tam Damodaran analizi"""
    row = quarters[donems[-1]].get(kod, {})
    if not row: return {}

    yd    = yasam_dongusu(quarters, donems, kod)
    uc_p  = uc_p_testi(row, yd)
    icsel = icsel_deger_hesapla(row, quarters, donems, kod, yd)
    fiyat = fiyat_deger_analizi(row, icsel)
    risk  = risk_analizi(row, yd)
    karar = nihai_karar(yd, uc_p, fiyat, risk)

    return {
        "kod": kod,
        "sektor": row.get(C_SEKTOR, ''),
        "yasam_dongusu": yd,
        "uc_p": uc_p,
        "icsel": icsel,
        "fiyat": fiyat,
        "risk": risk,
        "karar": karar,
    }


print("Damodaran Engine hazir!")
