import streamlit as st
import pandas as pd
import json
import io
from datetime import datetime

from damodaran_engine import (
    read_excel_bytes, donem_from_filename, fmt_milyon, safe_float,
    tam_analiz, firsat_skoru, yasam_dongusu, uc_p_testi, icsel_deger_hesapla,
    fiyat_deger_analizi, risk_analizi, nihai_karar, hesapla_pd,
    C_SEKTOR, C_NS_BUY, C_FAVOK_MRJ, C_ROIC, C_BETA, C_PEG,
    C_FV_FAVOK, C_PDDD, C_ROE, C_PIOTROSKI, C_BODE, C_CARI,
    C_NAKIT, C_EFK, C_NS, C_FAVOK, C_OZKAYNAK, C_NETBORC_F,
    RISKSIZ_FAIZ, PIYASA_PRM, ULKE_RISK_PRM,
)

st.set_page_config(
    page_title="Damodaran Yatirim Sistemi",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700&display=swap');
* { font-family: 'Space Grotesk', sans-serif; }
[data-testid="stAppViewContainer"] { background:#080E17; }
[data-testid="stSidebar"] { background:#060D18; border-right:1px solid #0F2040; }
[data-testid="stSidebar"] * { color:#94A3B8 !important; }
h1,h2,h3 { color:#E2E8F0 !important; }
p, li { color:#94A3B8 !important; }
.sb-brand { font-size:20px;font-weight:800;letter-spacing:-0.5px;padding:12px 0 2px }
.sb-brand span { color:#A78BFA; }
.sb-sub { font-size:10px;color:#1E3448 !important;letter-spacing:2px;text-transform:uppercase;margin-bottom:16px }
.ph { padding:18px 24px 14px;margin-bottom:16px;border-bottom:1px solid #0F2040 }
.ph-badge { display:inline-block;font-size:9px;font-weight:700;letter-spacing:2px;
  text-transform:uppercase;padding:3px 10px;border-radius:20px;margin-bottom:6px }
.ph-title { font-size:22px;font-weight:800;color:#E2E8F0;margin:0 }
.ph-sub { font-size:12px;color:#475569;margin-top:4px }
.mrow { display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap }
.mc { flex:1;min-width:80px;background:#0D1926;border:1px solid #0F2040;
  border-radius:10px;padding:12px 14px;text-align:center }
.mc-num { font-size:22px;font-weight:900;color:#E2E8F0 }
.mc-lbl { font-size:9px;color:#475569;text-transform:uppercase;letter-spacing:1px;margin-top:3px }
</style>
""", unsafe_allow_html=True)

# Session state
for k, v in [('quarters', {}), ('donems', []), ('son_donem', None),
              ('aktif_sayfa', None), ('hisse_git', None)]:
    if k not in st.session_state: st.session_state[k] = v

SAYFALAR = [
    "📖 Tanitim",
    "🏠 Genel Bakis",
    "🔄 Yasam Dongusu",
    "🔍 Hisse Tarayici",
    "📊 Detay Analizi",
    "⚙️ Ayarlar",
]

def git_detay(kod):
    st.session_state.hisse_git = kod
    st.session_state.aktif_sayfa = "📊 Detay Analizi"
    st.rerun()

# ── SiDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("<div class='sb-brand'>DAMODARAN<br><span>YATİRİM</span></div>", unsafe_allow_html=True)
    st.markdown("<div class='sb-sub'>Kurumsal Yasam Dongusu · DCF · Risk</div>", unsafe_allow_html=True)

    if st.session_state.aktif_sayfa:
        idx = SAYFALAR.index(st.session_state.aktif_sayfa) if st.session_state.aktif_sayfa in SAYFALAR else 0
        st.session_state.aktif_sayfa = None
    else:
        idx = 0

    page = st.radio("", SAYFALAR, index=idx, label_visibility="collapsed", key="page_radio")

    st.markdown("<hr>", unsafe_allow_html=True)

    # Veri yukle
    st.markdown("<div style='font-size:10px;color:#475569;text-transform:uppercase;letter-spacing:1px'>Veri Yukle</div>", unsafe_allow_html=True)
    yuklu = st.file_uploader("", type=['xlsx'], accept_multiple_files=True,
                              label_visibility="collapsed", key="xl_uploader")
    if yuklu:
        yeni = {}
        for f in yuklu:
            donem = donem_from_filename(f.name)
            if donem:
                data = read_excel_bytes(f.read())
                if data: yeni[donem] = data
        if yeni:
            st.session_state.quarters.update(yeni)
            st.session_state.donems = sorted(st.session_state.quarters.keys())
            st.session_state.son_donem = st.session_state.donems[-1]
            st.success(f"{len(yeni)} donem yuklendi")

    if st.session_state.donems:
        st.markdown(f"<div style='background:#0D1926;border:1px solid #0F2040;border-radius:8px;padding:8px 12px;margin-top:8px'>"
                    f"<div style='font-size:9px;color:#475569'>YUKLEMELi VERi</div>"
                    f"<div style='font-size:16px;font-weight:800;color:#A78BFA'>{len(st.session_state.donems)} donem</div>"
                    f"<div style='font-size:10px;color:#475569'>Son: {st.session_state.son_donem[:4]}/{st.session_state.son_donem[4:]}</div>"
                    f"</div>", unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-size:9px;color:#1E3448'>Türkiye Parametreleri<br>"
                f"Risksiz Faiz: %{RISKSIZ_FAIZ*100:.0f} | ERP: %{PIYASA_PRM*100:.0f} | Ülke Premi: %{ULKE_RISK_PRM*100:.0f}</div>",
                unsafe_allow_html=True)

# Veri yoksa dur
quarters = st.session_state.quarters
donems   = st.session_state.donems
son_d    = st.session_state.son_donem

def bos():
    st.markdown("<div style='text-align:center;padding:60px;color:#475569'>"
                "<div style='font-size:48px'>📂</div>"
                "<div style='font-size:16px;margin-top:12px'>Sidebar'dan Excel dosyalarını yükle</div>"
                "</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SAYFA 1: GENEL BAKIS
# ══════════════════════════════════════════════════════════════════════════════
if page == "📖 Tanitim":
    st.markdown("""<div class='ph'>
    <div class='ph-badge' style='background:#0A1020;color:#A78BFA;border:1px solid #4C1D95'>HOŞGELDİNİZ</div>
    <div class='ph-title'>Damodaran Yatırım Sistemi</div>
    <div class='ph-sub'>Kurumsal Yasam Dongusu · DCF Degerleme · Risk Analizi · Firsat Tespiti</div>
    </div>""", unsafe_allow_html=True)

    # Damodaran alıntısı
    st.markdown(
        "<div style='background:#0A1020;border-left:4px solid #A78BFA;padding:16px 20px;border-radius:8px;margin-bottom:20px'>"
        "<div style='font-size:14px;color:#E2E8F0;font-style:italic;line-height:1.7'>"
        '"Bir hissedarin sinikliginin en iyi tanimi: Her seyin fiyatini bilen, '
        "hicbir seyin degerini bilmeyen.' — Oscar Wilde<br><br>"
        "Bu sistem sizi fiyatci degil, yatirmci yapar."
        "</div><div style='font-size:11px;color:#475569;margin-top:8px'>— Damodaran, Investment Valuation</div></div>",
        unsafe_allow_html=True
    )

    # Sistemin amacı
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(
            "<div style='background:#0D1926;border:1px solid #0F2040;border-radius:12px;padding:18px'>"
            "<div style='color:#4ADE80;font-size:14px;font-weight:800;margin-bottom:12px'>🎯 Bu Sistem Ne Yapar?</div>"
            "<p style='color:#64748B;font-size:12px;line-height:1.8'>"
            "• <b style='color:#E2E8F0'>Yasam Dongusunu</b> tespit eder — sirket buyume mi, olgunlasma mi, dusus mu?<br>"
            "• <b style='color:#E2E8F0'>Icsel Degeri</b> hesaplar — DCF ile gercek deger nedir?<br>"
            "• <b style='color:#E2E8F0'>Piyasa Fiyatini</b> karsilastirir — ucuz mu, pahali mi?<br>"
            "• <b style='color:#E2E8F0'>Riski</b> olcer — is, buyume ve piyasa riski<br>"
            "• <b style='color:#E2E8F0'>Firsatlari</b> isaretler — 7 Damodaran sinyali</p></div>",
            unsafe_allow_html=True
        )
    with col2:
        st.markdown(
            "<div style='background:#0D1926;border:1px solid #0F2040;border-radius:12px;padding:18px'>"
            "<div style='color:#38BDF8;font-size:14px;font-weight:800;margin-bottom:12px'>📖 Damodaran Ne Diyor?</div>"
            "<p style='color:#64748B;font-size:12px;line-height:1.8'>"
            "• Piyasalar hata yapar ama bu hatalar <b style='color:#E2E8F0'>pencereler halinde</b> acilip kapanir<br>"
            "• <b style='color:#E2E8F0'>Duygu yok</b> — sadece sayilar ve hikaye<br>"
            "• Iyi degerleme bir hikaye anlatir, iyi hikaye sayilarla desteklenir<br>"
            "• Bir sey cok iyi gorunuyorsa <b style='color:#E2E8F0'>buyuk ihtimalle degil</b><br>"
            "• Piyasanin her zaman dogru oldugunu varsay, sonra ikna et</p></div>",
            unsafe_allow_html=True
        )

    st.markdown("<hr>", unsafe_allow_html=True)

    # Kullanim Akisi
    st.markdown("<h3 style='color:#E2E8F0;font-size:15px;margin-bottom:12px'>🔄 Nasıl Kullanılır?</h3>", unsafe_allow_html=True)
    adimlar = [
        ("1", "#4ADE80", "Excel Yukle", "Sidebar'dan 25 donem Excel dosyalarini yukle. Ne kadar cok donem, o kadar isabetli analiz."),
        ("2", "#38BDF8", "Yasam Dongusu", "🔄 Yasam Dongusu sekmesine gec. BIST'teki tum hisselerin 6 asamaya dagilimini gor. Bar Mitzvah adaylarini incele."),
        ("3", "#A78BFA", "Firsat Tara", "🔍 Hisse Tarayici'da filtrele. Asama 2-3 + Dusuk Risk kombinasyonu Damodaran'in tercih ettigi penceredir."),
        ("4", "#FCD34D", "Detay Analiz", "📊 Ilginc hisseleri Detay Analizi'nde incele. 4 kart + DCF + 7 Firsat Sinyali tam resmi goster."),
        ("5", "#FB923C", "Takip Et", "Her donem yeni Excel yukle, asamasi degisen hisseleri izle. Buyume baslarken erken girmek hedeftir."),
    ]
    adim_cols = st.columns(5)
    for (num, renk, baslik, acik), col in zip(adimlar, adim_cols):
        col.markdown(
            f"<div style='background:#0D1926;border:1px solid {renk};border-radius:10px;padding:14px;text-align:center;height:180px'>"
            f"<div style='font-size:24px;font-weight:900;color:{renk}'>{num}</div>"
            f"<div style='font-size:12px;font-weight:700;color:#E2E8F0;margin:6px 0'>{baslik}</div>"
            f"<div style='font-size:10px;color:#475569;line-height:1.4'>{acik}</div></div>",
            unsafe_allow_html=True
        )

    st.markdown("<hr>", unsafe_allow_html=True)

    # 7 Fırsat Sinyali
    st.markdown("<h3 style='color:#E2E8F0;font-size:15px;margin-bottom:12px'>🎯 7 Damodaran Fırsat Sinyali</h3>", unsafe_allow_html=True)
    sinyaller = [
        ("🎓", "Bar Mitzvah", "Genc Buyume asamasinda, EFK pozitife donuyor. En degerli gecis noktasi."),
        ("💎", "ROIC > Maliyet", "Sirket sermayesini maliyetinin ustunde kullanabiliyor. Deger URATIYOR."),
        ("📈", "PEG < 1", "Buyume icin odenen fiyat dusuk. Buyumeye gore ucuz."),
        ("🎯", "Dusuk PD/DD + Yuksek ROE", "Defter degerinin altinda fiyat ama yuksek ozsermaye getirisi."),
        ("🏆", "Piotroski 7+", "Finansal saglik ve muhasebe kalitesi dogrulanmis."),
        ("🚀", "Buyume + PD/Satis<1", "Hizli buyurken satis bazinda ucuz. Piyasa buyumeyi fiyatlamamis."),
        ("🔄", "Contrarian", "Dusus asamasinda ama EFK pozitif. Piyasa asiri tepki vermis."),
    ]
    s_cols = st.columns(7)
    for (em, baslik, acik), col in zip(sinyaller, s_cols):
        col.markdown(
            f"<div style='background:#0D1926;border:1px solid #0F2040;border-radius:10px;padding:12px;text-align:center'>"
            f"<div style='font-size:22px'>{em}</div>"
            f"<div style='font-size:10px;font-weight:700;color:#E2E8F0;margin:6px 0'>{baslik}</div>"
            f"<div style='font-size:9px;color:#475569;line-height:1.3'>{acik}</div></div>",
            unsafe_allow_html=True
        )

    st.markdown("<hr>", unsafe_allow_html=True)

    # Damodaran 10 Kuralı
    st.markdown("<h3 style='color:#E2E8F0;font-size:15px;margin-bottom:12px'>📜 Damodaran'ın 10 Kuralı</h3>", unsafe_allow_html=True)
    kurallar = [
        ("1", "Modeli degil, ilkeleri koruyun", "Model degisebilir ama temel prensipler degismez."),
        ("2", "Piyasayi dinle, ona tapma", "Piyasa fiyatina saygi goster ama kole olma."),
        ("3", "Risk degeri etkiler", "Belirsizlik arttikca beklenen getiri artmali."),
        ("4", "Buyume bedavaya gelmez", "Yuksek buyume icin yuksek yatirim gerekir."),
        ("5", "Her sey biter", "Buyume de sonunda yavaslayacak, bunu fiyatla."),
        ("6", "Batma riskini goz ardi etme", "Bircok sirket hayatta kalamaz."),
        ("7", "Degerleme önyargili olabilir", "Her analizde kendi duygularini kontrol et."),
        ("8", "Basit modeller iyidir", "Karmasik model ≠ dogru analiz."),
        ("9", "Hikaye + Sayi birlikte", "Rakamlar olmadan hikaye fantezi, hikaye olmadan rakam anlamsiz."),
        ("10", "Hata yapilabilir", "Hatadan kork degil, süreci dogru yonet."),
    ]
    for num, baslik, acik in kurallar:
        st.markdown(
            f"<div style='display:flex;align-items:flex-start;gap:12px;padding:8px 0;border-bottom:1px solid #0F2040'>"
            f"<div style='font-size:14px;font-weight:900;color:#A78BFA;min-width:24px'>{num}</div>"
            f"<div><div style='font-size:12px;font-weight:700;color:#E2E8F0'>{baslik}</div>"
            f"<div style='font-size:11px;color:#475569'>{acik}</div></div></div>",
            unsafe_allow_html=True
        )

elif page == "🏠 Genel Bakis":
    st.markdown("""<div class='ph'>
    <div class='ph-badge' style='background:#0A1020;color:#A78BFA;border:1px solid #4C1D95'>DAMODARAN</div>
    <div class='ph-title'>Genel Bakis</div>
    <div class='ph-sub'>BIST Damodaran Analizi — Yasam Dongusu + DCF + Risk</div>
    </div>""", unsafe_allow_html=True)

    if not quarters: bos(); st.stop()

    son_data = quarters[son_d]

    # Tum hisseleri analiz et
    with st.spinner("Analiz yapiliyor..."):
        sonuclar = []
        for kod, row in son_data.items():
            s = tam_analiz(kod, quarters, donems)
            if s and s.get('fiyat', {}).get('guvenlik_marji') is not None:
                sonuclar.append(s)

    # Karar dagilimi
    karar_sayim = {"GUCLU AL": 0, "AL": 0, "TUT/IZLE": 0, "DIKKATLI": 0, "PAHALI": 0}
    for s in sonuclar:
        k = s['karar']['karar']
        if "GUCLU AL" in k: karar_sayim["GUCLU AL"] += 1
        elif "AL" in k:     karar_sayim["AL"] += 1
        elif "TUT" in k:    karar_sayim["TUT/IZLE"] += 1
        elif "DIKKAT" in k: karar_sayim["DIKKATLI"] += 1
        else:               karar_sayim["PAHALI"] += 1

    # Asama dagilimi
    asama_sayim = {i: 0 for i in range(1, 7)}
    for s in sonuclar:
        a = s['yasam_dongusu'].get('asama')
        if a: asama_sayim[a] += 1

    # Ozet kartlar
    toplam = len(sonuclar)
    st.markdown(
        f"<div class='mrow'>"
        f"<div class='mc'><div class='mc-num' style='color:#E2E8F0'>{toplam}</div><div class='mc-lbl'>Toplam Hisse</div></div>"
        f"<div class='mc'><div class='mc-num' style='color:#4ADE80'>{karar_sayim['GUCLU AL']}</div><div class='mc-lbl'>Guclu Al</div></div>"
        f"<div class='mc'><div class='mc-num' style='color:#86EFAC'>{karar_sayim['AL']}</div><div class='mc-lbl'>Al</div></div>"
        f"<div class='mc'><div class='mc-num' style='color:#FCD34D'>{karar_sayim['TUT/IZLE']}</div><div class='mc-lbl'>Tut/izle</div></div>"
        f"<div class='mc'><div class='mc-num' style='color:#FB923C'>{karar_sayim['DIKKATLI']}</div><div class='mc-lbl'>Dikkatli</div></div>"
        f"<div class='mc'><div class='mc-num' style='color:#F87171'>{karar_sayim['PAHALI']}</div><div class='mc-lbl'>Pahali</div></div>"
        f"</div>", unsafe_allow_html=True
    )

    # En iyi firsatlar
    st.markdown("<h3 style='color:#E2E8F0;font-size:15px;margin:16px 0 8px'>🏆 En İyi Firsatlar (Guvenlik Marjina Gore)</h3>", unsafe_allow_html=True)
    en_iyi = sorted([s for s in sonuclar if s['fiyat'].get('guvenlik_marji', -999) > 0],
                    key=lambda x: x['karar']['puan'], reverse=True)[:15]

    if en_iyi:
        df_iyi = pd.DataFrame([{
            "Kod":    s['kod'],
            "Sektor": s['sektor'][:20],
            "Asama":  f"{s['yasam_dongusu'].get('emoji','')} {s['yasam_dongusu'].get('label','')}",
            "PD":     s['fiyat'].get('pd_fmt', '-'),
            "İcsel D.": s['fiyat'].get('id_fmt', '-'),
            "GM%":    s['fiyat'].get('guvenlik_marji', '-'),
            "Risk":   s['risk'].get('seviye', '-'),
            "Karar":  s['karar']['karar'][:12],
            "Firsat": f"{s['firsat']['puan']}/7 {s['firsat']['seviye'][:6]}",
            "Puan":   s['karar']['puan'],
        } for s in en_iyi])
        st.dataframe(df_iyi, hide_index=True, use_container_width=True, height=500)

        bt_cols = st.columns(8)
        for i, s in enumerate(en_iyi[:16]):
            with bt_cols[i % 8]:
                if st.button(s['kod'], key=f"gb_{s['kod']}"):
                    git_detay(s['kod'])

# ══════════════════════════════════════════════════════════════════════════════
# SAYFA 2: YASAM DONGUSU
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🔄 Yasam Dongusu":
    import plotly.graph_objects as go
    st.markdown("""<div class='ph'>
    <div class='ph-badge' style='background:#0A1020;color:#A78BFA;border:1px solid #4C1D95'>YASAM DONGUSU</div>
    <div class='ph-title'>Damodaran 6 Asama Analizi</div>
    <div class='ph-sub'>Baslangic · Genc Buyume · Yuksek Buyume · Olgun Buyume · Stabil · Dusus</div>
    </div>""", unsafe_allow_html=True)

    if not quarters: bos(); st.stop()

    ASAMA_META = {
        1: ("🌱", "Baslangic",    "#94A3B8", "EFK istikrarsiz. Hikaye one cikar."),
        2: ("🧒", "Genc Buyume",  "#38BDF8", "Bar Mitzvah esigi. Gelir hizli, EFK pozitife doniyor."),
        3: ("🚀", "Yuksek Buyume","#4ADE80", "EFK guçlu, gelir hizli, marj yukseliyor."),
        4: ("💪", "Olgun Buyume", "#A78BFA", "Buyume devam ediyor, olgunlasiyor."),
        5: ("🏛️", "Olgun/Stabil", "#FCD34D", "Dusuk buyume, yuksek marj, fazla nakit."),
        6: ("📉", "Dusus",        "#F87171", "Gelir ve marj geriliyor."),
    }

    son_data = quarters[son_d]
    dagilim  = {i: [] for i in range(1, 7)}

    for kod, row in son_data.items():
        pd_val = hesapla_pd(row)
        if not pd_val or pd_val <= 0: continue
        yd = yasam_dongusu(quarters, donems, kod)
        a  = yd.get("asama")
        if not a: continue
        dagilim[a].append({
            "kod": kod, "sektor": row.get(C_SEKTOR, ""),
            "ns_buy": yd.get("ns_buy"), "marj": yd.get("marj_son"),
            "yy_buy": yd.get("yy_buy"), "efk_poz": yd.get("efk_poz"),
            "roic":   yd.get("son_roic"), "pd_val": pd_val,
            "aciklama": yd.get("aciklama",""),
        })

    toplam = sum(len(v) for v in dagilim.values())

    # Kartlar
    kart_html = "<div class='mrow'>"
    for a in range(1, 7):
        em, lbl, renk, _ = ASAMA_META[a]
        n = len(dagilim[a])
        kart_html += (f"<div class='mc' style='border-top:3px solid {renk}'>"
                      f"<div style='font-size:18px'>{em}</div>"
                      f"<div class='mc-num' style='color:{renk}'>{n}</div>"
                      f"<div class='mc-lbl'>{lbl}</div>"
                      f"<div style='font-size:9px;color:#475569'>%{n/toplam*100:.0f}</div></div>")
    kart_html += "</div>"
    st.markdown(kart_html, unsafe_allow_html=True)

    # Pasta + Açıklama
    col1, col2 = st.columns([1, 1])
    with col1:
        fig = go.Figure(go.Pie(
            labels=[f"{ASAMA_META[i][1]} ({len(dagilim[i])})" for i in range(1,7)],
            values=[len(dagilim[i]) for i in range(1,7)],
            marker=dict(colors=[ASAMA_META[i][2] for i in range(1,7)],
                        line=dict(color="#080E17", width=2)),
            hole=0.45, textinfo="percent", textfont=dict(size=11, color="white"),
        ))
        fig.update_layout(paper_bgcolor="#080E17", plot_bgcolor="#080E17",
                          font=dict(color="#94A3B8", size=10),
                          margin=dict(l=10,r=10,t=10,b=10), height=260,
                          legend=dict(font=dict(size=9), bgcolor="rgba(0,0,0,0)"))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.markdown(
            "<div style='background:#0A1020;border:1px solid #4C1D95;border-radius:10px;padding:16px 18px'>"
            "<div style='color:#A78BFA;font-weight:800;font-size:14px;margin-bottom:10px'>🎓 Bar Mitzvah Testi</div>"
            "<p style='color:#64748B;font-size:12px;line-height:1.7'>"
            "En kritik geçiş: <b style='color:#38BDF8'>Genç Büyüme → Yüksek Büyüme</b><br><br>"
            "Gelir hızla büyürken EFK pozitife dönüyor, marjlar yükseliyor. "
            "Bu geçişi erken yakalamak en yüksek getiriyi sağlar.<br><br>"
            "<b style='color:#38BDF8'>🧒 Genç Büyüme</b> aşamasındaki hisseler "
            "tam bu eşikte — yakından takip edilmeli.</p></div>",
            unsafe_allow_html=True
        )

    st.markdown("<hr>", unsafe_allow_html=True)

    # 6 Tab
    tabs = st.tabs([f"{ASAMA_META[i][0]} {ASAMA_META[i][1]} ({len(dagilim[i])})" for i in range(1,7)])
    for asama, tab in zip(range(1, 7), tabs):
        with tab:
            em, lbl, renk, acik = ASAMA_META[asama]
            liste = sorted(dagilim[asama], key=lambda x: x.get("ns_buy") or 0, reverse=True)
            st.markdown(f"<div style='background:#0D1926;border-left:3px solid {renk};"
                        f"border-radius:6px;padding:10px 14px;margin-bottom:8px;"
                        f"font-size:11px;color:#64748B'>{acik}</div>", unsafe_allow_html=True)
            if not liste:
                st.info("Bu asamada hisse yok.")
                continue
            df = pd.DataFrame([{
                "Kod": r["kod"], "Sektor": r["sektor"][:20],
                "NS Buy%": round(r["ns_buy"],0) if r.get("ns_buy") is not None else None,
                "Marj%": round(r["marj"],0) if r.get("marj") is not None else None,
                "ROIC": round(r["roic"],0) if r.get("roic") is not None else None,
                "EFK ist%": round(r["efk_poz"],0) if r.get("efk_poz") is not None else None,
                "PD": fmt_milyon(r["pd_val"]),
            } for r in liste])
            st.dataframe(df, hide_index=True, use_container_width=True, height=min(40+len(liste)*35, 450))
            bt = st.columns(8)
            for i, r in enumerate(liste[:24]):
                with bt[i%8]:
                    if st.button(r["kod"], key=f"yd_{asama}_{r['kod']}"):
                        git_detay(r["kod"])

# ══════════════════════════════════════════════════════════════════════════════
# SAYFA 3: HiSSE TARAYICI
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🔍 Hisse Tarayici":
    st.markdown("""<div class='ph'>
    <div class='ph-badge' style='background:#071A0F;color:#4ADE80;border:1px solid #166534'>TARAYICI</div>
    <div class='ph-title'>Hisse Tarayici</div>
    <div class='ph-sub'>Damodaran metodolojisi ile tum BIST analizi</div>
    </div>""", unsafe_allow_html=True)

    if not quarters: bos(); st.stop()

    # Filtreler
    c1, c2, c3 = st.columns(3)
    with c1:
        asama_filtre = st.multiselect("Yaşam Aşaması", [1,2,3,4,5,6],
            format_func=lambda x: {1:"🌱 Baslangic",2:"🧒 Genc",3:"🚀 Yuksek Buy",
                                    4:"💪 Olgun Buy",5:"🏛️ Stabil",6:"📉 Dusus"}[x],
            default=[2,3])
    with c2:
        min_gm = st.slider("Min Güvenlik Marjı (%)", -100, 100, 0)
    with c3:
        risk_filtre = st.multiselect("Risk", ["DUSUK","ORTA","YUKSEK"], default=["DUSUK","ORTA"])

    with st.spinner("Analiz yapiliyor..."):
        sonuclar = []
        for kod, row in quarters[son_d].items():
            s = tam_analiz(kod, quarters, donems)
            if not s: continue
            gm = s['fiyat'].get('guvenlik_marji')
            if gm is None: continue
            a  = s['yasam_dongusu'].get('asama')
            rv = s['risk'].get('seviye')
            if asama_filtre and a not in asama_filtre: continue
            if gm < min_gm: continue
            if risk_filtre and rv not in risk_filtre: continue
            sonuclar.append(s)

    sonuclar.sort(key=lambda x: x['karar']['puan'], reverse=True)
    st.markdown(f"<p style='font-size:11px;color:#475569'>{len(sonuclar)} hisse bulundu</p>", unsafe_allow_html=True)

    if sonuclar:
        df = pd.DataFrame([{
            "Kod":     s['kod'],
            "Sektor":  s['sektor'][:20],
            "Asama":   f"{s['yasam_dongusu'].get('emoji','')} {s['yasam_dongusu'].get('label','')[:10]}",
            "PD":      s['fiyat'].get('pd_fmt','-'),
            "İcsel D": s['fiyat'].get('id_fmt','-'),
            "GM%":     s['fiyat'].get('guvenlik_marji'),
            "WACC%":   s['icsel'].get('wacc'),
            "3P":      s['uc_p'].get('toplam'),
            "Risk":    s['risk'].get('seviye'),
            "Karar":   s['karar']['karar'][:12],
            "Puan":    s['karar']['puan'],
        } for s in sonuclar])
        st.dataframe(df, hide_index=True, use_container_width=True, height=500)

        bt = st.columns(8)
        for i, s in enumerate(sonuclar[:24]):
            with bt[i%8]:
                if st.button(s['kod'], key=f"tr_{s['kod']}"):
                    git_detay(s['kod'])

        # Excel
        buf = io.BytesIO()
        df.to_excel(buf, index=False)
        st.download_button("⬇️ Excel İndir", data=buf.getvalue(),
                           file_name=f"Damodaran_{son_d}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ══════════════════════════════════════════════════════════════════════════════
# SAYFA 4: DETAY ANALiZi
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📊 Detay Analizi":
    st.markdown("""<div class='ph'>
    <div class='ph-badge' style='background:#0A1020;color:#A78BFA;border:1px solid #4C1D95'>DETAY</div>
    <div class='ph-title'>Tek Hisse Damodaran Analizi</div>
    <div class='ph-sub'>Yasam Dongusu · 3P Testi · DCF · Risk · Karar</div>
    </div>""", unsafe_allow_html=True)

    if not quarters: bos(); st.stop()

    # Hisse sec
    hisse_git = st.session_state.hisse_git
    if hisse_git: st.session_state.hisse_git = None
    tum_kodlar = sorted(quarters[son_d].keys())
    baslangic  = tum_kodlar.index(hisse_git) if hisse_git and hisse_git in tum_kodlar else 0
    secilen    = st.selectbox("Hisse Seç", tum_kodlar, index=baslangic)

    sonuc = tam_analiz(secilen, quarters, donems)
    if not sonuc:
        st.warning("Bu hisse analiz edilemiyor.")
        st.stop()

    row   = quarters[son_d].get(secilen, {})
    yd    = sonuc['yasam_dongusu']
    uc_p  = sonuc['uc_p']
    icsel = sonuc['icsel']
    fiyat = sonuc['fiyat']
    risk  = sonuc['risk']
    karar = sonuc['karar']

    # NiHAi KARAR KARTI
    st.markdown(
        f"<div style='background:{karar['renk']}22;border:2px solid {karar['renk']};"
        f"border-radius:14px;padding:16px 24px;margin-bottom:20px;text-align:center'>"
        f"<div style='font-size:28px;font-weight:900;color:{karar['renk']}'>{karar['karar']}</div>"
        f"<div style='font-size:12px;color:#94A3B8;margin-top:6px'>{karar['aciklama']}</div>"
        f"<div style='font-size:11px;color:{karar['renk']};margin-top:4px;font-weight:700'>Puan: {karar['puan']}/100</div>"
        f"</div>", unsafe_allow_html=True
    )

    # Terim aciklamalari
    ACIKLAMALAR = {
        "Yasam Dongusu": "Damadoran 6 Asama: Sirketin buyume evresini gosterir. Erken asama = yuksek potansiyel.",
        "3P Testi": "Possible + Plausible + Probable. Is modelinin gerceklesmesi ne kadar mumkun?",
        "DCF Degerleme": "Gelecek nakit akislarinin bugunki degeri. Guvenlik Marji = (Icsel Deger - PD) / Icsel Deger",
        "Risk": "Is riski + Buyume riski + Piyasa riski kombinasyonu. Dusuk = iyi.",
    }

    # 4 Kolon: YD + 3P + ICsel Deger + Risk
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        renk = yd.get('renk','#94A3B8')
        st.markdown(
            f"<div style='background:#0D1926;border:1px solid {renk};border-radius:10px;padding:14px'>"
            f"<div style='font-size:10px;color:#475569;text-transform:uppercase'>Yasam Dongusu</div>"
            f"<div style='font-size:22px;font-weight:800;color:{renk};margin:6px 0'>{yd.get('emoji','')} {yd.get('label','')}</div>"
            f"<div style='font-size:10px;color:#64748B'>{yd.get('aciklama','')[:80]}</div>"
            f"<div style='margin-top:8px;font-size:10px;color:#475569'>"
            f"NS Buy: <b style='color:#E2E8F0'>{yd.get('ns_buy','?')}%</b> &nbsp; "
            f"Marj: <b style='color:#E2E8F0'>{yd.get('marj_son','?')}%</b></div>"
            f"</div>", unsafe_allow_html=True
        )

    with c2:
        toplam_3p = uc_p.get('toplam', 0)
        renk_3p = "#4ADE80" if toplam_3p >= 70 else "#FCD34D" if toplam_3p >= 45 else "#F87171"
        st.markdown(
            f"<div style='background:#0D1926;border:1px solid {renk_3p};border-radius:10px;padding:14px'>"
            f"<div style='font-size:10px;color:#475569;text-transform:uppercase'>3P Testi</div>"
            f"<div style='font-size:22px;font-weight:800;color:{renk_3p};margin:6px 0'>{toplam_3p}/99</div>"
            f"<div style='font-size:10px;color:#64748B'>"
            f"Possible: {uc_p.get('possible',0)}/33<br>"
            f"Plausible: {uc_p.get('plausible',0)}/33<br>"
            f"Probable: {uc_p.get('probable',0)}/33</div>"
            f"</div>", unsafe_allow_html=True
        )

    with c3:
        gm = fiyat.get('guvenlik_marji')
        renk_gm = "#4ADE80" if gm and gm > 30 else "#FCD34D" if gm and gm > 0 else "#F87171"
        st.markdown(
            f"<div style='background:#0D1926;border:1px solid {renk_gm};border-radius:10px;padding:14px'>"
            f"<div style='font-size:10px;color:#475569;text-transform:uppercase'>DCF Degerleme</div>"
            f"<div style='font-size:22px;font-weight:800;color:{renk_gm};margin:6px 0'>"
            f"{'+' if gm and gm>0 else ''}{gm:.0f}% GM</div>"
            f"<div style='font-size:10px;color:#64748B'>"
            f"PD: {fiyat.get('pd_fmt','-')}<br>"
            f"İcsel: {fiyat.get('id_fmt','-')}<br>"
            f"WACC: %{icsel.get('wacc','?')}</div>"
            f"</div>", unsafe_allow_html=True
        )

    with c4:
        renk_r = risk.get('renk','#FCD34D')
        st.markdown(
            f"<div style='background:#0D1926;border:1px solid {renk_r};border-radius:10px;padding:14px'>"
            f"<div style='font-size:10px;color:#475569;text-transform:uppercase'>Risk</div>"
            f"<div style='font-size:22px;font-weight:800;color:{renk_r};margin:6px 0'>{risk.get('seviye','?')}</div>"
            f"<div style='font-size:10px;color:#64748B'>"
            f"İş: {risk.get('is_risk','?')}/100<br>"
            f"Büyüme: {risk.get('buy_risk','?')}/100<br>"
            f"Piyasa: {risk.get('piyasa_risk','?')}/100</div>"
            f"</div>", unsafe_allow_html=True
        )

    # Firsat Sinyalleri
    firsat = sonuc.get("firsat", {})
    if firsat.get("sinyaller"):
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown(
            f"<div style='background:#0A1020;border:1px solid {firsat['renk']};border-radius:10px;padding:12px 16px;margin-bottom:12px'>"
            f"<div style='color:{firsat['renk']};font-weight:800;font-size:13px'>🎯 Damodaran Firsat Sinyalleri: {firsat['puan']}/7 — {firsat['seviye']}</div></div>",
            unsafe_allow_html=True
        )
        f_cols = st.columns(min(len(firsat["sinyaller"]), 4))
        for i, s in enumerate(firsat["sinyaller"]):
            with f_cols[i % 4]:
                st.markdown(
                    f"<div style='background:#0D1926;border:1px solid #4ADE80;border-radius:8px;padding:10px;margin-bottom:8px'>"
                    f"<div style='font-size:18px'>{s['emoji']}</div>"
                    f"<div style='font-size:11px;font-weight:700;color:#4ADE80'>{s['baslik']}</div>"
                    f"<div style='font-size:9px;color:#475569;margin-top:4px'>{s['aciklama'][:80]}</div></div>",
                    unsafe_allow_html=True
                )

    # DCF Detay
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("<h3 style='color:#E2E8F0;font-size:14px;margin-bottom:8px'>📐 DCF Parametreleri</h3>", unsafe_allow_html=True)
    dc1, dc2, dc3, dc4, dc5 = st.columns(5)
    DCF_ACIK = {
        "Büyüme Oranı": "Sirketin yasam dongusune gore beklenen yillik gelir artisi",
        "WACC": "Agirlikli Ortalama Sermaye Maliyeti — kazancin en az bu kadar olmasi gerekir",
        "Süre": "Bu buyume oraninin kac yil surdurulebilecegi tahmini",
        "YY Oranı": "Kazancin yeni yatirima harcanan yuzdesI — yuksekse sirket buyumeye odakli",
        "Terminal PV": "Analiz suresinin otesindeki tum nakit akislarinin bugunki degeri",
    }
    for col, lbl, val in [
        (dc1, "Büyüme Oranı", f"%{icsel.get('buyume','?')}"),
        (dc2, "WACC",         f"%{icsel.get('wacc','?')}"),
        (dc3, "Süre",         f"{icsel.get('sure','?')} Yıl"),
        (dc4, "YY Oranı",     f"%{icsel.get('yy_orani','?')}"),
        (dc5, "Terminal PV",  fmt_milyon(icsel.get('pv_terminal'))),
    ]:
        col.markdown(
            f"<div style='background:#0D1926;border:1px solid #0F2040;border-radius:8px;padding:10px;text-align:center'>"
            f"<div style='font-size:9px;color:#475569;text-transform:uppercase'>{lbl}</div>"
            f"<div style='font-size:18px;font-weight:700;color:#E2E8F0'>{val}</div>"
            f"<div style='font-size:9px;color:#475569;margin-top:4px;line-height:1.3'>{DCF_ACIK.get(lbl,'')}</div>"
            f"</div>",
            unsafe_allow_html=True
        )

    # Temel finansal metrikler
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("<h3 style='color:#E2E8F0;font-size:14px;margin-bottom:8px'>📊 Temel Metrikler</h3>", unsafe_allow_html=True)
    METRIK_ACIK = {
        "ROIC":     "Yatirilan Sermaye Getirisi — yuksekse sirket sermayeyi verimli kullaniyor",
        "ROE":      "Ozsermaye Karliligi — hissedara donen kar yuzdesi",
        "Beta":     "Piyasaya gore risk — 1'den yuksekse piyasadan daha volatil",
        "PEG":      "FK / Buyume orani — 1'in altiysa buyumeye gore ucuz",
        "Piotroski":"Finansal saglik skoru 0-9 — 7 ve uzeri guclu",
        "FV/FAVÖK": "Firma Degeri / FAVOK — dusukse ucuz, yuksekse pahali",
        "PD/DD":    "Piyasa Degeri / Defter Degeri — 1'in altiysa varliklardan ucuz",
        "Cari Oran":"Dongu Varlik / KV Borc — 1'in ustuyse kisa vade odeme gucu var",
    }
    metrikler = [
        ("ROIC", safe_float(row.get('Roic','')), "%", "#4ADE80"),
        ("ROE", safe_float(row.get('Özsermaye Karlılığı (ROE) Yıllık (%)','')),"%" , "#38BDF8"),
        ("Beta", safe_float(row.get('Beta','')), "", "#A78BFA"),
        ("PEG", safe_float(row.get('Peg Oranı','')), "", "#FCD34D"),
        ("Piotroski", safe_float(row.get('Piotroski F Skor','')), "/9", "#4ADE80"),
        ("FV/FAVÖK", safe_float(row.get('Firma Değeri / FAVÖK','')), "x", "#94A3B8"),
        ("PD/DD", safe_float(row.get('Piyasa Değeri / Defter Değeri','')), "x", "#94A3B8"),
        ("Cari Oran", safe_float(row.get('Cari Oran','')), "x", "#38BDF8"),
    ]
    m_cols = st.columns(8)
    for (lbl, val, birim, renk), col in zip(metrikler, m_cols):
        val_str = f"{val:.1f}{birim}" if val is not None else "-"
        col.markdown(
            f"<div style='background:#0D1926;border:1px solid #0F2040;border-radius:8px;padding:10px;text-align:center'>"
            f"<div style='font-size:9px;color:#475569;text-transform:uppercase'>{lbl}</div>"
            f"<div style='font-size:16px;font-weight:700;color:{renk}'>{val_str}</div>"
            f"<div style='font-size:9px;color:#475569;margin-top:4px;line-height:1.3'>{METRIK_ACIK.get(lbl,'')}</div>"
            f"</div>",
            unsafe_allow_html=True
        )

# ══════════════════════════════════════════════════════════════════════════════
# SAYFA 5: AYARLAR
# ══════════════════════════════════════════════════════════════════════════════
elif page == "⚙️ Ayarlar":
    st.markdown("""<div class='ph'>
    <div class='ph-badge' style='background:#0D1926;color:#64748B;border:1px solid #1E3448'>AYARLAR</div>
    <div class='ph-title'>Sistem Ayarlari</div>
    </div>""", unsafe_allow_html=True)

    st.markdown("<h3 style='color:#E2E8F0;font-size:14px'>Türkiye DCF Parametreleri</h3>", unsafe_allow_html=True)
    st.markdown(
        f"<div style='background:#0D1926;border:1px solid #0F2040;border-radius:10px;padding:16px'>"
        f"<p style='color:#64748B;font-size:12px'>"
        f"• <b style='color:#E2E8F0'>Risksiz Faiz:</b> %{RISKSIZ_FAIZ*100:.0f} (TCMB politika faizi baz)<br>"
        f"• <b style='color:#E2E8F0'>Piyasa Risk Primi:</b> %{PIYASA_PRM*100:.0f}<br>"
        f"• <b style='color:#E2E8F0'>Ülke Risk Primi:</b> %{ULKE_RISK_PRM*100:.0f}<br>"
        f"• <b style='color:#E2E8F0'>Sürdürülebilir Büyüme:</b> %20 (enflasyon dahil uzun vade)<br><br>"
        f"WACC = Beta × ERP + Risksiz Faiz + Ülke Premi</p></div>",
        unsafe_allow_html=True
    )

    if st.button("🗑️ Tüm Veriyi Temizle"):
        st.session_state.quarters = {}
        st.session_state.donems   = []
        st.session_state.son_donem = None
        st.rerun()
