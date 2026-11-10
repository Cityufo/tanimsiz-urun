#!/usr/bin/env python3
"""MASTER Tanımsız Ürün raporundan data.json üretir.

Kullanım: python3 tools/guncelle.py RAPOR.xlsx data.json [GG.AA.YYYY]

Sütunlar sırayla değil başlık adıyla bulunur; rapordaki sütun düzeni değişse de çalışır.
Gönderici, alıcı ve sevkiyat numarası yayınlanmaz.
"""
import sys, json, re, datetime, os
import openpyxl

SRC, OUT = sys.argv[1], sys.argv[2]
UPDATED = sys.argv[3] if len(sys.argv) > 3 else datetime.date.today().strftime('%d.%m.%Y')

def norm(h):
    return re.sub(r'\s+', ' ', str(h or '')).strip().upper()

def colmap(ws):
    hdr = {norm(c.value): i for i, c in enumerate(next(ws.iter_rows(min_row=1, max_row=1)))}
    def find(*names, required=True):
        for n in names:
            for h, i in hdr.items():
                if h == n or h.startswith(n):
                    return i
        if required:
            raise SystemExit(f'{ws.title}: "{names[0]}" başlığı bulunamadı')
        return None
    return find

def tarih(v):
    if isinstance(v, datetime.datetime): v = v.date()
    if isinstance(v, datetime.date): return v.strftime('%d.%m.%Y')
    return str(v or '').strip()

def metin(v):
    return '' if v is None else re.sub(r'\s+', ' ', str(v)).strip()

wb = openpyxl.load_workbook(SRC, data_only=True)
rows = []
for rapor in ('UFO', 'UTO', 'URO'):
    ws = wb[f'{rapor} TAKİP']
    f = colmap(ws)
    c_t, c_d, c_a = f('KAYIT TARİHİ'), f('DOSYA NUMARASI'), f('DİJİTAL ARŞİV')
    c_s, c_st = f('SONUÇ'), f('SONUÇLANDIRMA TARİHİ')
    c_ds, c_ac, c_tp = f('ARAŞTIRMA DURUMU'), f('AÇIKLAMA'), f('TESPİT YÖNTEMİ')
    if rapor == 'UTO':
        c_u, c_x, c_g = f('LASTİK EBATI'), f('LASTİK ADI'), f('LASTİK SINIFI')
    else:
        c_u, c_x, c_g = f('ÜRÜN GRUBU'), f('ÜRÜN İÇERİK'), None
    c_k = f('ÜRÜN GELİŞ KAYNAĞI') if rapor == 'UFO' else None
    c_y = f('BÖLGE') if rapor == 'URO' else None
    for r in ws.iter_rows(min_row=2):
        d = metin(r[c_d].value)
        if not d or r[c_t].value in (None, ''):
            continue
        s = metin(r[c_s].value)
        if s == 'ARAŞTIRILIYOR': s = ''
        g = metin(r[c_g].value) if c_g is not None else ''
        link = r[c_a].hyperlink.target if r[c_a].hyperlink is not None and r[c_a].hyperlink.target else ''
        rows.append({
            'd': d, 'r': rapor, 't': tarih(r[c_t].value),
            'u': metin(r[c_u].value), 'x': metin(r[c_x].value),
            'g': f'Lastik · {g}' if g else '',
            'y': (metin(r[c_y].value) or 'Bölge') if c_y is not None else 'Merkez depo',
            's': s, 'st': tarih(r[c_st].value) if s else '',
            'ds': metin(r[c_ds].value), 'ac': metin(r[c_ac].value), 'tp': metin(r[c_tp].value),
            'k': metin(r[c_k].value) if c_k is not None else '',
            'a': link,
        })

# güvenlik: yanlış ya da yarım dosya okunduysa mevcut veriyi ezme
if os.path.exists(OUT):
    eski = json.load(open(OUT, encoding='utf-8'))['rows']
    if len(rows) < 0.9 * len(eski):
        raise SystemExit(f'Okunan kayıt sayısı ({len(rows)}) mevcut veriden ({len(eski)}) çok az; data.json değiştirilmedi.')
json.dump({'updated': UPDATED, 'rows': rows}, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
print(f'{len(rows)} kayıt yazıldı ({UPDATED})')
