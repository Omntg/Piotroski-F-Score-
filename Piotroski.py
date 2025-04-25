import pandas as pd
import re
from xlsxwriter.utility import xl_col_to_name  # xl_col_to_name fonksiyonunu içe aktarıyoruz

# Sabit çeyrek sütun sırası (header), istege gore yeni ceyreklik donemler eklenip cikarilabilir
quarters_order = [
    "2021/6", "2021/9", "2021/12",
    "2022/3", "2022/6", "2022/9", "2022/12",
    "2023/3", "2023/6", "2023/9", "2023/12",
    "2024/3", "2024/6", "2024/9", "2024/12" 
]

# Belirtilen metrik için, itemDescTr eşleşmesiyle ilgili satırın,
# belirtilen çeyrek sütunlarındaki verilerini döndüren fonksiyon
def get_metric(data, metric_name, quarters):
    mask = data['itemDescTr'].astype(str).str.strip() == metric_name
    row = data[mask]
    if row.empty:
        raise ValueError(f"'{metric_name}' kalemi bulunamadı!")
    return row[quarters].iloc[0].to_dict()

# 'finansallar.xlsx' dosyasını açıyoruz
xls = pd.ExcelFile('finansallar.xlsx')
results = []  # Her hisse için hesaplanan sonuçların tutulacağı liste

for sheet in xls.sheet_names:
    print(f"Hesaplama yapılıyor: {sheet}")
    df = pd.read_excel(xls, sheet_name=sheet)
    df.columns = [str(col).strip() for col in df.columns]

    quarter_cols = [col for col in df.columns if "/" in col and re.search(r'\d{4}/\d{1,2}$', col)]
    try:
        quarter_cols = [col for col in quarter_cols if int(col.split('/')[0]) >= 2021]
    except Exception:
        continue
    quarter_cols = sorted(quarter_cols, key=lambda x: (int(x.split('/')[0]), int(x.split('/')[1])))
    if not quarter_cols:
        continue

    try:
        net_income = get_metric(df, "Dönem Net Kar/Zararı", quarter_cols)
    except ValueError:
        try:
            net_income = get_metric(df, "DÖNEM KARI (ZARARI)", quarter_cols)
        except ValueError:
            continue

    try:
        operating_cash = get_metric(df, "İşletme Faaliyetlerinden Kaynaklanan Net Nakit", quarter_cols)
        total_assets   = get_metric(df, "TOPLAM VARLIKLAR", quarter_cols)
        current_assets = get_metric(df, "Dönen Varlıklar", quarter_cols)
        short_liabilities = get_metric(df, "Kısa Vadeli Yükümlülükler", quarter_cols)
        sales          = get_metric(df, "Satış Gelirleri", quarter_cols)
        cost_of_sales  = get_metric(df, "Satışların Maliyeti (-)", quarter_cols)
    except ValueError:
        continue

    long_term_items = [
        "Finansal Borçlar", "Diğer Finansal Yükümlülükler", "Ticari Borçlar", "Diğer Borçlar",
        "Müşteri Söz.Doğan Yük.", "Finans Sektörü Faaliyetlerinden Borçlar", "Devlet Teşvik ve Yardımları",
        "Ertelenmiş Gelirler (Müşteri Söz.Doğan Yük. Dış.Kal.)", "Uzun vadeli karşılıklar",
        "Çalışanlara Sağlanan Faydalara İliş.Karş.", "Ertelenmiş Vergi Yükümlülüğü", "Diğer Uzun Vadeli Yükümlülükler"
    ]
    long_term_debt = {}
    for col in quarter_cols:
        total_debt = 0
        for item in long_term_items:
            mask = df['itemDescTr'].astype(str).str.strip() == item
            row = df[mask]
            if not row.empty:
                total_debt += row[col].values[0]
        long_term_debt[col] = total_debt

    roa = {}
    current_ratio = {}
    lt_debt_ratio = {}
    gross_margin = {}
    asset_turnover = {}
    for col in quarter_cols:
        ta = total_assets[col]
        ni = net_income[col]
        roa[col] = ni / ta if ta != 0 else None
        st = short_liabilities[col]
        ca = current_assets[col]
        current_ratio[col] = ca / st if st != 0 else None
        debt = long_term_debt[col]
        lt_debt_ratio[col] = debt / ta if ta != 0 else None
        s = sales[col]
        c = cost_of_sales[col]
        gross_margin[col] = (s - c) / s if s != 0 else None
        asset_turnover[col] = s / ta if ta != 0 else None

    scores_dict = {}
    prev_col = None
    for col in quarter_cols:
        if prev_col is None:
            scores_dict[col] = None
            prev_col = col
            continue
        score = 0
        if net_income[col] > 0:
            score += 1
        if operating_cash[col] > 0:
            score += 1
        if roa[col] is not None and roa[prev_col] is not None and roa[col] > roa[prev_col]:
            score += 1
        if operating_cash[col] > net_income[col]:
            score += 1
        if lt_debt_ratio[col] is not None and lt_debt_ratio[prev_col] is not None and lt_debt_ratio[col] < lt_debt_ratio[prev_col]:
            score += 1
        if current_ratio[col] is not None and current_ratio[prev_col] is not None and current_ratio[col] > current_ratio[prev_col]:
            score += 1
        if gross_margin[col] is not None and gross_margin[prev_col] is not None and gross_margin[col] > gross_margin[prev_col]:
            score += 1
        if asset_turnover[col] is not None and asset_turnover[prev_col] is not None and asset_turnover[col] > asset_turnover[prev_col]:
            score += 1
        scores_dict[col] = score
        prev_col = col

    row_data = {'Hisse': sheet}
    for q in quarters_order:
        row_data[q] = scores_dict.get(q, None)

    valid_scores = [scores_dict[q] for q in quarters_order if scores_dict.get(q) is not None]
    if valid_scores:
        import numpy as np
        avg_score = np.mean(valid_scores)
        exp_avg_score = pd.Series(valid_scores).ewm(alpha=0.5).mean().iloc[-1]
        std_dev = np.std(valid_scores, ddof=0)
        rel_std = (std_dev / avg_score) * 100 if avg_score != 0 else None
    else:
        avg_score = None
        exp_avg_score = None
        rel_std = None

    row_data['Ortalama'] = avg_score
    row_data['Üstel Ortalama'] = exp_avg_score
    row_data['Sapma'] = rel_std

    results.append(row_data)

final_df = pd.DataFrame(results)
final_df = final_df[['Hisse'] + quarters_order + ['Ortalama', 'Üstel Ort.', 'Sapma']]

writer = pd.ExcelWriter('piotroski_skorlari2.xlsx', engine='xlsxwriter')
final_df.to_excel(writer, sheet_name='Skorlar', index=False)

workbook  = writer.book
worksheet = writer.sheets['Skorlar']

# 1) Tüm sütunlar için genişlik = 9
worksheet.set_column(0, final_df.shape[1]-1, 9)

# 3) A kolonunu (ilk kolon) header hariç renk, font rengi beyaz
soft_a_format = workbook.add_format({'bg_color': '#110082', 'font_color': '#FFFFFF'})
num_data_rows = final_df.shape[0]
worksheet.conditional_format(f'A2:A{num_data_rows+1}', {'type': 'no_errors', 'format': soft_a_format})

# 4) İlk satır (header) hücreleri rengi, font rengi beyaz
header_format = workbook.add_format({'bg_color': '#2F005B', 'font_color': '#FFFFFF'})
for col_num in range(final_df.shape[1]):
    worksheet.write(0, col_num, final_df.columns[col_num], header_format)

# Sayısal formatlamalar
# - Çeyrek skorlar 
int_format = workbook.add_format({'num_format': '0'})
for idx, col_name in enumerate(quarters_order, start=1):
    worksheet.set_column(idx, idx, 9, int_format)

# - Ortalama, Üstel Ortalama ve Sapma için virgülden sonra 1 basamak
dec_format = workbook.add_format({'num_format': '0.0'})
for col_name in ['Ortalama', 'Üstel Ortalama', 'Sapma']:
    idx = final_df.columns.get_loc(col_name)
    worksheet.set_column(idx, idx, 9, dec_format)

# Koşullu formatlama
# - Çeyrek skorlar ve Ortalama, Üstel Ortalama: Kırmızı-Sarı-Yeşil
for idx, col_name in enumerate(quarters_order + ['Ortalama', 'Üstel Ortalama'], start=1):
    col_letter = xl_col_to_name(idx)
    cell_range = f"{col_letter}2:{col_letter}{num_data_rows+1}"
    worksheet.conditional_format(cell_range, {
        'type':      '3_color_scale',
        'min_color': '#FF0000',
        'mid_color': '#FFFF00',
        'max_color': '#00FF00'
    })

# - Sapma için ters gradyan: Yeşil-Sarı-Kırmızı (düşükten yükseğe)
sapma_idx = final_df.columns.get_loc('Sapma')
sapma_letter = xl_col_to_name(sapma_idx)
cell_sapma = f"{sapma_letter}2:{sapma_letter}{num_data_rows+1}"
worksheet.conditional_format(cell_sapma, {
    'type':      '3_color_scale',
    'min_color': '#00FF00',
    'mid_color': '#FFFF00',
    'max_color': '#FF0000'
})

writer.close()
print("Piotroski skorları 'piotroski_skorlari2.xlsx' dosyasına kaydedildi.")
