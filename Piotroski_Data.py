from isyatirimhisse import Financials
import pandas as pd

# Hisse sembollerinin listesi, dilediginiz hisseyi "" icerisinde belirtebilirsiniz, hisseler arasi virgul koyunuz.
symbols = [
    "AGESA", "AGHOL", "AKBNK", "AKFGY", "AKGRT", "AKMGY", "AKSGY",
    "ALBRK", "ALCTL", "ANHYT", "ANSGR", "ARDYZ", "ARTMS", "ASELS", "ASTOR", "ATATP", "ATLAS", "AVGYO", "AVHOL", "AYEN", "AYGAZ", "AZTEK",
    "BANVT", "BASGZ", "BEGYO", "BFREN", "BIGCH", "BIMAS", "BRISA", "BRLSM", "BURCE", "BURVA",
    "CCOLA", "CIMSA", "CLEBI", "CRDFA", "CUSAN", "CVKMD", "DARDL", "DERIM", "DESA", "DESPC", "DGATE", "DNISI", "DOCO", "DOHOL", "DYOBY", "DZGYO",
    "EBEBK", "EDIP", "EFORC", "EGEEN", "EGGUB", "EGPRO", "EKGYO", "ELITE", "EMKEL", "ENERY", "ENJSA", "ENKAI", "EREGL", "ETILR", "EUPWR", "EURWN", "EYGYO",
    "FADE", "FMIZP", "FONET", "FORTE", "FRIGO", "FROTO", "GARAN", "GARFA", "GENIL", "GEREL", "GESAN", "GLCVY",
    "GLYHO", "GMTAS", "GOKNR", "GRSEL", "GSDDE", "GSDHO", "GUBRF", "GZNMI", "HALKB", "HTTBT", "HUNER", "IMASM", "INGRM", "IPEKE", "ISCTR", "ISDMR",
    "ISFIN", "ISGSY", "ISGYO", "ISKPL", "ISMEN", "ISYAT", "KATMR", "KERVT", "KLKIM", "KLSYN", "KONTR", "KOPOL", "KOZAA", "KRPLS", "KRSTL", "KUTPO",
    "KUYAS", "LIDFA", "LINK", "LKMNH", "LOGO", "LRSHO", "MACKO", "MAKTK", "MARTI", "MAVI", "MERIT", "MERKO", "MGROS", "MIATK", "MPARK", "MRGYO", "MTRKS",
    "NTGAZ", "NTHOL", "NUHCM", "OBASE", "ORGE", "OYAKC", "OZKGY", "PAGYO", "PAPIL", "PASEU", "PCILT", "PENTA", "PETUN", "PGSUS", "PINSU", "PLTUR",
    "PRKME", "RNPOL", "RYGYO", "RYSAS", "SAFKR", "SAHOL", "SANEL", "SEKFK", "SELEC", "SILVR", "SUNTK", "SUWEN", "TAVHL", "TBORG", "TCELL", "TEZOL",
    "TGSAS", "THYAO", "TKFEN", "TNZTP", "TRGYO", "TRILC", "TSKB", "TTKOM", "TUKAS", "TURSG", "ULKER", "ULUFA", "VAKBN", "VAKFN", "VERUS", "VKGYO",
    "YEOTK", "YKBNK", "ZRGYO"
]

# Financials sınıfından nesne oluşturuluyor
financials = Financials()

# 2021'den günümüze kadar veriler çekiliyor (end_year belirtilmezse güncel yıl kullanılır)
data = financials.get_data(
    symbols=symbols,
    start_year='2021',
    exchange='TRY'
)

# Tüm hisse verilerini tek bir Excel dosyasında, her hisseye ait veriyi ayrı sayfada saklayalım.
with pd.ExcelWriter('finansallar.xlsx') as writer:
    for symbol, df in data.items():
         df.to_excel(writer, sheet_name=symbol, index=False)

print("Finansal veriler 'finansallar.xlsx' dosyasına kaydedildi.")
