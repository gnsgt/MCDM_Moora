import pandas as pd
import numpy as np

    #Excel dosyasını oku ve DataFrame'e dönustur
df_excel = pd.read_excel("C:/Users/gokha/OneDrive/Masaüstü/Tez/3.1-Moora.xlsx")
maliyet = df_excel.iloc[1:2].copy()
agirlik = df_excel.iloc[:1].copy()
df = df_excel.iloc[2:]  

    #Kac satır kaç sutun?
satir_sayisi = df.shape[0] 
sutun_sayisi = df.shape[1]


    #Kareler Toplamları ve Karekok
karelertoplamlari = []
karekok = []
for j in range(0,sutun_sayisi):
    karetop = 0
    for i in range(0,satir_sayisi):
        karetop = karetop + (df.iloc[i,j]**2)
    karelertoplamlari.append(karetop)
    karekok.append(karetop**0.5)


    #Normalize islemi
normalize_df = df.copy()
normalize_rj_min = []
normalize_rj_max = []
normalize_rj = []
moora_oran = []  
for i in range(sutun_sayisi):
    normalize_df.iloc[:, i] = df.iloc[:, i] / karekok[i]
    
    if maliyet.iloc[0, i] == "maliyet":
        normalize_rj_min.append(min(normalize_df.iloc[:, i]))
        normalize_rj.append(min(normalize_df.iloc[:, i]))
    else:
        normalize_rj_max.append(max(normalize_df.iloc[:, i]))
        normalize_rj.append(max(normalize_df.iloc[:, i]))

    #Moora-Oran Hesaplama
for j in range(satir_sayisi):
    moora_oran_toplam = 0  
    for i in range(sutun_sayisi):
        if maliyet.iloc[0, i] == "maliyet":
            moora_oran_toplam -= normalize_df.iloc[j, i]
        else:
            moora_oran_toplam += normalize_df.iloc[j, i]
    moora_oran.append(moora_oran_toplam)
    
    #MOORA-Oran Yazdirma
for i in range(satir_sayisi):
    max_deger = max(moora_oran)
    max_index = moora_oran.index(max_deger)
    print(f"MOORA-ORAN Değer: {max_deger}, İndex: {max_index + 1}")
    moora_oran[max_index] = -99999



    #Agirlikli Normalize Islemi
anormalize_df = normalize_df.copy()  # normalize_df verilerini kopyala
for i in range(sutun_sayisi):
    anormalize_df.iloc[:, i] = normalize_df.iloc[:, i] * agirlik.iloc[0, i]  


    #Moora-Onem Hesaplama
moora_onem = [] 
print("\n")
for j in range(satir_sayisi):
    moora_onem_toplam = 0  
    for i in range(sutun_sayisi):
        if maliyet.iloc[0, i] == "maliyet":
            moora_onem_toplam -= anormalize_df.iloc[j, i]
        else:
            moora_onem_toplam += anormalize_df.iloc[j, i]
    moora_onem.append(moora_onem_toplam)
    
    #MOORA-Onem Yazdirma
for i in range(satir_sayisi):
    max_deger = max(moora_onem)
    max_index = moora_onem.index(max_deger)
    print(f"MOORA-ÖNEM Değer: {max_deger}, İndex: {max_index + 1}")
    moora_onem[max_index] = -99999


    #Mutlak dj
mutlak_dj = normalize_df.copy()
for j in range(sutun_sayisi): 
    for i in range(satir_sayisi): 
        mutlak_dj.iloc[i,j] = abs(normalize_rj[j] - normalize_df.iloc[i,j])

    #MOORA Referans
moora_referans = []
for i in range(satir_sayisi):
    moora_referans.append(max(mutlak_dj.iloc[i,:]))

    #MOORA Referans Yazdirma
print("\n")
for i in range(satir_sayisi):
    min_deger = min(moora_referans)
    min_index = moora_referans.index(min_deger)
    print(f"MOORA-Referans Değer: {min_deger}, İndex: {min_index + 1}")
    moora_referans[min_index] = 99999


    #Moora tam carpim
ai_star = []
bi_star = []
for j in range(satir_sayisi):
    ai = 1
    bi = 1
    for i in range(sutun_sayisi):
        if maliyet.iloc[0,i] == "maliyet":
            bi = df.iloc[j,i] * bi
        else:
            ai = df.iloc[j,i] * ai
    ai_star.append(ai)
    bi_star.append(bi)
    
u = []
for i in range(len(ai_star)):
    u.append(ai_star[i] / bi_star[i])

print("\n")
for i in range(satir_sayisi):
    max_deger = max(u)
    max_index = u.index(max_deger)
    print(f"MOORA-Tam Çarpım Değer: {max_deger}, İndex: {max_index + 1}")
    u[max_index] = -99999



            

            







