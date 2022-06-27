import math
import openpyxl
import pandas as pd

#Dikdörtgen Kirişler İçin Donatı Tasarımı

#Kesit Özellikleri
d=35#mm Paspayı
slab_thickness=150 #mm Döşeme Kalınlığı İsteğe Bağlı Kiriş Boyut Kontrolü İçin Gerekli

#Malzeme Özellikleri
fck=24#MPa
fyk=420 #MPa

###################################################################################
dosyaadi="Kiris_Donati_Hesabi.xlsx" #.xlsx formatında olmalı
###################################################################################

#CALCULATİON#

Gama_C = 1.5
Gama_S = 1.15

TasimaGucu=[]
Span_Area=[]        #Açıklık İçin Seçilen Donatıların Alanları
Span_R=[]           #Açıklık İçin Seçilen Donatılar
basinc_donatisi=[]  #Basınç Donatıları
Cal_Support_Area=[] #Hesaplanan Mesnet Donatı Alanları

def Dimensions_Control(width,height,slab_thickness): #TBDY2018
    if width<250:
        print("Kiriş Genişliği 250mm küçük olamaz!!")
    elif 300<=height and height<=3.5*width and height>=3*slab_thickness:
        print(15*"*","Kiriş Boyutları Yeterlidir.",15*"*")
    else:
        print("Boyutlar Yönetmelik Şartına Uymuyor!!")

def Reinforcement_Area(As):
    global Alanc,n1,d1
    Donati_Alanlari = []
    for d in range(12,16, 2):  # İstediğiniz Donatı Çapını Burdan ve;
        for n in range(2, 20):
            Alan_ = n * (math.pi * d ** 2) / 4
            Donati_Alanlari.append(Alan_)
    Alan_Dict = {}
    for item in Donati_Alanlari:
        if item>=As:
            Alan_Dict[item] = (item-As)
    #print(Alan_Dict)
    for key, value in Alan_Dict.items():
        if value == min(Alan_Dict.values()):
            Alan = (key * 4) / math.pi
            for d in range(12,16, 2):   # burdan ayarlamalısınız.....
                for n in range(2, 20):
                    Adet = n * d ** 2
                    if round(Adet,2) == round(Alan,2):
                        Alanc=key
                        n1=n
                        d1=d
                        break
    return Alanc,n1,d1
def Calculate_Beam_Reinforcement(Moment,isSupportMoment,width,height,d,fck,fyk):
    global As,a,Mr,M_tasima,Kiris_Tasima
    d_Height = height - d  # Faydalı Yükseklik
    fcd=fck/Gama_C              #Tasarım Basınç Dayanımı
    fyd=fyk/Gama_S              #Donatının Tasarım Akma Dayanımı
    fctk=0.35*math.sqrt(fck)    #Karakteristik çekme dayanımı
    fctd=fctk/Gama_C            #Tasarım çekme dayanımı

    def calculate_a (Moment):  #Beton Basınç Bloğu Derinlik Hesabı
        global a
        # Fs*z=Fc*z
        #a= d_Height-((d_Height**2)-((2*Moment*1e6)/(0.85*fcd*width)))**0.5

        x=(0.85*fcd*width)/2
        y=0.85*fcd*width*d_Height
        delta=y**2-(4*x*Moment*10**6)
        if delta>0:
            a1=(y+math.sqrt(delta))/(2*x)
            a2=(y-math.sqrt(delta))/(2*x)
            if a1>height:
                a=a2
            else:
                a=a1
        else:
            a1=y/(2*x)
            a=a1

        return a

    def Kiris_Tasima(Alan):
        global M_tasima
        a = (Alan * fyd) / (0.85 * fcd * width)
        M_tasima = (Alan * fyd * (d_Height - (a / 2))) * 10 ** -6
        return M_tasima
    calculate_a(Moment)

    #Fc=Fs
    Fc=0.85*fcd*width*a
    As=Fc/fyd

    if fck<25:
        k1=0.85
    else:
        k1=0.85-(fck-25)*0.006

    q = As / (width * d_Height) #Donatı Oranı
    qb = 0.85 * k1 * (fcd / fyd) * (600 / (600 + fyd))
    qmax = 0.85 * qb
    qmin = 0.80 * (fctd / fyd)
    qsehim = 0.235 * (fcd / fyd)
    qdep = 0.6 * qb

    def reinforcement_ratio(ro,Mr,q):
        global As
        if q <= qmin:
            As = qmin * width * d_Height
            print("Asmin:{} mm2".format(round(As, 2)))
            Reinforcement_Area(As)
            basinc_donatisi.append(0)
            print("Seçilen Donatı: {}Ø{}".format(n1, d1), "--", "Alanı: {} mm2".format(round(Alanc, 3)))
        elif q < ro:
            As = q * width * d_Height
            print("Çekme Donatı Alanı:{} mm2".format(round(As, 2)))
            Reinforcement_Area(As)
            basinc_donatisi.append(0)
            print("Seçilen Donatı: {}Ø{}".format(n1, d1), "--", "Alanı: {} mm2".format(round(Alanc, 3)))
        elif q > ro:
            if ro==qsehim:
                print("Sehim Şartı Sağlanmıyor, Çift Donatılı Hesap")
            else:
                print("Deprem Şartı Sağlanmıyor, Çift Donatılı Hesap")
            Aso=ro*width*d_Height
            a=(Aso*fyd)/(0.85*fcd*width)
            Mro=(Aso*fyd*(d_Height-(a/2)))/10**6
            Delta_M=Mr-Mro
            As1_=(Delta_M*10**6)/(fyd*(d_Height-d))  #Basınç Donatısı Aktığı Kabul Edilmiştir.
            if As1_<=226:  #2fi12 minimum donatı alanı
                As1_=226
            print(10*"*","\nBasınç Donatı Alanı: {} mm2".format(round(As1_,3)))
            Reinforcement_Area(As1_)
            basinc_donatisi.append(Alanc)
            print("Seçilen Donatı: {}Ø{}".format(n1, d1), "--", "Alanı: {} mm2".format(round(Alanc, 3)))
            print(10*"*")
            As=Aso+Alanc
            print("Çekme Donatı Alanı: {} mm2".format(round(As,3)))
            Reinforcement_Area(As)
            print("Seçilen Donatı: {}Ø{}".format(n1, d1), "--", "Alanı: {} mm2".format(round(Alanc, 3)))
        else:
            print("q > qmax,Maksimum Donatı Oranı Aşılıyor. Kesit Büyütülmeli")
        return As

    if isSupportMoment:
        reinforcement_ratio(qdep,Moment,q)
        Cal_Support_Area.append(As)
    else:
        reinforcement_ratio(qsehim,Moment,q)
        Span_Area.append(Alanc)
        Span_R.append(str("{}Ø{}".format(n1, d1)))

    return As

data=pd.read_excel(dosyaadi,engine="openpyxl")
#engine="openpyxl",columns=["Kiris","Section","Length","M_span","M_support_left","M_support_right","V"])
#names=["Kiris","Section","Length","M_span","M_support_left","M_support_right","V"])
Width=[]
Height=[]

Width_=[i.split('x', 1)[0] for i in data.Section[::]]
Width__=([i.split('B', 1)[1] for i in Width_])
Height_=[i.split('x', 1)[1] for i in data.Section[::]]

M_span=[] #Açıklık Momenti
M_support_left=[]  #Sol Mesnete Gelen Moment
M_support_right=[] #Sağ Mesnete Gelen Moment
Length=[] #Kiriş Boyları
Vdy=[]
VG_Q_DVE=[]
Vd=[]
Vdy_2=[]

for item in range(0,len(data.M_span)):
    M_support_left.append(abs(data.M_support_left[item]))
    M_support_right.append(abs(data.M_support_right[item]))
    M_span.append(abs(data.M_span[item]))
    Length.append(data.Length[item])
    Vdy.append((data.Vdy[item]))
    VG_Q_DVE.append((data.VE[item]))
    Vd.append(data.Vd[item])
    Vdy_2.append(data.Vdy2[item])

for i in range(0,len(M_span)):
    a=int(Width__[i])
    b=int(Height_[i])
    Width.append(10*a)
    Height.append(10*b)

Dimensions_Control(Width[0], Height[0],slab_thickness)

for item in range(0,len(M_span)):
    print(50 * "-")
    print(5*"*"," {} Kirişi".format(data.Kiris[item]),5*"*")
    print(2* "=", "Açıklık İçin Donatı Hesabı",2* "=")
    Calculate_Beam_Reinforcement(M_span[item], False, Width[item], Height[item],d, fck, fyk)
    print("\n",2* "=", "Sol Mesnet İçin Donatı Hesabı", 2* "=")
    Calculate_Beam_Reinforcement(M_support_left[item], True, Width[item], Height[item],d,fck, fyk)
    print("\n", 2 * "=", "Sağ Mesnet İçin Donatı Hesabı", 2 * "=")
    Calculate_Beam_Reinforcement(M_support_right[item], True, Width[item], Height[item],d,fck, fyk)

Cal_Support_Area_R=[] #Hesaplanan Sağ Mesnet Donatı Alanı
Cal_Support_Area_L=[] #Hesaplanan Sol Mesnet Donatı Alanı

Span_Upper=[]      #Açıklık Üst Donatısı Max Mesnetin 1/4'ü

Span_Upper_Area=[] #Açıklık ÜSt Donatı Seçilen Donatının Alanları
Span_Upper_R=[]    #Açıklık Üst Donatısı Seçilen Donatılar
Span_Basinc=[]     #Açıklık Basınc Donatısı

Left_Basinc=[]     #Sol Mesnet Basınç Donatısı
Right_Basinc=[]    #Sağ Mesnet Basınç Donatısı

Support_Left=[]    #Seçilen Sol Üst Mesnet Donatı Alanı
Support_Left_R=[]  #Seçilen Sol Üst Mesnet Donatıları
Support_Right=[]   #Seçilen Sağ Üst Mesnet Donatı Alanı
Support_Right_R=[] #Seçilen Sağ Üst Mesnet Donatıları

Support_Under_Left=[]  #Seçilen Sol Alt Mesnet Donatı Alanı
Support_Under_Left_R=[] #Seçilen Sol Alt Mesnet Donatıları
Support_Under_Right=[] #Seçilen Sağ Alt Mesnet Donatı Alanı
Support_Under_Right_R=[] #Seçilen Sağ Alt Mesnet Donatıları

Support_Under_Left_Bearing=[] #Sol Alt Mesnet Donatılarının Taşıma Kapasitesi
Support_Upper_Left_Bearing=[] #Sol Üst Mesnet Donatılarının Taşıma Kapasitesi
Support_Under_Right_Bearing=[] #Sağ Alt Mesnet Donatılarının Taşıma Kapasitesi
Support_Upper_Right_Bearing=[] #Sağ Üst Mesnet Donatılarının Taşıma Kapasitesi

for i in range(0,2*len(M_span),2):
    Cal_Support_Area_L.append(round(Cal_Support_Area[i],2))
for i in range(1, 2 * len(M_span), 2):
    Cal_Support_Area_R.append(round(Cal_Support_Area[i],2))
for i in range(0,len(basinc_donatisi),3):
    Span_Basinc.append(basinc_donatisi[i])
for i in range(1,len(basinc_donatisi),3):
    Left_Basinc.append(basinc_donatisi[i])
for i in range(2,len(basinc_donatisi),3):
    Right_Basinc.append(basinc_donatisi[i])

for i in range(0,len(M_span)):
    global Alanc,Kiris_Tasima,M_tasima,n1,d1
    Span_Upper_=(max(Cal_Support_Area_L[i]/4,Cal_Support_Area_R[i]/4)) #Açıklık Üst Donatı: Sağ ve Sol Mesnet Donatılarının 1/4'ü
    if Span_Upper_>Span_Basinc[i]:
        Span_Upper.append(Span_Upper_)
    else:
        Span_Upper.append(Span_Basinc[i])
    Reinforcement_Area(Span_Upper[i])
    Span_Upper_Area.append(round(Alanc,3))  #Açıklık Üst Donatı Alanı
    Span_Upper_R.append(str("{}Ø{}".format(n1, d1))) #Açıklık Üst Donatı Seçimi
    Kiris_Tasima(Alanc + Span_Area[i])

    Support_Left_ =(Cal_Support_Area_L[i] - Span_Upper_Area[i])  #Sol Üst Mesnet Donatı Hesabı
    if Support_Left_<0:
        Support_Left.append("None")
        Support_Left_R.append("None")
        Kiris_Tasima(Span_Upper_Area[i])
        Support_Upper_Left_Bearing.append(M_tasima)
        Kiris_Tasima(Span_Upper_Area[i]+Span_Area[i])

    else:
        Reinforcement_Area(Support_Left_)
        Support_Left.append(round(Alanc,3))
        Support_Left_R.append(str("{}Ø{}".format(n1, d1)))
        Kiris_Tasima(Alanc + Span_Upper_Area[i])
        Support_Upper_Left_Bearing.append(M_tasima)
        Kiris_Tasima(Alanc+Span_Upper_Area[i] + Span_Area[i])


    Support_Right_ = (Cal_Support_Area_R[i] - Span_Upper_Area[i])  #Sağ Üst Mesnet Donatı Hesabı
    if Support_Right_<0:
        Support_Right.append("None")
        Support_Right_R.append("None")
        Kiris_Tasima(Span_Upper_Area[i])
        Support_Upper_Right_Bearing.append(M_tasima)
    else:
        Reinforcement_Area(Support_Right_)
        Support_Right.append(round(Alanc,3))
        Support_Right_R.append(str("{}Ø{}".format(n1, d1)))
        Kiris_Tasima(Alanc + Span_Upper_Area[i])
        Support_Upper_Right_Bearing.append(M_tasima)

    if Left_Basinc[i]>(Cal_Support_Area_L[i]/2):  #Sol Alt Mesnet Sol Üst Mesnetin 1/2'si
        if Left_Basinc[i]>Span_Area[i]:
            Support_Under_Left_=(Left_Basinc[i]-Span_Area[i])
            Reinforcement_Area(Support_Under_Left_)
            Support_Under_Left.append(round(Alanc,3))
            Support_Under_Left_R.append(str("{}Ø{}".format(n1, d1)))
            Kiris_Tasima(Alanc+Span_Area[i])
            Support_Under_Left_Bearing.append(M_tasima)
        else:
            Support_Under_Left.append("None")
            Support_Under_Left_R.append("None")
            Kiris_Tasima(Span_Area[i])
            Support_Under_Left_Bearing.append(M_tasima)
    else:
        if (Cal_Support_Area_L[i]/2)>Span_Area[i]:
            Support_Under_Left_ =(Cal_Support_Area_L[i]-Span_Area[i])
            Reinforcement_Area(Support_Under_Left_)
            Support_Under_Left.append(round(Alanc, 3))
            Support_Under_Left_R.append(str("{}Ø{}".format(n1, d1)))
            Kiris_Tasima(Alanc + Span_Area[i])
            Support_Under_Left_Bearing.append(M_tasima)
        else:
            Support_Under_Left.append("None")
            Support_Under_Left_R.append("None")
            Kiris_Tasima(Span_Area[i])
            Support_Under_Left_Bearing.append(M_tasima)


    if Right_Basinc[i] > (Cal_Support_Area_R[i]/2):
        if Right_Basinc[i] > Span_Area[i]:
            Support_Under_Right_ = (Right_Basinc[i]-Span_Area[i])
            Reinforcement_Area(Support_Under_Right_)
            Support_Under_Right.append(round(Alanc, 3))
            Support_Under_Right_R.append(str("{}Ø{}".format(n1, d1)))
            Kiris_Tasima(Alanc + Span_Area[i])
            Support_Under_Right_Bearing.append(M_tasima)
        else:
            Support_Under_Right.append("None")
            Support_Under_Right_R.append("None")
            Kiris_Tasima(Span_Area[i])
            Support_Under_Right_Bearing.append(M_tasima)
    else:
        if (Cal_Support_Area_R[i]/2) > Span_Area[i]:
            Support_Under_Right_ = (Cal_Support_Area_R[i]-Span_Area[i])
            Reinforcement_Area(Support_Under_Right_)
            Support_Under_Right.append(round(Alanc, 3))
            Support_Under_Right_R.append(str("{}Ø{}".format(n1, d1)))
            Kiris_Tasima(Alanc + Span_Area[i])
            Support_Under_Right_Bearing.append(M_tasima)
        else:
            Support_Under_Right.append("None")
            Support_Under_Right_R.append("None")
            Kiris_Tasima(Span_Area[i])
            Support_Under_Right_Bearing.append(M_tasima)

def Etriye_Hesabi(Ve, VG_Q_DVE,Vd,Vdy,width,height):
    global Etriye_Capi,Sk,So,Vc
    d_Height=height-d
    Etriye_Capi = 8
    n_kol = 2  # Etriye Kol Sayısı
    fi_min=12
    fydw = fyk / Gama_S
    fctk = 0.35 * math.sqrt(fck)  # Karakteristik çekme dayanımı
    fctd = fctk / Gama_C  # Tasarım çekme dayanımı
    Asw=(n_kol * math.pi * Etriye_Capi ** 2) / 4
    Vdmax = (0.85 * width * d_Height * math.sqrt(fck)) / 1e3
    Vcr = (0.65 * fctd * width * d_Height) / 1e3
    Vc=0.8*Vcr
    if Ve > VG_Q_DVE:
        Ve = VG_Q_DVE

    def OrtaBolge():
        global So
        Vw=abs(Ve-Vc)
        S=Asw*fydw*d_Height/(Vw*1e3)
        if Vd<= 3 * Vcr:
            if  S< d_Height / 2:
                So=S
                #print("Orta Bölge Adım Mesafesi:",S)
            else:
                So=d_Height / 2
        elif Vd > 3 * Vcr:
            if S < d_Height / 4:
                So=S
                #print("Orta Bölge Adım Mesafesi:", So)
            else:
                So=d_Height / 4
        return So
    def SarilmaBolge():
        global Sk,Vc
        if Ve - Vdy > (VG_Q_DVE / 2):  # Beton Katkısı
            Vc = 0
        Vw = abs(Ve - Vc)
        S = Asw * fydw * d_Height / (Vw * 1e3)
        if S<=d_Height/4 and S<=8*fi_min and S<=150:
            Sk=S
        else:
            S_=min(d_Height/4,8*fi_min,150)
            Sk=S_
        return Sk
    if Vdmax<=Ve:
        print("ERRORR!!!Kesit Boyutları Artırılmalı")
    else:
        OrtaBolge()
        SarilmaBolge()
        #print("Ø{}/{}/{}".format(Etriye_Capi, round(Sk, 0), round(So, 0)))
    return Etriye_Capi,Sk,So

Etriye=[]
global Etriye_Capi,Sk,So
for i in range(0,len(M_span)):
    Mr_1=(1.4*Support_Upper_Right_Bearing[i]+1.4*Support_Under_Left_Bearing[i])/Length[i]  #Sağ Üst Sol Alt
    Mr_2=(1.4*Support_Upper_Left_Bearing[i]+1.4*Support_Under_Right_Bearing[i])/Length[i]  #Sağ ALt Sol Üst
    Ve=max(Vdy[i]+Mr_1,abs(Vdy_2[i]-Mr_1),Vdy[i]+Mr_2,abs(Vdy_2[i]-Mr_2))
    Vdy_max=max(Vdy[i],abs(Vdy_2[i])) #G+Q
    Etriye_Hesabi(Ve, VG_Q_DVE[i], Vd[i],Vdy_max, Width[i],Height[i])
    Etriye.append(str("Ø{}/{}/{}").format(Etriye_Capi,round(Sk,0),round(So,0)))

data_write=pd.DataFrame({"Kiriş":data.Kiris[::],"Length":Length[::],"Açıklık Alt":Span_R[::],"Açıklık Alt A":Span_Area[::],"Açıklık Üst":Span_Upper_R[::],
                         "Açıklık Üst A":Span_Upper_Area[::],"Sol Üst":Support_Left_R[::],
                         "Sol Üst A":Support_Left[::],"Sol ALt":Support_Under_Left_R[::],
                         "Sol Alt A":Support_Under_Left[::],"Sağ Üst":Support_Right_R[::],
                         "Sağ Üst A":Support_Right[::],"Sağ ALt":Support_Under_Right_R[::],
                         "Sağ Alt A":Support_Under_Right[::],"Etriye":Etriye[::]})

book=openpyxl.load_workbook(dosyaadi)
writer=pd.ExcelWriter(dosyaadi,engine="openpyxl")
writer.book=book
data_write.to_excel(writer,sheet_name="Donatilar")
writer.save()
#data_Save=pd.ExcelWriter("C:/Users/Demir/Desktop/Yeni klasör/deneme.xlsx",engine="xlsxwriter")
#data_write.to_excel(data_Save, sheet_name="Kiris")
#data_Save.save()


