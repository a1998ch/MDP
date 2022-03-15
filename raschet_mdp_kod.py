import function_mdp as fm
import pandas as pd
import win32com.client
import sys

# Величина амплитуды нерегулярных колебаний.
p_nk = 30

# Осуществляем расчёт режима.
kod = fm.rastr.rgm('p')

# МДП по критерию обеспечения нормативного коэффициента запаса
# статической апериодической устойчивости по активной мощности
# в контролируемом сечении в нормальной схеме.
while kod == 0:
    fm.utyazhelenie()
    kod = fm.rastr.rgm('p')
    fm.peretok_v_sechenii()

p_pred1 = fm.peretok_v_sechenii()
p_mdp1 = 0.8 * p_pred1 - p_nk

fm.vozvrat_k_ishodnomu_regimu()
kod = fm.rastr.rgm('p')

# МДП по критерию обеспечения нормативного коэффициента запаса
# статической устойчивости по напряжению
# в узлах нагрузки в нормальной схеме.
while kod == 0:
    kod = fm.rastr.rgm('p')
    for j in range(fm.nodes_count):
        if fm.uras.Z(j) < (fm.unom.Z(j) * 0.7) * 1.15:
            break
    if fm.uras.Z(j) > (fm.unom.Z(j) * 0.7) * 1.15:
        fm.utyazhelenie()
        kod = fm.rastr.rgm('p')
    else:
        fm.peretok_v_sechenii()
        break
else:
    fm.peretok_v_sechenii()

p_pred2 = fm.peretok_v_sechenii()
p_mdp2 = p_pred2 - p_nk

fm.vozvrat_k_ishodnomu_regimu()
kod = fm.rastr.rgm('p')

# МДП по критерию обеспечения нормативного коэффициента запаса
# статической апериодической устойчивости по активной мощности
# в контролируемом сечении в послеаварийных режимах
# после нормативных возмущений.
pred = {}
pred_sta = {}
index_vozmush_statica = list(fm.index_vozmush)
listic_sta_faults_statica = fm.df_faults.sta.T.tolist()

for j in range(len(index_vozmush_statica)):
    for (index, sta_faults) in zip(
         index_vozmush_statica, listic_sta_faults_statica):
        fm.sta.SetZ(index, sta_faults)
        break
    index_vozmush_statica.remove(index)
    listic_sta_faults_statica.remove(sta_faults)
    while kod == 0:
        kod = fm.rastr.rgm('p')
        if kod == 0:
            fm.utyazhelenie()
        else:
            pred[index] = fm.peretok_v_sechenii() * 0.92
            pred_sta[sta_faults] = fm.peretok_v_sechenii() * 0.92
            fm.sta.SetZ(index, abs(sta_faults - 1))
            kod = fm.rastr.rgm('p')
            break
    while kod == 0:
        kod = fm.rastr.rgm('p')
        if fm.peretok_v_sechenii() > fm.p_sech_nach and \
                len(index_vozmush_statica) > 0:
            fm.obratnoe_utyazhelenie()
            kod = fm.rastr.rgm('p')
            summa = 0
        else:
            break

for key, value in pred.items():
    min_pred_index = min(pred, key=pred.get)
min_pred = pred[min_pred_index]

for key, value in pred_sta.items():
    min_pred_sta = min(pred_sta, key=pred_sta.get)

# Повторное разутяжеление до наименьшего
# предельного перетока с 8% запасом
# из рассматриваемых возмущений
kod = fm.rastr.rgm('p')
while kod == 0:
    for i in range(fm.nodes_count):
        fm.sta.SetZ(min_pred_index, min_pred_sta)
        break
    kod = fm.rastr.rgm('p')
    if fm.peretok_v_sechenii() > min_pred:
        fm.obratnoe_utyazhelenie()
        kod = fm.rastr.rgm('p')
    else:
        fm.sta.SetZ(min_pred_index, abs(min_pred_sta - 1))
        fm.rastr.rgm('p')
        fm.peretok_v_sechenii()
        break

p_pred3 = fm.peretok_v_sechenii()
p_mdp3 = p_pred3 - p_nk

fm.vozvrat_k_ishodnomu_regimu()
kod = fm.rastr.rgm('p')

# МДП по критерию обеспечения нормативного коэффициента запаса
# статической устойчивости по напряжению в узлах нагрузки
# в послеаварийных режимах после нормативных возмущений.
pred = {}
index_vozmush_naprygenie = list(fm.index_vozmush)
listic_sta_faults_naprygenie = fm.df_faults.sta.T.tolist()
kod = fm.rastr.rgm('p')

for j in range(len(index_vozmush_naprygenie)):
    for (index, sta_faults) in zip(
            index_vozmush_naprygenie, listic_sta_faults_naprygenie):
        fm.sta.SetZ(index, sta_faults)
        break
    index_vozmush_naprygenie.remove(index)
    listic_sta_faults_naprygenie.remove(sta_faults)
    while kod == 0:
        kod = fm.rastr.rgm('p')
        for j in range(fm.nodes_count):
            if fm.uras.Z(j) < (fm.unom.Z(j) * 0.7) * 1.1:
                break
        if fm.uras.Z(j) > (fm.unom.Z(j) * 0.7) * 1.1:
            fm.utyazhelenie()
            kod = fm.rastr.rgm('p')
        else:
            fm.sta.SetZ(index, abs(sta_faults - 1))
            kod = fm.rastr.rgm('p')
            pred[index] = fm.peretok_v_sechenii()
            fm.peretok_v_sechenii()
            break
    else:
        fm.sta.SetZ(index, abs(sta_faults - 1))
        kod = fm.rastr.rgm('p')
        pred[index] = fm.peretok_v_sechenii()
        fm.peretok_v_sechenii()
    while kod == 0:
        kod = fm.rastr.rgm('p')
        if fm.peretok_v_sechenii() > fm.p_sech_nach and \
                len(index_vozmush_naprygenie) > 0:
            fm.obratnoe_utyazhelenie()
            kod = fm.rastr.rgm('p')
        else:
            break

for key, value in pred.items():
    min_pred_index = min(pred, key=pred.get)
min_pred = pred[min_pred_index]

p_pred4 = min_pred
p_mdp4 = p_pred4 - p_nk

fm.vozvrat_k_ishodnomu_regimu()
kod = fm.rastr.rgm('p')

# МДП по критерию обеспечения допустимой токовой нагрузки
# линий электропередачи и электросетевого оборудования
# в нормальной схеме.
while kod == 0:
    kod = fm.rastr.rgm('p')
    for j in range(fm.nodes_count):
        if (fm.zag_i.Z(j) * 1000) >= 100 and (fm.zag_it.Z(j) * 1000) >= 100:
            break
    if (fm.zag_i.Z(j) * 1000) < 100 or (fm.zag_it.Z(j) * 1000) < 100:
        fm.utyazhelenie()
        kod = fm.rastr.rgm('p')
    else:
        fm.peretok_v_sechenii()
        break

p_pred5 = fm.peretok_v_sechenii()
p_mdp5 = p_pred5 - p_nk

fm.vozvrat_k_ishodnomu_regimu()
kod = fm.rastr.rgm('p')

# МДП по критерию обеспечения допустимой токовой нагрузки
# линий электропередачи и электросетевого оборудования
# в послеаварийных режимах после нормативных возмущений.
pred = {}
index_vozmush_tok = list(fm.index_vozmush)
listic_sta_faults_tok = fm.df_faults.sta.T.tolist()

for j in range(len(index_vozmush_tok)):
    for (index, sta_faults) in zip(index_vozmush_tok, listic_sta_faults_tok):
        fm.sta.SetZ(index, sta_faults)
        break
    index_vozmush_tok.remove(index)
    listic_sta_faults_tok.remove(sta_faults)
    while kod == 0:
        kod = fm.rastr.rgm('p')
        for j in range(fm.nodes_count):
            if (fm.zag_i_av.Z(j) * 1000) >= 100 and \
               (fm.zag_it_av.Z(j) * 1000) >= 100:
                break
        if (fm.zag_i_av.Z(j) * 1000) < 100 or \
                (fm.zag_it_av.Z(j) * 1000) < 100:
            fm.utyazhelenie()
            kod = fm.rastr.rgm('p')
        else:
            fm.sta.SetZ(index, abs(sta_faults - 1))
            kod = fm.rastr.rgm('p')
            pred[index] = fm.peretok_v_sechenii()
            fm.peretok_v_sechenii()
            break
    else:
        fm.sta.SetZ(index, abs(sta_faults - 1))
        kod = fm.rastr.rgm('p')
        pred[index] = fm.peretok_v_sechenii()
        fm.peretok_v_sechenii()
    while kod == 0:
        kod = fm.rastr.rgm('p')
        if fm.peretok_v_sechenii() > fm.p_sech_nach and \
                len(index_vozmush_tok) > 0:
            fm.obratnoe_utyazhelenie()
            kod = fm.rastr.rgm('p')
        else:
            break

for key, value in pred.items():
    min_pred_index = min(pred, key=pred.get)
min_pred = pred[min_pred_index]

p_pred6 = min_pred
p_mdp6 = p_pred6 - p_nk


df = pd.DataFrame({
                    'Предельный переток': [
                     p_pred1, p_pred2, p_pred3,
                     p_pred4, p_pred5, p_pred6],
                    'МДП': [p_mdp1, p_mdp2, p_mdp3, p_mdp4, p_mdp5, p_mdp6],
})
df.index = ['Обеспечение нормативного коэффициента запаса \
             статической апериодической устойчивости в КС в нормальной схеме',
            'Обеспечение нормативного коэффициента запаса \
             статической устойчивости по напряжению в КС в нормальной схеме',
            'Обеспечение нормативного коэффициента запаса \
             статической апериодической устойчивости в КС в ПАР',
            'Обеспечение нормативного коэффициента запаса \
             статической устойчивости по напряжению в КС в ПАР',
            'Обеспечение допустимой токовой нагрузки ЛЭП и \
             электросетевого оборудования в КС в нормальной схеме',
            'Обеспечение допустимой токовой нагрузки ЛЭП и \
             электросетевого оборудования в КС в ПАР']
print(df)
