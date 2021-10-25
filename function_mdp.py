import sys
import pandas as pd
import win32com.client
rastr = win32com.client.Dispatch('Astra.Rastr')


# Осуществляем загрузку файлов режима, траектории утяжеления,
# контролируемого сечения и нормативных возмущений.
rastr.Load(1, 'regime.rg2', '')
flowgate = pd.read_json('flowgate.json')
faults = pd.read_json('faults.json')
vector = pd.read_csv('vector.csv')
regime = rastr.rgm('p')


# Производим создание переменных,
# которые будем использовать в дальнейшем.
node = rastr.Tables("node")
ny = node.Cols("ny")
pg = node.Cols("pg")
pn = node.Cols("pn")
qn = node.Cols("qn")
tg = node.Cols("tg_phi")
name = node.Cols("name")
tip = node.Cols("tip")
unom = node.Cols("uhom")
uras = node.Cols("vras")

vetv = rastr.Tables("vetv")
name_v = vetv.Cols("name")
n_nach = vetv.Cols("ip")
n_kon = vetv.Cols("iq")
np = vetv.Cols("np")
p_nach = vetv.Cols("pl_ip")
q_nach = vetv.Cols("ql_ip")
sta = vetv.Cols("sta")
zag_i = vetv.Cols("zag_i")
zag_it = vetv.Cols("zag_it")
zag_i_av = vetv.Cols("zag_i_av")
zag_it_av = vetv.Cols("zag_it_av")


# Траектория утяжеления
df = pd.DataFrame(vector)

listic_node_type = df.variable.T.tolist()
listic_node = df.node.T.tolist()
listic_value = df.value.T.tolist()
listic_tg = df.tg.T.tolist()


# Определяем номера узлов в заданном режиме
nodes_count = node.size

listic_nodes_num = []
for i in range(nodes_count):
    listic_nodes_num.append(ny.Z(i))


# Определяем индексы узлов в заданном режиме
listic_nodes_index = []
for i in listic_node:
    for j, num in enumerate(listic_nodes_num, start=0):
        if i == num:
            listic_nodes_index.append(j)


# Отделим генераторные узлы от нагрузочных,
# а также приращение мощности для данных типов узлов.
listic_pn = []
listic_pg = []
listic_pn_value = []
listic_pg_value = []
for (a, b, c) in zip(listic_node_type, listic_nodes_index, listic_value):
    if a == 'pn':
        listic_pn.append(b)
        listic_pn_value.append(c)
    else:
        listic_pg.append(b)
        listic_pg_value.append(c)


# Расчёт тангенса нагрузки
for i in range(nodes_count):
    if tip.Z(i) == 1 and pn.Z(i) != 0:
        tg.SetZ(i, qn.Z(i) / pn.Z(i))


# Определим количество узлов где необходимо сохранять тангенс нагрузки,
# для корректного расчёта оставшиеся узлы заменим нулями
listic_tg_true = []
for (a, b) in zip(listic_tg, listic_nodes_index):
    if a == 1:
        listic_tg_true.append(b)
    else:
        listic_tg_true.append(0)


# Контролируемое сечение
df2 = pd.DataFrame(flowgate)
df2 = df2.transpose()

listic_ip = df2.ip.T.tolist()
listic_iq = df2.iq.T.tolist()
listic_np = df2.np.T.tolist()


# Перечень нормативных возмущений
df3 = pd.DataFrame(faults)
df3 = df3.transpose()

listic_ip_faults = df3.ip.T.tolist()
listic_iq_faults = df3.iq.T.tolist()
listic_np_faults = df3.np.T.tolist()
listic_sta_faults = df3.sta.T.tolist()


# Переток в сечении в исходном режиме
sechenie_nach = {}
index_vetv_sech = []
for i in range(nodes_count):
    for (a, b, c) in zip(listic_ip, listic_iq, listic_np):
        if n_nach.Z(i) == a and n_kon.Z(i) == b and np.Z(i) == c:
            sechenie_nach[name_v.Z(i)] = p_nach.Z(i)
            index_vetv_sech.append(i)
sechenie_nach = pd.DataFrame(list(sechenie_nach.items()), columns=[
                'Номер ветви', 'Переток активной мощности'])
p_sech_nach = sechenie_nach['Переток активной мощности'].sum()


# Нормативные возмущения
index_vozmush = []
for i in range(nodes_count):
    for (a, b, c, d) in zip(
     listic_ip_faults, listic_iq_faults, listic_np_faults, listic_sta_faults):
        if n_nach.Z(i) == a and n_kon.Z(i) == b and np.Z(i) == c:
            index_vozmush.append(i)


# Определение исходного состояния ветвей
listic_sta_faults_vkluchit = []
for i in listic_sta_faults:
    if i == 0:
        listic_sta_faults_vkluchit.append(1)
    else:
        listic_sta_faults_vkluchit.append(0)


# Функция утяжеления
def utyazhelenie():
    """Функция осуществляет изменение мощности нагрузки
       и генерации в соответствии с заданной траекторией
       утяжеления

    Аргументы:
        listic_pn: список названий узлов, где необходимо
                   изменить мощность нагрузки.
        listic_pn_value: список, где находятся величины,
                         на которые необходимо изменить
                         мощность нагрузки.
        listic_tg_true: список, который содержит информацию
                        о узлах, в которых необходимо
                        сохранять тангенс нагрузки.
        listic_pg: список названий узлов, где необходимо
                   изменить мощность генерации.
        listic_pg_value: список, где находятся величины,
                         на которые необходимо изменить
                         мощность генерации.
        pn: активная мощность нагрузки.
        qn: реактивная мощность генерации.
        pg: активная мощность генерации.
    Возвращает:
        Изменённый режим.
    """
    for (a, b, c) in zip(listic_pn, listic_pn_value, listic_tg_true):
        pn.SetZ(a, pn.Z(a) + b)
        if a == c:
            qn.SetZ(a, tg.Z(a) * pn.Z(a))
        else:
            qn.SetZ(a, qn.Z(a))
    for (c, d) in zip(listic_pg, listic_pg_value):
        pg.SetZ(c, pg.Z(c) + d)


# Функция обратного утяжеления
def obratnoe_utyazhelenie():
    """Функция осуществляет изменение мощности нагрузки
       и генерации противоположно заданной траекторией
       утяжеления

    Аргументы:
        listic_pn: список названий узлов, где необходимо
                   изменить мощность нагрузки.
        listic_pn_value: список, где находятся величины,
                         на которые необходимо изменить
                         мощность нагрузки.
        listic_tg_true: список, который содержит информацию
                        о узлах, в которых необходимо
                        сохранять тангенс нагрузки.
        listic_pg: список названий узлов, где необходимо
                   изменить мощность генерации.
        listic_pg_value: список, где находятся величины,
                         на которые необходимо изменить
                         мощность генерации.
        pn: активная мощность нагрузки.
        qn: реактивная мощность генерации.
        pg: активная мощность генерации.
    Возвращает:
        Изменённый режим.
    """
    for (a, b, c) in zip(listic_pn, listic_pn_value, listic_tg_true):
        pn.SetZ(a, pn.Z(a) - b)
        if a == c:
            qn.SetZ(a, tg.Z(a) * pn.Z(a))
        else:
            qn.SetZ(a, qn.Z(a))
    for (c, d) in zip(listic_pg, listic_pg_value):
        pg.SetZ(c, pg.Z(c) - d)


# Функция возврата к исходному режиму
def vozvrat_k_ishodnomu_regimu():
    """Функция осуществляет изменение мощности нагрузки
       и генерации до достижения исходного режима

    Аргументы:
        listic_pn: список названий узлов, где необходимо
                   изменить мощность нагрузки.
        listic_pn_value: список, где находятся величины,
                         на которые необходимо изменить
                         мощность нагрузки.
        listic_tg_true: список, который содержит информацию
                        о узлах, в которых необходимо
                        сохранять тангенс нагрузки.
        listic_pg: список названий узлов, где необходимо
                   изменить мощность генерации.
        listic_pg_value: список, где находятся величины,
                         на которые необходимо изменить
                         мощность генерации.
        pn: активная мощность нагрузки.
        qn: реактивная мощность генерации.
        pg: активная мощность генерации.
        index_vetv_sech: индексы ветвей, входящих в сечение
        p_sech_nach: переток в сечении в исходном режиме
    Возвращает:
        Исходный режим.
    """
    for i in range(sys.maxsize):
        kod = rastr.rgm('p')
        sechenie = 0
        for j in index_vetv_sech:
            sechenie += p_nach.Z(j)
        if sechenie > p_sech_nach:
            obratnoe_utyazhelenie()
            kod = rastr.rgm('p')
            sechenie = 0
        else:
            break


# Функция определения перетока в сечении
def peretok_v_sechenii():
    """Функция осуществляет определение
       перетока в сечении.

    Аргументы:
        index_vetv_sech: индексы ветвей, входящих в сечение
        p_sech_nach: переток в сечении в исходном режиме
    Возвращает:
        Величину перетока в сечении.
    """
    sechenie = 0
    for j in index_vetv_sech:
        sechenie += p_nach.Z(j)
    return sechenie


# МДП по критерию обеспечения нормативного коэффициента запаса
# статической апериодической устойчивости по активной мощности
# в контролируемом сечении в нормальной схеме.
for i in range(sys.maxsize):
    kod = rastr.rgm('p')
    if kod == 0:
        utyazhelenie()
        kod = rastr.rgm('p')
    else:
        peretok_v_sechenii()
        break

p_pred1 = peretok_v_sechenii()


vozvrat_k_ishodnomu_regimu()


# МДП по критерию обеспечения нормативного коэффициента запаса
# статической устойчивости по напряжению
# в узлах нагрузки в нормальной схеме.
for i in range(sys.maxsize):
    kod = rastr.rgm('p')
    for j in range(nodes_count):
        if uras.Z(j) < (unom.Z(j) * 0.7) * 1.15:
            break
    if kod == 0 and uras.Z(j) > (unom.Z(j) * 0.7) * 1.15:
        utyazhelenie()
        kod = rastr.rgm('p')
    else:
        peretok_v_sechenii()
        break

p_pred2 = peretok_v_sechenii()


vozvrat_k_ishodnomu_regimu()


# МДП по критерию обеспечения нормативного коэффициента запаса
# статической апериодической устойчивости по активной мощности
# в контролируемом сечении в послеаварийных режимах
# после нормативных возмущений.
pred = {}
pred_sta = {}
index_vozmush_statica = list(index_vozmush)
listic_sta_faults_statica = list(listic_sta_faults)

for j in range(len(index_vozmush_statica)):
    for (a, b) in zip(index_vozmush_statica, listic_sta_faults_statica):
        sta.SetZ(a, b)
        k = a
        k2 = abs(b - 1)
        break
    index_vozmush_statica.remove(a)
    listic_sta_faults_statica.remove(b)
    for i in range(sys.maxsize):
        kod = rastr.rgm('p')
        if kod == 0:
            utyazhelenie()
        else:
            summa = 0
            for j in index_vetv_sech:
                summa += p_nach.Z(j)
            pred[k] = summa * 0.92
            pred_sta[abs(k2 - 1)] = summa * 0.92
            sta.SetZ(k, k2)
            kod = rastr.rgm('p')
            break
    for i in range(sys.maxsize):
        kod = rastr.rgm('p')
        summa = 0
        for j in index_vetv_sech:
            summa += p_nach.Z(j)
        if summa > p_sech_nach and len(index_vozmush_statica) > 0:
            obratnoe_utyazhelenie()
            kod = rastr.rgm('p')
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
for i in range(sys.maxsize):
    for i in range(nodes_count):
        sta.SetZ(min_pred_index, min_pred_sta)
        k = min_pred_index
        k2 = abs(min_pred_sta - 1)
        break
    kod = rastr.rgm('p')
    for j in index_vetv_sech:
        summa += p_nach.Z(j)
    if summa > min_pred:
        obratnoe_utyazhelenie()
        kod = rastr.rgm('p')
        summa = 0
    else:
        sta.SetZ(k, k2)
        rastr.rgm('p')
        peretok_v_sechenii()
        break

p_pred3 = peretok_v_sechenii()


vozvrat_k_ishodnomu_regimu()


# МДП по критерию обеспечения нормативного коэффициента запаса
# статической устойчивости по напряжению в узлах нагрузки
# в послеаварийных режимах после нормативных возмущений.
pred = {}
index_vozmush_naprygenie = list(index_vozmush)
listic_sta_faults_naprygenie = list(listic_sta_faults)

for j in range(len(index_vozmush_naprygenie)):
    for (a, b) in zip(index_vozmush_naprygenie, listic_sta_faults_naprygenie):
        sta.SetZ(a, b)
        k = a
        k2 = abs(b - 1)
        break
    index_vozmush_naprygenie.remove(a)
    listic_sta_faults_naprygenie.remove(b)
    for i in range(sys.maxsize):
        kod = rastr.rgm('p')
        for j in range(nodes_count):
            if uras.Z(j) < (unom.Z(j) * 0.7) * 1.1:
                break
        if kod == 0 and uras.Z(j) > (unom.Z(j) * 0.7) * 1.1:
            utyazhelenie()
            kod = rastr.rgm('p')
        else:
            sta.SetZ(k, k2)
            kod = rastr.rgm('p')
            summa = 0
            for j in index_vetv_sech:
                summa += p_nach.Z(j)
            pred[k] = summa
            peretok_v_sechenii()
            break
    for i in range(sys.maxsize):
        kod = rastr.rgm('p')
        summa = 0
        for j in index_vetv_sech:
            summa += p_nach.Z(j)
        if summa > p_sech_nach and len(index_vozmush_naprygenie) > 0:
            obratnoe_utyazhelenie()
            kod = rastr.rgm('p')
            summa = 0
        else:
            break

for key, value in pred.items():
    min_pred_index = min(pred, key=pred.get)
min_pred = pred[min_pred_index]

p_pred4 = min_pred


vozvrat_k_ishodnomu_regimu()


# МДП по критерию обеспечения допустимой токовой нагрузки
# линий электропередачи и электросетевого оборудования
# в нормальной схеме.
for i in range(sys.maxsize):
    kod = rastr.rgm('p')
    for j in range(nodes_count):
        if (zag_i.Z(j) * 1000) >= 100 and (zag_it.Z(j) * 1000) >= 100:
            break
    if kod == 0 and (zag_i.Z(j) * 1000) < 100 or (zag_it.Z(j) * 1000) < 100:
        utyazhelenie()
        kod = rastr.rgm('p')
    else:
        peretok_v_sechenii()
        break

p_pred5 = peretok_v_sechenii()


vozvrat_k_ishodnomu_regimu()


# МДП по критерию обеспечения допустимой токовой нагрузки
# линий электропередачи и электросетевого оборудования
# в послеаварийных режимах после нормативных возмущений.
pred = {}
index_vozmush_tok = list(index_vozmush)
listic_sta_faults_tok = list(listic_sta_faults)

for j in range(len(index_vozmush_tok)):
    for (a, b) in zip(index_vozmush_tok, listic_sta_faults_tok):
        sta.SetZ(a, b)
        k = a
        k2 = abs(b - 1)
        break
    index_vozmush_tok.remove(a)
    listic_sta_faults_tok.remove(b)
    for i in range(sys.maxsize):
        kod = rastr.rgm('p')
        for j in range(nodes_count):
            if (zag_i_av.Z(j) * 1000) >= 100 and \
               (zag_it_av.Z(j) * 1000) >= 100:
                break
        if kod == 0 and (zag_i_av.Z(j) * 1000) < 100 or \
                        (zag_it_av.Z(j) * 1000) < 100:
            utyazhelenie()
            kod = rastr.rgm('p')
        else:
            sta.SetZ(k, k2)
            kod = rastr.rgm('p')
            summa = 0
            for j in index_vetv_sech:
                summa += p_nach.Z(j)
            pred[k] = summa
            peretok_v_sechenii()
            break
    for i in range(sys.maxsize):
        kod = rastr.rgm('p')
        summa = 0
        for j in index_vetv_sech:
            summa += p_nach.Z(j)
        if summa > p_sech_nach and len(index_vozmush_tok) > 0:
            obratnoe_utyazhelenie()
            kod = rastr.rgm('p')
            summa = 0
        else:
            break

for key, value in pred.items():
    min_pred_index = min(pred, key=pred.get)
min_pred = pred[min_pred_index]

p_pred6 = min_pred
