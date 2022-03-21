import sys
from matplotlib.pyplot import flag
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

listic_nodes = []
slovaric_nodes_index = {}
index_vetv_sech = []
index_vozmush = []
listic_sta_faults_ischodnii = []
p_sech_nach = []

# Осуществляем расчёт режима.
kod = rastr.rgm('p')

# Количество узлов и ветвей в заданном режиме
nodes_count = node.size
vetv_count = vetv.size

# Траектория утяжеления
df_vector = pd.DataFrame(vector)

# Контролируемое сечение
df_flowgate = pd.DataFrame(flowgate).transpose()

# Перечень нормативных возмущений
df_faults = pd.DataFrame(faults).transpose()


# Определяем номера узлов в заданном режиме
# Расчёт тангенса нагрузки
def nodes_number():
    """Функция осуществляет подсчёт узлов в
       заданном режиме, а также производит
       расчёт тангенса нагрузки

    Аргументы:
        ny: номер узла.
        pn: активная мощность нагрузки.
        qn: реактивная мощность генерации.
        tip: тип узла.
        tg: тангенс угла нагрузки
    Возвращает:
        список узлов режима.
    """
    if len(listic_nodes) == 0:
        for i in range(nodes_count):
            listic_nodes.append(ny.Z(i))
            if tip.Z(i) == 1 and pn.Z(i) != 0:
                tg.SetZ(i, qn.Z(i) / pn.Z(i))
        return listic_nodes
    else:
        return listic_nodes


# Определяем индексы узлов в заданном режиме
def nodes_index():
    """Функция определяет индексы узлов режима

    Аргументы:
        nodes_number: Функция, которая осуществляет
                      подсчёт узлов в заданном режиме.
    Возвращает:
        словарь узлов и их индексов режима.
    """
    if len(slovaric_nodes_index) == 0:
        for i, num in enumerate(nodes_number()):
            slovaric_nodes_index[num] = i
        return slovaric_nodes_index
    else:
        return slovaric_nodes_index


# Определение индексов ветвей в сечении.
def index_vetv_sech_f():
    """Функция определяет индексы ветвей в сечении

    Аргументы:
        n_nach: номер начала ветви.
        n_kon: номер конца ветви.
        np: номер параллельности ветви.
    Возвращает:
        словарь узлов и их индексов режима.
    """
    if len(index_vetv_sech) == 0:
        for i in range(vetv_count):
            for j in range(1, len(df_flowgate.index) + 1):
                if n_nach.Z(i) == df_flowgate.at[f'line_{j}', 'ip'] and \
                    n_kon.Z(i) == df_flowgate.at[f'line_{j}', 'iq'] and \
                        np.Z(i) == df_flowgate.at[f'line_{j}', 'np']:
                    index_vetv_sech.append(i)
        return index_vetv_sech
    else:
        return index_vetv_sech


# Определение перетока в сечении в исходном режиме.
def p_sech_nach_f():
    """Функция определяет переток в сечении
       в исходном режиме

    Аргументы:
        n_nach: номер начала ветви.
        n_kon: номер конца ветви.
        np: номер параллельности ветви.
    Возвращает:
        переток в сечении в исходном режиме.
    """
    if len(p_sech_nach) == 0:
        p_sech_nach_p = 0
        for i in range(vetv_count):
            for j in range(1, len(df_flowgate.index) + 1):
                if n_nach.Z(i) == df_flowgate.at[f'line_{j}', 'ip'] and \
                    n_kon.Z(i) == df_flowgate.at[f'line_{j}', 'iq'] and \
                        np.Z(i) == df_flowgate.at[f'line_{j}', 'np']:
                    p_sech_nach_p += p_nach.Z(i)
        p_sech_nach.append(p_sech_nach_p)
        return p_sech_nach[0]
    else:
        return p_sech_nach[0]


# Определение индексов ветвей, эксплуатационное состояние
# которых требуется изменить.
def index_vozmush_f():
    """Функция определяет индексы ветвей,
       которые входят в нормативные возмущения

    Аргументы:
        n_nach: номер начала ветви.
        n_kon: номер конца ветви.
        np: номер параллельности ветви.
    Возвращает:
        индексы ветвей, которые входят
        в нормативные возмущения.
    """
    if len(index_vozmush) == 0:
        for i in range(vetv_count):
            for k in range(len(df_faults.index)):
                if n_nach.Z(i) == df_faults.iloc[k]['ip'] and \
                    n_kon.Z(i) == df_faults.iloc[k]['iq'] and \
                        np.Z(i) == df_faults.iloc[k]['np']:
                    index_vozmush.append(i)
                    listic_sta_faults_ischodnii.append(sta.Z(i))
        return index_vozmush
    else:
        return index_vozmush


# Функция утяжеления
def utyazhelenie():
    """Функция осуществляет изменение мощности нагрузки
       и генерации в соответствии с заданной траекторией
       утяжеления

    Аргументы:
        pn: активная мощность нагрузки.
        qn: реактивная мощность генерации.
        pg: активная мощность генерации.
        tg: тангенс угла нагрузки
    Возвращает:
        Изменённый режим.
    """
    for j, i in enumerate(df_vector.node.T.tolist()):
        if df_vector.iloc[j]['variable'] == 'pn':
            pn.SetZ(
                nodes_index().get(i), pn.Z(
                    nodes_index().get(i)) + df_vector.iloc[j]['value'])
            if df_vector.iloc[j]['tg'] == 1:
                qn.SetZ(
                    nodes_index().get(i), tg.Z(
                        nodes_index().get(i)) * pn.Z(
                            nodes_index().get(i)))
        else:
            pg.SetZ(
                nodes_index().get(i), pg.Z(
                    nodes_index().get(i)) + df_vector.iloc[j]['value'])


# Функция обратного утяжеления
def obratnoe_utyazhelenie():
    """Функция осуществляет изменение мощности нагрузки
       и генерации в соответствии с заданной траекторией
       утяжеления

    Аргументы:
        pn: активная мощность нагрузки.
        qn: реактивная мощность генерации.
        pg: активная мощность генерации.
        tg: тангенс угла нагрузки
    Возвращает:
        Изменённый режим.
    """
    for j, i in enumerate(df_vector.node.T.tolist()):
        if df_vector.iloc[j]['variable'] == 'pn':
            pn.SetZ(
                nodes_index().get(i), pn.Z(
                    nodes_index().get(i)) - df_vector.iloc[j]['value'])
            if df_vector.iloc[j]['tg'] == 1:
                qn.SetZ(
                    nodes_index().get(i), tg.Z(
                        nodes_index().get(i)) * pn.Z(
                            nodes_index().get(i)))
        else:
            pg.SetZ(
                nodes_index().get(i), pg.Z(
                    nodes_index().get(i)) - df_vector.iloc[j]['value'])


# Функция возврата к исходному режиму
def vozvrat_k_ishodnomu_regimu():
    """Функция осуществляет изменение мощности нагрузки
       и генерации до достижения исходного режима

    Аргументы:
        index_vetv_sech: индексы ветвей, входящих в сечение
        p_sech_nach: переток в сечении в исходном режиме
    Возвращает:
        Исходный режим.
    """
    for i in range(sys.maxsize):
        kod = rastr.rgm('p')
        if peretok_v_sechenii() > p_sech_nach_f():
            obratnoe_utyazhelenie()
            kod = rastr.rgm('p')
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
    for j in index_vetv_sech_f():
        sechenie += p_nach.Z(j)
    return sechenie


# Определение предельного перетока по критерию обеспечения
# нормативного коэффициента запаса статической апериодической
# устойчивости по активной мощности
# в контролируемом сечении в нормальной схеме.
def pred_1(utyazhelenie, peretok_v_sechenii, kod):
    """Функция осуществляет определение предельного перетока
       по критерию обеспечения нормативного коэффициента
       запаса статической апериодической устойчивости по
       активной мощности в контролируемом сечении
       в нормальной схеме.

    Аргументы:
        utyazhelenie: осуществляет утяжеление режима
        peretok_v_sechenii: определяет переток в сечении
        kod: осуществляет расчёт режима
    Возвращает:
        Предельный переток в сечении.
    """
    while kod == 0:
        utyazhelenie()
        kod = rastr.rgm('p')
        peretok_v_sechenii()
    return peretok_v_sechenii()


# Определение предельного перетока по критерию
# обеспечения нормативного коэффициента запаса
# статической устойчивости по напряжению
# в узлах нагрузки в нормальной схеме.
def pred_2(utyazhelenie, peretok_v_sechenii, kod):
    """Функция осуществляет определение предельного перетока
       по критерию обеспечения нормативного коэффициента
       запаса статической устойчивости по напряжению
       в узлах нагрузки в нормальной схеме.

    Аргументы:
        utyazhelenie: осуществляет утяжеление режима
        peretok_v_sechenii: определяет переток в сечении
        kod: осуществляет расчёт режима
    Возвращает:
        Предельный переток в сечении.
    """
    while kod == 0:
        kod = rastr.rgm('p')
        for j in range(nodes_count):
            if uras.Z(j) < (unom.Z(j) * 0.7) * 1.15:
                break
        if uras.Z(j) > (unom.Z(j) * 0.7) * 1.15:
            utyazhelenie()
            kod = rastr.rgm('p')
        else:
            peretok_v_sechenii()
            break
    else:
        peretok_v_sechenii()
    return peretok_v_sechenii()


# Определение предельного перетока по критерию обеспечения
# нормативного коэффициента запаса статической
# апериодической устойчивости по активной мощности
# в контролируемом сечении в послеаварийных режимах
# после нормативных возмущений.
def pred_3(utyazhelenie, peretok_v_sechenii, kod):
    """Функция осуществляет определение предельного перетока
       по критерию обеспечения нормативного коэффициента запаса
       статической апериодической устойчивости по
       активной мощности в контролируемом сечении
       в послеаварийных режимах после нормативных возмущений.

    Аргументы:
        utyazhelenie: осуществляет утяжеление режима
        peretok_v_sechenii: определяет переток в сечении
        kod: осуществляет расчёт режима
    Возвращает:
        Предельный переток в сечении.
    """
    pred = {}
    pred_sta = {}
    index_vozmush_statica = list(index_vozmush_f())
    listic_sta_faults_statica = df_faults.sta.T.tolist()

    for j in range(len(index_vozmush_statica)):
        for (index, sta_faults) in zip(
                index_vozmush_statica, listic_sta_faults_statica):
            sta.SetZ(index, sta_faults)
            break
        index_vozmush_statica.remove(index)
        listic_sta_faults_statica.remove(sta_faults)
        while kod == 0:
            kod = rastr.rgm('p')
            if kod == 0:
                utyazhelenie()
            else:
                pred[index] = peretok_v_sechenii() * 0.92
                pred_sta[sta_faults] = peretok_v_sechenii() * 0.92
                sta.SetZ(index, abs(sta_faults - 1))
                kod = rastr.rgm('p')
                break
        while kod == 0:
            kod = rastr.rgm('p')
            if peretok_v_sechenii() > p_sech_nach_f() and \
                    len(index_vozmush_statica) > 0:
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
    kod = rastr.rgm('p')
    while kod == 0:
        for i in range(nodes_count):
            sta.SetZ(min_pred_index, min_pred_sta)
            break
        kod = rastr.rgm('p')
        if peretok_v_sechenii() > min_pred:
            obratnoe_utyazhelenie()
            kod = rastr.rgm('p')
        else:
            sta.SetZ(min_pred_index, abs(min_pred_sta - 1))
            rastr.rgm('p')
            peretok_v_sechenii()
            break
    return peretok_v_sechenii()


# Определение предельного перетока по критерию
# обеспечения нормативного коэффициента запаса
# статической устойчивости по напряжению
# в узлах нагрузки в послеаварийных режимах
# после нормативных возмущений.
def pred_4(utyazhelenie, peretok_v_sechenii, kod):
    """Функция осуществляет определение предельного перетока
       по критерию обеспечения нормативного коэффициента
       запаса статической устойчивости по напряжению
       в узлах нагрузки в послеаварийных режимах
       после нормативных возмущений.

    Аргументы:
        utyazhelenie: осуществляет утяжеление режима
        peretok_v_sechenii: определяет переток в сечении
        kod: осуществляет расчёт режима
    Возвращает:
        Предельный переток в сечении.
    """
    pred = {}
    index_vozmush_naprygenie = list(index_vozmush_f())
    listic_sta_faults_naprygenie = df_faults.sta.T.tolist()
    kod = rastr.rgm('p')

    for j in range(len(index_vozmush_naprygenie)):
        for (index, sta_faults) in zip(
                index_vozmush_naprygenie, listic_sta_faults_naprygenie):
            sta.SetZ(index, sta_faults)
            break
        index_vozmush_naprygenie.remove(index)
        listic_sta_faults_naprygenie.remove(sta_faults)
        while kod == 0:
            kod = rastr.rgm('p')
            for j in range(nodes_count):
                if uras.Z(j) < (unom.Z(j) * 0.7) * 1.1:
                    break
            if uras.Z(j) > (unom.Z(j) * 0.7) * 1.1:
                utyazhelenie()
                kod = rastr.rgm('p')
            else:
                sta.SetZ(index, abs(sta_faults - 1))
                kod = rastr.rgm('p')
                pred[index] = peretok_v_sechenii()
                peretok_v_sechenii()
                break
        else:
            sta.SetZ(index, abs(sta_faults - 1))
            kod = rastr.rgm('p')
            pred[index] = peretok_v_sechenii()
            peretok_v_sechenii()
        while kod == 0:
            kod = rastr.rgm('p')
            if peretok_v_sechenii() > p_sech_nach_f() and \
                    len(index_vozmush_naprygenie) > 0:
                obratnoe_utyazhelenie()
                kod = rastr.rgm('p')
            else:
                break

    for key, value in pred.items():
        min_pred_index = min(pred, key=pred.get)
    min_pred = pred[min_pred_index]
    return min_pred


# Определение предельного перетока по критерию обеспечения
# допустимой токовой нагрузки линий электропередачи и
# электросетевого оборудования в нормальной схеме.
def pred_5(utyazhelenie, peretok_v_sechenii, kod):
    """Функция осуществляет определение предельного перетока
       по критерию обеспечения допустимой токовой нагрузки
       линий электропередачи и электросетевого оборудования
       в нормальной схеме.

    Аргументы:
        utyazhelenie: осуществляет утяжеление режима
        peretok_v_sechenii: определяет переток в сечении
        kod: осуществляет расчёт режима
    Возвращает:
        Предельный переток в сечении.
    """
    while kod == 0:
        kod = rastr.rgm('p')
        for j in range(nodes_count):
            if (zag_i.Z(j) * 1000) >= 100 and (zag_it.Z(j) * 1000) >= 100:
                break
        if (zag_i.Z(j) * 1000) < 100 or (zag_it.Z(j) * 1000) < 100:
            utyazhelenie()
            kod = rastr.rgm('p')
        else:
            peretok_v_sechenii()
            break
    return peretok_v_sechenii()


# Определение предельного перетока по критерию обеспечения
# допустимой токовой нагрузки линий электропередачи и
# электросетевого оборудования в послеаварийных режимах
# после нормативных возмущений.
def pred_6(utyazhelenie, peretok_v_sechenii, kod):
    """Функция осуществляет определение предельного перетока
       по критерию обеспечения допустимой токовой нагрузки
       линий электропередачи и электросетевого оборудования
       в послеаварийных режимах после нормативных возмущений.

    Аргументы:
        utyazhelenie: осуществляет утяжеление режима
        peretok_v_sechenii: определяет переток в сечении
        kod: осуществляет расчёт режима
    Возвращает:
        Предельный переток в сечении.
    """
    pred = {}
    index_vozmush_tok = list(index_vozmush_f())
    listic_sta_faults_tok = df_faults.sta.T.tolist()

    for j in range(len(index_vozmush_tok)):
        for (index, sta_faults) in zip(
                index_vozmush_tok, listic_sta_faults_tok):
            sta.SetZ(index, sta_faults)
            break
        index_vozmush_tok.remove(index)
        listic_sta_faults_tok.remove(sta_faults)
        while kod == 0:
            kod = rastr.rgm('p')
            for j in range(nodes_count):
                if (zag_i_av.Z(j) * 1000) >= 100 and \
                        (zag_it_av.Z(j) * 1000) >= 100:
                    break
            if (zag_i_av.Z(j) * 1000) < 100 or \
                    (zag_it_av.Z(j) * 1000) < 100:
                utyazhelenie()
                kod = rastr.rgm('p')
            else:
                sta.SetZ(index, abs(sta_faults - 1))
                kod = rastr.rgm('p')
                pred[index] = peretok_v_sechenii()
                peretok_v_sechenii()
                break
        else:
            sta.SetZ(index, abs(sta_faults - 1))
            kod = rastr.rgm('p')
            pred[index] = peretok_v_sechenii()
            peretok_v_sechenii()
        while kod == 0:
            kod = rastr.rgm('p')
            if peretok_v_sechenii() > p_sech_nach_f() and \
                    len(index_vozmush_tok) > 0:
                obratnoe_utyazhelenie()
                kod = rastr.rgm('p')
            else:
                break

    for key, value in pred.items():
        min_pred_index = min(pred, key=pred.get)
    min_pred = pred[min_pred_index]
    return min_pred
