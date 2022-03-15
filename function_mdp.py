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
listic_nodes = []
for i in range(nodes_count):
    listic_nodes.append(ny.Z(i))
    if tip.Z(i) == 1 and pn.Z(i) != 0:
        tg.SetZ(i, qn.Z(i) / pn.Z(i))

# Определяем индексы узлов в заданном режиме
slovaric_nodes_index = {}
for i, num in enumerate(listic_nodes):
    slovaric_nodes_index[num] = i

# Переток в сечении в исходном режиме,
# индексы ветвей в сечении,
# индексы ветвей, эксплуатационное состояние
# которых требуется изменить
# Определение исходного состояния ветвей
index_vetv_sech = []
index_vozmush = []
listic_sta_faults_ischodnii = []
p_sech_nach = 0

for i in range(vetv_count):
    for j in range(1, len(df_flowgate.index) + 1):
        if n_nach.Z(i) == df_flowgate.at[f'line_{j}', 'ip'] and \
         n_kon.Z(i) == df_flowgate.at[f'line_{j}', 'iq'] and \
         np.Z(i) == df_flowgate.at[f'line_{j}', 'np']:
            p_sech_nach += p_nach.Z(i)
            index_vetv_sech.append(i)
            j += 1
    for k in range(len(df_faults.index)):
        if n_nach.Z(i) == df_faults.iloc[k]['ip'] and \
         n_kon.Z(i) == df_faults.iloc[k]['iq'] and \
         np.Z(i) == df_faults.iloc[k]['np']:
            index_vozmush.append(i)
            listic_sta_faults_ischodnii.append(sta.Z(i))
            k += 1


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
                slovaric_nodes_index.get(i), pn.Z(
                    slovaric_nodes_index.get(i)) + df_vector.iloc[j]['value'])
            if df_vector.iloc[j]['tg'] == 1:
                qn.SetZ(
                    slovaric_nodes_index.get(i), tg.Z(
                        slovaric_nodes_index.get(i)) * pn.Z(
                            slovaric_nodes_index.get(i)))
        else:
            pg.SetZ(
                slovaric_nodes_index.get(i), pg.Z(
                    slovaric_nodes_index.get(i)) + df_vector.iloc[j]['value'])


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
                slovaric_nodes_index.get(i), pn.Z(
                    slovaric_nodes_index.get(i)) - df_vector.iloc[j]['value'])
            if df_vector.iloc[j]['tg'] == 1:
                qn.SetZ(
                    slovaric_nodes_index.get(i), tg.Z(
                        slovaric_nodes_index.get(i)) * pn.Z(
                            slovaric_nodes_index.get(i)))
        else:
            pg.SetZ(
                slovaric_nodes_index.get(i), pg.Z(
                    slovaric_nodes_index.get(i)) - df_vector.iloc[j]['value'])


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
        if peretok_v_sechenii() > p_sech_nach:
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
    for j in index_vetv_sech:
        sechenie += p_nach.Z(j)
    return sechenie
