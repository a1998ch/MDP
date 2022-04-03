import function_mdp as fm
import pandas as pd
import win32com.client
rastr = win32com.client.Dispatch('Astra.Rastr')


# Осуществляем загрузку файлов режима, траектории утяжеления,
# контролируемого сечения и нормативных возмущений.
regime = rastr.Load(1, 'regime.rg2', '')
flowgate = pd.read_json('flowgate.json').transpose()
faults = pd.read_json('faults.json').transpose()
vector = pd.read_csv('vector.csv')

# Количество узлов и ветвей в заданном режиме
nodes_count = rastr.Tables("node").size
vetv_count = rastr.Tables("vetv").size

# Величина амплитуды нерегулярных колебаний.
p_nk = 30

# МДП по критерию обеспечения нормативного коэффициента запаса
# статической апериодической устойчивости по активной мощности
# в контролируемом сечении в нормальной схеме.
p_pred1 = fm.pred_1(nodes_count, rastr, vector, vetv_count, flowgate)
p_mdp1 = 0.8 * p_pred1 - p_nk


fm.vozvrat_k_ishodnomu_regimu(nodes_count, rastr, vector, vetv_count, flowgate)
rastr.rgm('p')


# МДП по критерию обеспечения нормативного коэффициента запаса
# статической устойчивости по напряжению
# в узлах нагрузки в нормальной схеме.
p_pred2 = fm.pred_2(nodes_count, rastr, vector, vetv_count, flowgate)
p_mdp2 = p_pred2 - p_nk


fm.vozvrat_k_ishodnomu_regimu(nodes_count, rastr, vector, vetv_count, flowgate)
rastr.rgm('p')


# МДП по критерию обеспечения нормативного коэффициента запаса
# статической апериодической устойчивости по активной мощности
# в контролируемом сечении в послеаварийных режимах
# после нормативных возмущений.
p_pred3 = fm.pred_3(nodes_count, rastr, vector, vetv_count, flowgate, faults)
p_mdp3 = p_pred3 - p_nk


fm.vozvrat_k_ishodnomu_regimu(nodes_count, rastr, vector, vetv_count, flowgate)
rastr.rgm('p')


# МДП по критерию обеспечения нормативного коэффициента запаса
# статической устойчивости по напряжению в узлах нагрузки
# в послеаварийных режимах после нормативных возмущений.
p_pred4 = fm.pred_4(nodes_count, rastr, vector, vetv_count, flowgate, faults)
p_mdp4 = p_pred4 - p_nk


fm.vozvrat_k_ishodnomu_regimu(nodes_count, rastr, vector, vetv_count, flowgate)
rastr.rgm('p')


# МДП по критерию обеспечения допустимой токовой нагрузки
# линий электропередачи и электросетевого оборудования
# в нормальной схеме.
p_pred5 = fm.pred_5(nodes_count, rastr, vector, vetv_count, flowgate)
p_mdp5 = p_pred5 - p_nk


fm.vozvrat_k_ishodnomu_regimu(nodes_count, rastr, vector, vetv_count, flowgate)
rastr.rgm('p')


# МДП по критерию обеспечения допустимой токовой нагрузки
# линий электропередачи и электросетевого оборудования
# в послеаварийных режимах после нормативных возмущений.
p_pred6 = fm.pred_6(nodes_count, rastr, vector, vetv_count, flowgate, faults)
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
