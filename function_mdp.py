def nodes_index(nodes_count, rastr):
    """Функция определяет индексы узлов режима

    Аргументы:
        nodes_count: Количество узлов в заданном режиме.
        rastr: библиотека rastrWin.
    Возвращает:
        словарь узлов и их индексов режима.
    """
    if not hasattr(nodes_index, '_state'):
        slovaric_nodes_index = {}
        ny = rastr.Tables("node").Cols("ny")
        pn = rastr.Tables("node").Cols("pn")
        qn = rastr.Tables("node").Cols("qn")
        tg = rastr.Tables("node").Cols("tg_phi")
        tip = rastr.Tables("node").Cols("tip")
        for i in range(nodes_count):
            slovaric_nodes_index[ny.Z(i)] = i
            if tip.Z(i) == 1 and pn.Z(i) != 0:
                tg.SetZ(i, qn.Z(i) / pn.Z(i))
        nodes_index._state = slovaric_nodes_index
    return nodes_index._state


def index_vetv_sech(rastr, vetv_count, flowgate):
    """Функция определяет индексы ветвей в сечении

    Аргументы:
        vetv_count: Количество ветвей в заданном режиме.
        rastr: библиотека rastrWin.
        flowgate: контролируемое сечение
    Возвращает:
        словарь узлов и их индексов режима.
    """
    if not hasattr(index_vetv_sech, '_state'):
        index_vetv = []
        n_nach = rastr.Tables("vetv").Cols("ip")
        n_kon = rastr.Tables("vetv").Cols("iq")
        np = rastr.Tables("vetv").Cols("np")
        for i in range(vetv_count):
            for j in range(1, len(flowgate.index) + 1):
                if n_nach.Z(i) == flowgate.at[f'line_{j}', 'ip'] and \
                    n_kon.Z(i) == flowgate.at[f'line_{j}', 'iq'] and \
                        np.Z(i) == flowgate.at[f'line_{j}', 'np']:
                    index_vetv.append(i)
        index_vetv_sech._state = index_vetv
    return index_vetv_sech._state


def p_sech_nach(rastr, vetv_count, flowgate):
    """Функция определяет переток в сечении
       в исходном режиме

    Аргументы:
        vetv_count: Количество ветвей в заданном режиме.
        rastr: библиотека rastrWin.
        flowgate: контролируемое сечение
    Возвращает:
        переток в сечении в исходном режиме.
    """
    if not hasattr(p_sech_nach, '_state'):
        p_sech = 0
        n_nach = rastr.Tables("vetv").Cols("ip")
        n_kon = rastr.Tables("vetv").Cols("iq")
        np = rastr.Tables("vetv").Cols("np")
        p_nach = rastr.Tables("vetv").Cols("pl_ip")
        for i in range(vetv_count):
            for j in range(1, len(flowgate.index) + 1):
                if n_nach.Z(i) == flowgate.at[f'line_{j}', 'ip'] and \
                    n_kon.Z(i) == flowgate.at[f'line_{j}', 'iq'] and \
                        np.Z(i) == flowgate.at[f'line_{j}', 'np']:
                    p_sech += p_nach.Z(i)
        p_sech_nach._state = p_sech
    return p_sech_nach._state


def index_vozmush(rastr, vetv_count, faults):
    """Функция определяет индексы ветвей,
       которые входят в нормативные возмущения

    Аргументы:
        vetv_count: Количество ветвей в заданном режиме.
        rastr: библиотека rastrWin.
        faults: нормативные возмущения
    Возвращает:
        индексы ветвей, которые входят
        в нормативные возмущения.
    """
    if not hasattr(index_vozmush, '_state'):
        index_vozmush_p = []
        n_nach = rastr.Tables("vetv").Cols("ip")
        n_kon = rastr.Tables("vetv").Cols("iq")
        np = rastr.Tables("vetv").Cols("np")
        for i in range(vetv_count):
            for k in range(len(faults.index)):
                if n_nach.Z(i) == faults.iloc[k]['ip'] and \
                    n_kon.Z(i) == faults.iloc[k]['iq'] and \
                        np.Z(i) == faults.iloc[k]['np']:
                    index_vozmush_p.append(i)
        index_vozmush._state = index_vozmush_p
    return index_vozmush._state


def utyazhelenie(nodes_count, rastr, vector):
    """Функция осуществляет изменение мощности нагрузки
       и генерации в соответствии с заданной траекторией
       утяжеления

    Аргументы:
        nodes_count: Количество узлов в заданном режиме.
        rastr: библиотека rastrWin.
        vector: траектория утяжеления.
    Возвращает:
        Изменённый режим.
    """
    pn = rastr.Tables("node").Cols("pn")
    pg = rastr.Tables("node").Cols("pg")
    qn = rastr.Tables("node").Cols("qn")
    tg = rastr.Tables("node").Cols("tg_phi")
    for j, i in enumerate(vector.node.T.tolist()):
        if vector.iloc[j]['variable'] == 'pn':
            pn.SetZ(
                nodes_index(nodes_count, rastr).get(i), pn.Z(
                    nodes_index(
                        nodes_count, rastr).get(i)) + vector.iloc[j]['value'])
            if vector.iloc[j]['tg'] == 1:
                qn.SetZ(
                    nodes_index(nodes_count, rastr).get(i), tg.Z(
                        nodes_index(nodes_count, rastr).get(i)) * pn.Z(
                            nodes_index(nodes_count, rastr).get(i)))
        else:
            pg.SetZ(
                nodes_index(nodes_count, rastr).get(i), pg.Z(
                    nodes_index(
                        nodes_count, rastr).get(i)) + vector.iloc[j]['value'])


def obratnoe_utyazhelenie(nodes_count, rastr, vector):
    """Функция осуществляет изменение мощности нагрузки
       и генерации в соответствии с заданной траекторией
       утяжеления

    Аргументы:
        nodes_count: Количество узлов в заданном режиме.
        rastr: библиотека rastrWin.
        vector: траектория утяжеления.
    Возвращает:
        Изменённый режим.
    """
    pn = rastr.Tables("node").Cols("pn")
    pg = rastr.Tables("node").Cols("pg")
    qn = rastr.Tables("node").Cols("qn")
    tg = rastr.Tables("node").Cols("tg_phi")
    for j, i in enumerate(vector.node.T.tolist()):
        if vector.iloc[j]['variable'] == 'pn':
            pn.SetZ(
                nodes_index(nodes_count, rastr).get(i), pn.Z(
                    nodes_index(
                        nodes_count, rastr).get(i)) - vector.iloc[j]['value'])
            if vector.iloc[j]['tg'] == 1:
                qn.SetZ(
                    nodes_index(nodes_count, rastr).get(i), tg.Z(
                        nodes_index(nodes_count, rastr).get(i)) * pn.Z(
                            nodes_index(nodes_count, rastr).get(i)))
        else:
            pg.SetZ(
                nodes_index(nodes_count, rastr).get(i), pg.Z(
                    nodes_index(
                        nodes_count, rastr).get(i)) - vector.iloc[j]['value'])


def vozvrat_k_ishodnomu_regimu(
        nodes_count, rastr, vector, vetv_count, flowgate):
    """Функция осуществляет изменение мощности нагрузки
       и генерации до достижения исходного режима

    Аргументы:
        nodes_count: Количество узлов в заданном режиме.
        rastr: библиотека rastrWin.
        vector: траектория утяжеления.
        vetv_count: Количество ветвей в заданном режиме.
        flowgate: контролируемое сечение
    Возвращает:
        Исходный режим.
    """
    while True:
        rastr.rgm('p')
        if peretok_v_sechenii(
            rastr, vetv_count, flowgate) > p_sech_nach(
                rastr, vetv_count, flowgate):
            obratnoe_utyazhelenie(nodes_count, rastr, vector)
            rastr.rgm('p')
        else:
            break


def peretok_v_sechenii(rastr, vetv_count, flowgate):
    """Функция осуществляет определение
       перетока в сечении.

    Аргументы:
        rastr: библиотека rastrWin.
        vetv_count: Количество ветвей в заданном режиме.
        flowgate: контролируемое сечение
    Возвращает:
        Величину перетока в сечении.
    """
    p_nach = rastr.Tables("vetv").Cols("pl_ip")
    sechenie = 0
    for j in index_vetv_sech(rastr, vetv_count, flowgate):
        sechenie += p_nach.Z(j)
    return sechenie


def pred_1(nodes_count, rastr, vector, vetv_count, flowgate):
    """Функция осуществляет определение предельного перетока
       по критерию обеспечения нормативного коэффициента
       запаса статической апериодической устойчивости по
       активной мощности в контролируемом сечении
       в нормальной схеме.

    Аргументы:
        nodes_count: Количество узлов в заданном режиме.
        rastr: библиотека rastrWin.
        vector: траектория утяжеления.
        vetv_count: Количество ветвей в заданном режиме.
        flowgate: контролируемое сечение
    Возвращает:
        Предельный переток в сечении.
    """
    p_sech_nach(rastr, vetv_count, flowgate)
    while rastr.rgm('p') == 0:
        utyazhelenie(nodes_count, rastr, vector)
        rastr.rgm('p')
    return peretok_v_sechenii(rastr, vetv_count, flowgate)


def pred_2(nodes_count, rastr, vector, vetv_count, flowgate):
    """Функция осуществляет определение предельного перетока
       по критерию обеспечения нормативного коэффициента
       запаса статической устойчивости по напряжению
       в узлах нагрузки в нормальной схеме.

    Аргументы:
        nodes_count: Количество узлов в заданном режиме.
        rastr: библиотека rastrWin.
        vector: траектория утяжеления.
        vetv_count: Количество ветвей в заданном режиме.
        flowgate: контролируемое сечение
    Возвращает:
        Предельный переток в сечении.
    """
    unom = rastr.Tables("node").Cols("uhom")
    uras = rastr.Tables("node").Cols("vras")
    while rastr.rgm('p') == 0:
        rastr.rgm('p')
        for j in range(nodes_count):
            if uras.Z(j) < (unom.Z(j) * 0.7) * 1.15:
                break
        if uras.Z(j) > (unom.Z(j) * 0.7) * 1.15:
            utyazhelenie(nodes_count, rastr, vector)
            rastr.rgm('p')
        else:
            break
    return peretok_v_sechenii(rastr, vetv_count, flowgate)


def pred_3(nodes_count, rastr, vector, vetv_count, flowgate, faults):
    """Функция осуществляет определение предельного перетока
       по критерию обеспечения нормативного коэффициента запаса
       статической апериодической устойчивости по
       активной мощности в контролируемом сечении
       в послеаварийных режимах после нормативных возмущений.

    Аргументы:
        nodes_count: Количество узлов в заданном режиме.
        rastr: библиотека rastrWin.
        vector: траектория утяжеления.
        vetv_count: Количество ветвей в заданном режиме.
        flowgate: контролируемое сечение
        faults: нормативные возмущения
    Возвращает:
        Предельный переток в сечении.
    """
    sta = rastr.Tables("vetv").Cols("sta")
    pred = []
    index_vozmush_statica = list(index_vozmush(rastr, vetv_count, faults))
    listic_sta_faults_statica = faults.sta.T.tolist()

    for j in range(len(index_vozmush_statica)):
        sta.SetZ(
            index_vozmush_statica[0], listic_sta_faults_statica[0])
        while rastr.rgm('p') == 0:
            rastr.rgm('p')
            utyazhelenie(nodes_count, rastr, vector)
        else:
            pred.append(peretok_v_sechenii(
                rastr, vetv_count, flowgate) * 0.92)
            sta.SetZ(index_vozmush_statica[0], abs(
                listic_sta_faults_statica[0] - 1))
            index_vozmush_statica.pop(0)
            listic_sta_faults_statica.pop(0)
            rastr.rgm('p')
        if len(index_vozmush_statica) > 0:
            vozvrat_k_ishodnomu_regimu(
                nodes_count, rastr, vector, vetv_count, flowgate)
    return min(pred)


def pred_4(nodes_count, rastr, vector, vetv_count, flowgate, faults):
    """Функция осуществляет определение предельного перетока
       по критерию обеспечения нормативного коэффициента
       запаса статической устойчивости по напряжению
       в узлах нагрузки в послеаварийных режимах
       после нормативных возмущений.

    Аргументы:
        nodes_count: Количество узлов в заданном режиме.
        rastr: библиотека rastrWin.
        vector: траектория утяжеления.
        vetv_count: Количество ветвей в заданном режиме.
        flowgate: контролируемое сечение
        faults: нормативные возмущения
    Возвращает:
        Предельный переток в сечении.
    """
    sta = rastr.Tables("vetv").Cols("sta")
    unom = rastr.Tables("node").Cols("uhom")
    uras = rastr.Tables("node").Cols("vras")
    pred = []
    index_vozmush_naprygenie = list(index_vozmush(rastr, vetv_count, faults))
    listic_sta_faults_naprygenie = faults.sta.T.tolist()

    for j in range(len(index_vozmush_naprygenie)):
        sta.SetZ(
            index_vozmush_naprygenie[0], listic_sta_faults_naprygenie[0])
        while rastr.rgm('p') == 0:
            rastr.rgm('p')
            for j in range(nodes_count):
                if uras.Z(j) < (unom.Z(j) * 0.7) * 1.1:
                    break
            if uras.Z(j) > (unom.Z(j) * 0.7) * 1.1:
                utyazhelenie(nodes_count, rastr, vector)
                rastr.rgm('p')
            else:
                sta.SetZ(index_vozmush_naprygenie[0], abs(
                    listic_sta_faults_naprygenie[0] - 1))
                rastr.rgm('p')
                pred.append(peretok_v_sechenii(rastr, vetv_count, flowgate))
                index_vozmush_naprygenie.pop(0)
                listic_sta_faults_naprygenie.pop(0)
                break
        else:
            sta.SetZ(index_vozmush_naprygenie[0], abs(
                listic_sta_faults_naprygenie[0] - 1))
            rastr.rgm('p')
            pred.append(peretok_v_sechenii(rastr, vetv_count, flowgate))
            index_vozmush_naprygenie.pop(0)
            listic_sta_faults_naprygenie.pop(0)
        if len(index_vozmush_naprygenie) > 0:
            vozvrat_k_ishodnomu_regimu(
                nodes_count, rastr, vector, vetv_count, flowgate)
    return min(pred)


def pred_5(nodes_count, rastr, vector, vetv_count, flowgate):
    """Функция осуществляет определение предельного перетока
       по критерию обеспечения допустимой токовой нагрузки
       линий электропередачи и электросетевого оборудования
       в нормальной схеме.

    Аргументы:
        nodes_count: Количество узлов в заданном режиме.
        rastr: библиотека rastrWin.
        vector: траектория утяжеления.
        vetv_count: Количество ветвей в заданном режиме.
        flowgate: контролируемое сечение
    Возвращает:
        Предельный переток в сечении.
    """
    zag_i = rastr.Tables("vetv").Cols("zag_i")
    zag_it = rastr.Tables("vetv").Cols("zag_it")
    while rastr.rgm('p') == 0:
        rastr.rgm('p')
        for j in range(nodes_count):
            if (zag_i.Z(j) * 1000) >= 100 and (zag_it.Z(j) * 1000) >= 100:
                break
        if (zag_i.Z(j) * 1000) < 100 or (zag_it.Z(j) * 1000) < 100:
            utyazhelenie(nodes_count, rastr, vector)
            rastr.rgm('p')
        else:
            break
    return peretok_v_sechenii(rastr, vetv_count, flowgate)


def pred_6(nodes_count, rastr, vector, vetv_count, flowgate, faults):
    """Функция осуществляет определение предельного перетока
       по критерию обеспечения допустимой токовой нагрузки
       линий электропередачи и электросетевого оборудования
       в послеаварийных режимах после нормативных возмущений.

    Аргументы:
        nodes_count: Количество узлов в заданном режиме.
        rastr: библиотека rastrWin.
        vector: траектория утяжеления.
        vetv_count: Количество ветвей в заданном режиме.
        flowgate: контролируемое сечение
        faults: нормативные возмущения
    Возвращает:
        Предельный переток в сечении.
    """
    sta = rastr.Tables("vetv").Cols("sta")
    zag_i_av = rastr.Tables("vetv").Cols("zag_i_av")
    zag_it_av = rastr.Tables("vetv").Cols("zag_it_av")
    pred = []
    index_vozmush_tok = list(index_vozmush(rastr, vetv_count, faults))
    listic_sta_faults_tok = faults.sta.T.tolist()

    for j in range(len(index_vozmush_tok)):
        sta.SetZ(index_vozmush_tok[0], listic_sta_faults_tok[0])
        while rastr.rgm('p') == 0:
            rastr.rgm('p')
            for j in range(nodes_count):
                if (zag_i_av.Z(j) * 1000) >= 100 and \
                        (zag_it_av.Z(j) * 1000) >= 100:
                    break
            if (zag_i_av.Z(j) * 1000) < 100 or \
                    (zag_it_av.Z(j) * 1000) < 100:
                utyazhelenie(nodes_count, rastr, vector)
                rastr.rgm('p')
            else:
                sta.SetZ(index_vozmush_tok[0], abs(
                    listic_sta_faults_tok[0] - 1))
                rastr.rgm('p')
                pred.append(
                    peretok_v_sechenii(rastr, vetv_count, flowgate))
                index_vozmush_tok.pop(0)
                listic_sta_faults_tok.pop(0)
                break
        else:
            sta.SetZ(index_vozmush_tok[0], abs(
                listic_sta_faults_tok[0] - 1))
            rastr.rgm('p')
            pred.append(peretok_v_sechenii(rastr, vetv_count, flowgate))
            index_vozmush_tok.pop(0)
            listic_sta_faults_tok.pop(0)
        if len(index_vozmush_tok) > 0:
            vozvrat_k_ishodnomu_regimu(
                nodes_count, rastr, vector, vetv_count, flowgate)
    return min(pred)
