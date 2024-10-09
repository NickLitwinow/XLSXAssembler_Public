import pandas as pd
import re

def assemble_data(all_sheets_data):
    cad_df1 = pd.DataFrame()
    cad_df2 = pd.DataFrame()
    ecad_df1 = pd.DataFrame()
    ecad_df2 = pd.DataFrame()
    cae_df1 = pd.DataFrame()
    cae_df2 = pd.DataFrame()
    capp_df1 = pd.DataFrame()
    capp_df2 = pd.DataFrame()
    cam_df1 = pd.DataFrame()
    cam_df2 = pd.DataFrame()
    pdm_df1 = pd.DataFrame()
    pdm_df2 = pd.DataFrame()
    erp_df1 = pd.DataFrame()
    erp_df2 = pd.DataFrame()
    subu_df1 = pd.DataFrame()
    subu_df2 = pd.DataFrame()
    sb_df1 = pd.DataFrame()
    sb_df2 = pd.DataFrame()
    supr_df1 = pd.DataFrame()
    supr_df2 = pd.DataFrame()
    sup_df1 = pd.DataFrame()
    sup_df2 = pd.DataFrame()
    mrp2_df1 = pd.DataFrame()
    mrp2_df2 = pd.DataFrame()
    ils_df1 = pd.DataFrame()
    ils_df2 = pd.DataFrame()
    iatr_df1 = pd.DataFrame()
    iatr_df2 = pd.DataFrame()
    mdm_df1 = pd.DataFrame()
    mdm_df2 = pd.DataFrame()
    sad_df1 = pd.DataFrame()
    sad_df2 = pd.DataFrame()
    eam_df1 = pd.DataFrame()
    eam_df2 = pd.DataFrame()

    reglamenty = pd.DataFrame()
    kommunikazii = pd.DataFrame()
    cody = pd.DataFrame()
    skt = pd.DataFrame()
    obshesistemnoe_po_df1 = pd.DataFrame()
    obshesistemnoe_po_df2 = pd.DataFrame()
    intergracia_oborudovaniya = pd.DataFrame()
    sistemy_monitoringa = pd.DataFrame()
    standarty = pd.DataFrame()
    bi_sistemy = pd.DataFrame()
    ORD = pd.DataFrame()
    kd = pd.DataFrame()
    mzk = pd.DataFrame()
    kadry_1 = pd.DataFrame()
    kadry_2 = pd.DataFrame()
    bim = pd.DataFrame()
    ib = pd.DataFrame()

    pattern = r"\s*ИТОГ[ОA0]\s*\(\s*считается\s+[аa@]втоматически,\s+не\s+заполнять\s*\)\s*"
    pattern_ext = r"\s*Итог[оO0]\s*\(\s*заполняется\s+автоматически,\s+не\s+вводить\s+данные\s+вручную\s*\)\s*"
    pattern_amount = r"\s*Количество\s+уникальных\s+наименований\s+отечественного\s+ПО\s*\(\s*считается\s+автоматически,\s+не\s+заполнять\s*\)\s*"

    # Вывод информации о загруженных датафреймах
    for key, dataframe in all_sheets_data.items():
        if key.endswith("1.CAD"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index

            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            cad_df1 = pd.concat([cad_df1, df.loc[2:end_index - 1].iloc[:, :9]], ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                f'не найден индекс итого для {key}')  # Второе появление
            cad_df2 = pd.concat([cad_df2, df.loc[start_index:end_index - 1].iloc[:, :9]], ignore_index=True)

        elif key.endswith("2.ECAD"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index

            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            ecad_df1 = pd.concat([ecad_df1, df.loc[2:end_index - 1].iloc[:, :9]], ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                f'не найден индекс итого для {key}')  # Второе появление
            ecad_df2 = pd.concat([ecad_df2, df.loc[start_index:end_index - 1].iloc[:, :9]], ignore_index=True)

        elif key.endswith("3.CAE"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            cae_df1 = pd.concat([cae_df1,
                                 pd.concat(
                                     [df.loc[2:end_index - 1].iloc[:, :9], df.loc[2:end_index - 1].iloc[:, 11]],
                                     axis=1)], ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                f'не найден индекс итого для {key}')  # Второе появление
            cae_df2 = pd.concat([cae_df2, pd.concat(
                [df.loc[start_index:end_index - 1].iloc[:, :9], df.loc[start_index:end_index - 1].iloc[:, 11]],
                axis=1)], ignore_index=True)

        elif key.endswith("4.CAPP"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            capp_df1 = pd.concat([capp_df1, df.loc[2:end_index - 1].iloc[:, :9]], ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                f'не найден индекс итого для {key}')  # Второе появление
            capp_df2 = pd.concat([capp_df2, df.loc[start_index:end_index - 1].iloc[:, :9]], ignore_index=True)

        elif key.endswith("5.CAM"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'Не найден индекс итого для {key}')  # Первое появление
            cam_df1 = pd.concat([cam_df1, df.loc[2:end_index - 1].iloc[:, :9]], ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                f'не найден индекс итого для {key}')  # Второе появление
            cam_df2 = pd.concat([cam_df2, df.loc[start_index:end_index - 1].iloc[:, :9]], ignore_index=True)

        elif key.endswith("6.PDM"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            pdm_df1 = pd.concat([pdm_df1, df.loc[2:end_index - 1].iloc[:, :9]], ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                f'не найден индекс итого для {key}')  # Второе появление
            pdm_df2 = pd.concat([pdm_df2, df.loc[start_index:end_index - 1].iloc[:, :9]], ignore_index=True)

        elif key.endswith("7.ERP"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            erp_df1 = pd.concat([erp_df1, df.loc[2:end_index - 1].iloc[:, :9]], ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                f'не найден индекс итого для {key}')  # Второе появление
            erp_df2 = pd.concat([erp_df2, df.loc[start_index:end_index - 1].iloc[:, :9]], ignore_index=True)

        elif key.endswith("8.СУБУ"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            subu_df1 = pd.concat([subu_df1, df.loc[2:end_index - 1].iloc[:, :9]], ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                f'не найден индекс итого для {key}')  # Второе появление
            subu_df2 = pd.concat([subu_df2, df.loc[start_index:end_index - 1].iloc[:, :9]], ignore_index=True)

        elif key.endswith("9.СБ"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            sb_df1 = pd.concat([sb_df1, df.loc[2:end_index - 1].iloc[:, :9]], ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                f'не найден индекс итого для {key}')  # Второе появление
            sb_df2 = pd.concat([sb_df2, df.loc[start_index:end_index - 1].iloc[:, :9]], ignore_index=True)

        elif key.endswith("10.СУПР"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            supr_df1 = pd.concat([supr_df1, df.loc[2:end_index - 1].iloc[:, :9]], ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            if len(end_indexes) > 1:
                end_index = end_indexes[1]
            else:
                end_indexes = df[df.apply(
                    lambda row: row.astype(str).str.contains(pattern_amount, flags=re.IGNORECASE, regex=True).any(),
                    axis=1)].index
                end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                    f'не найден индекс итого для {key}')  # Второе появление
            supr_df2 = pd.concat([supr_df2, df.loc[start_index:end_index - 1].iloc[:, :9]], ignore_index=True)

        elif key.endswith("11.СУП"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            sup_df1 = pd.concat([sup_df1, df.loc[2:end_index - 1].iloc[:, :9]], ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                f'не найден индекс итого для {key}')  # Второе появление
            sup_df2 = pd.concat([sup_df2, df.loc[start_index:end_index - 1].iloc[:, :9]], ignore_index=True)

        elif key.endswith("12.MRPII"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            mrp2_df1 = pd.concat([mrp2_df1, df.loc[2:end_index - 1].iloc[:, :9]], ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                f'не найден индекс итого для {key}')  # Второе появление
            mrp2_df2 = pd.concat([mrp2_df2, df.loc[start_index:end_index - 1].iloc[:, :9]], ignore_index=True)

        elif key.endswith("13.ILS"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            ils_df1 = pd.concat([ils_df1, df.loc[2:end_index - 1].iloc[:, :9]], ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                f'не найден индекс итого для {key}')  # Второе появление
            ils_df2 = pd.concat([ils_df2, df.loc[start_index:end_index - 1].iloc[:, :9]], ignore_index=True)

        elif key.endswith("14.ПО для ИЭТР"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            iatr_df1 = pd.concat([iatr_df1, df.loc[2:end_index - 1].iloc[:, :9]], ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                f'не найден индекс итого для {key}')  # Второе появление
            iatr_df2 = pd.concat([iatr_df2, df.loc[start_index:end_index - 1].iloc[:, :9]], ignore_index=True)

        elif key.endswith("15.MDM"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            mdm_df1 = pd.concat([mdm_df1, pd.concat(
                [df.loc[2:end_index - 1].iloc[:, :11], df.loc[2:end_index - 1].iloc[:, 14:18]], axis=1)],
                                ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                f'не найден индекс итого для {key}')  # Второе появление
            mdm_df2 = pd.concat([mdm_df2, pd.concat(
                [df.loc[start_index:end_index - 1].iloc[:, :11], df.loc[start_index:end_index - 1].iloc[:, 14:18]],
                axis=1)], ignore_index=True)

        elif key.endswith("16.СЭД"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            sad_df1 = pd.concat([sad_df1, df.loc[2:end_index - 1].iloc[:, :9]], ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                f'не найден индекс итого для {key}')  # Второе появление
            sad_df2 = pd.concat([sad_df2, df.loc[start_index:end_index - 1].iloc[:, :9]], ignore_index=True)

        elif key.endswith("17.EAM"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            eam_df1 = pd.concat([eam_df1, df.loc[2:end_index - 1].iloc[:, :9]], ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                f'не найден индекс итого для {key}')  # Второе появление
            eam_df2 = pd.concat([eam_df2, df.loc[start_index:end_index - 1].iloc[:, :9]], ignore_index=True)

        elif key.endswith("18.Регламенты"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            reglamenty = pd.concat([reglamenty, df.loc[1:end_index - 1]], ignore_index=True)

        elif key.endswith("19.Коммуникации"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            kommunikazii = pd.concat([kommunikazii, df.loc[0:5]], ignore_index=True)
            # Добавляем две пустые строки в конец
            kommunikazii = pd.concat([kommunikazii, pd.DataFrame([[None] * len(kommunikazii.columns)])],
                                     ignore_index=True)

        elif key.endswith("20.ЦОДы"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            cody = pd.concat([cody, df.loc[1:end_index - 1].iloc[:, :15]], ignore_index=True)

        elif key.endswith("21.СКТ"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            skt = pd.concat(
                [skt,
                 pd.concat([df.loc[1:end_index - 1].iloc[:, :7], df.loc[1:end_index - 1].iloc[:, 8:10]], axis=1)],
                ignore_index=True)

        elif key.endswith("22.Общесистемное ПО"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index

            # Отечественное ПО
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            obshesistemnoe_po_df1 = pd.concat([obshesistemnoe_po_df1, df.loc[2:end_index - 1].iloc[:, :8]],
                                              ignore_index=True)

            # Зарубежное ПО
            start_index = df[df.iloc[:, 0] == 'Зарубежное ПО'].index[0] + 1
            end_index = end_indexes[1] if len(end_indexes) > 1 else print(
                f'не найден индекс итого для {key}')  # Второе появление
            obshesistemnoe_po_df2 = pd.concat(
                [obshesistemnoe_po_df2, df.loc[start_index:end_index - 1].iloc[:, :8]],
                ignore_index=True)

        elif key.endswith("23. Интеграция оборудования"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_index = df.index[-1]
            intergracia_oborudovaniya = pd.concat([intergracia_oborudovaniya, df.loc[1:end_index].iloc[:, :10]],
                                                  ignore_index=True)

        elif key.endswith("24.Системы мониторинга"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(
                    lambda row: row.astype(str).str.contains(pattern_ext, flags=re.IGNORECASE, regex=True).any(),
                    axis=1)].index
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            sistemy_monitoringa = pd.concat([sistemy_monitoringa, df.loc[2:end_index - 1].iloc[:, :51]],
                                            ignore_index=True)

        elif key.endswith("25. Стандарты"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            standarty = pd.concat([standarty, df.loc[1:end_index - 1].iloc[:, :5]], ignore_index=True)

        elif key.endswith("26.BI-системы"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_index = df.index[-1]
            bi_sistemy = pd.concat([bi_sistemy, df.loc[1:end_index].iloc[:, :11]], ignore_index=True)

        elif key.endswith("27.ОРД"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            ORD = pd.concat([ORD, df.loc[1:end_index - 1].iloc[:, :7]], ignore_index=True)

        elif key.endswith("28.КД"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            kd = pd.concat([kd, df.loc[1:end_index - 1].iloc[:, :10]], ignore_index=True)

        elif key.endswith("29.МЗК"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            mzk = pd.concat([mzk, df.loc[1:end_index - 1].iloc[:, :9]], ignore_index=True)

        elif key.endswith("30.Кадры 1"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            kadry_1 = pd.concat([kadry_1, df.loc[1:end_index - 1].iloc[:, :18]], ignore_index=True)

        elif key.endswith("31.Кадры 2"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            kadry_2 = pd.concat([kadry_2, df.loc[1:end_index - 1].iloc[:, :4]], ignore_index=True)

        elif key.endswith("32.BIM"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            bim = pd.concat([bim, df.loc[1:end_index - 1].iloc[:, :11]], ignore_index=True)

        elif key.endswith("33.ИБ"):
            print(f"DataFrame '{key}':")
            df = all_sheets_data[key]
            end_indexes = df[
                df.apply(lambda row: row.astype(str).str.contains(pattern, flags=re.IGNORECASE, regex=True).any(),
                         axis=1)].index
            end_index = end_indexes[0] if len(end_indexes) > 0 else print(
                f'не найден индекс итого для {key}')  # Первое появление
            ib = pd.concat([ib, df.loc[1:end_index - 1].iloc[:, :26]], ignore_index=True)

        else:
            continue

    dataframes = [cad_df1, cad_df2,
                  ecad_df1, ecad_df2,
                  cae_df1, cae_df2,
                  capp_df1, capp_df2,
                  cam_df1, cam_df2,
                  pdm_df1, pdm_df2,
                  erp_df1, erp_df2,
                  subu_df1, subu_df2,
                  sb_df1, sb_df2,
                  supr_df1, supr_df2,
                  sup_df1, sup_df2,
                  mrp2_df1, mrp2_df2,
                  ils_df1, ils_df2,
                  iatr_df1, iatr_df2,
                  mdm_df1, mdm_df2,
                  sad_df1, sad_df2,
                  eam_df1, eam_df2,
                  obshesistemnoe_po_df1, obshesistemnoe_po_df2]

    # Перебор каждого датафрейма в списке
    for df in dataframes:
        # Условие, которое проверяет, пустые ли первые четыре столбца
        mask = df.iloc[:, 0:4].isnull().all(axis=1)
        # Удаление строк, где условие истинно
        df.drop(index=df[mask].index, inplace=True)
        # Дополнительно: удаление строк, где все значения пусты
        df.dropna(how='all', inplace=True)

    dataframes_base = [reglamenty, kommunikazii, cody, skt, intergracia_oborudovaniya, sistemy_monitoringa, standarty, bi_sistemy,
                       ORD, kd, mzk, kadry_1, kadry_2, bim, ib]

    for df in dataframes_base:
        # Удаление строк, где первый столбец пуст
        df.dropna(subset=[df.columns[0]], inplace=True)
        # Дополнительно: удаление строк, где все значения пусты
        df.dropna(how='all', inplace=True)

    dfs = [dataframes, dataframes_base]
    return dfs
