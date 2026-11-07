import streamlit as st
import pandas as pd
import math
import io
from openpyxl.styles import Font, Alignment

st.set_page_config(page_title="Сборка заказа Додо", layout="wide")
st.title("Автоматизация заказов: Додо Магадан")
st.markdown("Загрузите выгрузки Dodo IS и бланк поставщика, чтобы рассчитать заказ.")

if 'calculated_data' not in st.session_state:
    st.session_state['calculated_data'] = None

st.header("Шаг 1. Загрузка данных")
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("**Пиццерии**")
    file_m1 = st.file_uploader("Магадан-1", type=["xlsx"])
    file_m2 = st.file_uploader("Магадан-2", type=["xlsx"])
    file_m3 = st.file_uploader("Магадан-3", type=["xlsx"])

with col2:
    st.markdown("**Склады ПРЦ (Заказы)**")
    file_prc1 = st.file_uploader("ПРЦ-1 (Заказы)", type=["xlsx"])
    file_prc2 = st.file_uploader("ПРЦ-2 (Заказы)", type=["xlsx"])

with col3:
    st.markdown("**Справочники**")
    file_supplier = st.file_uploader("Бланк поставщика", type=["xlsx"])
    file_warehouse = st.file_uploader("Складские остатки (для сверки)", type=["xlsx"],
                                      help="Файл для проверки расхождений и отрицательных остатков")

if st.button("🚀 Рассчитать заказ", type="primary"):
    if not all([file_m1, file_m2, file_m3, file_prc1, file_prc2, file_supplier]):
        st.warning("Пожалуйста, загрузите все основные 6 файлов!")
    else:
        with st.spinner('Считаем...'):
            # 1. ЧТЕНИЕ
            df_m1 = pd.read_excel(file_m1, skiprows=9)
            df_m2 = pd.read_excel(file_m2, skiprows=9)
            df_m3 = pd.read_excel(file_m3, skiprows=9)
            df_supplier = pd.read_excel(file_supplier, skiprows=12)
            df_prc1 = pd.read_excel(file_prc1, skiprows=9)
            df_prc2 = pd.read_excel(file_prc2, skiprows=9)

            df_supplier.columns = df_supplier.columns.astype(str).str.strip()

            # 2. ОБЪЕДИНЕНИЕ
            df_all = pd.concat([df_m1, df_m2, df_m3], ignore_index=True)
            df_all = df_all[df_all['Надо заказать'] > 0]
            df_combined = df_all.groupby(['Код тары', 'Сырьё'], as_index=False).agg({
                'Надо заказать': 'sum',
                'Вместимость': 'max'
            })

            # 3. ОСТАТКИ ПРЦ
            df_prc_all = pd.concat([df_prc1, df_prc2], ignore_index=True)
            df_prc = df_prc_all.groupby('Код тары', as_index=False)['Текущий остаток'].sum()

            df_combined['Код тары'] = df_combined['Код тары'].astype(str).str.strip()
            df_prc['Код тары'] = df_prc['Код тары'].astype(str).str.strip()

            df_combined = pd.merge(df_combined, df_prc, on='Код тары', how='left')
            df_combined['Текущий остаток'] = df_combined['Текущий остаток'].fillna(0)

            # --- СВЕРКА ---
            df_combined['Складской остаток'] = 0.0
            if file_warehouse:
                try:
                    wh_sheets = pd.read_excel(file_warehouse, sheet_name=None)
                    df_wh = pd.concat(wh_sheets.values(), ignore_index=True)
                    df_wh.columns = df_wh.columns.astype(str).str.strip()
                    df_wh['Наименование товара'] = df_wh['Наименование товара'].astype(str).str.strip().str.lower()
                    df_wh = df_wh.dropna(subset=['Остаток'])
                    df_wh = df_wh.groupby('Наименование товара', as_index=False)['Остаток'].sum()

                    df_combined['Сырьё_lower'] = df_combined['Сырьё'].astype(str).str.strip().str.lower()
                    df_combined = pd.merge(df_combined, df_wh, left_on='Сырьё_lower', right_on='Наименование товара',
                                           how='left')
                    df_combined['Складской остаток'] = df_combined['Остаток'].fillna(0)
                    df_combined = df_combined.drop(columns=['Сырьё_lower', 'Наименование товара', 'Остаток'])
                except Exception as e:
                    st.warning(f"Не удалось прочитать файл складских остатков: {e}")

            df_combined['Отрицательный остаток'] = (df_combined['Текущий остаток'] < 0) | (
                        df_combined['Складской остаток'] < 0)
            df_combined['Расхождение источников'] = (df_combined['Текущий остаток'] != df_combined[
                'Складской остаток']) & (df_combined['Складской остаток'] != 0)

            df_combined['Текущий остаток'] = df_combined['Текущий остаток'].clip(lower=0)
            df_combined['Чистая потребность'] = (df_combined['Надо заказать'] - df_combined['Текущий остаток']).apply(
                lambda x: max(0, x))
            df_combined = df_combined[df_combined['Чистая потребность'] > 0]

            # 4. СВЯЗЬ С БЛАНКОМ
            df_supplier['Товар №'] = df_supplier['Товар №'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            sup_map_1 = df_supplier[
                ['Код товара поставщика', 'Наименование', 'All', 'Зона хранения', 'Объем ЕИ отгрузки']].rename(
                columns={'Код товара поставщика': 'Код тары'})
            sup_map_2 = df_supplier[['Товар №', 'Наименование', 'All', 'Зона хранения', 'Объем ЕИ отгрузки']].rename(
                columns={'Товар №': 'Код тары'})

            supplier_map = pd.concat([sup_map_1, sup_map_2], ignore_index=True).dropna(subset=['Код тары'])
            supplier_map['Код тары'] = supplier_map['Код тары'].astype(str).str.strip()
            supplier_map = supplier_map.drop_duplicates(subset=['Код тары'])

            missing_in_supplier = df_combined[~df_combined['Код тары'].isin(supplier_map['Код тары'])]
            df_final = pd.merge(df_combined, supplier_map, on='Код тары', how='inner')

            unavailable = df_final[pd.to_numeric(df_final['All'], errors='coerce') != 1.0]
            df_final = df_final[pd.to_numeric(df_final['All'], errors='coerce') == 1.0]

            # 5. РАСЧЕТ
            df_final['Вместимость'] = pd.to_numeric(df_final['Вместимость'], errors='coerce').fillna(1)
            df_final['Коробок к заказу'] = df_final.apply(
                lambda row: math.ceil(row['Чистая потребность'] / row['Вместимость']), axis=1)
            df_final['Объем ЕИ отгрузки'] = pd.to_numeric(df_final['Объем ЕИ отгрузки'], errors='coerce').fillna(0)


            # 6. ПРОВЕРКА (возвращаем полные тексты)
            def check_row(row):
                reasons = []
                if row['Отрицательный остаток']: reasons.append("⚠️ Отрицательный остаток")
                if row['Расхождение источников']: reasons.append("⚠️ Расхождение ПРЦ")
                if row['Коробок к заказу'] > 100: reasons.append("⚠️ Большой объем")
                if row['Вместимость'] == 1: reasons.append("⚠️ Нет фасовки")
                return " | ".join(reasons) if reasons else "✅ ОК"


            df_final['Проверка'] = df_final.apply(check_row, axis=1)
            df_final = df_final.rename(
                columns={'Код тары': 'Код DDI', 'Чистая потребность': 'Нужно (шт/кг)', 'Вместимость': 'Фасовка'})

            st.session_state['calculated_data'] = {
                'df_final': df_final,
                'missing': missing_in_supplier,
                'unavailable': unavailable
            }
            st.success("Расчет готов!")

if st.session_state['calculated_data'] is not None:
    data = st.session_state['calculated_data']
    df_final = data['df_final']

    st.write("### 🚨 Проверка рисков")
    if not data['missing'].empty:
        st.warning(f"Внимание: {len(data['missing'])} позиций заказа не найдены в бланке поставщика!")
        with st.expander("Посмотреть потерянные позиции"):
            st.dataframe(data['missing'][['Код тары', 'Сырьё', 'Чистая потребность']])

    if not data['unavailable'].empty:
        st.error(f"Внимание: {len(data['unavailable'])} позиций недоступны для РФ.")
        with st.expander("Посмотреть недоступные товары"):
            st.dataframe(data['unavailable'][['Код тары', 'Наименование']])

    st.write("### 📦 Готовый расчет заказа")
    tab1, tab2 = st.tabs(["Общий список", "Загрузка контейнеров"])

    with tab1:
        st.info(
            "💡 Дважды кликните на 'Коробок к заказу', чтобы изменить вручную. Позиции со статусом ⚠️ требуют внимания.")

        df_display = df_final[
            ['Код DDI', 'Наименование', 'Нужно (шт/кг)', 'Фасовка', 'Коробок к заказу', 'Складской остаток',
             'Проверка']].sort_values('Наименование')

        edited_df = st.data_editor(
            df_display,
            use_container_width=True,
            disabled=['Код DDI', 'Наименование', 'Нужно (шт/кг)', 'Фасовка', 'Складской остаток', 'Проверка'],
            column_config={
                "Код DDI": st.column_config.TextColumn("Код DDI", width="small"),
                "Наименование": st.column_config.TextColumn("Наименование", width="large"),
                "Нужно (шт/кг)": st.column_config.NumberColumn("Нужно (шт/кг)", width="small", format="%.2f"),
                "Фасовка": st.column_config.NumberColumn("Фасовка", width="small"),
                "Коробок к заказу": st.column_config.NumberColumn("Коробок к заказу", min_value=0, step=1,
                                                                  width="small"),
                "Складской остаток": st.column_config.NumberColumn("Склад (сверка)", width="small"),
                "Проверка": st.column_config.TextColumn("Проверка", width="large")
            },
            key="editor"
        )

    with tab2:
        df_final_edited = df_final.copy()
        for idx, row in edited_df.iterrows():
            if row['Коробок к заказу'] != \
                    df_final_edited.loc[df_final_edited['Код DDI'] == row['Код DDI'], 'Коробок к заказу'].values[0]:
                df_final_edited.loc[df_final_edited['Код DDI'] == row['Код DDI'], 'Коробок к заказу'] = row[
                    'Коробок к заказу']

        df_containers = df_final_edited.groupby('Зона хранения').agg({
            'Коробок к заказу': 'sum',
            'Объем ЕИ отгрузки': lambda x: (x * df_final_edited.loc[x.index, 'Коробок к заказу']).sum()
        }).reset_index()
        df_containers.columns = ['Тип контейнера', 'Всего коробок', 'Общий объем (м³)']
        df_containers['Общий объем (м³)'] = df_containers['Общий объем (м³)'].round(2)

        for index, row in df_containers.iterrows():
            c_type = row['Тип контейнера']
            vol = row['Общий объем (м³)']
            limit = 49.5 if c_type in ['МОРОЗ', 'СУХОЙ'] else 23.0
            percent = min(vol / limit, 1.0)

            st.write(f"**{c_type}** — {row['Всего коробок']} коробок, {vol} м³ (из {limit} м³)")
            st.progress(percent)
            if vol > limit:
                st.error(f"⚠️ Превышение лимита контейнера {c_type}!")

    df_export = edited_df[['Код DDI', 'Коробок к заказу', 'Наименование']].copy()
    df_export.columns = ['Код товара', 'Количество', 'Название товара (не обязательно)']
    df_export = df_export[df_export['Количество'] > 0]

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Заказ')
        worksheet = writer.sheets['Заказ']
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
        worksheet.column_dimensions['A'].width = 20
        worksheet.column_dimensions['B'].width = 15
        worksheet.column_dimensions['C'].width = 85
        for row in worksheet.iter_rows(min_row=2, min_col=2, max_col=2):
            for cell in row:
                cell.alignment = Alignment(horizontal='center')

    st.download_button(
        label="📥 Скачать готовый заказ в Excel",
        data=buffer.getvalue(),
        file_name="Заказ_поставщику_Магадан.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key="unique_download_btn"
    )