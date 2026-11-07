import streamlit as st
import pandas as pd
import math
import io

st.set_page_config(page_title="Сборка заказа Додо", layout="wide")
st.title("Автоматизация заказов: Додо Магадан")

st.header("Шаг 1. Загрузка данных")
col1, col2 = st.columns(2)

with col1:
    file_m1 = st.file_uploader("Магадан-1", type=["xlsx"])
    file_m2 = st.file_uploader("Магадан-2", type=["xlsx"])
    file_m3 = st.file_uploader("Магадан-3", type=["xlsx"])

with col2:
    file_prc1 = st.file_uploader("ПРЦ-1", type=["xlsx"])
    file_prc2 = st.file_uploader("ПРЦ-2", type=["xlsx"])
    file_supplier = st.file_uploader("Бланк поставщика", type=["xlsx"])

if st.button("🚀 Рассчитать заказ", type="primary"):
    if not all([file_m1, file_m2, file_m3, file_prc1, file_prc2, file_supplier]):
        st.warning("Пожалуйста, загрузите все 6 файлов!")
    else:
        st.success("Файлы успешно загружены! Считаем...")

        # 1. ЧТЕНИЕ ВСЕХ ФАЙЛОВ
        df_m1 = pd.read_excel(file_m1, skiprows=9)
        df_m2 = pd.read_excel(file_m2, skiprows=9)
        df_m3 = pd.read_excel(file_m3, skiprows=9)
        df_supplier = pd.read_excel(file_supplier, skiprows=12)
        df_prc1 = pd.read_excel(file_prc1, skiprows=9)
        df_prc2 = pd.read_excel(file_prc2, skiprows=9)

        # Чистим названия колонок от случайных пробелов
        df_supplier.columns = df_supplier.columns.astype(str).str.strip()

        # 2. ОБЪЕДИНЕНИЕ ПОТРЕБНОСТЕЙ
        df_all = pd.concat([df_m1, df_m2, df_m3], ignore_index=True)
        df_all = df_all[df_all['Надо заказать'] > 0]
        df_combined = df_all.groupby(['Код тары', 'Сырьё'], as_index=False).agg({
            'Надо заказать': 'sum',
            'Вместимость': 'max'
        })

        # 3. ОСТАТКИ ПРЦ
        df_prc_all = pd.concat([df_prc1, df_prc2], ignore_index=True)
        df_prc = df_prc_all.groupby('Код тары', as_index=False)['Текущий остаток'].sum()

        # Приводим коды к строке
        df_combined['Код тары'] = df_combined['Код тары'].astype(str).str.strip()
        df_prc['Код тары'] = df_prc['Код тары'].astype(str).str.strip()

        # Вычитаем склады
        df_combined = pd.merge(df_combined, df_prc, on='Код тары', how='left')
        df_combined['Текущий остаток'] = df_combined['Текущий остаток'].fillna(0).apply(lambda x: max(0, x))
        df_combined['Чистая потребность'] = (df_combined['Надо заказать'] - df_combined['Текущий остаток']).apply(
            lambda x: max(0, x))
        df_combined = df_combined[df_combined['Чистая потребность'] > 0]

        # --- ИСПРАВЛЕНИЕ: МАППИНГ ПО ДВУМ КОЛОНКАМ ---
        # Готовим колонку "Товар №" (убираем .0 в конце)
        df_supplier['Товар №'] = df_supplier['Товар №'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

        # Делаем два справочника: один по DDI, второй по числовому коду
        sup_map_1 = df_supplier[
            ['Код товара поставщика', 'Наименование', 'All', 'Зона хранения', 'Объем ЕИ отгрузки']].copy()
        sup_map_1 = sup_map_1.rename(columns={'Код товара поставщика': 'Код тары'})

        sup_map_2 = df_supplier[['Товар №', 'Наименование', 'All', 'Зона хранения', 'Объем ЕИ отгрузки']].copy()
        sup_map_2 = sup_map_2.rename(columns={'Товар №': 'Код тары'})

        # Объединяем их в один большой справочник
        supplier_map = pd.concat([sup_map_1, sup_map_2], ignore_index=True)
        supplier_map = supplier_map.dropna(subset=['Код тары'])
        supplier_map['Код тары'] = supplier_map['Код тары'].astype(str).str.strip()
        supplier_map = supplier_map.drop_duplicates(subset=['Код тары'])

        # 4. УМНЫЕ ПРОВЕРКИ (Предупреждения)
        st.write("### 🚨 Проверка рисков")

        # Ищем, чего нет в бланке
        missing_in_supplier = df_combined[~df_combined['Код тары'].isin(supplier_map['Код тары'])]
        if not missing_in_supplier.empty:
            st.warning(
                f"Внимание: {len(missing_in_supplier)} позиций заказа не найдены в бланке поставщика! Они исключены из расчета.")
            with st.expander("Посмотреть, какие позиции потерялись"):
                st.dataframe(missing_in_supplier[['Код тары', 'Сырьё', 'Чистая потребность']])

        # Сопоставление через новый справочник
        df_final = pd.merge(df_combined, supplier_map, on='Код тары', how='inner')

        # Ищем, что недоступно для РФ (All != 1.0)
        unavailable = df_final[pd.to_numeric(df_final['All'], errors='coerce') != 1.0]
        if not unavailable.empty:
            st.error(
                f"Внимание: {len(unavailable)} позиций недоступны для заказа в РФ (колонка All = 0). Они исключены.")
            with st.expander("Посмотреть недоступные товары"):
                st.dataframe(unavailable[['Код тары', 'Наименование']])

        # Оставляем только доступные
        df_final = df_final[pd.to_numeric(df_final['All'], errors='coerce') == 1.0]

        # 5. РАСЧЕТ КОРОБОК
        df_final['Вместимость'] = pd.to_numeric(df_final['Вместимость'], errors='coerce').fillna(1)
        df_final['Коробок к заказу'] = df_final.apply(
            lambda row: math.ceil(row['Чистая потребность'] / row['Вместимость']), axis=1
        )

        df_final['Объем ЕИ отгрузки'] = pd.to_numeric(df_final['Объем ЕИ отгрузки'], errors='coerce').fillna(0)
        df_final = df_final.rename(
            columns={'Код тары': 'Код DDI', 'Чистая потребность': 'Нужно (шт/кг)', 'Вместимость': 'Фасовка'})

        # 6. ВЫВОД НА ЭКРАН (С ВКЛАДКАМИ)
        st.write("### 📦 Готовый расчет заказа")
        tab1, tab2 = st.tabs(["Общий список", "Загрузка контейнеров"])

        with tab1:
            st.dataframe(
                df_final[['Код DDI', 'Наименование', 'Нужно (шт/кг)', 'Фасовка', 'Коробок к заказу']].sort_values(
                    'Наименование'),
                use_container_width=True
            )

        with tab2:
            df_containers = df_final.groupby('Зона хранения').agg({
                'Коробок к заказу': 'sum',
                'Объем ЕИ отгрузки': lambda x: (x * df_final.loc[x.index, 'Коробок к заказу']).sum()
            }).reset_index()
            df_containers.columns = ['Тип контейнера', 'Всего коробок', 'Общий объем (м³)']
            df_containers['Общий объем (м³)'] = df_containers['Общий объем (м³)'].round(2)

            for index, row in df_containers.iterrows():
                c_type = row['Тип контейнера']
                vol = row['Общий объем (м³)']

                if c_type == 'МОРОЗ':
                    limit = 49.5
                elif c_type == 'СУХОЙ':
                    limit = 49.5
                elif c_type == 'ХОЛОД':
                    limit = 23.0
                else:
                    limit = 49.5

                percent = min(vol / limit, 1.0)

                st.write(f"**{c_type}** — {row['Всего коробок']} коробок, {vol} м³ (из {limit} м³)")
                st.progress(percent)

                if vol > limit:
                    st.error(f"⚠️ Превышение лимита контейнера {c_type}! Нужен еще один контейнер.")

        # 7. КНОПКА СКАЧИВАНИЯ EXCEL
        df_export = df_final[['Код DDI', 'Коробок к заказу', 'Наименование']].copy()
        df_export.columns = ['Код товара', 'Количество', 'Название товара (не обязательно)']

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, sheet_name='Заказ')

        st.download_button(
            label="📥 Скачать готовый заказ в Excel",
            data=buffer.getvalue(),
            file_name="Заказ_поставщику_Магадан.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )