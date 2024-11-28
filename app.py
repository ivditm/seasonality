import streamlit as st
import pandas as pd
from io import BytesIO


def load_and_preprocess_data(file):
    df = pd.read_excel(file)
    df_transposed = df.set_index('Фраза').transpose()
    df_transposed.index.name = 'date'
    df_transposed.reset_index(inplace=True)
    df_transposed['date'] = df_transposed['date'].apply(get_month_name)
    return df_transposed


def get_month_name(date_str):
    return pd.to_datetime(date_str).strftime('%B')


def calculate_category_seasonality(df_transposed):
    df_transposed['Monthly Total'] = df_transposed.sum(axis=1, numeric_only=True)
    average_monthly_demand = df_transposed['Monthly Total'].sum() / 24
    df_transposed['Average Monthly Demand'] = average_monthly_demand
    df_transposed['Seasonality Index of category'] = df_transposed[
        'Monthly Total'] / average_monthly_demand
    global_damnd = df_transposed['Monthly Total'].sum() / 2
    df = (df_transposed.groupby(by='date')[
        'Seasonality Index of category']
        .mean().reset_index().set_index('date')
        .T.reset_index().rename(columns={'index': 'Phrase'}))
    df['Yearly Demand'] = global_damnd
    df['Phrase'] = 'Seasonality Index of category'
    df = df[['Phrase', 'Yearly Demand', 'April', 'August', 'December',
             'February', 'January', 'July', 'June', 'March', 'May',
             'November', 'October', 'September']]
    return df


def process_column(data, column):
    monthly = data.groupby(by='date')[column].mean()
    yearly = data[column].sum() / (len(data) // 12)
    if yearly != 0:
        season = monthly / yearly
    else:
        season = pd.Series([0] * len(monthly), index=monthly.index)
    season = season.reset_index().set_index('date').T
    season['Yearly Demand'] = yearly
    season = season[['Yearly Demand', 'April', 'August',
                     'December', 'February', 'January', 'July',
                     'June', 'March', 'May', 'November',
                     'October', 'September']]
    season = season.reset_index().rename(columns={'index': 'Phrase'})
    season['Phrase'] = column
    return season


def generate_final_df(file):
    df_transposed = load_and_preprocess_data(file)
    category_seasonality_df = calculate_category_seasonality(df_transposed)
    result_df = category_seasonality_df
    for column in set(df_transposed.columns) - set(['date','Monthly Total',
       'Average Monthly Demand', 'Seasonality Index of category']):
        season = process_column(df_transposed[[column, 'date']], column)
        result_df = pd.concat([result_df, season], ignore_index=True)
    return result_df


def process_file(file):
    df = generate_final_df(file)
    st.write("Первые 5 строк измененного файла:")
    st.write(df.head())

    def convert_df_to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
        processed_data = output.getvalue()
        return processed_data

    excel_data = convert_df_to_excel(df)
    st.download_button(
        label="Скачать измененный файл",
        data=excel_data,
        file_name='seasonality.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


def main():
    st.title("Загрузка и обработка файла")
    uploaded_file = st.file_uploader("Загрузите Excel файл",  type=["xlsx"])
    if uploaded_file is not None:
        process_file(uploaded_file)
    st.text("Загрузите файл для анализа и изменения.")


if __name__ == "__main__":
    main()
