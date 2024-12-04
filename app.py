import pandas as pd
from pandas import DataFrame

# Загрузите данные из файла `sales_data.csv` в pandas DataFrame.
df = pd.read_csv('sales_data.csv')


# Task1
# Постройте сводную таблицу, отображающую общую выручку (`Revenue`) по каждому
# магазину за каждую дату, и сохраните результаты на первой странице Excel-файла.
def task1(data_frame: DataFrame, mode: str = "w"):
    pivot = data_frame.pivot_table(values='Revenue', index=['Date'], columns=['Store'], aggfunc='sum')

    with pd.ExcelWriter('sales_result.xlsx', mode=mode) as writer:
        pivot.to_excel(writer, sheet_name='task1')


# Task2
# Рассчитайте среднюю цену продукции для каждого продукта во всех магазинах и
# сохраните результаты на второй странице Excel-файла.
def task2(data_frame: DataFrame, mode: str = "w"):
    average_price = data_frame.groupby('Product')['Price'].mean()

    with pd.ExcelWriter('sales_result.xlsx', mode=mode) as writer:
        average_price.to_excel(writer, sheet_name='task2')


# Task3
# Найдите магазин, в котором было продано максимальное количество продукции за
# весь период данных, и сохраните результаты на третьей странице Excel-файла.
def task3(data_frame: DataFrame, mode: str = "w"):
    total_quantity_by_store = data_frame.groupby('Store')['Quantity'].sum()
    index = total_quantity_by_store.idxmax()
    max_value = total_quantity_by_store.max()

    with pd.ExcelWriter('sales_result.xlsx', mode=mode) as writer:
        pd.DataFrame({'Store': [index], 'Value': [max_value]}).to_excel(writer, sheet_name='task3')


# Task4
# Постройте мультииндекс, используя `Date` и `Store`, и отобразите общее
# количество проданных продуктов (`Quantity`) для каждой комбинации даты и магазина,
# отсортировав результат по убыванию количества, и сохраните результаты на четвёртой
# странице Excel-файла.
def task4(data_frame: DataFrame, mode: str = "w"):
    multi_index = data_frame.groupby(['Date', 'Store'])['Quantity'].sum().sort_values(ascending=False)

    with pd.ExcelWriter('sales_result.xlsx', mode=mode) as writer:
        multi_index.to_excel(writer, sheet_name='task4')


# Task5
# Рассчитайте 25-й, 50-й (медиана) и 75-й перцентили для столбца `Quantity` в данных и
# отобразите их значения, а затем сохраните результаты на пятой странице Excel-файла.
def task5(data_frame: DataFrame, mode: str = "w"):
    percentiles = (data_frame['Quantity'].quantile([0.25, 0.5, 0.75]))
    new_df = pd.DataFrame({'Percentiles': ['25%%', '50%%', '75%%'],
                           'Quantity': percentiles.values})

    with pd.ExcelWriter('sales_result.xlsx', mode=mode) as writer:
        new_df.to_excel(writer, sheet_name='task5')


if __name__ == '__main__':
    task1(df)
    task2(df, 'a')
    task3(df, 'a')
    task4(df, 'a')
    task5(df, 'a')
