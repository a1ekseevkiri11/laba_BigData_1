import openpyxl

NUMBER_LIST = 2

COLUMN_COUNTRY = 1
COLUMN_CONTINENT = 3
COLUMN_YEAR = [4, 5, 6, 7, 8, 9, 10, 11]
COLUMN_AREA = 12

COUNT_YEAR = 52


wb = openpyxl.open("laba.xlsx")
table_1 = wb.worksheets[NUMBER_LIST]


continents = {}
countAllPeople = 0
countCountry = 0
sumPopulationDensity = 0
countryes = {}

for row in range(2, table_1.max_row + 1):
    countCountry += 1
    areaCountry = table_1[row][COLUMN_AREA].value
    countryes[table_1[row][COLUMN_COUNTRY].value] = table_1[row][COLUMN_YEAR[0]].value - table_1[row][COLUMN_YEAR[len(COLUMN_YEAR) - 1]].value
    for year in COLUMN_YEAR:
        nameContinent = table_1[row][COLUMN_CONTINENT].value
        cellValue = table_1[row][year].value
        sumPopulationDensity += cellValue / areaCountry
        countAllPeople += cellValue
        if nameContinent in continents:
            continents[nameContinent] += cellValue
        else:
            continents[nameContinent] = cellValue


averageNumberOfResidentsByYear = countAllPeople / COUNT_YEAR

print("Среднее количество жителей по годам:", f'{averageNumberOfResidentsByYear:,.2f}'.replace(',', ' '), "\n")

print("Среднее количество жителей по континентам:")
for continent in continents:
    continents[continent] /= COUNT_YEAR
    print(continent, ":", f'{continents[continent]:,.2f}'.replace(',', ' '))


print("\n")

print("Средняя плотность населения на планете за", COUNT_YEAR ,"года:", sumPopulationDensity / (countCountry * COUNT_YEAR))

print("\n")

print("Динамика плотности населения для страны:")
for country in countryes:
    print(country, ":", f'{countryes[country]:,.2f}'.replace(',', ' '))