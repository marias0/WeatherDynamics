from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup

import time
from datetime import date, timedelta, datetime
from collections import Counter

from openpyxl import Workbook
from openpyxl.styles import Font, Side, Border, Alignment, NamedStyle, PatternFill
from openpyxl.chart import Reference, StockChart, LineChart

options = Options()
options.add_argument("--headless")
options.add_argument("--disable-gpu")

driver = webdriver.Chrome(
    executable_path="C:\\Program Files\\chromedriver.exe", options=options
)


def Data(website, find_tag, find_class, findAll_tag, findAll_class):

    driver.get(website)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    info = soup.find(find_tag, class_=find_class).findAll(
        findAll_tag, class_=findAll_class
    )

    return list(map(lambda x: x.text.strip(), info))


tenDaysWeather = Data(
    "https://www.gismeteo.ru/weather-sankt-peterburg-4079/10-days/",
    "div",
    "values",
    "span",
    "unit unit_temperature_c",
)

weatherMax = tenDaysWeather[::2]
weatherMin = tenDaysWeather[1::2]


TODAY = date.today()
tenDays = []
for i in range(10):
    date = TODAY + timedelta(i + 1)
    tenDays.append(date)


def table(days, max_weather, min_weather):
    weather_table = []
    for i in range(len(days)):
        weather_table.append([days[i], max_weather[i], min_weather[i]])
    return weather_table


weatherForecast = table(tenDays, weatherMax, weatherMin)
# print(weatherForecast)


monthsAll = []
for i in range(9):
    date = TODAY - timedelta(i + 18)
    current_date = str(date)
    year, month, day = current_date.split("-")
    monthsAll.append(int(month))

months = sorted(Counter(monthsAll))
print(months)

weatherPast = []

for month in months:
    path = "https://www.gismeteo.ru/diary/4079/2021/{query}/"
    website_path = path.format(query=month)

    MonthDays = Data(
        website_path,
        "div",
        "container",
        "td",
        "first",
    )

    MonthWeather = Data(
        website_path,
        "div",
        "container",
        "td",
        ["first_in_group positive", "first_in_group"],
    )

    MonthWeatherMax = MonthWeather[::2]
    MonthWeatherMin = MonthWeather[1::2]

    MonthDate = []
    for day in MonthDays:
        Dt = datetime(int(2021), int(month), int(day))
        D = datetime.date(Dt)
        MonthDate.append(D)

    i = 0
    for i in range(len(MonthDate)):
        weatherPast.append([MonthDate[i], MonthWeatherMax[i], MonthWeatherMin[i]])

print(weatherPast)

weatherInPast = []
for k in range(9):
    date = TODAY - timedelta(k + 18)
    current_date = str(date)
    for row in weatherPast:
        if row[0] == date:
            weatherInPast.append(row)
print(weatherInPast)

wb = Workbook()
destination_filename = "Weather.xlsx"

ws = wb.active
ws.title = "Weather Dynamics"

ws["A1"] = "Date"
ws["B1"] = "Temperature, C"
ws["B2"] = "Max"
ws["C2"] = "Min"
ws.merge_cells("A1:A2")
ws.merge_cells("B1:C1")

for row in reversed(weatherInPast):
    ws.append(row)

for row in weatherForecast:
    ws.append(row)

ws.column_dimensions["A"].width = 11

dateStyle = NamedStyle(name="dateStyle", number_format="dd.mm.yy")
for row in ws["A3:A21"]:
    for cell in row:
        cell.style = dateStyle

cellBorder = Side(style="thin", color="808080")

for row in ws["A1:C21"]:
    for cell in row:
        cell.font = Font(name="Malgun Gothic")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            top=cellBorder, bottom=cellBorder, left=cellBorder, right=cellBorder
        )

today_fill = PatternFill(start_color="45F045", end_color="45F045", fill_type="solid")
ws["A12"].fill = today_fill
ws["B12"].fill = today_fill
ws["C12"].fill = today_fill

chartObject = StockChart()

labels_reference = Reference(ws, min_row=3, min_col=1, max_row=21)
values_reference = Reference(ws, min_row=3, min_col=2, max_row=21, max_col=3)

chartObject.add_data(values_reference)
chartObject.set_categories(labels_reference)

ws.add_chart(chartObject, "E3")

chartObject.title = "Weather"
chartObject.x_axis.title = "Date"
chartObject.y_axis.title = "Temperature"

chartObject.style = 5
chartObject.legend = None
chartObject.series[0].graphicalProperties.line.solidFill = "FF4500"
chartObject.series[1].graphicalProperties.line.solidFill = "1E90FF"

chartObject.height = 9
chartObject.width = 19

wb.save(destination_filename)