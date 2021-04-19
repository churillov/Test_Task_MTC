import win32com.client as win32
import time

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException

# Создаем СОМ объект и делаем его видимым
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True

# Создаем объект chrome
path_to_chromedriver = r'A:\Project\lesson_python\chromedriver'
driver = webdriver.Chrome(executable_path=path_to_chromedriver)
driver.get("https://fssp.gov.ru")
driver.find_element_by_link_text("Расширенный поиск").click()

# Получаем доступ к листу по ссылке
pathToExcel = input("Введите полный путь до файла Excel\n")
workbook = excel.Workbooks.Open(pathToExcel)
excel.ScreenUpdating = False
book = excel.Workbooks.Add()
sheet = book.Worksheets(1)

# Считываем первую страницу Excel
readData = workbook.Worksheets(1)

# Количество задействованых колонок и столбцов
lastCol = readData.UsedRange.Columns.Count
lastRow = readData.UsedRange.Rows.Count

# Счетчик начинается с 1 из за шапки в Excel
row = 2
row2 = 2
col = [1, 2, 3, 4]
s = True

while row <= lastRow:
    surname = readData.Cells(row, col[0]).Text.replace(' ', '')
    name = readData.Cells(row, col[1]).Text.replace(' ', '')
    patronymic = readData.Cells(row, col[2]).Text.replace(' ', '')
    dateOfBirth = readData.Cells(row, col[3]).Text.replace(' ', '')

    driver.find_element_by_xpath("//input[@name='is[last_name]'][@type='text']").clear()
    driver.find_element_by_xpath("//input[@name='is[last_name]'][@type='text']").send_keys(surname)
    driver.find_element_by_xpath("//input[@name='is[first_name]'][@type='text']").clear()
    driver.find_element_by_xpath("//input[@name='is[first_name]'][@type='text']").send_keys(name)
    driver.find_element_by_xpath("//input[@name='is[patronymic]'][@type='text']").clear()
    driver.find_element_by_xpath("//input[@name='is[patronymic]'][@type='text']").send_keys(patronymic)
    driver.find_element_by_xpath("//input[@name='is[date]'][@type='text']").clear()
    driver.find_element_by_xpath("//input[@name='is[date]'][@type='text']").send_keys(dateOfBirth)

    if row == 2:
        driver.find_element_by_css_selector('[class="btn btn-primary"]').submit()
        input("Пройди капчу и нажми Enter")
    else:
        driver.find_element_by_id("btn-sbm").click()

    try:
        time.sleep(3)
        x1 = driver.find_element_by_xpath("//div[@class='results-frame']").find_element_by_xpath("./table/tbody/tr")
        if s:
            for i in range(len(driver.find_elements_by_xpath("//div[@class='results-frame']//table/tbody/tr/th"))):
                sheet.Cells(row2-1, i+1).Value = x1.find_element_by_xpath("./th[{}]".format(str(i+1))).text
                sheet.Cells(row2 - 1, i + 1).WrapText = True
            s = False
        for _ in range(3, len(driver.find_elements_by_xpath("//div[@class='results-frame']//table/tbody/tr"))+1):
            for col2 in range(len(driver.find_elements_by_xpath("//div[@class='results-frame']//table/tbody/tr[{}]/td".format(str(_))))):
                sheet.Cells(row2, col2+1).Value = driver.find_element_by_xpath("//div[@class='results-frame']//table/tbody/tr[{}]/td[{}]".format(_, col2+1)).text
                sheet.Cells(row2, col2 + 1).WrapText = True
                excel.Worksheets(1).Rows(row2).RowHeight = 80
            row2 += 1
    except NoSuchElementException:
        pass
    row += 1

book.SaveAs(input("Введите полный путь для сохранения файла Excel\nНапример A:\Project\data.xlsx\n"))
excel.ScreenUpdating = True
driver.close()
driver.quit()
# Закрываем СОМ объект
excel.Application.Quit()
