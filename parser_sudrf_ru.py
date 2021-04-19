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
driver.get("https://sudrf.ru/index.php?id=300#sp")

# Получаем доступ к листу по ссылке
pathToExcel = input("Введите полный путь до файла Excel\n")  # "A:\Project\lesson_python\data1.xlsx"
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
col = [1, 2, 3]
s = True

while row <= lastRow:
    surname = readData.Cells(row, col[0]).Text.replace(' ', '')
    name = readData.Cells(row, col[1]).Text.replace(' ', '')
    patronymic = readData.Cells(row, col[2]).Text.replace(' ', '')

    driver.find_element_by_xpath("//div[1]/div[4]/div[6]/form/table/tbody/tr[8]/td[2]/input[2]").click()
    driver.find_element_by_xpath("//input[@name='f_name'][@id='f_name']").send_keys("{} {} {}".format(surname, name, patronymic))
    driver.find_element_by_xpath('//div[1]/div[4]/div[6]/form/table/tbody/tr[1]/td[2]/select').click()
    driver.find_element_by_xpath('//div[1]/div[4]/div[6]/form/table/tbody/tr[1]/td[2]/select/option[12]').click()
    driver.find_element_by_xpath('//div[1]/div[4]/div[6]/form/table/tbody/tr[8]/td[2]/input[1]').click()

    time.sleep(7)
    try:
        if s:
            for i in range(len(driver.find_elements_by_xpath(
                    "//div[1]/div[4]/div[6]/form/table/tbody/tr[9]/td/div/div[4]/table/tbody/tr[1]/td"))):
                sheet.Cells(row2 - 1, i + 1).Value = driver.find_element_by_xpath(
                    "//div[1]/div[4]/div[6]/form/table/tbody/tr[9]/td/div/div[4]/table/tbody/tr[1]/td[{}]".format(
                        i + 1)).text
                sheet.Cells(row2 - 1, i + 1).WrapText = True
            s = False

        for _ in range(2, len(driver.find_elements_by_xpath(
                    "//div[1]/div[4]/div[6]/form/table/tbody/tr[9]/td/div/div[4]/table/tbody/tr"))+1):
            for col2 in range(len(driver.find_elements_by_xpath("//div[1]/div[4]/div[6]/form/table/tbody/tr[9]/td/div/div[4]/table/tbody/tr[{}]/td".format(str(_))))):
                sheet.Cells(row2, col2 + 1).Value = driver.find_element_by_xpath(
                    "//div[1]/div[4]/div[6]/form/table/tbody/tr[9]/td/div/div[4]/table/tbody/tr[{}]/td[{}]".format(_,
                                                                                                        col2 + 1)).text
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
