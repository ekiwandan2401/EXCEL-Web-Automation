from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

from openpyxl import load_workbook
import time
import os

print("ex Lokasi Excel : C:\ekitools\data-tod.xlsx")
print("ex Lokasi Chromedriver : C:\ekitools\CSV-Web-Automation\chromedriver.exe")
print("==================================")
# Load data Chromedriver dan excel kalian
loadexcel = input("Masukan Lokasi Excel : ")
loadchdrv = input("Masukan Lokasi Chromedriver : ")

wb =  load_workbook(filename=loadexcel)
sheetRange = wb['Sheet1']

driver = webdriver.Chrome(executable_path=loadchdrv)
# Ganti dengan URL kalian
url = "http://127.0.0.1:8000/mahasiswa"  
driver.get(url)
driver.maximize_window()
driver.implicitly_wait(10)

os.system('cls')
os.system('color a')

#looping cok

i = 2

while i <= len(sheetRange['A']):
	Nama = sheetRange['A'+str(i)].value
	NIM = sheetRange['B'+str(i)].value
	Jurusan = sheetRange['C'+str(i)].value
	JK = sheetRange['D'+str(i)].value
	Alamat = sheetRange['E'+str(i)].value

	driver.find_element(By.XPATH, '//*[@id="tomboltambah"]').click()

	try:
		WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="exampleModal"]/div/div')))

		driver.find_element(By.XPATH, '//*[@id="nama"]').send_keys(Nama)
		driver.find_element(By.XPATH, '//*[@id="nim"]').send_keys(NIM)
		driver.find_element(By.XPATH, '//*[@id="jurusan"]').send_keys(Jurusan)
		driver.find_element(By.XPATH, '//*[@id="jenis_kelamin"]').send_keys(JK)
		driver.find_element(By.XPATH, '//*[@id="alamat"]').send_keys(Alamat)
		driver.find_element(By.XPATH, '//*[@id="exampleModal"]/div/div/div[3]/button[2]').click()
		print("Data "+ Nama +" Berhasil di input")

	except TimeoutException:
		print("Form Gak Muncul TOD")
		pass

	time.sleep(1)
	i = i + 1

print("Semua data berhasil di input")