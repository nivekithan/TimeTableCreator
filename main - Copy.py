import time
from selenium import webdriver
from bs4 import BeautifulSoup
import re
import xlsxwriter


url_portal = "https://sappro.delhi.upes.ac.in:8443/sap/bc/webdynpro/sap/zupes_student_portal#"
url_time = "https://sappro.delhi.upes.ac.in:8443/sap/bc/webdynpro/sap/zupes_timetable#"

driver = webdriver.Edge("H:\driver\msedgedriver")


# login to site

driver.get(url_portal)

time.sleep(10)

driver.find_element_by_name("WD2E").send_keys("username")  # typing username
driver.find_element_by_name("WD34").send_keys("password")  # typing password
driver.find_element_by_id("WD3F").click()  # clicking login

time.sleep(5)

driver.find_element_by_id("WD0272").click()  # clicking Timetable

time.sleep(3)

driver.switch_to.window(driver.window_handles[-1])  # switching to that tab

pagesource = driver.page_source
bs = BeautifulSoup(pagesource, "html.parser")

#timetable

time_table = bs.find_all("td",{"class":"urSTC urST3TD urSTTDRo2 urColorPaleBlue urColorPaleBlueIconBar"})


with open("time_table.txt","w+") as f:
    f.truncate()

for table in time_table:
    with open("time_table.txt", "a") as f:
        f.write(str(table) + "\n")


out = {}




with open("time_table.txt","r") as f:
    i = 1
    for line in f:
        out[i] = {}

        #checking for what day
        col = int(re.search('cc=.(.).', line).group(1))
        module = re.search('Module : (.*)<br/>Room', line).group(1)
        room_no = re.search('Room : VR_B_(\d\d\d\d)', line).group(1)
        start_time = re.search("Start Time : (\d\d:\d\d:\d\d)", line).group(1)
        end_time = re.search("End Time : (\d\d:\d\d:\d\d)",line).group(1)

        #adding day to out[i]

        if col == 1:
            out[i]["day"] = "monday"
        elif col ==2:
            out[i]["day"] = "tuesday"
        elif col == 3:
            out[i]["day"] = "wednesday"
        elif col == 4:
            out[i]["day"] = "thursday"
        elif col == 5:
            out[i]["day"] = "friday"
        elif col == 6:
            out[i]["day"] = "saturday"
        elif col == 7:
            out[i]["day"] = "sunday"

        # adding module to out[i]

        out[i]["subject"] = module

        # adding room number to out[i]

        out[i]["room"] = room_no

        # adding starttime to out[i]

        out[i]["start"] = start_time

        # adding endtime to out[i]

        out[i]["end"] = end_time

        # increasing the value of i
        i +=1





class ExcelCreator:


    col_row ={
        "monday":{
            "row": 7,
            "col": 0,
            "a_row":0,
            "a_col":0
        },

        "tuesday":{
            "row":7,
            "col":3,
            "a_row":0,
            "a_col":0
        },

        "wednesday":{
            "row":7,
            "col":6,
            "a_row":0,
            "a_col":0
        },

        "thursday":{
            "row":7,
            "col":9,
            "a_row":0,
            "a_col":0
        },

        "friday":{
            "row":7,
            "col":12,
            "a_row":0,
            "a_col":0
        },

        "saturday":{
            "row":7,
            "col":15,
            "a_row":0,
            "a_col":0
        },

        "sunday":{
            "row":7,
            "col":18,
            "a_row":0,
            "a_col":0
        }
    }

    day_list = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]

    def __init__(self, name):
        self.wb = xlsxwriter.Workbook("{0}.xlsx".format(name))
        self.sheet = self.wb.add_worksheet("sheet1")

        row = 6
        col = 1
        for day in self.day_list:
            self.sheet.write(row,col,day)
            col +=3




    def row_col(self, day):
        row = self.col_row[day]["row"] + self.col_row[day]["a_row"]
        col = self.col_row[day]["col"]

        return (row, col)

    def add_subject(self, name, day):

        row, col = self.row_col(day)
        self.sheet.write(row + 1, col +1, name)

    def add_room(self, name, day):

        row, col = self.row_col(day)
        self.sheet.write(row + 1, col + 2, name)

    def add_start(self, name, day):

        row, col = self.row_col(day)
        self.sheet.write(row + 2, col + 1, name)

    def add_end(self, name, day):

        row, col = self.row_col(day)
        self.sheet.write(row + 2, col + 2, name)

    def close(self):
        self.wb.close()

    def add_everythin(self, i, day):
        self.add_subject(out[i]["subject"], day)
        self.add_room(out[i]["room"], day)
        self.add_start(out[i]["start"], day)
        self.add_end(out[i]["end"], day)

        self.col_row[day]["a_row"] +=3


class Run:

    def __init__(self,name):
        self.f = ExcelCreator(name)

        for i in range(1, len(out) + 1):

            if out[i]["day"] == "monday":
                day = "monday"
                self.f.add_everythin(i,day)

            elif out[i]["day"] == "tuesday":
                day = "tuesday"
                self.f.add_everythin(i, day)

            elif out[i]["day"] == "wednesday":
                day = "wednesday"
                self.f.add_everythin(i, day)

            elif out[i]["day"] == "thursday":
                day = "thursday"
                self.f.add_everythin(i, day)

            elif out[i]["day"] == "friday":
                day = "friday"
                self.f.add_everythin(i, day)

            elif out[i]["day"] == "saturday":
                day = "saturday"
                self.f.add_everythin(i, day)

            elif out[i]["day"] == "sunday":
                day = "sunday"
                self.f.add_everythin(i, day)

            else:
                print("something is wrong")

        self.f.close()


check = Run("timetable_2")