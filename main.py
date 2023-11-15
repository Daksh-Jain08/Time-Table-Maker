import openpyxl
import datetime

myData = 'MyDatabase.xlsx'
wb = openpyxl.load_workbook(myData)
ws1 = wb['Courses']
ws2 = wb['Sections']

data = 'data.xlsx'
wb2 = openpyxl.load_workbook(data)
dataSheet = wb2['Courses']

class Course:

    def __init__(self, code, name, section, examDate, examTime):
        self.__code = code
        self.__name = name
        self.__sectionList = section      #This whould be a list of section object (aggregation)
        self.__examDates = examDate
        self.__examTime = examTime

    def get_code(self):
        return self.__code

    def get_sections(self):
        return self.__sectionList

    def print_all_sections(self):
        for i in self.__sectionList:
            print(i)

    def get_info(self):
        l = {
            "code":self.__code,
            "name":self.__name,
            "section":self.__sectionList,
            "examDate":self.__examDates,
            "examTime":self.__examTime}
        return l

    def populate_section(self, type, number, password):
        tryPass = input("Enter the admin password: ")
        if tryPass == password:
            for i in range (number):
                sectionNumber = type[0] + str(len(self.__sectionList)+1)
                nHour = int(input("Enter the number of hours: "))
                hourList = []
                hourStr = ''
                hour = int(input("Enter the starting Hour: "))
                for i in range(nHour):
                    hourList.append(hour)
                    hourStr += str(hour) +','
                    hour+=1
                nDays = int(input("Enter the number of days in the section: "))
                daysList = []
                daysStr = ''
                for j in range(nDays):
                    day = input("Enter the day " + str(j+1) + " for " + type + " " + str(i+1) +": ")
                    daysList.append(day)
                    daysStr += day + ','
                venue = int(input("Enter the venue for " + type + " " + str(i+1) +": "))
                nProff = int(input("Enter the number of professors in the section: "))
                proffList = []
                proffStr = ''
                for j in range(nProff):
                    proff = input("Enter the professor " + str(j+1) + "for" + type + " " + str(i+1) +": ")
                    proffList.append(proff)
                    proffStr += proff + ','
                section = Section(self.__code, sectionNumber, hourList, daysList, venue, proffList)
                self.__sectionList.append(section)
                row = [self.__code, type, hourStr, daysStr, venue, proffStr]
                ws2.append(row)
                wb.save(myData)
        else:
            print("\nWrong Pin.\n")

class Section():

    def __init__(self, code, sectionNo, hour, day, venue, proff):
        self.__code = code
        self.__sectionNo = sectionNo
        self.__timing = hour    #1st hour is 8 to 9, 2nd is 9 to 10 and so on.
        self.__day = day    #list of lecture days in a week.
        self.__venue = venue
        self.__proff = proff

    def get_info(self):
        data = {"code":self.__code,
                "sectionNo":self.__sectionNo,
                "timing":self.__timing,
                "day":self.__day,
                "venue":self.__venue,
                "proff":self.__proff}
        return data

    def __str__(self):
        return '[' + str(self.__code) + ',' + str(self.__sectionNo) + ',' + str(self.__timing) + ',' + str(self.__day) + ',' + str(self.__venue) + ',' + str(self.__proff) + ']'

class TimeTable:

    def __init__(self):
        self.table = [[0]*9]*6
        self.__listOfCourses = []

    def enroll_subject(self,course,section):
        courseDetails = course.get_info()
        sectionDetails = section.get_info()
        code = courseDetails["code"]
        days = sectionDetails["day"]
        week = {'M':0 , 'T':1 , 'W':2 , 'Th':3 , 'F':4 , 'S':5}
        for i in self.__listOfCourses:
            if i.get_info()["examDate"] == courseDetails["examDate"] and i.get_info()["examTime"] == courseDetails["examTime"]:
                print("\nClash of Examination. " + i.get_info()["name"] + " has examination at the same time.\n")
        for i in range(0,len(days)):
            days[i] = week[days[i]]
        hours = sectionDetails["timing"]
        indices = []
        for i in hours:
            for j in days:
                if self.table[i][j]==0:
                    indices.append([i,j])
                else:
                    print("\nClash Between Courses. " + str(self.table[i][j]) + " already at this slot.\n")
                    return
        for i,j in indices:
            self.table[i][j] = courseDetails["code"]
        print("\nSuccessfully Enrolled.\n")
        self.__listOfCourses.append(course)

    def withdraw_subject(self,course):
        for i in range(0,9):
            for j in range(0,6):
                if self.table[i][j] == course.__code:
                    self.table[i][j] = 0
                    print("You have sucessfully withdrawn from the course")
                else:
                    print("You were not enrolled into this course.")

    def show_courses(self):
        for i in self.__listOfCourses:
            print(i.get_info())

    def __str__(self):
        pass

    def export_to_csv(self,file):
        pass

def populate_course(password):
    tryPass = input("Enter the admin password: ")
    if tryPass == password:
        counter=1
        availableCourses = []    # List of the courses that are not added to personal database
        print("The available courses are: ")
        # iterating throgh all the rows in the database
        for i in range (2,dataSheet.max_row+1):
            flag=1
            code = dataSheet.cell(row=i,column=2).value
            #checking if the course already exists in personal database
            for j in range(2,ws1.max_row+1):
                if(ws1.cell(row=j,column=1).value == code):
                    flag=0
            #displaying all the available courses
            if (flag == 1):
                availableCourses.append(code)
                print(str(counter) + ' ' + code)
                counter+=1
        choice = int(input("Enter Your Choice: "))
        # selecting the required row from the database and inserting it to the personal database
        for i in range (2,dataSheet.max_row+1):
            code = dataSheet.cell(row=i,column=2).value
            if(code == availableCourses[choice-1]):
                # making the row to be appended
                row = []
                for j in range(2,dataSheet.max_column+1):
                    row.append(dataSheet.cell(row=i,column=j).value)
                ws1.append(row)
                wb.save(myData)
    else:
        print("Wrong Pin.")

def menu():
    print('Welcome to the Time Table Manager.')
    password = input("Please Set your admin Password: ")
    print('''
1. Add Courses to the database.
2. Add Section(s) to a course.
''')