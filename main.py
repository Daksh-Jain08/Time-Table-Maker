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

    def __init__(self, code):
        self.__code = code
        # self.__name = name
        # self.__sectionList = section      #This whould be a list of section object (aggregation)
        # self.__examDates = examDate
        # self.__examTime = examTime
        self.__sectionList = []

        for code,name,examDate,examTime in ws1.rows:
            if code.value == code:
                self.__name = name.value
                self.__examDates = examDate.value
                self.__examTime = examTime.value
                self.__sectionList = []

        for i in range(2,ws2.max_row+1):
            l = []
            if ws2.cell(row=i,column=1) == code:
                for j in range(1,3):
                    d = j.value
                    l.append(d)
            if(len(l)!=0):
                section = Section(l[0],l[1])
                self.__sectionList.append(section)

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
                row = [self.__code, sectionNumber, hourStr, daysStr, venue, proffStr]
                ws2.append(row)
                wb.save(myData)
                section = Section(self.__code, sectionNumber)
                self.__sectionList.append(section)
                print("Section has been added successfully.")
        else:
            print("\nWrong Pin.\n")

class Section():

    def __init__(self, code, sectionNo):
        self.__code = code
        self.__sectionNo = sectionNo
        for code1, sectionNumber, hour, day, venue, proff in ws2.rows:
            if code1.value == code and sectionNumber.value == sectionNo :
                self.__timing = hour.value.split(',')
                self.__day = day.value.split(',')
                self.__venue = venue.value
                self.__proff = proff.value.split(',')

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

def get_all_courses():
    l=[]
    for i in range(2,ws1.max_row+1):
        a = ws1.cell(row=i,column=1)
        l.append(a.value)
    return l

def menu():
    print('Welcome to the Time Table Manager.')
    password = input("Please Set your admin Password: ")
    print("\nYour admin password has been set.\n")
    nav = ('''
1.  Add Courses to the database.
2.  Add Section(s) to a Course.
3.  Enroll to a Course.
4.  Withdraw Form a Course.
5.  See list of enrolled Courses.
6.  See List of Available Courses.
7.  See List of all the sections of a given Course.
8.  See Details of a Course.
9.  See Details of a Section of a Course.
10. See Current Time Table.
11. Get Time Table CSV File.
12. Exit
''')
    print(nav)
    choice = int(input("Enter your choice: "))
    
    if(choice == 1):
        populate_course(password)

    elif(choice == 2):
        print("These are the Courses: ")
        L = get_all_courses()
        print(L)
        code = input("Which Course do you want to add a section in: ")
        course1 = Course(code)
        sectionType = input("What type of section do you want to add?\nL for Lecture\nT for Tutorial\nP for Laboratory. ")
        nSection = int(input("How many sections do you want to add? "))
        course1.populate_section(sectionType,nSection,password)

    elif(choice == 3):
        pass

menu()