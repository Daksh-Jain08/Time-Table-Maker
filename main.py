import openpyxl

wb = openpyxl.load_workbook('TimeTable.xlsx')
ws1 = wb['Courses']
ws2 = wb['Sections']

class Course:

    def __init__(self, code, name, section, examDate, examTime):
        self.__code = code
        self.__name = name
        self.__section = section      #This whould be a list of section object (aggregation)
        self.__examDates = examDate
        self.__examTime = examTime

    def get_section(self):
        for i in self.__section:
            print(i)

    def get_info(self):
        pass

class Section():

    def __init__(self, code, type, hour, day, venue, proff):
        self.__code = code
        self.__type = type
        self.__timing = hour    #1st hour is 8 to 9, 2nd is 9 to 10 and so on.
        self.__day = day    #list of lecture days in a week.
        self.__venue = venue
        self.__proff = proff

    def __str__(self):
        return '[' + str(self.__code) + ',' + str(self.__type) + ',' + str(self.__timing) + ',' + str(self.__day) + ',' + str(self.__venue) + ',' + str(self.__proff) + ']'

class TimeTable:

    def __init__(self):
        pass

    def enroll_subject(self,course,section):
        pass

    def check_for_clashes(self):
        pass

    def export_to_csv(self,file):
        pass

def populate_section(code,type,number):
    listOfSections = []
    for i in range (number):
        hour = input("Enter the hour for" + type + " " + str(i) +": ")
        days = input("Enter the list of days for" + type + " " + str(i) +": ")
        venue = input("Enter the venue for" + type + " " + str(i) +": ")
        proff = input("Enter the list of proff for" + type + " " + str(i) +": ")
        section = Section(code, type, hour, days, venue, proff)
        listOfSections.append(section)
        ws2.append([code, type, hour, days, venue, proff])
        wb.save('TimeTable.xlsx')
    return listOfSections

def populate_course():
    wb = openpyxl.load_workbook('TimeTable.xlsx')
    ws1 = wb['Courses']
    ws2 = wb['Sections']
    sections = []
    code = input("Enter Course Code: ")
    name = input("Enter Course Name: ")
    examDate = input("Enter Course Exam Date: ")
    examTime = input("Enter Course Exam Hour: ")
    nLecture = int(input("Enter number of lecture Sections: "))
    nTutorial = int(input("Enter number of tutorial Sections: "))
    nLabs = int(input("Enter number of Laboratory Sections: "))
    sections.append(populate_section(code, "Lecture", nLecture))
    sections.append(populate_section(code, "Tutorial", nTutorial))
    sections.append(populate_section(code, "Laboratory", nLabs))
    course1 = Course(code, name, sections, examDate, examTime)
    ws1.append([code, name, examDate, examTime])
    wb.save('TimeTable.xlsx')
