import openpyxl

wb = openpyxl.load_workbook('Time Table.xlsx')
ws1 = wb['Courses']
ws2 = wb['Sections']

class Course:

    def __init__(self, code, section, examDates=0):
        self.__code = code
        self.__section = section      #This whould be a section object (aggregation)
        self.__examDates = examDates

    def get_info(self):
        pass

class Section():

    def __init__(self, hour, day, venue, proff):
        self.__timing = hour    #1st hour is 8 to 9, 2nd is 9 to 10 and so on.
        self.__day = day    #list of lecture days in a week.
        self.__venue = venue
        self.__proff = proff

class TimeTable:

    def __init__(self):
        pass

    def enroll_subject(self,course,section):
        pass

    def check_for_clashes(self):
        pass

    def export_to_csv(self,file):
        pass

def populate_subject():
    pass






































