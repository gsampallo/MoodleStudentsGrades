# -*- coding: utf-8 -*-
import xlwt
from datetime import datetime
import csv
import os,io,json

class StudentGrades:

    def load_parameters(self):
        self.__parameters = ""
        with open("config.json") as f:
            self.__parameters = json.load(f)   


    def create_columns_grades(self):
        self.entries = os.listdir(self.parameters["input"]) 
        self.entries.sort()       
        col_number = 2
        for entry in self.entries:
            #print(entry)
            col_name = entry.split(".")[0]
            self.ws.write(0, col_number, col_name)
            col_number += 1
            
            # with open(self.parameters["input"]+"/"+entry, mode='r') as csv_file:
            #     csv_reader = csv.reader(csv_file, delimiter=',')
            #     line_count = 0

            #     for row in csv_reader:
            #         if line_count == 0:
            #             print(f'Column names are {", ".join(row)}')
            #             line_count += 1
            #         else:
            #             print(f'\t{row[0]},{row[1]} {row[2]} - {row[7]}.')
            #             line_count += 1
                    
                
    def list_of_students(self,students_file):
        #self.row_xls = 1
        with open(students_file, mode='r') as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            line_count = 0

            for row in csv_reader:
                if line_count == 0:
                    line_count += 1
                else:
                    print(f'\t{row[0]},{row[1]} {row[2]}')

                    self.ws.write(line_count,0,row[0]+","+row[1])
                    self.ws.write(line_count,1,row[2])
                    col_grade = 2
                    for entry in self.entries:
                        grades = self.get_grades(row[2],entry)
                        self.ws.write(line_count,col_grade,grades)
                        col_grade += 1
                        #print(grades, end =" ") 

                    line_count += 1

    def get_grades(self,mail,file):
        with open(self.parameters["input"]+"/"+file, mode='r') as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            line_count = 0

            for row in csv_reader:
                if line_count == 0:
                    line_count += 1
                if(row[2] == mail):
                    return row[7]

    def create_xls(self):
        output_xls = self.parameters["output"]+"/"+datetime.now().strftime("%Y%m%d-%H%M%S")+".xls"
        self.wb = xlwt.Workbook()
        self.ws = self.wb.add_sheet("Notas",cell_overwrite_ok=True)
        self.ws.write(0, 0, "Apellido(s),Nombre")
        self.ws.write(0, 1, "Dirección Email")

    def save_xls(self):
        output_xls = self.parameters["output"]+"/"+datetime.now().strftime("%Y%m%d-%H%M%S")+".xls"
        self.wb.save(output_xls)

    def __init__(self):
        inicio = datetime.now()
        print(inicio)
        self.load_parameters()

        self.create_xls()
        self.create_columns_grades()
        self.list_of_students('alumnos.csv')

        self.save_xls()





    @property
    def parameters(self):
        return self.__parameters

if __name__ == "__main__":

    grades = StudentGrades()
    #grades.load_parameters()
    #grades.create_xls()
    #grades.list_of_files()
    
    #col_grade = 7

