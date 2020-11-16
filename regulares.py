# -*- coding: utf-8 -*-
import xlwt
from datetime import datetime
import csv
import os,io,json

class RegularStudent:

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
                    
    def add_student(self,row_number,data):
        col_number = 0
        for item in data:
            self.ws.write(row_number,col_number,item)
            col_number += 1


    def list_of_students(self,students_file):
        #self.row_xls = 1
        with open(students_file, mode='r') as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            line_count = 0

            for row in csv_reader:
                if line_count == 0:
                    line_count += 1
                else:
                    print(f'\t{row[1]},{row[0]} {row[2]}')
                    data = []
                    data.append(row[1]+","+row[0])
                    data.append(row[2])
                    #self.ws.write(line_count,0,)
                    #self.ws.write(line_count,1,row[2])
                    col_grade = 2
                    for archivo in self.lista:

                        #entry = self.parameters["input"]+"/"+archivo[0]
                        grades = self.get_grades(row[2],archivo[0])
                        if(grades > 6):
                            data[col_grade] = grades
                            #self.ws.write(line_count,col_grade,grades)
                            col_grade += 1
                        else:
                            #entry = self.parameters["input"]+"/"+archivo[1]
                            grades = self.get_grades(row[2],archivo[1])
                            if(grades > 6):
                                #self.ws.write(line_count,col_grade,grades)
                                data[col_grade] = grades
                                col_grade += 1
                        col_grade += 1
                        #print(grades, end =" ") 
                    self.add_student(line_count,data)
                    line_count += 1

    def get_grades(self,mail,file):
        grades = 0
        with open(self.parameters["input"]+"/"+file, mode='r') as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            line_count = 0

            for row in csv_reader:
                if line_count == 0:
                    line_count += 1
                if(row[2] == mail):
                    grades = row[7]
                    break
                
        #print(type(grades))
        if(isinstance(grades,str)):
            print ("Tipo STR "+grades)
            try:
                grades = float(grades)
            except ValueError as err:
                grades = 0
            
        else:
            print(grades)
        return grades


    def create_xls(self):
        output_xls = self.parameters["output"]+"/"+datetime.now().strftime("%Y%m%d-%H%M%S")+".xls"
        self.wb = xlwt.Workbook()
        self.ws = self.wb.add_sheet("Notas",cell_overwrite_ok=True)
        self.ws.write(0, 0, "Apellido(s),Nombre")
        self.ws.write(0, 1, "Dirección Email")
        col_num = 2
        for col_name in self.lista:
            self.ws.write(0, col_num, col_name[0])
            col_num += 1

    def save_xls(self):
        output_xls = self.parameters["output"]+"/regulares_"+datetime.now().strftime("%Y%m%d-%H%M%S")+".xls"
        self.wb.save(output_xls)

    def __init__(self):
        inicio = datetime.now()
        print(inicio)
        self.load_parameters()

        self.lista = [
            ["Info-Evp1-Sistemas numericos-calificaciones.csv","Info-Recuperatorio EVP1-calificaciones.csv",1],
            ["Info-Evp2 - Estructura Secuencial-calificaciones.csv","Info-Recuperatorio EVP2-calificaciones.csv",1],
            ["Info-EVP3 - decision simple-calificaciones.csv","Info-Recuperatorio EVP3-calificaciones.csv",1],
            ["Info-EVP4 - decision doble y multiple-calificaciones.csv","Info-Recuperatorio EVP4-calificaciones.csv",1],
            ["Info-EVP5 - iteracion-calificaciones.csv","Info-Recuperatorio EVP5-calificaciones.csv",1]
        ]

        # for item in self.lista:
        #     print(item[0])

        #print(self.lista[0][0])

        self.create_xls()
        #self.create_columns_grades()
        self.list_of_students('students.csv')

        self.save_xls()





    @property
    def parameters(self):
        return self.__parameters

if __name__ == "__main__":

    regular = RegularStudent()


