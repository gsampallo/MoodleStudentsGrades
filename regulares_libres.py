# -*- coding: utf-8 -*-
import xlwt
from datetime import datetime
import csv
import os,io,json

class GradeStudents:

    def __init__(self):
        inicio = datetime.now()
        print(inicio)
        self.load_parameters()

        self.create_xls()

        self.list_of_students()

        self.save_xls()


    def list_of_students(self):

        with open(self.parameters["list_estudents"], mode='r',encoding='UTF-8') as csv_file:
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

                    col_grade = 2
                    
                    evpAprobada = 0
                    integrador = 0

                    for col_name in self.parameters["columnas"]:

                        grades = self.get_grades(row[2],col_name["examen"])
                        data.append(grades)
                        if(grades < 6):
                            grades = self.get_grades(row[2],col_name["recuperatorio"])
                            data.append(grades)
                        else:
                            data.append('')

                        if(col_name["integradora"] == 0 and grades >= 6):
                            evpAprobada += 1
                        if(col_name["integradora"] == 1 and grades >= 6):
                            integrador += 1
                        
                    if(evpAprobada >= 4 and integrador == 1):
                        self.ws.write(line_count,22,"REGULAR")

                    else:
                        self.ws.write(line_count,22,"LIBRE")


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

        if(isinstance(grades,str)):
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
        for col_name in self.__parameters["columnas"]:
            self.ws.write(0, col_num, col_name["name"])
            col_num += 1
            self.ws.write(0, col_num, "Recuperatorio "+col_name["name"])
            col_num += 1
        
        self.ws.write(0, col_num, "ESTADO")

    def save_xls(self):
        output_xls = self.parameters["output"]+"/estado"+datetime.now().strftime("%Y%m%d-%H%M%S")+".xls"
        self.wb.save(output_xls)

    def add_student(self,row_number,data):
        col_number = 0
        for item in data:
            self.ws.write(row_number,col_number,item)
            col_number += 1

    def load_parameters(self):
        self.__parameters = ""
        with open("config.json") as f:
            self.__parameters = json.load(f)   

    @property
    def parameters(self):
        return self.__parameters

if __name__ == "__main__":

    grades = GradeStudents()            