# -*- coding: utf-8 -*-
"""
Created on Sun Dec 24 02:24:45 2017

@author: Faraj Jasim
"""
from openpyxl import Workbook
import xlsxwriter
import xlrd
import json
import ast
import pandas as pd
import numpy
from pylab import *   
file_location = 'D:/theFinalData_movies_dataBase/SimilarityMatrix/raingDirectorsGenres3.xlsx';
#file_location2 = 'C:/Users/haide/Desktop/datafoundation/Group Coursework/Data set/IMDb_movies_write.xlsx';


wd = xlrd.open_workbook(file_location)
wr = xlsxwriter.Workbook(file_location)

worksheet = wd.sheet_by_name('raingDirectorsGenres3')
worksheet1 = wr.add_worksheet()



# col0_director_name = 0
col1_movie_name = 1


# Creates a list containing 5 lists, each of 8 items, all set to 0
movies_len = worksheet.nrows
w, h = movies_len, movies_len;
SimMatrix = [[0 for x in range(w)] for y in range(h)] 

noOfFeatures = 2
colDirectorName = 4
colGenre0 = 5
for i in range(1, worksheet.nrows):
    for j in range(1, worksheet.nrows):
        # CHECK THE SIMILARITY FOR EACH FEATURE IN THE COLUMN WITH THE OTHER
        # check if the directors are the same
        str1 = worksheet.cell_value(i,colDirectorName)
        str2 = worksheet.cell_value(j,colDirectorName)
        if str1 == str2:
            weightDirector = 1
        else:
            weightDirector = 0


        # check if the Genres are the same
        weightgenres = 0 # intialise with zero
        countReleventGenres = 0
        for genre in range(colGenre0, colGenre0+14):
            str1 = worksheet.cell_value(i,genre)
            str2 = worksheet.cell_value(j,genre)
            if str1 != "None" or str2 != "None":
                countReleventGenres = countReleventGenres + 1
                if str1 == str2:
                    weightgenres = weightgenres + 1
        weightgenres = weightgenres/countReleventGenres; # here to take the average weight for all the relevent genres
        # check other features here
        
        # Calacualte the similarity from the weights for the all the feature we considered
        SimMatrix[i][j] = (weightDirector + weightgenres) /noOfFeatures;
        
a = numpy.asarray(SimMatrix)
#a.astype(numpy.int64)
numpy.savetxt("SimilarityMatrix.csv", a, delimiter=",")

        
                    
            


#        #str1 = worksheet.cell_value(row+1,col0_director_name)

# tronsform the workbook to a list of dictionnary
#data =[]
#i = 0
#row = 0
##read from the first row
#str1 = worksheet.cell_value(row,col0_director_name)
#str2 = worksheet.cell_value(row,col1_movie_name)
#
#worksheet1.write(row, 0,str1)
#worksheet1.write(row, 1,str2)
#j = 1
#for row in range(1, worksheet.nrows):
#    data = 'str'
#    #rad from the second row and so on
#    str3 = worksheet.cell_value(row,col0_director_name)
#    str4 = worksheet.cell_value(row,col1_movie_name)
#    if not str3 and not str4:
#        #str1 = worksheet.cell_value(row+1,col0_director_name)
#        j = j
#    elif not str3:
#        worksheet1.write(j, 0,str1)
#        worksheet1.write(j, 1,str4)
#        j=j+1
#    else:
#        worksheet1.write(j, 0,str3)
#        worksheet1.write(j, 1,str4)
#        str1 = str3
#        j=j+1
