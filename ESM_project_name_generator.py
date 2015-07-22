from Tkinter import Tk
from tkFileDialog import askopenfilename
import xlrd, os, datetime, re

class generateProjectName():
    def getIPPFilePath():
        Tk().withdraw()
    
        options = {}
        options['initialdir'] = "B:\\Production Control\\Planning\\IPP (Individual Publication Plan)\\ESM individual publication plan"
        options['title'] = 'ESM Project File Name Generator'
        options['filetypes'] = [('Excel files', ".xls")]
    
        filePath = askopenfilename(**options)
        return filePath 
    
    def generateFileName(filePath):
        wb = xlrd.open_workbook(filePath, formatting_info=True)
        sh = wb.sheet_by_index(0)
        
        project_type = "ESM"
        model = sh.cell(3, 5).value  # get the model code
        nissan_project_code = sh.cell(10, 5).value  # get the project code used by nissan
        if nissan_project_code == "":
            nissan_project_code = sh.cell(10, 7).value
        mebv_project_code = nissan_project_code[nissan_project_code.find("(") + 1:nissan_project_code.find(")")]  # extract only the string located between paranthesis
        publication_code = sh.cell(16, 28).value  # get the publication code
        if publication_code == "MEX" or publication_code == "":
            publication_code = sh.cell(15, 26).value
            if publication_code == "NL" or publication_code == "":
                publication_code = sh.cell(15, 28).value
                if publication_code == "NL" or publication_code == "":
                    publication_code = sh.cell(16, 26).value
        print "Pub. Code: %s" %publication_code
        today = datetime.datetime.now()  # get the current date 
        current_month = today.strftime("%B")  # return the month's name
        current_year = today.year
        
        string = "%s %s %s %s %s %s" % (project_type, model, mebv_project_code, publication_code, current_month, current_year)
        
        print(string)
          
        langrow = 16
#        print langrow
        if sh.cell(langrow, 6).value == "Europe":
            langrow = 15
#        print langrow # test purpose

        print "=" * 50
        print "Languages for this project"
        print "=" * 50
        
        for i in range(6, 27, 2):
            xfx = sh.cell_xf_index(langrow, i)
            xf = wb.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
#            print bgx
            if bgx != 55 and bgx != 14 and bgx != 23:
                language = sh.cell(langrow, i).value
                print ("!!! --- Excluded language: %s" % language)
#                print "Column: %s" %i # test purpose
            else:
                language = sh.cell(langrow, i).value
                print (language)
    
    generateFileName(getIPPFilePath())
