import sys

import os
import pyexcel as p

class XLSCheck:
    @staticmethod
    def checkInput(ipFile):
        # Check Excel format
        fileName = ipFile
        splitName = fileName.split('.')
        splitName[1] = fileName[-3:]
        if (splitName[1] == 'xls'):
            targetInputFile = splitName[0] + '.xlsx'
            print('Its a xls file. Converting to .xlsx file..')
            source_f="C:\\Users\\Vignesh\\PycharmProjects\\AutoExcel\\src\\uploads\\"+fileName
            dest_f = "C:\\Users\\Vignesh\\PycharmProjects\\AutoExcel\\src\\uploads\\" + targetInputFile
            p.save_book_as(file_name=source_f, dest_file_name=dest_f)
            XLSCheck.clean() #delete old file
            return targetInputFile
        elif (splitName[1] == 'lsx'):
            print('Its a xlsx file')
            targetInputFile = fileName
            return targetInputFile
        else:
            print('Invalid file format! Use either .xls or .xlsx')
            sys.exit()

    @staticmethod
    def clean():
        for the_file in os.listdir('uploads'):
            file_path = os.path.join('uploads', the_file)
            try:
                if the_file.endswith('xls'):
                    os.unlink(file_path)
            except Exception as e:
                print(e)