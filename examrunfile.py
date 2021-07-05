import os
import pandas as pd
import openpyxl as pyxl

sep = os.path.sep
exam_file_folder = {'rename': 'examfile_src' + sep + 'rename' + sep,
                    'subject': 'examfile_src' + sep + 'subject' + sep,
                    'teacher': 'examfile_src' + sep + 'teacher' + sep,
                    }


class ExamFolder:
    def __init__(self):
        self.exam = ['ut1', 'exam1', 'ut2', 'exam2', 'mock']
        self.subfolder = ['websams_src',
                          'websams_import',
                          'pdf',
                          'merge',
                          'markfile_src',
                          'db',
                          'analysis']



class ExamRoot:
    def __init__(self, exam):
        self.exam_list = ['ut1', 'exam1', 'ut2', 'exam2', 'mock']
        self.exam = Exam(exam)
        self.exam_run_fist = ExamRunFile  # need to further init


class Exam:
    def __init__(self, exam):
        self.exam = exam
        print('should load file list')
        print('define directory structure and create them')

    def create(self):
        print('create')

    def rename(self):
        print('rename')

    def merge(self):
        print('merge')
        ''' copy from merge'''
