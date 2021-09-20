import pandas as pd
import os
import pathlib
from model.examclass import ExamClass

db_header = ['regno', 'enname', 'chname', 'sex',  # 0-3
             'classlevel', 'classcode', 'classno', 'subject', 'xgroup',  # 4-8
             'ut1', 'daily1', 'exam1', 'total1',  # 9-12
             'ut2', 'daily2', 'exam2', 'total2', 'final',  # 13-17
             't1comp1', 't1comp2', 't1comp3', 't1comp4', 't1comp5',  # 18-22
             't2comp1', 't2comp2', 't2comp3', 't2comp4', 't2comp5',  # 23-27
             'fcomp1', 'fcomp2', 'fcomp3', 'fcomp4', 'fcomp5',  # 28-32
             ]

mark_col_dict = {'cat1': {'ut1': ['ut1'], 'exam1': ['daily1', 'exam1', 'total1'],
                          'ut2': ['ut2'], 'exam2': ['daily2', 'exam2', 'total2']},
                 'cat2': {'ut1': ['ut1'], 'exam1': ['t1comp1', 't1comp2', 't1comp3', 't1comp4', 't1comp5', 'total1'],
                          'ut2': ['ut2'], 'exam2': ['t1comp1', 't1comp2', 't1comp3', 't1comp4', 't1comp5', 'total2']},
                 'cat3': {'ut1': [], 'exam1': [], 'ut2': [], 'exam2': ['total2']}}

comp_col = {'ut1': [0],
            'exam1': ['t1comp1', 't1comp2', 't1comp3', 't1comp4', 't1comp5'],
            'ut2': [0],
            'exam2': ['t2comp1', 't2comp2', 't2comp3', 't2comp4', 't2comp5'],
            'final': ['fcomp1', 'fcomp2', 'fcomp3', 'fcomp4', 'fcomp5'],
            'mock': ['fcomp1', 'fcomp2', 'fcomp3', 'fcomp4', 'fcomp5'],
            }

check_mark_col = {'ut1': ['ut1'],
                  'exam1': ['daily1', 'exam1', 'total1'],
                  'ut2': ['ut2'],
                  'exam2': ['daily2', 'exam2', 'total2'],
                  'final': ['daily2', 'exam2', 'total2'],
                  'mock': ['daily2', 'exam2', 'total2'],
                  }

stat_mark_col = {'ut1': 'ut1',
                 'exam1': 'total1',
                 'ut2': 'ut2',
                 'exam2': 'total2',
                 'final': 'total2',
                 'mock': 'total2',
                 }

db_filter_header = {'ut1': db_header[0:9] + db_header[9:10],
                    'exam1': db_header[0:9] + db_header[12:13] + db_header[18:22],
                    'ut2': db_header[0:9] + db_header[13:14],
                    'exam2': db_header[0:9] + db_header[16:17] + db_header[23:27],
                    'mock': db_header[0:9] + db_header[17:18] + db_header[28:32],
                    }


class ExamFile:

    # to load the exam class data
    # on exam run sheet
    # can add weight data and load weight later

    def __init__(self, exam_class: ExamClass, file_path: str, assessment_name: str):
        self.exam_class = exam_class
        self.classlevel = exam_class.classlevel
        self.classcode = exam_class.classcode
        self.class_type = exam_class.class_type
        self.groupcode = exam_class.groupcode
        self.subject = exam_class.subject
        self.subj_code = exam_class.subj_code
        self.subj_key = exam_class.subj_key
        self.path = exam_class.path
        self.tch = exam_class.tch
        self.teacher = exam_class.teacher
        self.room = exam_class.room
        self.basename = exam_class.basename
        self.exam_file_name = assessment_name + exam_class.basename
        self.file_path = file_path
        self.file_to_rename = ''
        self.file_state = -1

        # exam_type: ut1 / exam1 / ut2 / final / mock
        self.exam_type = assessment_name[4:]
        self.check_mark_columns = check_mark_col[self.exam_type]
        self.comp_columns = comp_col[self.exam_type]
        self.mark_column = stat_mark_col[self.exam_type]

        if self.subject in ['mth', 'm1', 'm2', 'chs', 'hst', 'geo', 'eco', 'isc',
                            'phy', 'chm', 'bio', 'ict', 'baf', 'pth']:
            self.cat = 'cat1'
        elif self.subject in ['eng', 'chi']:
            self.cat = 'cat2'
        else:
            self.cat = 'cat3'

        self.mark_column_dict = mark_col_dict[self.cat][self.exam_type]

        self.db_df = None
        self.pass_mark = 40 if self.classlevel.lower() in ['s5', 's6'] else 50
        self.statistics = {}

        # folder setting here
        # self.home_folder = os.path.abspath(os.getcwd())
        # self.template_folder = os.path.join(self.home_folder, 'template')
        # self.html_template_folder = os.path.join(self.template_folder, 'html')

    def check_file(self):
        temp_file_path = pathlib.Path(self.file_path)

        if not temp_file_path.exists():
            print('{}:{} does not exist.'.format(self.teacher, self.file_path))
            self.file_state = 0
        elif not temp_file_path.is_file():
            print('{}:{} is not a file.'.format(self.teacher, self.file_path))
            self.file_state = 0
        else:
            self.file_state = 1

        return self.file_state

    def file_state_list(self):
        file_state_list = [self.classlevel, self.classcode, self.class_type, self.groupcode, self.subject,
                           self.subj_code, self.subj_key, self.path, self.tch, self.teacher,
                           self.room, self.exam_file_name, self.file_path, self.file_state]
        return file_state_list

    def load_db_df(self):
        rename_header_list = {'UT1': 'ut1',
                              'Daily1': 'daily1',
                              'Exam1': 'exam1',
                              'Total1': 'total1',
                              'UT2': 'ut2',
                              'Daily2': 'daily2',
                              'Exam2': 'exam2',
                              'Total2': 'total2',
                              'Final': 'final',
                              'T1Comp1': 't1comp1',
                              'T1Comp2': 't1comp2',
                              'T1Comp3': 't1comp3',
                              'T1Comp4': 't1comp4',
                              'T1Comp5': 't1comp5',
                              'T2Comp1': 't2comp1',
                              'T2Comp2': 't2comp2',
                              'T2Comp3': 't2comp3',
                              'T2Comp4': 't2comp4',
                              'T2Comp5': 't2comp5',
                              'FComp1': 'fcomp1',
                              'FComp2': 'fcomp2',
                              'FComp3': 'fcomp3',
                              'FComp4': 'fcomp4',
                              'FComp5': 'fcomp5',
                              }

        self.db_df = pd.read_excel(self.file_path, sheet_name='db', index_col=False)
        self.db_df.dropna(subset=['regno'], inplace=True)
        self.db_df.rename(columns=rename_header_list, inplace=True)

    def to_dict(self):
        if self.db_df is None:
            self.load_db_df()
        return self.db_df.to_dict('records')

    def db_to_print(self):
        if self.db_df is None:
            self.load_db_df()
        # self.db_df.to_excel('db.xlsx')
        self.db_df['key'] = self.db_df['classcode'] + self.db_df['classno'].apply(lambda x: '('+str(int(x)).zfill(2)+')')
        self.db_df['name2'] = self.db_df['enname'] + ' (' + self.db_df['chname'] + ')'
        # self.db_df['mark'] = self.db_df[self.mark_column].apply(lambda x: '{:.2f}'.format(x))

        temp_col = []
        for col in self.mark_column_dict:
            self.db_df[col + '_f'] = self.db_df[col].apply(lambda x: '{:.2f}'.format(x))
            self.db_df[col + '_p'] = self.db_df[col] >= self.pass_mark
            temp_col.append(col + '_f')
            temp_col.append(col + '_p')

        db_filter_columns = ['key', 'name2'] + temp_col
        return self.db_df[db_filter_columns]

    def get_class_statistics(self):
        if not self.statistics:

            if self.db_df is None:
                self.load_db_df()
            # self.db_df['class_rank'] = self.db_df[self.mark_column].rank(ascending=False)

            statistics_item_list1 = ['No of Ss', 'No of Pass', 'NoFail', 'NoZero', 'Passing%',
                                     'Mean', 'SD', 'Max', 'Q3', 'Q2', 'Q1', 'Min']
            statistics_item_list2 = ['0 - 10 (excluding 10)', '10 - 20 (excluding 20)',
                                     '20 - 30 (excluding 30)', '30 - 40 (excluding 40)',
                                     '40 - 50 (excluding 50)', '50 - 60 (excluding 60)',
                                     '60 - 70 (excluding 70)', '70 - 80 (excluding 80)',
                                     '80 - 90 (excluding 90)', '90 - 100']

            # bin = [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100]

            temp_df = self.db_df[self.mark_column]
            self.statistics['No of Ss'] = self.db_df['regno'].count()
            self.statistics['No of Pass'] = temp_df[temp_df >= self.pass_mark].count()
            self.statistics['NoFail'] = temp_df[temp_df < self.pass_mark].count()
            self.statistics['NoZero'] = temp_df[temp_df == 0].count()
            self.statistics['Passing%'] = round(self.statistics['No of Pass']/self.statistics['No of Ss']*100, 2)
            self.statistics['Mean'] = round(self.db_df[self.mark_column].mean(), 2)
            self.statistics['SD'] = round(self.db_df[self.mark_column].std(), 2)
            self.statistics['Max'] = round(self.db_df[self.mark_column].max(), 2)
            self.statistics['Q3'] = round(self.db_df[self.mark_column].quantile(0.75), 2)
            self.statistics['Q2'] = round(self.db_df[self.mark_column].quantile(0.5), 2)
            self.statistics['Q1'] = round(self.db_df[self.mark_column].quantile(0.25), 2)
            self.statistics['Min'] = round(self.db_df[self.mark_column].min(), 2)

        return self.statistics



