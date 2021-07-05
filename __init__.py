from websams.model.examrun import ExamRun
from websams.model.examclass import ExamClass
from websams.model.examfile import ExamFile
from websams.model.assessment import Assessment
from websams.model.student import Student
from websams.model.student import ClassList
from websams.model.dic import Conduct
from websams.model.report import Report
#from websams.model.reportcard import RecordCard
#from websams.model.detention import Detention
'''

ExamRun :  Define how to run exam
ExamClass : Define Each exam class details
ExamFile : Define Each File of each exam class
Assessment : Define how to run assessment
ClassList : Define class list and generate all lists of students
Conduct : Define conduct file and generate conduct files
Report : Define reportp files and generate school report
Detention :

'''

__all__ = ['ExamRun',
           'ExamClass',
           'ExamFile',
           'Assessment',
           'Student',
           'ClassList',
           'Conduct',
           'Report',
           #'RecordCard',
           #'Detention'
           ]

