
class ExamClass:

    # to load the exam class data
    # on exam run sheet

    def __init__(self, examyear,
                 ind, select, key, classlevel, classcode, classtype, groupcode, subject, subj_code, subj_key,
                 path, tch, teacher, room, ut1, exam1, ut2, exam2, create, rename,
                 basename, create_folder, rename_folder):
        self.ind = ind
        self.select = select
        self.key = key
        self.classlevel = classlevel
        self.classcode = classcode
        self.class_type = classtype
        self.groupcode = groupcode
        self.subject = subject
        self.subj_code = subj_code
        self.subj_key = subj_key
        self.path = path
        self.tch = tch
        self.teacher = teacher
        self.room = room
        self.ut1 = ut1
        self.exam1 = exam1
        self.ut2 = ut2
        self.exam2 = exam2
        self.create = create
        self.rename = rename
        self.basename = basename
        self.create_folder = create_folder
        self.rename_folder = rename_folder
        # no in the run sheet
        self.exam_year = examyear


