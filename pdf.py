# pdf_merger2.py
import glob
from PyPDF2 import PdfFileMerger


def merger(output_path, input_paths):
    pdf_merger = PdfFileMerger()
    file_handles = []

    for path in input_paths:
        pdf_merger.append(path)

    with open(output_path, 'wb') as file_object:
        pdf_merger.write(file_object)


# if __name__ == '__main__':
#    paths = glob.glob('fw9_*.pdf')
#    paths.sort()
#    merger('pdf_merger2.pdf', paths)

