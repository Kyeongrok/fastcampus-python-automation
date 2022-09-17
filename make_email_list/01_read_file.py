"""
파일 읽고
"""


class CreateEmailList:

    def __init__(self, filename, title, text, path='data/'):
        self.filename = filename #data_file
        self.title = title#제목
        self.text = text#본문
        self.path = path

    def set_filenames(self):

        print('ee')
        self.filenames = []
