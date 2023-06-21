
import pandas as pd
from docx import Document


class DataFrame:
    def __init__(self, datafile):
        self.df = pd.read_excel(datafile) 
        self.key_words = list(self.df.columns)  
        self.len_columns = (len(self.key_words))
    
    def row_values(self):
        return self.df.values
    
    def replace_words(self):
        return self.key_words

class OpenDoc:
    def __init__(self, filename):
        self.filename = filename
        self._file = None
        self.save = None

    def __enter__(self):
        self._file = Document(self.filename)
        return self._file
    
    def __exit__(self, exceptclass, exception_, traceback_): 
        ...