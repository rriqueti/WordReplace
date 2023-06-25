import tkinter as tk
from tkinter import Tk, Button, filedialog
from objetcts_files import *

root = tk.Tk()
root.title('DocWord Replace')
root.geometry('500x300')
             
class ButtonWord:
    def __init__(self): 
        self.word_file = tk.Button(root, text='Pesquisar arquivo Word', command=self.word_filedialog)
        self.word_file.place(x=50,y=50)
        self.word_file_dialog = None

    def word_filedialog(self):
        word_file_dialog = filedialog.askopenfilename()
        self.label_word_anexado(word_file_dialog)
        self.word_file_dialog = word_file_dialog
    
    def label_word_anexado(self, diretorio):
        self.label1 = tk.Label(root, text=diretorio)
        self.label1.place(x=190,y=50)

class ButtonExcel(ButtonWord):
    def __init__(self):
        super().__init__()
        self.excel_file = tk.Button(root, text='Pesquisar arquivo Excel', command=self.excel_filedialog)
        self.excel_file.place(x=50,y=100)
        self.excel_file_dialog = None
        
    def excel_filedialog(self):
        excel_file_dialog = filedialog.askopenfilename()
        self.label_excel_anexado(excel_file_dialog) ### Label do arquivo anexado
        self.excel_file_dialog = excel_file_dialog
    
    def label_excel_anexado(self, diretorio):
        self.label2 = tk.Label(root, text=diretorio)
        self.label2.place(x=190,y=100)

class ButtonSaveDiretory(ButtonExcel):
    def __init__(self):
        super().__init__()
        self.readme_button = tk.Button(root, text='Salvar em:', command=self.save_filedialog)
        self.readme_button.place(x=120,y=150)
        
    def save_filedialog(self):
        save_file_dialog = filedialog.askdirectory()
        self.label_save_anexado(save_file_dialog) ### Label do arquivo anexado
        self.save_file_dialog = save_file_dialog
        
    
    def label_save_anexado(self, diretorio):
        self.label2 = tk.Label(root, text=diretorio)
        self.label2.place(x=190,y=152)

    def save_file(self, files, file_name):
        diretory_filename = self.save_file_dialog + "/" + file_name + ".docx"
        print(diretory_filename)
        files.save(diretory_filename) 
        
class ExecuteButton(ButtonSaveDiretory):
    def __init__(self):
        super().__init__()
        self.execute_button = tk.Button(root, text='Substituir palavras', command=self.execute)
        self.execute_button.place(x=75,y=200) 
        
    def conclued(self):
        concluded_window = tk.Toplevel()
        concluded_window.title('Concluido com Sucesso')
        concluded_window.geometry('200x50')
        
        concluded_label = tk.Label(concluded_window, text='Palavras substituidas com sucesso')
        concluded_label.pack()
        
        concluded_label.mainloop()  
    
    def exception(self):
        except_window = tk.Toplevel()
        except_window.title('ERROR')
        except_window.geometry('200x50')
        
        except_label = tk.Label(except_window, text='Formato de arquivo nao aceito')
        except_label.pack()
        
        except_label.mainloop() 
     
    def execute(self):
        """validação se foi anexada uma planilha"""
        try:
            data_xlsx = DataFrame(exec.excel_file_dialog)
        except:
            self.exception()
            
        rows = data_xlsx.row_values()
        contador = 0 

        for r in rows:
            """Nome do arquivo salvo, é uma tupla com os valores da primeira e segunda coluna"""
            file_name = r[0] + " " + r[1]
            with OpenDoc(exec.word_file_dialog) as files:
                """Base de dados"""
                df_excel = data_xlsx.df
                """Head será a lista das palavras chaves procuradas no documento para substituir"""
                head = data_xlsx.replace_words()
                """Variável que dará todos os paragráfos do documento"""
                paragraphs_doc = files.paragraphs
                """ W representará as palavras chaves a serem substituidas, utilizando a primeira linha de cada coluna"""
                for w in head:
                    """Cada palavra W passará por todos os paragrafos porcurando substituição"""
                    for linha in paragraphs_doc:
                        """Validação se a palavra a ser substituida aparece no pagrafo"""
                        if w in linha.text:
                            linha.text = linha.text.replace(str(w), str(df_excel.loc[contador, w]))
                    exec.save_file(files, file_name)
            contador += 1      

        self.conclued()

exec = ExecuteButton()
root.mainloop()






