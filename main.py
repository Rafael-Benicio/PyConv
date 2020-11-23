import subprocess
import PySimpleGUI as sg

sg.theme('Dark Amber')

lay=[
    [sg.InputText(text_color='black' ,key='-dir-',background_color='white'),sg.FileBrowse( file_types=(('ALL Files', '*.doc'),))],
    [sg.Button('Converter'),sg.Button('Cancelar')]]

try:
    from comtypes import client
except ImportError:
    client = None

def doc2pdf_linux(doc):

    # convert a doc/docx document to pdf format (linux only, requires libreoffice)
    # :param doc: path to document

    cmd = 'libreoffice --convert-to pdf'.split() + [doc]
    p = subprocess.Popen(cmd, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
    p.wait(timeout=10)
    stdout, stderr = p.communicate()
    if stderr:
        raise subprocess.SubprocessError(stderr)

window=sg.Window('Convert',lay)

while True:
    event, values=window.read()
    if event == sg.WIN_CLOSED or event == 'Cancelar':
        break
    if event == 'Converter':
        if values['-dir-']!='':
            doc2pdf_linux(values['-dir-'])
            break
    
window.close()