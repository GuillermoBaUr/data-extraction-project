import sys
import zipfile
import win32com.client
import os
import re

def read_word(route):

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(route)

    # Getting Body Content
    body_content = doc.Content.Text
    body_content = body_content.replace("\r", "\n").replace("\x07", "\n")

    # Getting Header Content
    section = doc.Sections(1)
    header = section.Headers(1)
    header_text = header.Range.Text.strip()
    header_text = header_text.replace("\r", "").replace("\x07", "").replace("\x01", "")

    text_complete = (header_text + " "+ body_content).lower()

    doc.Close()
    word.Quit()

    return text_complete

def get_revision_number_string(text):
    if "revisión:" in text:
        start_index = text.find("revisión:") + len("revisión:")
        revision_str = ""
        for char in text[start_index:].strip():
            if char.isdigit():
                revision_str += char
            else:
                break
        if revision_str:
            return revision_str
    return None

def get_professor_name(text):
    match = re.search(r"docente\s*:\s*(?:\([^)]*\)_*)*_*(?P<name>[a-záéíóúñ\s][^\t_\n\r]+)", text)
    if match:
        return match.group("name").strip()
    return None

def get_subject_name(text):
    # match = re.search(r"grupo\s*:\s*_*(?P<name>[^\t_\n\r]+?)(?=_{2,}|\s{2,}|[^\w\s]|$)", text)
    # match = re.search(r'grupo\s*:\s(?:\([^)]*\)_*)*_*(?P<name>[^\t_\n\r]+?)(?=_{2,}|\s{2,}|[^\w\s"\'\-“”]|$|(?:\s_+|_+\s))', text)
    match = re.search(
        r'grupo\s*:\s(?:\([^)]*\)_*)*_*(?P<name>[^\t_\n\r]+?)(?=_{2,}|\s{2,}|[^\w\s"\'\-“”]|$|(?:\s_+|_+\s)|(?<=[a-zA-Z])_+)', text)
    if match:
        return match.group("name").split("cantidad")[0].strip()
    return None

def get_students_quantity(text):
    match = re.search(r"inscritos\s*:\s(?:\([^)]*\)_*)*_*(?P<name>[^\t_\n\r]+)_*", text)
    if match:
        return match.group("name").strip()
    return None

def get_modules(text, course_information):
    backup_list = course_information[:]

    pattern = re.compile(r'\n(10|[1-9])\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n')

    matches = re.findall(pattern, text)

    for match in matches:
        for index in range(len(match)):
            if match[index]:
                course_information.append(match[index])
            else:
                course_information.append('sin informacion')

        data_list.append(course_information)
        course_information = backup_list[:]

def get_data(text):
    revision_number = get_revision_number_string(text)
    teacher = get_professor_name(text)
    subject = get_subject_name(text)
    student_quantity = get_students_quantity(text)
    course_data = [revision_number, teacher, subject, student_quantity]

    get_modules(text, course_data)

def read_zip(zip_path):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        for file_info in  zip_ref.infolist():
            if file_info.filename.endswith('.docx') or \
                file_info.filename.endswith('.doc'):
                with zip_ref.open(file_info.filename) as file:
                    temp_path = os.path.join(os.getcwd(), f"temp_{os.path.basename(file_info.filename)}")
                    print(temp_path)
                    with open(temp_path, 'wb') as temp_file:
                        temp_file.write(file.read())
                    content = read_word(temp_path)
                    get_data(content)
                    os.remove(temp_path)

def process_zip_files_in_directory(directory_path):
    for filename in os.listdir(directory_path):
        if filename.endswith('.zip'):
            element_path = os.path.join(directory_path, filename)
            read_zip(element_path)

def create_excel():
    # Crear una instancia de Excel
    excel = win32com.client.Dispatch("Excel.Application")

    # Crear un nuevo libro
    workbook = excel.Workbooks.Add()

    # Seleccionar la primera hoja
    sheet = workbook.Worksheets(1)

    # Agregar los datos a la hoja
    for row_index, row in enumerate(data_list):
        for col_index, value in enumerate(row):
            sheet.Cells(row_index + 1, col_index + 1).Value = value

    # Guardar el archivo
    current_directory = os.getcwd()
    file_path = os.path.join(current_directory, "datos.xlsx")
    workbook.SaveAs(file_path)

    # Cerrar Excel
    excel.Quit()

data_list = [["Revisión", "Docente", "Asignatura", "Cantidad De Estudiantes Inscritos", "Unidades", "Captura de Calificaciones",
         "Aprobados", "Reprobados", "Evidencias de Aprendizaje", "% De Avance", "Producto No Conforme"]]

process_zip_files_in_directory(r"..\data")


for element in data_list:
    print(element)

create_excel()