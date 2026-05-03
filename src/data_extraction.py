import sys
import zipfile
import win32com.client
import os
import re
import screen
import threading
import pythoncom


data_list = [["Revisión", "Docente", "Asignatura", "Cantidad De Estudiantes Inscritos", "Unidades", "Captura de Calificaciones",
             "Aprobados", "Reprobados", "Evidencias de Aprendizaje", "% De Avance", "Producto No Conforme"]]

def read_word(route):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = None

    try:
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

        return text_complete

    finally:
        if doc:
            doc.Close()
        word.Quit()



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
    match = re.search(
        r'grupo\s*:\s(?:\([^)]*\)_*)*_*(?P<name>[^\t_\n\r]+?)(?=_{2,}|\s{2,}|[^\w\s"\'\-“”]|$|(?:\s_+|_+\s)|(?<=[a-zA-Z])_+)',
        text)

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

    pattern = re.compile(r'(\d+)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n(.*?)\n')

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

def read_zip(zip_path, console_screen):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        for file_info in  zip_ref.infolist():
            if file_info.filename.endswith('.docx') or \
                file_info.filename.endswith('.doc'):
                with zip_ref.open(file_info.filename) as file:
                    temp_path = os.path.join(os.getcwd(), f"temp_{os.path.basename(file_info.filename)}")
                    screen.print_text(console_screen, temp_path)
                    print(temp_path)
                    with open(temp_path, 'wb') as temp_file:
                        temp_file.write(file.read())
                    try:
                        content = read_word(temp_path)
                        get_data(content)
                    finally:
                        os.remove(temp_path)

def process_zip_files_in_directory(directory_path):
    for filename in os.listdir(directory_path):
        if filename.endswith('.zip'):
            element_path = os.path.join(directory_path, filename)
            read_zip(element_path)

def create_excel(path):
    pythoncom.CoInitialize()

    # Crear una instancia de Excel
    excel = win32com.client.DispatchEx("Excel.Application")

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

    normalized_path = os.path.normpath(path)
    file_path = os.path.join(normalized_path, "datos.xlsx")
    print(file_path)
    workbook.SaveAs(file_path)

    # Cerrar Excel
    excel.Quit()

    pythoncom.CoUninitialize()

def draft_function():
    process_zip_files_in_directory(r"..\data")


    for element in data_list:
        print(element)

def extract_data(data, console_screen, progress_bar):
    global data_list
    data_list = [["Revisión", "Docente", "Asignatura", "Cantidad De Estudiantes Inscritos", "Unidades",
                  "Captura de Calificaciones",
                  "Aprobados", "Reprobados", "Evidencias de Aprendizaje", "% De Avance", "Producto No Conforme"]]

    total_files = len(data.zip_files)
    progress_bar['maximum'] = total_files + 1
    if total_files > 0:
        for index, path in enumerate(data.zip_files, start=1):
            try:
                screen.print_text(console_screen, f"Path Number {index} {path}")
                read_zip(path, console_screen)
            except Exception as e:
                screen.print_text(console_screen, f"Error en archivo {path}: {e}")
            progress_bar.after(0, lambda val=index: progress_bar.config(value=val))
        if len(data.destination_path) > 0:
            screen.print_text(console_screen, str(data.destination_path))
            create_excel(data.destination_path)
            progress_bar.after(0, lambda val=total_files + 1: progress_bar.config(value=val))
        else:
            screen.print_text(console_screen, "Error with destination path")
    else:
        screen.print_text(console_screen, "Error with Files selection")
    data.zip_files.clear()
    data.destination_path = None

def run_extract_data_in_thread(data, console_screen, progress_bar):
    thread = threading.Thread(target=extract_data, args=(data, console_screen, progress_bar))
    thread.start()


