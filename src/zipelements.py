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
                    print(f"content of {file_info.filename} is: \n{content}\n")
                    os.remove(temp_path)