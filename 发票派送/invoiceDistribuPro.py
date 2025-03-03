import json
import os
import openpyxl
import zipfile

def read_config(config_path):
    try:
        if not os.path.exists(config_path):
            print(f"配置文件 '{config_path}' 不存在")
            return None
        with open(config_path, 'r', encoding='utf-8') as config_file:
            return json.load(config_file)
    except Exception as e:
        print(f"读取配置文件出错: {e}")
        return None

def create_directory(workbook_path, file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    keyword_list = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        to_email = row[0]
        keyword = row[1]
        keyword_list.append((to_email, keyword))
    return keyword_list

def create_directories(workbook_path, directory_list):
    for to_email, keyword in directory_list:
        directory_path = os.path.join(workbook_path, to_email)
        files = []
        for root, _, filenames in os.walk(workbook_path):
            for filename in filenames:
                if keyword in filename:
                    files.append(os.path.join(root, filename))
        if files:
            if not os.path.exists(directory_path):
                try:
                    os.makedirs(directory_path)
                    print(f"已成功创建目录 '{to_email}'。")
                except Exception as e:
                    print(f"创建目录出错: {e}")
            else:
                print(f"目录 '{to_email}' 已存在。")
            for file in files:
                new_file_path = os.path.join(directory_path, os.path.basename(file))
                try:
                    os.rename(file, new_file_path)
                    print(f"已成功移动文件 '{os.path.basename(file)}' 到目录 '{to_email}'。")
                except Exception as e:
                    print(f"移动文件出错: {e}")
        else:
            print(f"没有找到包含 '{keyword}' 的文件。")

def create_zip(directory_path):
    root_dir = directory_path
    for folder_name in os.listdir(root_dir):
        folder_path = os.path.join(root_dir, folder_name)
        if os.path.isdir(folder_path):
            zip_file_name = f"{folder_name}.zip"
            zip_file_path = os.path.join(root_dir, zip_file_name)
            with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for root, _, files in os.walk(folder_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        zip_file.write(file_path, os.path.relpath(file_path, root_dir))
    print("打包完成！")

def main():
    config_path = os.path.join(os.getcwd(), 'config.ini')  # 确保文件名正确
    config = read_config(config_path)
    if config is None:
        return
    file_path = config.get('file_path')
    directory_path = config.get('directory_path')
    workbook_path = os.path.dirname(file_path)
    directory_list = create_directory(workbook_path, file_path)
    if directory_list is not None:
        create_directories(workbook_path, directory_list)
    create_zip(directory_path)

if __name__ == "__main__":
    main()