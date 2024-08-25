import shutil
import os

pwd = os.getcwd()
root_dir = os.path.abspath(os.sep)

def copy_files(source_folder, destination_folder):

    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)

    sub_folders = {}
    for item in os.listdir(source_folder):
        item_path = os.path.join(source_folder, item)
        if os.path.isdir(item_path):
            sub_folders[item] = item_path

    for folder, folder_path in sub_folders.items():
        for item in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item)
            if os.path.isfile(item_path) and os.path.splitext(item)[1] == ".xlsx":
                shutil.copy(item_path, destination_folder)
                print(f"Copied {item_path} to {destination_folder}")

source_folder = os.path.join(pwd, 'test_folder')
destination_folder = os.path.join(pwd, 'excel_files')

copy_files(source_folder, destination_folder)
