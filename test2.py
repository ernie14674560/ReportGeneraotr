# #!/usr/bin/env python
# # _*_ coding:utf-8 _*_
import shutil
import os
from pathlib import Path

path = "C:\\Users\\ernie.chen\\Desktop\\BSA"
year_folders_paths = [f"""{path}\\{name}""" for name in os.listdir(path) if os.path.isdir(f"""{path}\\{name}""")]


def up_two_dir(file_path):
    try:
        # from Python 3.6
        new_name = file_path.split('\\')[6] + '.xls'
        parent_dir = Path(file_path).parents[2]
        # for Python 3.4/3.5, use str to convert the file_path to string
        # parent_dir = str(Path(file_path).parents[1])
        shutil.move(file_path, str(parent_dir) + '\\' + new_name)
    except IndexError:
        # no upper directory
        pass


def mass_change_loc():
    for p in year_folders_paths:
        dirs = [p + '\\' + f for f in os.listdir(p) if os.path.isdir(p + '\\' + f)]
        files_list = []
        for f in dirs:
            lst = os.listdir(f)
            lst1 = [f + '\\' + i for i in lst]
            files_list.append(lst1)
        for a in files_list:
            for b in a:
                if os.path.isfile(b):
                    up_two_dir(b)

        #
        # files_list = [os.listdir(f) for f in dirs]
        # # for file in files:
        # check_list = [p + '\\' + name for name in files if os.path.isfile(p + '\\' + name)]


def main():
    mass_change_loc()


if __name__ == '__main__':
    main()
