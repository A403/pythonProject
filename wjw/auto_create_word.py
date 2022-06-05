import os
import zipfile


def alter(file, old_str, new_str):
    """
    替换文件中的字符串
    :param file:文件名
    :param old_str:就字符串
    :param new_str:新字符串
    :return:

    """
    file_data = ""
    with open(file, "r", encoding="utf-8") as f:
        for line in f:
            if old_str in line:
                line = line.replace(old_str, new_str)
            file_data += line
    with open(file, "w", encoding="utf-8") as f:
        f.write(file_data)


def create_docx(filename):
    f = zipfile.ZipFile(f'{filename}.docx', 'w', zipfile.ZIP_DEFLATED)
    start_dir = "template"
    for dir_path, dir_names, filenames in os.walk(start_dir):
        for filename in filenames:
            f.write(os.path.join(dir_path, filename))


# alter("template/word/document.xml", "刘业晟", "杨雨晨")
if __name__ == "__main__":
    print('[*] 程序开始运行')
    create_docx("杨雨晨")
