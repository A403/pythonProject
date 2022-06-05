import os
import zipfile
import shutil
import random


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
    f = zipfile.ZipFile(f'../output/{filename}.docx', 'w', zipfile.ZIP_DEFLATED)
    start_dir = "."
    for dir_path, dir_names, filenames in os.walk(start_dir):
        for filename in filenames:
            if filename == 'auto_create_word.py':
                print("")
            else:
                f.write(os.path.join(dir_path, filename))


if __name__ == "__main__":
    print('[*] 程序开始运行')
    # student_name = ['赵博洋20', '田兴媛47', '夏广泽', '夏广泽2', '夏广泽3', '黄茗姿',
    #                 '黄茗姿2']
    #
    student_name = ['刘业晟', '刘业晟2', '刘业晟3', '刘业晟4', '刘业晟5', '滕子雨', '陆博涵', '陈泓轩', '秦朗',
                    '孙豫立', '何梓涵',
                    '陈天轶', '陈天轶2', '陈天轶3', '雷胜舟', '刘安然', '高天泽', '高括',
                    '高括2', '高括3', '高括4', '高括5', '秦朗2']
    # print(len(student_name))
    # print(len(student_count))
    target = ['增强下肢肌群力量', '增强下肢肌群力量', '增强单侧下肢稳定性', '增强单侧下肢稳定性', '增强平衡协调能力',
              '改善下肢异常姿势', '加强核心肌群力量', '增强核心力量']
    target1 = ['增强下肢力量', '加强体位转换能力', '增强单侧下肢稳定性', '增强单侧下肢稳定性', '加强步行稳定性',
               '加强步行稳定性', '加强步行稳定性', '增强核心力量']
    content = ['蹲起训练', '跳跃训练', '单腿站立', '单腿弯起', '平衡板上训练', '下肢PNF', '仰卧起坐训练', '悬吊手支撑']
    content1 = ['蹲起训练', '体位转换训练', '单腿站立', '单腿弯起', '负重步行训练', '负重上下楼梯', '跨越障碍物',
                '悬吊手支撑']
    status = ['蹲起训练，20个一组，完成2-3组，完成度888%', '跳跃3层高垫子，跳跃15次，完成度888%',
              '每侧单腿保持2min，完成度888%',
              '每侧下肢完成30次，完成度100%', '平衡板上抛接球，完成30-40次，完成度888%',
              '抗阻PNF训练，完成20-40次，完成度888%', '仰卧起坐无辅助下完成30-40次，完成度888%',
              '利用悬吊手支撑保持5分钟，完成度888%']
    status1 = ['蹲起训练，20个一组，完成2-3组，完成度888%', '从四点位到两点位到站立位体位转换，完成10-15组，完成度888%',
               '每侧单腿保持2min，完成度888%', '每侧下肢完成30次，完成度888%', '上下肢绑沙袋步行训练，行走15-20分钟',
               '辅助下负重上、下5层楼梯，完成度88%', '独立跨越低障碍物，来回5-10趟', '利用悬吊手支撑保持5分钟，完成度888%']
    ending = ['今日儿童课程较兴奋，注意力分散，训练完成质量一般。',
              '儿童今日课程注意力较集中，各项训练完成度高，规则感提高。',
              '.今日儿童情绪控制差，配合度低，给予惩罚时情绪严重，奖励配合度提高。',
              '儿童今日情绪良好，配合度较高，鼓励加少量辅助下，完成度高。',
              '儿童下肢肌力较弱，核心力量弱，在鼓励和少量辅助下完成训练。',
              '儿童上课配合度高，由于有奖励机制，训练计划完成度高。',
              '儿童配合度高，训练较之前有进步，训练中异常姿势较为明显。',
              '儿童情绪良好，喜欢被夸奖，各项训练辅助下完成度高。']
    for i in student_name:
        print(f'[*] {i}')
        # 产生随机数
        List = [0, 1, 2, 3, 4, 5, 6, 7]
        schedule = ['70', '80', '90', '100']
        random.shuffle(List)
        for k in range(0, 4):
            random.shuffle(schedule)
            print(List)
            abcd = ['a', 'b', 'c', 'd']
            print(target[List[0]])
            alter("word/document.xml", "aaa_" + abcd[k], target1[List[k]])
            print(content[List[0]])
            alter("word/document.xml", "bbb_" + abcd[k], content1[List[k]])
            print(status[List[0]].replace('888', schedule[k]))
            alter("word/document.xml", "ccc_" + abcd[k], status1[List[k]].replace('888', schedule[0]))
        random.shuffle(ending)
        alter("word/document.xml", "user_name", i)
        alter("word/document.xml", "ending", ending[0])
        create_docx(i)
        shutil.copy('../bak/document.xml', 'word/document.xml')
