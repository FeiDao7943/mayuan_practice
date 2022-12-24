import time
import pandas as pd
import random

Problem_num = 1657      #题目总数
Column = 14             #列表总列数
Question_clo = 2        #题干所在列
Type_clo = 1            #题目类型所在列（单选多选之类多）
Answer_clo = 3          #答案所在列
Choice_clo = 6          #选项个数所在列
excel_file = pd.ExcelFile(r'马原题库练习2021.1.6(1).xls')    #打开文件


Choice_list = [["A", "B", "C", "D", "E", "F", "G"],
               [7, 8, 9, 10, 11, 12, 13]]
read_excel = pd.read_excel(excel_file, header=0)


def single_question(num_now, num_require, type_require):

    real_type = "W"
    while real_type != type_require:
        Question = random.randint(1, Problem_num)
        real_type = read_excel.iat[Question, Type_clo]
        if real_type == "单选题":
            real_type = "A"
        if real_type == "多选题":
            real_type = "B"
        if real_type == "判断题":
            real_type = "C"
        if type_require == "D":
            real_type = "D"

    Question_text = read_excel.iat[Question, Question_clo]
    Answer_text = read_excel.iat[Question, Answer_clo]
    Answer_len = len(Answer_text)
    print("================= " + str(num_now + 1) + " / " + str(num_require) + " ================ " + "Question "
          + str(Question) + " =================")
    print(str(Question_text))
    Choice_num = read_excel.iat[Question, Choice_clo]
    for i in range(int(Choice_num)):
        Choice_text = read_excel.iat[Question, Choice_list[1][i]]
        print(str(Choice_list[0][i]) + ": " + str(Choice_text))
    if Answer_len == 1:
        print("你的选择是(单选): ", end="")
    else:
        print("你的选择是(多选): ", end="")
    Answer_input = str(input()).upper()

    if str(Answer_text) == str(Answer_input):
        print("正确!")
        correct_flag = 1

    else:
        print("错误!")
        correct_flag = 0
        print("答案是: ")

        for i in range(int(Choice_num)):
            for n in range(Answer_len):
                if Choice_list[0][i] == Answer_text[n]:
                    Choice_text = read_excel.iat[Question, Choice_list[1][i]]
                    print(str(Choice_list[0][i]) + ": " + str(Choice_text))
    return correct_flag


def main():
    print("请选择题型：")
    print("A: 单选题")
    print("B: 多选题")
    print("C: 判断题")
    print("D: 所有")
    require_type = str(input()).upper()
    print("请选择题数：")
    require_num = int(input())
    total_num = 0
    correct_num = 0
    for nums in range(require_num):
        correct_flag = single_question(nums, require_num, require_type)
        total_num += 1
        if correct_flag:
            correct_num += 1
        correct_rate = 100 * correct_num / total_num
        print("总题数: " + str(total_num) + " ,  正确题数: " + str(correct_num) + " ,  正确率: " + str(
            correct_rate) + "%")
        time.sleep(0.5)
        print("（按回车键下一题）")
        input_space = 'need input'
        while input_space != "":
            input_space = str(input())

        print("")
        print("")
    print("__________________已完成!__________________")
    print("总题数: " + str(total_num) + " ,  正确题数: " + str(correct_num) + " ,  正确率: " + str(
        correct_rate) + "%")


main()