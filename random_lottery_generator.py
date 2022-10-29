"""
@File    :   random_lottery_generator.py
@Time    :   2022/10/29
@Author  :   Kyochul Jang, Sangwon Seo, Dogun Lim
@Version :   3.10
@Contact :   gcj1234567890@gmail.com
@License :
@Desc    :   Python program to generate a set of lottery numbers
"""

import random
import xlsxwriter


def generator():
    row, col = 0, 0
    workbook = xlsxwriter.Workbook(f'test_lottery.xlsx')
    worksheet = workbook.add_worksheet()
    randnum_list = []

    for j in range(30):
        col = 0
        randNum = []

        for i in range(6):
            n = random.randint(1, 45)

            cnt = 0
            while n in randNum:
                n = random.randint(1, 45)

            randNum.append(n)
            worksheet.write(row, col, randNum[i])

            col += 1
        randnum_list.append(randNum)

        row += 1

    workbook.close()

    return randnum_list


if __name__ == "__main__":
    print(generator())
