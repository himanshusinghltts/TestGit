import os
from importxls import ClassData
import matplotlib.pyplot as plt
import pandas as pd
import scipy.stats
import statistics
import numpy as np


def main():
    filepath = os.getcwd()
    filepath = filepath + "/Mark_Sheet.xlsx"
    classmarks = ClassData(filepath)
    print("\n")
    print("\033[1m"+"\033[4m"+'\033[36m'+"Automation - Analysis and Visualisation of Result"+'\033[0m')
    print("\n"+"\033[92m"+"Press 1 to generate the whole class Report")
    print("Press 2 for analysing the marks of specific person"+"\033[0m")
    choice = input()
    if choice == str(1):
        callfunction(classmarks)
    elif choice == str(2):
        average_marks_three_subs(classmarks)
        pass
    else:
        print("\033[91m"+"Please Enter the Valid number"+"\033[0m")
        main()


def callfunction(classmarks):
    """
    The main function
    :return: void
    """
    print("\n"+"\033[4m"+"BEST PERFORMERS"+"\033[0m")
    maximummarksin_sdlc(classmarks)
    maximum_marksin_python(classmarks)
    maximum_marksin_mbse(classmarks)
    print("\n"+"\033[4m"+"AVERAGE MARKS"+"\033[0m")
    average_marks_sdlc(classmarks)
    average_marks_python(classmarks)
    average_marks_mbse(classmarks)
    print("\n"+"\033[4m""POOR PERFORMERS"+"\033[0m")
    minimum_marksin_mbse(classmarks)
    minimum_marksin_python(classmarks)
    minimum_marksin_sdlc(classmarks)
    print("\n"+"\033[4m"+"PERFORMANCE OF STUDENTS IN SDLC"+"\033[0m")
    meeting_expectationsin_sdlc(classmarks)
    print("\n"+"\033[4m"+"PERFORMANCE OF STUDENTS IN Python"+"\033[0m")
    meeting_expectationsin_python(classmarks)
    print("\n"+"\033[4m"+"PERFORMANCE OF STUDENTS IN MBSE"+"\033[0m")
    meeting_expectationsin_mbse(classmarks)
    sdlc_graph()
    python_graph()
    mbse_graph()
    average_graph()
    bell_curve()
    print("\n"+"\033[92m"+"Press 1 return to main menu")
    print("Press 2 to exit"+"\033[0m")
    cho = int(input())
    if cho == 1:
        main()
    else:
        pass


def maximummarksin_sdlc(classmarks):
    """
    it prints the name of the person who scored the maximum marks in SDLC module
    :param classmarks: object of ClassData
    :return: Void
    """
    ps_numbers = classmarks.listps()
    name = None
    maxmarks = 0
    for i in ps_numbers:
        currmarks = classmarks.getmarks_byps_sdlc(i)
        if currmarks > maxmarks:
            maxmarks = currmarks
            name = classmarks.getname(i)
        else:
            pass
    print(name + " scored maximum marks in SDLC")


def meeting_expectationsin_sdlc(classmarks):
    """
    it prints the name of the person meeting expectations or not in SDLC
    :param classmarks: object of ClassData
    :return: Void
    """
    ps_numbers = classmarks.listps()
    for i in ps_numbers:
        currmarks = classmarks.getmarks_byps_sdlc(i)
        name = classmarks.getname(i)
        if currmarks > 60:
            print(name, " is meeting expectations in SDLC")
        else:
            print(name, " is not meeting expectations in SDLC")


def meeting_expectationsin_python(classmarks):
    """
    it prints the name of the person meeting expectations or not in Python
    :param classmarks: object of ClassData
    :return: Void
    """
    ps_numbers = classmarks.listps()
    for i in ps_numbers:
        currmarks = classmarks.getmarks_byps_python(i)
        name = classmarks.getname(i)
        if currmarks > 60:
            print(name, " is meeting expectations in Python")
        else:
            print(name, " is not meeting expectations in Python")


def meeting_expectationsin_mbse(classmarks):
    """
    it prints the name of the person meeting expectations or not in Python
    :param classmarks: object of ClassData
    :return: Void
    """
    ps_numbers = classmarks.listps()
    for i in ps_numbers:
        currmarks = classmarks.getmarks_byps_mbse(i)
        name = classmarks.getname(i)
        if currmarks > 60:
            print(name, " is meeting expectations in MBSE")
        else:
            print(name, " is not meeting expectations in MBSE")


def maximum_marksin_python(classmarks):
    """
    it prints the name of the person who scored the maximum marks in Python module
    :param classmarks: object of ClassData
    :return: Void
    """
    ps_numbers = classmarks.listps()
    name = None
    maxmarks = 0
    for i in ps_numbers:
        currmarks = classmarks.getmarks_byps_python(i)
        if currmarks > maxmarks:
            maxmarks = currmarks
            name = classmarks.getname(i)
        else:
            pass
    print(name + " scored maximum marks in python")


def maximum_marksin_mbse(classmarks):
    """
    it prints the name of the person who scored the maximum marks in MBSE module
    :param classmarks: object of ClassData
    :return: Void
    """
    ps_numbers = classmarks.listps()
    name = None
    maxmarks = 0
    for i in ps_numbers:
        currmarks = classmarks.getmarks_byps_mbse(i)
        if currmarks > maxmarks:
            maxmarks = currmarks
            name = classmarks.getname(i)
        else:
            pass
    print(name + " scored maximum marks in MBSE")


def minimum_marksin_mbse(classmarks):
    """
    it prints the name of the person who scored the maximum marks in MBSE module
    :param classmarks: object of ClassData
    :return: Void
    """
    ps_numbers = classmarks.listps()
    name = None
    minmarks = 100
    for i in ps_numbers:
        currmarks = classmarks.getmarks_byps_mbse(i)
        if currmarks < minmarks:
            minmarks = currmarks
            name = classmarks.getname(i)
        else:
            pass
    print(name + " scored minimum marks in MBSE")


def minimum_marksin_python(classmarks):
    """
    it prints the name of the person who scored the maximum marks in Python module
    :param classmarks: object of ClassData
    :return: Void
    """
    ps_numbers = classmarks.listps()
    name = None
    minmarks = 100
    for i in ps_numbers:
        currmarks = classmarks.getmarks_byps_python(i)
        if currmarks < minmarks:
            minmarks = currmarks
            name = classmarks.getname(i)
        else:
            pass
    print(name + " scored minimum marks in Python")


def minimum_marksin_sdlc(classmarks):
    """
    it prints the name of the person who scored the maximum marks in SDLC module
    :param classmarks: object of ClassData
    :return: Void
    """
    ps_numbers = classmarks.listps()
    name = None
    minmarks = 100
    for i in ps_numbers:
        currmarks = classmarks.getmarks_byps_sdlc(i)
        if currmarks < minmarks:
            minmarks = currmarks
            name = classmarks.getname(i)
        else:
            pass
    print(name + " scored minimum marks in SDLC")


def average_marks_sdlc(classmarks):
    mylst = classmarks.getmarks_sdlc()

    avgsdlc = sum(mylst) / len(mylst)
    print(avgsdlc, " is the average marks of SLDC")
    return avgsdlc


def average_marks_python(classmarks):
    mylst = classmarks.getmarks_python()

    avgpython = sum(mylst) / len(mylst)
    print(avgpython, " is the average marks of python")
    return avgpython


def average_marks_mbse(classmarks):
    mylst = classmarks.getmarks_mbse()
    avgmbse = sum(mylst) / len(mylst)
    print(avgmbse, " is the average marks of MBSE")
    return avgmbse


def average_marks_three_subs(classmarks):
    psnum = int(input("\033[1m" + "Enter PS No. to know the result: " + "\033[0m"))
    mylist = classmarks.listps()
    if psnum in mylist:
        mylst = classmarks.getmarksbyps(psnum)
        avgsdlc = sum(mylst) / len(mylst)
        avgsdlc = "{:.2f}".format(avgsdlc)
        print(str(avgsdlc), " is the average marks of student in three subjects.")
        sdlc_marks = classmarks.getmarks_byps_sdlc(psnum)
        python_marks = classmarks.getmarks_byps_python(psnum)
        mbse_marks = classmarks.getmarks_byps_mbse(psnum)
        name = classmarks.getname(psnum)
        if sdlc_marks > 60:
            print(name, " is meeting expectations in SDLC")
        else:
            print(name, " is not meeting expectations in SDLC")

        if python_marks > 60:
            print(name, " is meeting expectations in Python")
        else:
            print(name, " is not meeting expectations in Python")

        if mbse_marks > 60:
            print(name, " is meeting expectations in MBSE")
        else:
            print(name, " is not meeting expectations in MBSE")
    else:
        print("\033[91m"+"Please Enter the Valid number"+"\033[0m")
    print("\n"+"\033[92m"+"Press 1 return to main menu")
    print("Press 2 to exit"+"\033[0m")
    cho = int(input())
    if cho == 1:
        main()
    else:
        pass


def sdlc_graph():
    xls = pd.ExcelFile('Mark_Sheet.xlsx')
    df1 = pd.read_excel(xls, 'Applied_SDLC')
    values = df1[['Emp Name', 'Marks']]

    values.plot.bar(x='Emp Name', y='Marks', rot=0)
    plt.title("SDLC marks")
    plt.xticks(rotation=90)
    plt.savefig('Graph_analysis/sdlc.png')
    plt.show()


def python_graph():
    xls = pd.ExcelFile('Mark_Sheet.xlsx')
    df2 = pd.read_excel(xls, 'Adv_Python')
    values = df2[['Emp Name', 'Marks']]
    values.plot.bar(x='Emp Name', y='Marks', rot=0)
    plt.title("Python marks")
    plt.xticks(rotation=90)
    plt.savefig('Graph_analysis/python.png')
    plt.show()


def mbse_graph():
    xls = pd.ExcelFile('Mark_Sheet.xlsx')
    df3 = pd.read_excel(xls, 'MBSE')
    values = df3[['Emp Name', 'Marks']]
    values.plot.bar(x='Emp Name', y='Marks', rot=0)
    plt.title("MBSE marks")
    plt.xticks(rotation=90)
    plt.savefig('Graph_analysis/MBSE.png')
    plt.show()


def average_graph():
    xls = pd.ExcelFile('Mark_Sheet.xlsx')
    df4 = pd.read_excel(xls, 'Average')
    values = df4[['Emp Name', 'AVERAGE']]
    values.plot.bar(x='Emp Name', y='AVERAGE', rot=0)
    plt.xticks(rotation=90)
    plt.savefig('Graph_analysis/Average.png')
    plt.show()


def bell_curve():
    xls = pd.ExcelFile('Mark_Sheet.xlsx')
    df5 = pd.read_excel(xls, 'Average')
    x_min = min(df5['AVERAGE'])
    x_max = max(df5['AVERAGE'])

    mean = statistics.mean(df5['AVERAGE'])
    std = statistics.stdev(df5['AVERAGE'])

    x = np.linspace(x_min, x_max, 100)

    y = scipy.stats.norm.pdf(x, mean, std)

    plt.plot(x, y, color='coral')
    plt.xlabel('Average of marks')
    plt.ylabel('Probability Density Function')
    plt.grid()

    plt.xlim(x_min, x_max)
    plt.show()
    plt.savefig('Graph_analysis/bell_curve.png')


if __name__ == "__main__":
    main()
