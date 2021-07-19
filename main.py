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
    ClassMarks = ClassData(filepath)
    print("\n")
    print("\033[1m"+"\033[4m"+'\033[36m'+"Automation - Analysis and Visualisation of Result"+'\033[0m')
    print("\n"+"\033[92m"+"Press 1 to generate the whole class Report")
    print("Press 2 for analysing the marks of specific person"+"\033[0m")
    choise = input()
    if(choise==str(1)):
        callFunction(ClassMarks)
    elif(choise== str(2) ):
        averageMarksThreeSubs(ClassMarks)
        pass
    else:
        print("\033[91m"+"Please Enter the Valid number"+"\033[0m")
        main()

def callFunction(ClassMarks):
    """
    The main function
    :return: void
    """
    print("BEST PERFORMERS")
    MaximumMarksinSDLC(ClassMarks)
    MaximumMarksinPython(ClassMarks)
    MaximumMarksinMBSE(ClassMarks)
    ClassMarks.printdata()
    print(ClassMarks.listps())
    print("AVERAGE MARKS")
    averagemarksSDLC(ClassMarks)
    averagemarkspython(ClassMarks)
    averagemarksMBSE(ClassMarks)
    print("POOR PERFORMERS")
    MinimumMarksinMBSE(ClassMarks)
    MinimumMarksinPython(ClassMarks)
    MinimumMarksinSDLC(ClassMarks)
    MeetingExpectationsinSDLC(ClassMarks)
    #averageSubjectGraph(ClassMarks)
    sdlcGraph(ClassMarks)
    pythonGraph(ClassMarks)
    mbseGraph(ClassMarks)
    averageGraph(ClassMarks)
    bellCurve(ClassMarks)
    print("\n"+"\033[92m"+"Press 1 return to main menu")
    print("Press 2 to exit"+"\033[0m")
    cho = int(input())
    if cho == 1:
        main()
    else:
        pass


def MaximumMarksinSDLC(ClassMarks):
    """
    it prints the name of the person who scored the maximum marks in SDLC module
    :param ClassMarks: object of ClassData
    :return: Void
    """
    PSnumbers = ClassMarks.listps()
    name = None
    maxmarks = 0
    for i in PSnumbers:
        currMarks = ClassMarks.getmarksbypsSDLC(i)
        if currMarks > maxmarks:
            maxmarks = currMarks
            name = ClassMarks.getname(i)
        else:
            pass
    print(name + " scored maximum marks in SDLC")


def MeetingExpectationsinSDLC(ClassMarks):
    """
    it prints the name of the person meeting expectations or not in SDLC
    :param ClassMarks: object of ClassData
    :return: Void
    """
    PSnumbers = ClassMarks.listps()
    name = None
    for i in PSnumbers:
        currMarks = ClassMarks.getmarksbypsSDLC(i)
        name = ClassMarks.getname(i)
        if currMarks > 60:
            print(name, " is meeting expectations in SDLC")
        else:
            print(name, " is not meeting expectations in SDLC")


def MaximumMarksinPython(ClassMarks):
    """
    it prints the name of the person who scored the maximum marks in SDLC module
    :param ClassMarks: object of ClassData
    :return: Void
    """
    PSnumbers = ClassMarks.listps()
    name = None
    maxmarks = 0
    for i in PSnumbers:
        currMarks = ClassMarks.getmarksbypspython(i)
        if currMarks > maxmarks:
            maxmarks = currMarks
            name = ClassMarks.getname(i)
        else:
            pass
    print(name + " scored maximum marks in python")


def MaximumMarksinMBSE(ClassMarks):
    """
    it prints the name of the person who scored the maximum marks in SDLC module
    :param ClassMarks: object of ClassData
    :return: Void
    """
    PSnumbers = ClassMarks.listps()
    name = None
    maxmarks = 0
    for i in PSnumbers:
        currMarks = ClassMarks.getmarksbypsMBSE(i)
        if currMarks > maxmarks:
            maxmarks = currMarks
            name = ClassMarks.getname(i)
        else:
            pass
    print(name + " scored maximum marks in MBSE")


def MinimumMarksinMBSE(ClassMarks):
    """
    it prints the name of the person who scored the maximum marks in SDLC module
    :param ClassMarks: object of ClassData
    :return: Void
    """
    PSnumbers = ClassMarks.listps()
    name = None
    minmarks = 100
    for i in PSnumbers:
        currMarks = ClassMarks.getmarksbypsMBSE(i)
        if currMarks < minmarks:
            minmarks = currMarks
            name = ClassMarks.getname(i)
        else:
            pass
    print(name + " scored minimum marks in MBSE")


def MinimumMarksinPython(ClassMarks):
    """
    it prints the name of the person who scored the maximum marks in SDLC module
    :param ClassMarks: object of ClassData
    :return: Void
    """
    PSnumbers = ClassMarks.listps()
    name = None
    minmarks = 100
    for i in PSnumbers:
        currMarks = ClassMarks.getmarksbypspython(i)
        if currMarks < minmarks:
            minmarks = currMarks
            name = ClassMarks.getname(i)
        else:
            pass
    print(name + " scored minimum marks in Python")


def MinimumMarksinSDLC(ClassMarks):
    """
    it prints the name of the person who scored the maximum marks in SDLC module
    :param ClassMarks: object of ClassData
    :return: Void
    """
    PSnumbers = ClassMarks.listps()
    name = None
    minmarks = 100
    for i in PSnumbers:
        currMarks = ClassMarks.getmarksbypsSDLC(i)
        if currMarks < minmarks:
            minmarks = currMarks
            name = ClassMarks.getname(i)
        else:
            pass
    print(name + " scored minimum marks in SDLC")


def averagemarksSDLC(ClassMarks):
    mylst = ClassMarks.getmarksSDLC()

    avgsdlc = sum(mylst) / len(mylst)
    print(avgsdlc, " is the average marks of SLDC")
    return avgsdlc


def averagemarkspython(ClassMarks):
    mylst = ClassMarks.getmarksPython()

    avgpython = sum(mylst) / len(mylst)
    print(avgpython, " is the average marks of python")
    return avgpython


def averagemarksMBSE(ClassMarks):
    mylst = ClassMarks.getmarksMBSE()

    avgMBSE = sum(mylst) / len(mylst)
    print(avgMBSE, " is the average marks of MBSE")
    return avgMBSE


def averageMarksThreeSubs(ClassMarks):
    psnum=int(input("\033[1m"+"Enter PS No. to know the result: "+"\033[0m"))
    mylist = ClassMarks.listps()
    if psnum in mylist:
        mylst = ClassMarks.getmarksbyps(psnum)
        avgsdlc = sum(mylst) / len(mylst)
        avgsdlc = "{:.2f}".format(avgsdlc)
        print(str(avgsdlc), " is the average marks of student in three subjects.")
        SDLCMarks = ClassMarks.getmarksbypsSDLC(psnum)
        PythonMarks = ClassMarks.getmarksbypspython(psnum)
        MBSEMarks = ClassMarks.getmarksbypsMBSE(psnum)
        name = ClassMarks.getname(psnum)
        if SDLCMarks > 60:
            print(name, " is meeting expectations in SDLC")
        else:
            print(name, " is not meeting expectations in SDLC")

        if PythonMarks > 60:
            print(name, " is meeting expectations in Python")
        else:
            print(name, " is not meeting expectations in Python")

        if MBSEMarks > 60:
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



def sdlcGraph(ClassMarks):
    xls = pd.ExcelFile('Mark_Sheet.xlsx')
    df1 = pd.read_excel(xls, 'Applied_SDLC')
    values = df1[['Emp Name', 'Marks']]

    ax = values.plot.bar(x='Emp Name', y='Marks', rot=0)
    plt.xticks(rotation=90)
    plt.show()




def sdlcGraph(ClassMarks):
    xls = pd.ExcelFile('Mark_Sheet.xlsx')
    df1 =  pd.read_excel(xls,'Applied_SDLC')
    # print(df1.head())
    values = df1[['Emp Name', 'Marks']]
    # print(values)

    ax = values.plot.bar(x='Emp Name', y = 'Marks', rot=0)
    plt.title("SDLC marks")
    plt.xticks(rotation = 90)
    plt.savefig('Graph_analysis/sdlc.png')
    plt.show()


def pythonGraph(ClassMarks):
    xls = pd.ExcelFile('Mark_Sheet.xlsx')
    df2 = pd.read_excel(xls, 'Adv_Python')
    # print(df2.head())
    values = df2[['Emp Name', 'Marks']]
    # print(values)

    ax = values.plot.bar(x='Emp Name', y = 'Marks', rot=0)
    plt.title("Python marks")
    plt.xticks(rotation = 90)
    plt.savefig('Graph_analysis/python.png')
    plt.show()


def mbseGraph(ClassMarks):
    xls = pd.ExcelFile('Mark_Sheet.xlsx')
    df3 =  pd.read_excel(xls,'MBSE')
    # print(df3.head())
    values = df3[['Emp Name', 'Marks']]
    # print(values)
    ax = values.plot.bar(x='Emp Name', y = 'Marks', rot=0)
    plt.title("MBSE marks")
    plt.xticks(rotation = 90)
    plt.savefig('Graph_analysis/MBSE.png')
    plt.show()


def averageGraph(ClassMarks):
    xls = pd.ExcelFile('Mark_Sheet.xlsx')
    df4 =  pd.read_excel(xls,'Average')
    # print(df4.head())
    values = df4[['Emp Name', 'AVERAGE']]
    # print(values)

    ax = values.plot.bar(x='Emp Name', y = 'AVERAGE', rot=0)
    plt.xticks(rotation = 90)
    plt.savefig('Graph_analysis/Average.png')
    plt.show()


def bellCurve(ClassMarks):
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
#    plt.savefig('bell_curve.png')


if (__name__) == "__main__":
    main()