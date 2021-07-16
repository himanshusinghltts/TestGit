import os
from importxls import ClassData


def main():
    """
    The main function
    :return: void
    """
    filepath = os.getcwd()
    filepath = filepath+"/Mark_Sheet.xlsx"
    ClassMarks = ClassData(filepath)
    MaximumMarksinSDLC(ClassMarks)
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
    print(name+" scored maximum marks in SDLC")


if __name__ == "__main__":
    main()
