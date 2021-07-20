"""
Class that imports the XLS file and have methods to access that data
"""
import pandas as pd


class ClassData:
    """
    The class stores the marks of specific list of the candidates
    """
    def __init__(self, filepath):
        xlfilethree = pd.read_excel(io=filepath, sheet_name="MBSE")
        xlfiletwo = pd.read_excel(io=filepath, sheet_name="Adv_Python")
        xlfileone = pd.read_excel(io=filepath, sheet_name="Applied_SDLC")
        self.newfile = pd.merge(xlfileone, xlfiletwo, on=("Emp PS #", "Sl#", "Emp Name"))
        self.newfile = pd.merge(self.newfile, xlfilethree, on=("Emp PS #", "Sl#", "Emp Name"))
        del self.newfile["Sl#"]

    def printdata(self):
        """
        print the xls file
        :return: Void
        """
        print(self.newfile)

    def listps(self):
        """
        :return: List of all PS numbers
        """
        mylist = self.newfile["Emp PS #"].values.tolist()
        return mylist

    def getmarks_sdlc(self):
        """
        :return: List of marks scored in SDLC of all candidates
        """
        mylist = self.newfile["Marks_x"]
        mylist = mylist.values.tolist()
        return mylist

    def getmarks_python(self):
        """
        :return: List of marks scored in Python of all candidates
        """
        mylist = self.newfile["Marks_y"]
        mylist = mylist.values.tolist()
        return mylist

    def getmarks_mbse(self):
        """
        :return: List of marks scored in MBSE of all candidates
        """
        mylist = self.newfile["Marks"]
        mylist = mylist.values.tolist()
        return mylist

    def position(self, psnumber):
        """
        :param psnumber: PS number of a specific candidate
        :return: position of the candidate's name in the list
        """
        mylist = self.listps()
        position = None
        # for i in range(0, len(mylist)):
        for i, content in enumerate(mylist):
            if content == psnumber:
                position = i
                break
        return position

    def getname(self, psnumber):
        """
        :param psnumber: PS number of a specific candidate
        :return: Name of the candidate
        """
        position = self.position(psnumber)
        getposition = self.newfile.iloc[position]
        getposition = getposition["Emp Name"]
        return getposition

    def getmarksbyps(self, psnumber):
        """
        :param psnumber: The PS number of the person
        :return: The list of the marks of specific person in all subjects
        """
        position = self.position(psnumber)
        getpos = self.newfile.iloc[position]
        mylist = [getpos["Marks_x"], getpos["Marks_y"], getpos["Marks"]]
        return mylist

    def getmarks_byps_sdlc(self, psnumber):
        """
        :param psnumber: The PS number of the person
        :return: The marks of the person in SDLC module
        """
        marks = self.getmarksbyps(psnumber)
        return marks[0]

    def getmarks_byps_python(self, psnumber):
        """
        :param psnumber: The PS number of the person
        :return: The marks of the person in Python module
        """
        marks = self.getmarksbyps(psnumber)
        return marks[1]

    def getmarks_byps_mbse(self, psnumber):
        """
        :param psnumber: The PS number of the person
        :return:  The marks of the person in SDLC module
        """
        marks = self.getmarksbyps(psnumber)
        return marks[2]
