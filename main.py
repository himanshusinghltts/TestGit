from abc import ABC, abstractmethod


class SayHello(ABC):
    """
    Abstract class """

    @abstractmethod
    def name(self, myname):
        """
        Say Myname
        :param myname:
        :return: None
        """
        print("Hi "+myname)


class SayHelloAgain(SayHello):
    """
    Normal class
    """
    def name(self, myname):
        """
        say my name
        :param myname:
        :return: None
        """
        print("Hi "+myname)


def main():
    """
    Main class
    :return: None
    """
    #absclas = SayHello()
    secclass = SayHelloAgain()
    print("Hello I am Himanshu,")
    input()
    print("This is just a test file for the git")
    input()
    print("I hope everything will go fine")
    input()
    print("Yes it is working fine.")
    name = input("Enter your name: ")
    #absclas.name(name)
    secclass.name(name)


if __name__ == "__main__":
    main()
