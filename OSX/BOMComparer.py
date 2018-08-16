"""This is the source code BOM Comparer Tools for OSX OS"""
import sys
import random
import time
import os.path
import numpy as np
import pandas as pd
import PyQt5.sip
from PyQt5.QtCore import QSize, Qt
from PyQt5.QtWidgets import (QApplication, QPushButton, QFileDialog, QAction, \
                             QMainWindow, QMessageBox, QDesktopWidget, QStyle, \
                             QMenu,)
from PyQt5.QtGui import QIcon, QImage, QPalette, QBrush


class Record(object):
    def __init__(self, data, idx_):
        """constructor for Record class.
        Inputs:
            data: related info for Record object, in the form of tuple,
                  data = (lvl, itm, des, qty, ref, seq).
            idx_: the index of this record.
        """
        super().__init__()
        assert(len(data)==7), "Illegal Constructor Inputs!"
        self.idx = idx_
        self.lvl = data[0]
        self.itm = data[1]
        self.des = data[2]
        self.qty = data[3]
        self.ref = data[4]
        self.seq = data[5]
        self.nme = data[6]
        self.children = None
        self.match = None


    def find_children(self, data):
        """find index for Record's children.
        Input:
            data: list of lvl.
        """
        i = 1
        while (self.idx+i<len(data)) and (data[self.idx+i] > self.lvl):
            i = i+1
        if i>1:
            temp = data[self.idx+1:self.idx+i]
            self.children = \
            [(x+self.idx+1)for x,val in enumerate(temp) if val==(self.lvl+1)]


def get_index(f, simple):
    """get the index of starting rows.
    Inputs:
        f: pandas Series object.
        simple: flag for simple BOMs. True means processing simple BOMs.
    Outputs:
        idx: start(, end) index of BOM.
    """
    # get index of row "BOM"
    line = f[f.columns[0]]
    end = 0

    if simple:
        start = -1
    else:
        start = 0
        while line[start]!="BOM":
            start = start+1

    # get index of row "Manufacturers"
    if "Manufacturers" in list(line):
        end = start+1
        while line[end]!="Manufacturers":
            end = end+1

    if end is not 0:
        return start, end
    return (start,)


def get_info(f, idx, simple):
    """get useful info.
    Inputs:
        f: pandas Series object.
        idx: start(, end) index of BOM.
        simple: flag of BOM types.
    Outputs:
        lvl: numpy array of Level.
        itm: numpy array of Item Number.
        des: numpy array of Item Description.
        qty: numpy array of Qty.
        ref: numpy array of Ref Des.
        seq: numpy array of Item Sequence.
        nme: numpy array of Item Series.
    """
    itm = None
    des = None
    qty = None
    ref = None
    seq = None
    lvl = None
    idx1 = idx[0]

    if len(idx) is 2:
        idx2 = idx[1]
    else:
        idx2 = 0

    # catch useful data, start from row "BOM",
    # (if necessary,) end at row "Manufacturers"
    # simple BOM fits in a differnet structure
    if simple:
        for j in range(0, len(f.columns)):
            if (lvl is None)&(f.columns[j] == "Level"):
                lvl = np.asarray(f[f.columns[j]][(idx1+2):(idx2-2)])
                continue
            if (itm is None)&(f.columns[j] == "Number"):
                itm = np.asarray(f[f.columns[j]][(idx1+2):(idx2-2)])
                continue
            if (des is None)&(f.columns[j] == "Description"):
                des = np.asarray(f[f.columns[j]][(idx1+2):(idx2-2)])
                continue
            if (qty is None)&(f.columns[j] == "BOM.Qty"):
                qty = np.asarray(f[f.columns[j]][(idx1+2):(idx2-2)])
                continue
            if (ref is None)&(f.columns[j] == "BOM.Ref Des"):
                ref = np.asarray(f[f.columns[j]][(idx1+2):(idx2-2)])
                continue
            if (seq is None)&(f.columns[j] == "BOM.Item Seq"):
                seq = np.asarray(f[f.columns[j]][(idx1+2):(idx2-2)])
            if (lvl is not None)&(itm is not None)&(des is not None)& \
                    (qty is not None)&(ref is not None)&(seq is not None):
                break
    else:
        for j in range(0, len(f.columns)):
            if (lvl is None)&(f[f.columns[j]][idx1+1] == "Level"):
                lvl = np.asarray(f[f.columns[j]][(idx1+2):(idx2-2)])
                continue
            if (itm is None)&(f[f.columns[j]][idx1+1] == "Item Number"):
                itm = np.asarray(f[f.columns[j]][(idx1+2):(idx2-2)])
                continue
            if (des is None)&(f[f.columns[j]][idx1+1] == "Item Description"):
                des = np.asarray(f[f.columns[j]][(idx1+2):(idx2-2)])
                continue
            if (qty is None)&(f[f.columns[j]][idx1+1] == "Qty"):
                qty = np.asarray(f[f.columns[j]][(idx1+2):(idx2-2)])
                continue
            if (ref is None)&(f[f.columns[j]][idx1+1] == "Ref Des"):
                ref = np.asarray(f[f.columns[j]][(idx1+2):(idx2-2)])
                continue
            if (seq is None)&(f[f.columns[j]][idx1+1] == "Item Seq"):
                seq = np.asarray(f[f.columns[j]][(idx1+2):(idx2-2)])
            if (lvl is not None)&(itm is not None)&(des is not None)& \
                    (qty is not None)&(ref is not None)&(seq is not None):
                break

    # omit blank rows, and set data type
    valid = []
    for i in range(0, len(itm)):
        if type(itm[i]) is str:
            valid.append(i)
    lvl = lvl[valid].astype("int")
    itm = itm[valid]
    des = des[valid]
    qty = qty[valid].astype("float")
    ref = ref[valid]
    seq = seq[valid].astype("int")

    for i,unit in enumerate(ref):
        if unit is np.nan:
            ref[i] = ""

    # split version number from item number
    nme = []
    for unit in itm:
        temp = unit.split("-")
        if len(temp)<3:
            nme.append(unit)
        else:
            nme.append(temp[0]+"-"+temp[1])
    nme = np.asarray(nme)

    return lvl, itm, des, qty, ref, seq, nme


def get_ancester(data1, data2):
    """
    Inputs:
        data1: numpy array of Record object.
        data2: numpy array of Record object.
    Outputs:
        ancester1: list of level 1 Record in data1.
        ancester2: list of level 1 Record in data2.
    """
    ancester1 = []
    ancester2 = []

    for i, unit in enumerate(data1):
        if unit.lvl == 1:
            ancester1.append(i)

    for i, unit in enumerate(data2):
        if unit.lvl == 1:
            ancester2.append(i)

    return ancester1, ancester2


def match(record1, record2):
    """find the item matched in BOM2 for item in BOM1.
    Inputs:
        record1: numpy array of Record in BOM 1.
        record2: numpy array of Record in BOM 2.
    Matching Rules:
        0. match code: 1, this Record belongs to record1, no Record matched it
                          in record2.
                       -1, this Record belongs to record2, no Record matched it
                          in record1.
                      integer, this Record belongs to record1, the integer
                               represents the index of matched Record in
                               record2.
        1. match the Item accroding to Item Series.
        2. if duplicates exist, match the Item according to Item Number.
    """

    # special cases
    if record1 is None:
        for unit in record2:
            unit.match = 1
        return
    if record2 is None:
        for unit in record1:
            unit.match = -1
        return

    # create list for item series
    pool1 = []
    pool2 = []
    origin_pool1 = []
    origin_pool2 = []
    for unit in record1:
        origin_pool1.append(unit.itm)
        pool1.append(unit.nme)
    pool1 = np.asarray(pool1)
    origin_pool1 = np.asarray(origin_pool1)
    for unit in record2:
        origin_pool2.append(unit.itm)
        pool2.append(unit.nme)
    pool2 = np.asarray(pool2)
    origin_pool2 = np.asarray(origin_pool2)

    # find matched items
    for unit in record1:
        if unit.nme in pool2:
            unit.match = np.where((pool2==unit.nme).astype("int")==1)
        else:
            unit.match = -1

    for unit in record2:
        if unit.nme not in pool1:
            unit.match = 1

    # find duplicates
    uni,i = np.unique(pool1, return_inverse = True)
    dup1 = uni[np.bincount(i)>1]

    uni,i = np.unique(pool2, return_inverse = True)
    dup2 = uni[np.bincount(i)>1]

    # rematch for duplicates
    for unit in record1:
        if (unit.nme in dup1) or (unit.nme in dup2):
            if unit.itm in origin_pool2:
                unit.match = np.where((origin_pool2==unit.itm).astype("int")==1)
            else:
                unit.match = -1

    for unit in record2:
        if (unit.nme in dup2) or (unit.nme in dup1):
            if unit.itm in origin_pool1:
                unit.match = None
            else:
                unit.match = 1


def get_compare(record1, record2, data1, data2, input):
    """comparison recursively.
    Inputs:
        record1: Record family1.
        record2: Record family2.
        data1: list of all data in BOM1.
        data2: list of all data in BOM2.
        input: dict to store the comparison result. format:
               {lvl,itm1,itm2,qty1,qty2,des1,des2,ref1,ref2,seq1,seq2}
    Output:
        output: dict to store the comparison result. same format with input.
    """
    match(record1, record2)
    (lvl,itm1,itm2,qty1,qty2,des1,des2,ref1,ref2,seq1,seq2) = input

    # special case
    if record2 is not None:
        for unit in record2:
            if unit.match is 1:
                lvl = np.append(lvl, unit.lvl)
                itm2 = np.append(itm2, unit.itm)
                itm1 = np.append(itm1, None)
                qty2 = np.append(qty2, unit.qty)
                qty1 = np.append(qty1, 0)
                des2 = np.append(des2, unit.des)
                des1 = np.append(des1, None)
                ref2 = np.append(ref2, unit.ref)
                ref1 = np.append(ref1, None)
                seq2 = np.append(seq2, unit.seq)
                seq1 = np.append(seq1, None)
                if unit.children is None:
                    continue
                temp = (lvl,itm1,itm2,qty1,qty2,des1,des2,ref1,ref2,seq1,seq2)
                (lvl,itm1,itm2,qty1,qty2,des1,des2,ref1,ref2,seq1,seq2) = \
                    get_compare(None, data2[unit.children], data1, data2, temp)

    if record1 is not None:
        for unit in record1:
            if unit.match is -1:
                lvl = np.append(lvl, unit.lvl)
                itm1 = np.append(itm1, unit.itm)
                itm2 = np.append(itm2, None)
                qty1 = np.append(qty1, unit.qty)
                qty2 = np.append(qty2, 0)
                des1 = np.append(des1, unit.des)
                des2 = np.append(des2, None)
                ref1 = np.append(ref1, unit.ref)
                ref2 = np.append(ref2, None)
                seq1 = np.append(seq1, unit.seq)
                seq2 = np.append(seq2, None)
                if unit.children is None:
                     continue
                temp = (lvl,itm1,itm2,qty1,qty2,des1,des2,ref1,ref2,seq1,seq2)
                (lvl,itm1,itm2,qty1,qty2,des1,des2,ref1,ref2,seq1,seq2) = \
                    get_compare(data1[unit.children], None, data1, data2, temp)

            elif unit.match is not None:
                if (unit.itm != record2[unit.match][0].itm) or \
                    (unit.qty != record2[unit.match][0].qty) or \
                    (unit.ref != record2[unit.match][0].ref):
                    lvl = np.append(lvl, unit.lvl)
                    itm1 = np.append(itm1, unit.itm)
                    itm2 = np.append(itm2, record2[unit.match][0].itm)
                    qty1 = np.append(qty1, unit.qty)
                    qty2 = np.append(qty2, record2[unit.match][0].qty)
                    des1 = np.append(des1, unit.des)
                    des2 = np.append(des2, record2[unit.match][0].des)
                    ref1 = np.append(ref1, unit.ref)
                    ref2 = np.append(ref2, record2[unit.match][0].ref)
                    seq1 = np.append(seq1, unit.seq)
                    seq2 = np.append(seq2, record2[unit.match][0].seq)
                    temp=(lvl,itm1,itm2,qty1,qty2,des1,des2,ref1,ref2,seq1,seq2)

                    if (unit.children is None) and \
                            (record2[unit.match][0].children is None):
                        continue

                    if (unit.children is None) and \
                            (record2[unit.match][0].children is not None):
                        (lvl,itm1,itm2,qty1,qty2,des1,des2,ref1,ref2,seq1, \
                        seq2)=get_compare(None, \
                        data2[record2[unit.match][0].children], \
                        data1, data2, temp)
                        continue

                    if (unit.children is not None) and \
                            (record2[unit.match][0].children is None):
                        (lvl,itm1,itm2,qty1,qty2,des1,des2,ref1,ref2,seq1, \
                        seq2)=get_compare(data1[unit.children], None, \
                        data1, data2, temp)
                        continue

                    if (unit.children is not None) and \
                            (record2[unit.match][0].children is not None):
                        (lvl,itm1,itm2,qty1,qty2,des1,des2,ref1,ref2,seq1, \
                        seq2)=get_compare(data1[unit.children], \
                        data2[record2[unit.match][0].children],data1,data2,temp)
                        continue

    return (lvl,itm1,itm2,qty1,qty2,des1,des2,ref1,ref2,seq1,seq2)


def main(path1, path2, simple, file_name=None):
    """generate comparison results.
    Inputs:
        path1: string, path of old BOM.
        path2: string, path of new BOM.
        simple: flag for BOM format. True for simple, False for standard.
        file_name: string, name of the file to be generated.
    """
    # set output file name
    if file_name is None:
        file_name = os.path.splitext(os.path.basename(path1))[0] + "_" + \
                    os.path.splitext(os.path.basename(path2))[0] + "_cmp.xlsx"

    # import excel file
    f1 = pd.read_excel(path1)
    f2 = pd.read_excel(path2)

    # capture BOM content from file
    idx1 = get_index(f1, simple)
    idx2 = get_index(f2, simple)

    # store related content in numpy arrays
    args1 = get_info(f1, idx1, simple)
    args2 = get_info(f2, idx2, simple)

    # construct numpy array of Record object.
    array1 = np.vstack(args1).transpose()
    array2 = np.vstack(args2).transpose()
    data1 = np.asarray([Record(unit, i) for i, unit in enumerate(array1)])
    data2 = np.asarray([Record(unit, i) for i, unit in enumerate(array2)])

    # find children for all Records
    for item in data1:
        item.find_children(args1[0])

    for item in data2:
        item.find_children(args2[0])

    # compare recursively
    ancester1, ancester2 = get_ancester(data1, data2)
    input = (np.array([]), np.array([]), np.array([]), np.array([]), \
             np.array([]), np.array([]), np.array([]), np.array([]), \
             np.array([]), np.array([]), np.array([]))
    (Level,Itm1,Itm2,Qty1,Qty2,Des1,Des2,Ref1,Ref2,Seq1,Seq2) = \
        get_compare(data1[ancester1], data2[ancester2], data1, data2, input)

    # store comparison result in excel file
    data = np.vstack((Level, Itm1, Itm2, Qty1, Qty2, Des1, Des2, Ref1, Ref2, \
                  Seq1, Seq2)).transpose()
    column = ["Level", "Original Item", "Updated Item", \
              "Original Qty.", "Updated Qty.", "Original Item Des.", \
              "Updated Item Des.", "Original Ref. Des.", \
              "Updated Ref. Des.", "Original Seq.", \
              "Updated Seq."]
    dic = pd.DataFrame.from_records(data)
    writer = pd.ExcelWriter(file_name)
    dic.to_excel(writer, sheet_name = "Comparison Report", index=False, \
                        header = column)
    writer.save()


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.file1 = None
        self.file2 = None
        self.file_name = None
        self.simple = True
        self.folder_path = ""
        self.blue_box = "color: white; font: 12pt Arial; \
                          font-weight: bold; background-color: rgb(5,188,228);"
        self.green_box = "color: white; font: 12pt Arial; \
                         font-weight: bold; background-color: rgb(25,140,69);"

        self.init_UI()

    def init_UI(self):

        #self.setFont(QFont("Arial"))
        self.setWindowFlags(Qt.WindowCloseButtonHint)

        # initiate file names
        self.statusBar().showMessage("Thanks for trusting BOM Comparer :)")
        iconLib = self.style()

        # add background image
        backgroundMAC = QImage(resource_path("backgroundOSX.png"))
        backgroundMAC.scaled(QSize(900,150))
        self.paletteMain = QPalette()
        self.paletteMain.setBrush(10, QBrush(backgroundMAC))
        self.setPalette(self.paletteMain)

        # add button "Import BOM1"
        inButton1 = QPushButton("Import Original BOM", self)
        inButton1.clicked.connect(self.showImport1)
        inButton1.setFixedWidth(150)
        inButton1.move(50, 50)
        inButton1.setFlat(True)
        inButton1.setStyleSheet("background-color: rgb(6,182,224); \
                                 color: white; \
                                 font: 12pt Arial; \
                                 font-weight: bold;")
        inButton1.setIcon(iconLib.standardIcon(getattr(QStyle,"SP_FileIcon")))

        # add button "Import BOM2"
        inButton2 = QPushButton("Import Updated BOM", self)
        inButton2.clicked.connect(self.showImport2)
        inButton2.setFixedWidth(150)
        inButton2.move(250, 50)
        inButton2.setFlat(True)
        inButton2.setStyleSheet("background-color: rgb(10,172,206); \
                                 color: white; \
                                 font: 12pt Arial; \
                                 font-weight: bold;")
        inButton2.setIcon(iconLib.standardIcon(getattr(QStyle,"SP_FileIcon")))

        # add button "Compare BOMs"
        cmp = QPushButton("Compare BOMs", self)
        cmp.clicked.connect(self.rUSure)
        cmp.resize(cmp.sizeHint())
        cmp.move(450, 50)
        cmp.setFixedWidth(150)
        cmp.setFlat(True)
        cmp.setStyleSheet("background-color: rgb(18,175,175); \
                                 color: white; \
                                 font: 12pt Arial; \
                                 font-weight: bold;")
        cmp.setIcon(iconLib.standardIcon(getattr(QStyle, \
                            "SP_ArrowForward")))


        # add menu option "Set Folder"
        setF = QAction("&Set Folder", self)
        setF.setStatusTip("Choose Folder to Store the Comparison Result")
        setF.triggered.connect(self.setFolder)
        # add menu option "Set BOM Types"
        setB = QMenu("&Set BOM Types", self)
        setB.setStatusTip("Choose BOM Types")
        setBSimple = QAction("&Simple BOMs", self)
        setBStandard = QAction("&Standard BOMs", self)
        setBSimple.triggered.connect(self.setBOMSimple)
        setBSimple.setStatusTip("To Process Simple BOMs")
        setBStandard.triggered.connect(self.setBOMStandard)
        setBStandard.setStatusTip("To Process Standard BOMs")
        setB.addAction(setBSimple)
        setB.addAction(setBStandard)
        # add Setting menu
        menubar = self.menuBar()
        settingMenu = menubar.addMenu("&Setting ")
        settingMenu.addAction(setF)
        settingMenu.addMenu(setB)

        # add menu option "Read Me"
        readMe = QAction("&Read Me", self)
        readMe.setStatusTip("See User Manual")
        readMe.triggered.connect(self.showHelp)
        # add menu option "Contact Me"
        contactMe = QAction("&Contact Me", self)
        contactMe.setStatusTip("Feel Free to Contact Developer Chen")
        contactMe.triggered.connect(self.showEmail)
        # add Help menu
        helpMenu = menubar.addMenu('&Help ')
        helpMenu.addAction(readMe)
        helpMenu.addAction(contactMe)

        """
        authMenu = QAction("&DECLARATION", self)
        authMenu.triggered.connect(self.showAuth)
        menubar.addAction(authMenu)
        """

        self.setStyleSheet("""QMenu{
                                    background-color: rgb(25,140,69);
                                    color: white;
                                    font: 12pt Arial;
                                    font-weight: bold;
                                    }
                              QMenu::item::selected{
                                    background-color:rgb(5,188,228);
                                    }
                              QMenuBar{
                                    background-color: rgb(25,140,69);
                                    color: white;
                                    font: 12pt Arial;
                                    font-weight: bold;
                                    }
                              QMenuBar::item::selected{
                                    background-color: rgb(5,188,228);
                              }""")

        # set window geometry
        self.resize(650, 150)
        self.setWindowTitle("BOM comparer")
        self.setWindowIcon(QIcon(iconLib.standardIcon(getattr(QStyle, \
                        "SP_ComputerIcon"))))
        self.center()
        self.show()

    # func to conter the window
    def center(self):
        mvCenter = self.frameGeometry()
        centerPlace = QDesktopWidget().availableGeometry().center()
        mvCenter.moveCenter(centerPlace)
        self.move(mvCenter.topLeft())

    # update statusbar while importing original BOM
    def showImport1(self):
        fname = QFileDialog.getOpenFileName(self, "Open", "\home")
        self.file1 = str(fname[0])
        self.statusBar().showMessage("Original BOM Imported: "+ \
                            os.path.basename(self.file1))

    # update statusbar while importing updated BOM
    def showImport2(self):
        fname = QFileDialog.getOpenFileName(self, "Open", "\home")
        self.file2 = str(fname[0])
        self.statusBar().showMessage("Updated BOM Imported: "+ \
                            os.path.basename(self.file2))


    def onClickCompare(self):
        string1 = os.path.splitext(os.path.basename(self.file1))[0]
        string2 = os.path.splitext(os.path.basename(self.file2))[0]
        timetag = time.strftime("%Y:%m:%d@%H:%M:%S", time.localtime())
        base = string1.replace(" ","_") + "_VS_" + string2.replace(" ","_") + \
                        "_" + timetag + ".xlsx"
        if self.folder_path is "":
            self.folder_path = os.path.join(os.path.expanduser("~"),"Desktop")
        self.file_name = os.path.join(self.folder_path, base)
        self.statusBar().showMessage("Processing...")
        main(self.file1, self.file2, self.simple, self.file_name)
        self.statusBar().showMessage("File Generated: "+ \
                                os.path.join(self.folder_path, self.file_name))


    def rUSure(self):
        if (self.file1 is not None) & (self.file2 is not None):
            if (os.path.isfile(self.file1)) & (os.path.isfile(self.file2)):
                string1 = os.path.splitext(os.path.basename(self.file1))[0]
                string2 = os.path.splitext(os.path.basename(self.file2))[0]
                type = "Simple" if self.simple else "Standard"
                startMessage = "Original BOM: " + string1 + "\n" + \
                          "Updated BOM: " + string2 + "\n" + \
                          "BOM Type: "+type + "\n" + \
                          "Generate comparison result?"
                start = QMessageBox()
                start.setWindowIcon(self.style().standardIcon(getattr(QStyle, \
                "SP_ComputerIcon")))
                start.setWindowTitle("Ready?")
                start.setText(startMessage)
                if random.randint(1,2)==1:
                    start.setStyleSheet(self.green_box)
                else:
                    start.setStyleSheet(self.blue_box)
                start.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
                start.buttonClicked.connect(self.startCheck)
                start.exec_()

            else:
                message1 = "Fail to import valid BOMs!\n" + \
                          "Please try again!"
                choice1 = QMessageBox()
                choice1.setWindowIcon(self.style().standardIcon(getattr(QStyle,\
                                    "SP_ComputerIcon")))
                choice1.setWindowTitle("Attention")
                if random.randint(1,2)==1:
                    choice1.setStyleSheet(self.green_box)
                else:
                    choice1.setStyleSheet(self.blue_box)
                choice1.setText(message1)
                choice1.exec_()
        else:
            message2 = "Please import BOMs first!\n" + \
                      "Click Help->Read Me to see User Manual."
            choice2 = QMessageBox()
            choice2.setWindowIcon(self.style().standardIcon(getattr(QStyle, \
                                "SP_ComputerIcon")))
            choice2.setWindowTitle("Attention")
            if random.randint(1,2)==1:
                choice2.setStyleSheet(self.green_box)
            else:
                choice2.setStyleSheet(self.blue_box)
            choice2.setText(message2)
            choice2.exec_()

    def startCheck(self, button):
        if button.text() == "OK":
            self.onClickCompare()


    # display help manual
    def showHelp(self):
        help = QMessageBox()
        help.setWindowIcon(self.style().standardIcon(getattr(QStyle, \
                        "SP_ComputerIcon")))
        if random.randint(1,2)==1:
            help.setStyleSheet(self.green_box)
        else:
            help.setStyleSheet(self.blue_box)
        help.setText("0. NO COMMERCIAL USE! NO ACCURACY GUARANTEE!\n"+ \
                    "1. Click 'Import Original BOM' to import old BOM.\n" + \
                    "2. Click 'Import Updated BOM' to import new BOM.\n"+ \
                    "3. Select 'Setting->Set Folder' to choose folder for "+ \
                    "generated files, default is Desktop\n"+ \
                    "4. Select 'Setting->Set BOM Types' to set BOM "+ \
                    "format, default is Simple.\n" + \
                    "5. Click 'Compare BOMs' to get comparison result.\n" +\
                    "*Restarting App restores all setting to dafault.")
        help.setWindowTitle("User Manual")
        help.exec_()

    # display contact email
    def showEmail(self):
        contact = QMessageBox()
        contact.setWindowIcon(self.style().standardIcon(getattr(QStyle, \
                        "SP_ComputerIcon")))
        contact.setText("Email Address:")
        if random.randint(1,2)==1:
            contact.setStyleSheet(self.green_box)
        else:
            contact.setStyleSheet(self.blue_box)
        contact.setText("chenw.pop@gmail.com")
        contact.setWindowTitle("Contact Me")
        contact.exec_()

    """
    # show Declaration
    def showAuth(self):
        auth = QMessageBox()
        auth.setWindowIcon(self.style().standardIcon(getattr(QStyle, \
                        "SP_ComputerIcon")))
        auth.setText("1. NO COMMERCIAL USE\n2. NO ACCURACY GUARANTEE")
        if random.randint(1,2)==1:
            auth.setStyleSheet(self.green_box)
        else:
            auth.setStyleSheet(self.blue_box)
        auth.setWindowTitle("DECLARATION")
        auth.exec_()
    """

    # set folder path to store comparison result.
    def setFolder(self):
        path = QFileDialog.getExistingDirectory(self,"Select Directory", \
                                                os.path.expanduser("~"))
        self.folder_path = str(path)
        self.statusBar().showMessage("Set Folder as: "+path)

    # set BOM Types
    def setBOMSimple(self):
        pass

    def setBOMStandard(self):
        self.simple = False


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Window()
    sys.exit(app.exec_())
