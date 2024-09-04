from imports.stuff import *

# 1120,680 max defult width
class BenefitsWindow(QWidget):
    def __init__(self):
        super().__init__()
    def set_arabic_format(self,cell):
        for paragraph in cell.paragraphs:
            pPr = paragraph._element.get_or_add_pPr()
            bidi = OxmlElement('w:bidi')
            bidi.set(qn('w:val'), '1')
            pPr.append(bidi)
            for run in paragraph.runs:
                run.font.name = 'Arial'
            lang = OxmlElement('w:lang')
            lang.set(qn('w:val'), 'ar-SA')
            pPr.append(lang)

    def adjustTableWidth(self):
        deserve = self.BenefitsTable.columnCount()
        for i in range(self.BenefitsTable.columnCount()):
            if self.showCheckBoxes and i == 0:
                deserve -=1
                self.BenefitsTable.setColumnWidth(i, 40)
                continue
            self.BenefitsTable.setColumnWidth(i, (self.BenefitsTable.width() - 20) // deserve)
    def showBenefits(self):
        try:
            self.destroyFrame(self.mainMenuFrame)
            self.destroyFrame(self.BenefitsFrame)
        except:
            pass
        self.showCheckBoxes = False

        self.BenefitsFrame = QFrame(self.mainFrame)
        
        self.closeButtonBenefetsFrame = QPushButton(self.BenefitsFrame)
        self.closeButtonBenefetsFrame.setStyleSheet("QPushButton {border:2px solid black;font-size:12px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 26px}QPushButton:hover {background-color:#c8c8c8;}")        
        self.closeButtonBenefetsFrame.clicked.connect(lambda x, frame=self.BenefitsFrame:self.destroyFrame(frame))

        self.comboSearchBenefitsBox = QComboBox(self.BenefitsFrame)
        self.comboSearchBenefitsBox.setGeometry(500,30,150,20)
        self.comboSearchBenefitsBox.addItems(["الكل", "رقم الهوية", "رقم الجوال", "الاسم"])
        self.comboSearchBenefitsBox.activated.connect(self.addEntrySearchBenefits)



        self.BenefitsFrame.setStyleSheet("background-color:white")

        self.BenefitsTable = QTableWidget(self.BenefitsFrame)

        self.BenefitsTable.setColumnCount(8)
        self.BenefitsTable.setHorizontalHeaderLabels(["اسم المستفيد","رقم الهوية","رقم الجوال","المدينة","الميلاد","الجنس","العمر","الجنسية"])

        self.contextMenuBenefitsTable = QMenu(self.BenefitsTable)
        self.contextMenuBenefitsTable.setStyleSheet("background-color:grey")
        self.createButtonBenefitsTable()

        self.BenefitsTable.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.BenefitsTable.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.BenefitsTable.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.BenefitsTable.customContextMenuRequested.connect(self.showMenuBenefitsTable)

        #Start minMenu Buttons style
        self.mainMenuFrame = QFrame(self.BenefitsFrame)
        self.mainMenuFrame.setStyleSheet("background-color:white;border:2px solid black")
        
        label = QLabel(self.mainMenuFrame,text="القائمة الرئيسيه")
        label.setStyleSheet("background-color:white;border-bottom:none;font: 14pt 'Arial';")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setGeometry(0,0,181,31)

        addBenefitButton = QPushButton(self.mainMenuFrame,text="اضافة مستفيد")
        addBenefitButton.setGeometry(0,50,181,31)
        addBenefitButton.setStyleSheet("QPushButton {background-color:white;border:2px solid black;font: 14pt 'Arial';} QPushButton:hover {background-color:#c8c8c8;}")
        addBenefitButton.clicked.connect(self.addBenefit)

        

        addExternalButton = QPushButton(self.mainMenuFrame,text="اضافة مستفيدين ملف خارجي")
        addExternalButton.setGeometry(0,90,181,31)
        addExternalButton.setStyleSheet("QPushButton {background-color:white;border:2px solid black;font: 14pt 'Arial';} QPushButton:hover {background-color:#c8c8c8;}")
        addExternalButton.clicked.connect(self.addBenefitsFromExternal)

        exportBenefits = QPushButton(self.mainMenuFrame,text="تصدير معلومات المستفيدين")
        exportBenefits.setGeometry(0,130,181,31)
        exportBenefits.setStyleSheet("QPushButton {background-color:white;border:2px solid black;font: 14pt 'Arial';} QPushButton:hover {background-color:#c8c8c8;}")
        exportBenefits.clicked.connect(self.exportAllBenefits)

        addBenefitsToTripsButton = QPushButton(self.mainMenuFrame,text="اضافة مستفيدين الى رحل")
        addBenefitsToTripsButton.setGeometry(0,170,181,31)
        addBenefitsToTripsButton.setStyleSheet("QPushButton {background-color:white;border:2px solid black;font: 14pt 'Arial';} QPushButton:hover {background-color:#c8c8c8;}")
        addBenefitsToTripsButton.clicked.connect(self.benefitsToTrips)

        #End minMenu Buttons style

        if self.width()-60 <= 847 or self.mainFrame.height()-60 <= 610:
            self.BenefitsFrame.setGeometry((self.width()-847)//2,(self.mainFrame.height()-610)//2,847,610)
            self.BenefitsTable.setGeometry(10,60,641,541)
            self.mainMenuFrame.setGeometry(660,60,181,261)
            self.closeButtonBenefetsFrame.setGeometry(800,10,41,31)
        else:
            self.BenefitsFrame.setGeometry(30,30,self.width()-60,self.mainFrame.height()-60)
            self.BenefitsTable.setGeometry(10,60,self.width()-181 - 89, self.BenefitsFrame.height() - 69)
            self.mainMenuFrame.setGeometry(self.BenefitsTable.width() + 19, 60,181,261)
            self.closeButtonBenefetsFrame.setGeometry((self.mainMenuFrame.x() + self.mainMenuFrame.width()) - 41, 10, 41,31)

            # self.BenefitsTable.resize(self.width()-self.mainMenuFrame.width() - 89, self.BenefitsFrame.height() - 69)

        self.adjustTableWidth()

        self.loadBenefits()
        self.BenefitsFrame.show()
    def addEntrySearchBenefits(self):
        if self.comboSearchBenefitsBox.currentText() == "الكل":
            try:
                self.searchEntryBenefits.destroy()
                self.searchEntryBenefits.hide()
            except:
                pass
            self.loadBenefits()
        else:
            try:
                self.searchEntryBenefits.destroy()
                self.searchEntryBenefits.hide()
            except:
                pass
            self.searchEntryBenefits = QLineEdit(self.BenefitsFrame)
            self.searchEntryBenefits.textChanged.connect(self.searchBenefitsFun)
            self.searchEntryBenefits.setGeometry(340,30,150,20)
            self.searchEntryBenefits.show()
    def searchBenefitsFun(self):
        if len(self.searchEntryBenefits.text())==0:
            self.loadBenefits()
            self.adjustTableWidth()
        else:
            self.BenefitsTable.setColumnCount(8)
            self.BenefitsTable.setRowCount(0)
            self.BenefitsTable.setHorizontalHeaderLabels(["اسم المستفيد","رقم الهوية","رقم الجوال","المدينة","الميلاد","الجنس","العمر","الجنسية"])
            tempThing = [] 

            if self.comboSearchBenefitsBox.currentText()=="رقم الهوية":
                cr.execute("SELECT identy FROM benefits")
            elif self.comboSearchBenefitsBox.currentText()=="رقم الجوال":
                cr.execute("SELECT phone FROM benefits")
            elif self.comboSearchBenefitsBox.currentText()=="الاسم":
                cr.execute("SELECT name FROM benefits")

            choices = cr.fetchall()
            posiple = []
            for o in choices:
                for n,i in enumerate(o):
                    try:
                        if o[n][:len(self.searchEntryBenefits.text())]==self.searchEntryBenefits.text():
                                if i not in posiple:
                                    posiple.append(i)
                    except:
                        pass

            for p in posiple:
                if self.comboSearchBenefitsBox.currentText()=="رقم الهوية":
                    cr.execute("SELECT * FROM benefits WHERE identy = ?", [p])
                elif self.comboSearchBenefitsBox.currentText()=="رقم الجوال":
                    cr.execute("SELECT * FROM benefits WHERE phone = ?", [p])
                elif self.comboSearchBenefitsBox.currentText()=="الاسم":
                    cr.execute("SELECT * FROM benefits WHERE name = ?", [p])

                for i in cr.fetchall():
                    tempThing.append(i)

            if self.showCheckBoxes:
                self.BenefitsTable.insertColumn(0)
                self.BenefitsTable.setHorizontalHeaderLabels(["","اسم المستفيد","رقم الهوية","رقم الجوال","المدينة","الميلاد","الجنس","العمر","الجنسية"])
                self.adjustTableWidth()

                for row,i in enumerate(tempThing):
                    self.BenefitsTable.insertRow(self.BenefitsTable.rowCount())
                    i = list(i)
                    i.insert(0,"")
                    for col,val in enumerate(i):
                        if col==0:
                            button = QCheckBox()
                            button.clicked.connect(lambda ch,BenefitId=i[2]:self.addBenefitIdToBenefitsList(BenefitId))
                            self.BenefitsTable.setIndexWidget(self.BenefitsTable.model().index(row,0),button)
                            if i[2] in self.beneditsIds:
                                button.setCheckState(Qt.CheckState.Checked)
                        else:
                            self.BenefitsTable.setItem(row,col,QTableWidgetItem(str(val)))
            else:
                for row,i in enumerate(tempThing):
                    self.BenefitsTable.insertRow(self.BenefitsTable.rowCount())
                    for col,val in enumerate(i):
                        self.BenefitsTable.setItem(row,col,QTableWidgetItem(str(val)))

    def benefitsToTrips(self):
        try:
            self.destroyFrame(self.mainMenuFrame)

        except:
            pass

        self.beneditsIds = []
        self.showCheckBoxes = True

        self.mainMenuFrame = QFrame(self.BenefitsFrame)
        self.mainMenuFrame.setStyleSheet("background-color:white;border:2px solid black")

        label = QLabel(parent=self.mainMenuFrame, text="يرجى اختيار المستفيدين من جدول\n المستفيدين ثم الضغط على زر الاضافة")
        label.setStyleSheet("border-bottom:none;font-size:11px;color:red;")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setGeometry(0,0,181,51)

        addBenefitsToTripButton = QPushButton(self.mainMenuFrame,text="اضافة")
        addBenefitsToTripButton.setGeometry(0,50,181,31)
        addBenefitsToTripButton.setStyleSheet("QPushButton {background-color:white;border:2px solid black;font: 14pt 'Arial';} QPushButton:hover {background-color:#c8c8c8;}")
        addBenefitsToTripButton.clicked.connect(self.addBenefitToTrip)

        backButton = QPushButton(self.mainMenuFrame,text="العودة")
        backButton.setGeometry(0,90,181,31)
        backButton.setStyleSheet("QPushButton {background-color:white;border:2px solid black;font: 14pt 'Arial';} QPushButton:hover {background-color:#c8c8c8;}")
        backButton.clicked.connect(self.showBenefits)

        if self.width()-60 <= 847 or self.mainFrame.height()-60 <= 610:
            self.mainMenuFrame.setGeometry(660,60,181,261)
        else:
            self.mainMenuFrame.setGeometry(self.BenefitsTable.width() + 19, 60,181,261)

        self.loadBenefits()
        self.mainMenuFrame.show()
    def createButtonBenefitsTable(self):
        if not self.showCheckBoxes:
            self.deleteButton = QAction(self.BenefitsTable)
            self.deleteButton.setIcon(QIcon("assests/deleteUser.png"))
            self.deleteButton.setText("حذف")
            self.deleteButton.setFont(QFont("Arial" , 12))
            self.deleteButton.triggered.connect(self.deleteBenefit)

            self.editButton = QAction(self.BenefitsTable)
            self.editButton.setIcon(QIcon("assests/edit.png"))
            self.editButton.setText("تعديل")
            self.editButton.setFont(QFont("Arial" , 12))
            self.editButton.triggered.connect(self.editBenefit)
            
            self.showTripsBenefits = QAction(self.BenefitsTable)
            self.showTripsBenefits.setText("اظهار الرحلات")
            self.showTripsBenefits.setFont(QFont("Arial" , 12))
            self.showTripsBenefits.triggered.connect(self.showTripsToBenefits)

            self.addToTrip = QAction(self.BenefitsTable)
            self.addToTrip.setText("اضافة لرحلة")
            self.addToTrip.setFont(QFont("Arial" , 12))
            self.addToTrip.triggered.connect(self.addBenefitToTrip)


            self.contextMenuBenefitsTable.addAction(self.deleteButton)
            self.contextMenuBenefitsTable.addAction(self.editButton)
            self.contextMenuBenefitsTable.addAction(self.showTripsBenefits)
            self.contextMenuBenefitsTable.addAction(self.addToTrip)
    def showMenuBenefitsTable(self,position):
        if not self.showCheckBoxes:
            indexes = self.BenefitsTable.selectedIndexes()
            for index in indexes:
                self.contextMenuBenefitsTable.exec(self.BenefitsTable.viewport().mapToGlobal(position))
    def showTripsToBenefits(self):
        try:
            self.destroyFrame(self.showBenefitsTripsFrame)
        except:
            pass
        self.BenefitIdToShowTrips = self.BenefitsTable.item(self.BenefitsTable.selectedIndexes()[0].row(),1).text()

        self.showBenefitsTripsFrame = QFrame(self.mainFrame)
        self.showBenefitsTripsFrame.setGeometry((self.mainFrame.width()-419)//2,(self.mainFrame.height()-319)//2,419,319)
        
        closeButton = QPushButton(self.showBenefitsTripsFrame)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:12px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 26px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(380,10,31,31)
        closeButton.clicked.connect(lambda x, frame=self.showBenefitsTripsFrame:self.destroyFrame(frame))

        self.showBenefitsTripsFrame.setStyleSheet("background-color:white")

        self.showBenefitsTripsTable = QTableWidget(self.showBenefitsTripsFrame)
        self.showBenefitsTripsTable.setGeometry(10,50,401,261)

        self.contextMenuShowBenefitsTripsTable = QMenu(self.showBenefitsTripsTable)
        self.contextMenuShowBenefitsTripsTable.setStyleSheet("background-color:grey")
        self.createButtonShowBenefitsTripsTable()

        self.showBenefitsTripsTable.setColumnCount(5)
        self.showBenefitsTripsTable.setHorizontalHeaderLabels(["id", "اسم الرحلة","تاريخ البداية", "تاريخ البداية هجري", "المتبرع"])
        self.showBenefitsTripsTable.setColumnHidden(0, True)
        
        self.showBenefitsTripsTable.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.showBenefitsTripsTable.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.showBenefitsTripsTable.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.showBenefitsTripsTable.customContextMenuRequested.connect(self.showMenuBenefitsTripsTable)

        self.showBenefitsTripsFrame.show()
        self.loadTripsBenefits()
    def showMenuBenefitsTripsTable(self,position):
        indexes = self.showBenefitsTripsTable.selectedIndexes()
        for index in indexes:
            self.contextMenuShowBenefitsTripsTable.exec(self.showBenefitsTripsTable.viewport().mapToGlobal(position))
    def createButtonShowBenefitsTripsTable(self):
        self.deleteButton = QAction(self.showBenefitsTripsTable)
        self.deleteButton.setIcon(QIcon("assests/trash.png"))
        self.deleteButton.setText("حذف")
        self.deleteButton.setFont(QFont("Arial" , 12))
        self.deleteButton.triggered.connect(self.deleteBenefitTrip)

        self.contextMenuShowBenefitsTripsTable.addAction(self.deleteButton)
    def deleteBenefitTrip(self):
        tripId = self.showBenefitsTripsTable.item(self.showBenefitsTripsTable.selectedIndexes()[0].row(),0).text()
        d = QMessageBox(parent=self,text=f"تأكيد حذف {self.showBenefitsTripsTable.item(self.showBenefitsTripsTable.selectedIndexes()[0].row(),1).text()}")
        d.setIcon(QMessageBox.Icon.Information)
        d.setWindowTitle("تأكيد")
        d.setStyleSheet("background-color:white")
        d.setStandardButtons(QMessageBox.StandardButton.Cancel|QMessageBox.StandardButton.Ok)
        important = d.exec()
        if important == QMessageBox.StandardButton.Ok:
            cr.execute("DELETE FROM trips_benfit WHERE benfit_id=? and trip_id=?", (self.BenefitIdToShowTrips, tripId))
            cr.execute("DELETE FROM end_date WHERE benefitId=? and trip_id=?", (self.BenefitIdToShowTrips, tripId))
            cr.execute("UPDATE trips SET lemit = lemit + 1 WHERE id = ?",[tripId])
            con.commit()
            self.loadTripsBenefits()
            d = QMessageBox(parent=self,text="تم الحذف بنجاح")
            d.setWindowTitle("نجاح")
            d.setIcon(QMessageBox.Icon.Information)
            d.setStyleSheet("background-color:white")
            d.exec()
    def loadTripsBenefits(self):
        self.showBenefitsTripsTable.setRowCount(0)
        valuesToAdd = []
        tripsIds = []
        cr.execute("SELECT trip_id FROM trips_benfit WHERE benfit_id = ?", [self.BenefitIdToShowTrips])
        for row, i in enumerate(cr.fetchall()):
            valuesToAdd.append([])
            for tripId in i:
                valuesToAdd[row].append(tripId)
                tripsIds.append(tripId)
        
        for row, tripId in enumerate(tripsIds):
            cr.execute("SELECT name, begin_date, begin_date_hijri, donorId FROM trips WHERE id = ?", [tripId])
            for tripInfo in cr.fetchall():
                for index in range(len(tripInfo)):
                    if index == 3:
                        cr.execute("SELECT name FROM donars WHERE id=?", [tripInfo[index]])
                        valuesToAdd[row].append(cr.fetchone()[0])
                    else:
                        valuesToAdd[row].append(tripInfo[index])

        for row in range(len(valuesToAdd)):
            self.showBenefitsTripsTable.insertRow(self.showBenefitsTripsTable.rowCount())
            for col in range(self.showBenefitsTripsTable.columnCount()):
                self.showBenefitsTripsTable.setItem(row,col,QTableWidgetItem(str(valuesToAdd[row][col])))

    def deleteBenefit(self):
        BenefitIdDelete = self.BenefitsTable.item(self.BenefitsTable.selectedIndexes()[0].row(),1).text()
        d = QMessageBox(parent=self,text=f"تأكيد حذف {self.BenefitsTable.item(self.BenefitsTable.selectedIndexes()[0].row(),0).text()}")
        d.setIcon(QMessageBox.Icon.Information)
        d.setWindowTitle("تأكيد")
        d.setStyleSheet("background-color:white")
        d.setStandardButtons(QMessageBox.StandardButton.Cancel|QMessageBox.StandardButton.Ok)
        important = d.exec()
        if important == QMessageBox.StandardButton.Ok:
            cr.execute("SELECT trip_id FROM trips_benfit WHERE benfit_id = ?",[BenefitIdDelete])
            for i in cr.fetchall():
                for j in i:
                    cr.execute("UPDATE trips SET lemit = lemit + 1 WHERE id = ?",[j])
            cr.execute("DELETE FROM benefits WHERE identy=?",[BenefitIdDelete])
            cr.execute("DELETE FROM end_date WHERE benefitId=?",[BenefitIdDelete])
            cr.execute("DELETE FROM trips_benfit WHERE benfit_id=?",[BenefitIdDelete])
            con.commit()

            self.loadBenefits()
            d = QMessageBox(parent=self,text="تم الحذف بنجاح")
            d.setWindowTitle("نجاح")
            d.setIcon(QMessageBox.Icon.Information)
            d.setStyleSheet("background-color:white")
            ret = d.exec()
    def editBenefit(self):
        try:
            self.destroyFrame(self.editBenefitFrame)
        except:
            pass
        self.editBenefitFrame = QFrame(self.mainFrame)
        self.editBenefitFrame.setGeometry((self.mainFrame.width()-350)//2,(self.mainFrame.height()-506)//2,350,506)
        self.editBenefitFrame.setStyleSheet("background-color:white;border:2px solid black")

        closeButton = QPushButton(self.editBenefitFrame)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:12px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 26px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(300,10,41,31)
        closeButton.clicked.connect(lambda x, frame=self.editBenefitFrame:self.destroyFrame(frame))
        
        #Start scrolAria
        
        self.frame = QFrame()

        layout = QVBoxLayout()
        self.frame.setLayout(layout)


        label = QLabel("اسم المستفيد")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.nameEntry = QLineEdit()

        layout.addWidget(label)
        layout.addWidget(self.nameEntry)

        label = QLabel("رقم الهوية")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")

        self.identyEntry = QLineEdit()

        layout.addWidget(label)
        layout.addWidget(self.identyEntry)

        label = QLabel("رقم الجوال")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.phoneEntry = QLineEdit()

        layout.addWidget(label)
        layout.addWidget(self.phoneEntry)

        label = QLabel("المدينة")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.cityEntry = QLineEdit()


        layout.addWidget(label)
        layout.addWidget(self.cityEntry)

        label = QLabel("الميلاد")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.birthEntry = QLineEdit()

        layout.addWidget(label)
        layout.addWidget(self.birthEntry)

        label = QLabel("الجنس")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.genderEntry = QComboBox()
        self.genderEntry.addItems(["ذكر","انثى"])

        layout.addWidget(label)
        layout.addWidget(self.genderEntry)

        label = QLabel("العمر")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.ageEntry = QLineEdit()

        layout.addWidget(label)
        layout.addWidget(self.ageEntry)

        label = QLabel("الجنسية")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.nationalityEntry = QLineEdit()

        layout.addWidget(label)
        layout.addWidget(self.nationalityEntry)

        editButton = QPushButton(text="تعديل")
        editButton.clicked.connect(self.completeEdit)
        editButton.setStyleSheet("QPushButton:hover {background-color:#c8c8c8;}")        

        layout.addWidget(editButton)

        self.scroolAria = QScrollArea(self.editBenefitFrame)
        self.scroolAria.setWidget(self.frame)
        self.scroolAria.setStyleSheet("border:1px solid gray")
        self.scroolAria.move(20,50)
        self.scroolAria.resize(321,431)

        self.scroolAria.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)


        self.scroolAria.setWidgetResizable(True)

        BenefitIdEdit = self.BenefitsTable.item(self.BenefitsTable.selectedIndexes()[0].row(),1).text()
        cr.execute("SELECT * FROM benefits WHERE identy = ?",[BenefitIdEdit])
        values = cr.fetchall()[0]

        self.nameEntry.setText(values[0])
        self.identyEntry.setText(values[1])
        self.identyEntry.setDisabled(True)
        self.phoneEntry.setText(values[2])
        self.cityEntry.setText(values[3])
        self.birthEntry.setText(values[4])
        self.genderEntry.setCurrentText(values[5])
        self.ageEntry.setText(values[6])
        self.nationalityEntry.setText(values[7])

        #End scrolAria
        self.editBenefitFrame.show()
    def completeEdit(self):
        cr.execute("UPDATE benefits set name=?, phone=?, city=?, birth=?, gender=?, age=?, nationality=? WHERE identy=?",(self.nameEntry.text(),self.phoneEntry.text(),self.cityEntry.text(),self.birthEntry.text(),self.genderEntry.currentText(),self.ageEntry.text(),self.nationalityEntry.text(),self.identyEntry.text()))
        con.commit()
        message = QMessageBox(parent=self,text="تم التعديل بنجاح")
        message.setIcon(QMessageBox.Icon.Information)
        message.setWindowTitle("نجاح")
        message.exec()
        self.loadBenefits()
    def addBenefit(self):
        try:
            self.destroyFrame(self.addBenefitFrame)
        except:
            pass
        self.addBenefitFrame = QFrame(self.mainFrame)
        self.addBenefitFrame.setGeometry((self.mainFrame.width()-350)//2,(self.mainFrame.height()-506)//2,350,506)
        self.addBenefitFrame.setStyleSheet("background-color:white;border:2px solid black")

        closeButton = QPushButton(self.addBenefitFrame)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:12px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 26px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(300,10,41,31)
        closeButton.clicked.connect(lambda x, frame=self.addBenefitFrame:self.destroyFrame(frame))
        
        #Start scrolAria
        
        self.frame = QFrame()

        layout = QVBoxLayout()
        self.frame.setLayout(layout)


        label = QLabel("اسم المستفيد")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.nameEntry = QLineEdit()

        layout.addWidget(label)
        layout.addWidget(self.nameEntry)

        label = QLabel("رقم الهوية")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")

        self.identyEntry = QLineEdit()

        layout.addWidget(label)
        layout.addWidget(self.identyEntry)

        label = QLabel("رقم الجوال")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.phoneEntry = QLineEdit()

        layout.addWidget(label)
        layout.addWidget(self.phoneEntry)

        label = QLabel("المدينة")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.cityEntry = QLineEdit()


        layout.addWidget(label)
        layout.addWidget(self.cityEntry)

        label = QLabel("الميلاد")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.birthEntry = QLineEdit()

        layout.addWidget(label)
        layout.addWidget(self.birthEntry)

        label = QLabel("الجنس")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.genderEntry = QComboBox()
        self.genderEntry.addItems(["ذكر","انثى"])

        layout.addWidget(label)
        layout.addWidget(self.genderEntry)

        label = QLabel("العمر")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.ageEntry = QLineEdit()

        layout.addWidget(label)
        layout.addWidget(self.ageEntry)

        label = QLabel("الجنسية")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.nationalityEntry = QLineEdit()

        layout.addWidget(label)
        layout.addWidget(self.nationalityEntry)

        addButton = QPushButton(text="اضافة")
        addButton.clicked.connect(self.completeAddBenefit)
        addButton.setStyleSheet("QPushButton:hover {background-color:#c8c8c8;}")        

        layout.addWidget(addButton)

        self.scroolAria = QScrollArea(self.addBenefitFrame)
        self.scroolAria.setWidget(self.frame)
        self.scroolAria.setStyleSheet("border:1px solid gray")
        self.scroolAria.move(20,50)
        self.scroolAria.resize(321,431)

        self.scroolAria.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)


        self.scroolAria.setWidgetResizable(True)


        #End scrolAria
        self.addBenefitFrame.show()
    def completeAddBenefit(self):
        if (len(self.nameEntry.text()) > 0 and len(self.identyEntry.text()) > 0 and len(self.phoneEntry.text()) > 0 and len(self.cityEntry.text()) > 0 and len(self.birthEntry.text()) > 0 and len(self.genderEntry.currentText()) > 0 and len(self.ageEntry.text()) > 0 and len(self.nationalityEntry.text())):
            cr.execute("SELECT name FROM benefits WHERE identy =?",[self.identyEntry.text()])
            if cr.fetchall() == []:
                cr.execute("INSERT INTO benefits (name,identy,phone,city,birth,gender,age,nationality) values (?,?,?,?,?,?,?,?)",(self.nameEntry.text(),self.identyEntry.text(),self.phoneEntry.text(),self.cityEntry.text(),self.birthEntry.text(),self.genderEntry.currentText(),self.ageEntry.text(),self.nationalityEntry.text()))
                con.commit()
                message = QMessageBox(parent=self,text="تمت الاضافة بنجاح")
                message.setIcon(QMessageBox.Icon.Information)
                message.setWindowTitle("نجاح")
                message.exec()
                self.loadBenefits()
            else:
                message = QMessageBox(parent=self,text="يوجد بالفعل مستفيد بنفس رقم الهوية")
                message.setIcon(QMessageBox.Icon.Critical)
                message.setWindowTitle("فشل")
                message.exec()
        else:
            message = QMessageBox(parent=self,text="يرجى تعبئة جميع الحقول")
            message.setIcon(QMessageBox.Icon.Critical)
            message.setWindowTitle("فشل")
            message.exec()
    def addBenefitToTrip(self):
        try:
            self.destroyFrame(self.addBenefitToTripFrame)
        except:
            pass
        self.currentIdTrip = None
        if not self.showCheckBoxes:
            self.beneditsIds = []
            self.beneditsIds.append(self.BenefitsTable.item(self.BenefitsTable.selectedIndexes()[0].row(),1).text())
        
        self.addBenefitToTripFrame = QFrame(self.mainFrame)
        self.addBenefitToTripFrame.setGeometry((self.mainFrame.width()-360)//2,(self.mainFrame.height()-506)//2,360,486)
        self.addBenefitToTripFrame.setObjectName("addBenefitToTripFrame")
        self.addBenefitToTripFrame.setStyleSheet("QFrame#addBenefitToTripFrame {background-color:white;border:2px solid black}")

        closeButton = QPushButton(self.addBenefitToTripFrame)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:12px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 26px;background-color:#fdfdfd}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(323,10,31,31)
        closeButton.clicked.connect(lambda x, frame=self.addBenefitToTripFrame:self.destroyFrame(frame))

        label = QLabel(parent=self.addBenefitToTripFrame, text="اختر الرحلة")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("font: 14pt 'Arial';background-color:white")
        label.setGeometry(108,70,151,31)

        self.comboBoxChoices = QComboBox(parent=self.addBenefitToTripFrame)
        self.comboBoxChoices.setStyleSheet("background-color:#fdfdfd")
        self.comboBoxChoices.addItems(["الكل","اسم الرحلة", "التاريخ"])
        self.comboBoxChoices.setGeometry(256,100,101,22)
        self.comboBoxChoices.currentIndexChanged.connect(self.addEntrySearchTripsBenefits)


        self.TripsTableToChoice = QTableWidget(parent=self.addBenefitToTripFrame)
        self.TripsTableToChoice.setStyleSheet("background-color:white")
        self.TripsTableToChoice.setColumnCount(4)
        self.TripsTableToChoice.setColumnHidden(0,True)
        self.TripsTableToChoice.setColumnWidth(1,40)
        self.TripsTableToChoice.setColumnWidth(2,140)
        self.TripsTableToChoice.setColumnWidth(3,140)

        self.TripsTableToChoice.setHorizontalHeaderLabels(["id", "", "اسم الرحلة", "التاريخ", "التاريخ هجري"])
        self.TripsTableToChoice.setGeometry(11,130,346,221)

        label = QLabel(parent=self.addBenefitToTripFrame, text="مدة الانتظار")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("font: 14pt 'Arial';background-color:white")
        label.setGeometry(108,370,151,31)

        self.WaitDuration = QLineEdit(parent=self.addBenefitToTripFrame)
        self.WaitDuration.setStyleSheet("background-color:#fdfdfd")
        self.WaitDuration.setValidator(QIntValidator(0,1000))
        self.WaitDuration.setText("6")
        self.WaitDuration.setGeometry(68,400,221,22)

        addBenefitToTripButton = QPushButton(parent=self.addBenefitToTripFrame, text="اضافة")
        addBenefitToTripButton.setStyleSheet("background-color:#fdfdfd")
        addBenefitToTripButton.clicked.connect(self.completeAddBenefitToTrip)
        addBenefitToTripButton.setGeometry(118,440,131,31)

        self.loadTripsToAddToBenefit()
        self.addBenefitToTripFrame.show()
    def completeAddBenefitToTrip(self):
        global numbers
        if self.currentIdTrip is not None:
            numbers = 0
            Names = []
            dateAfterIncresing = date.today()+ relativedelta(months=int(self.WaitDuration.text()))
            TripId = self.currentIdTrip

            cr.execute("SELECT lemit FROM trips WHERE id = ?",[TripId])
            if int(cr.fetchone()[0]) - len(self.beneditsIds) >=0:
                def addBenefitToTripComplete(idBenefit):
                    global numbers
                    cr.execute("SELECT benfit_id FROM trips_benfit WHERE benfit_id = ? and trip_id = ?",(idBenefit, TripId))
                    if cr.fetchall() == []:
                        cr.execute("UPDATE trips set lemit= lemit - 1 WHERE id = ?",[TripId])
                        cr.execute("INSERT INTO trips_benfit (benfit_id, trip_id) VALUES (?, ?)",(idBenefit, TripId))
                    cr.execute("DELETE FROM end_date WHERE benefitId= ?", [idBenefit])
                    cr.execute("INSERT INTO end_date (benefitId, end_date, trip_id) VALUES (?, ?, ?)",(idBenefit, dateAfterIncresing.strftime("%Y/%m/%d"), TripId))

                for i in self.beneditsIds:
                    cr.execute("SELECT name FROM benefits WHERE identy= ?", [i])
                    Names.append(cr.fetchone()[0])
                    Names.append("\n")
                    numbers+=1

                d = QMessageBox(parent=self,text=f"هل انت متأكد من اضافة {numbers} من المستفيدين الى الرحلة")
                d.setIcon(QMessageBox.Icon.Information)
                d.setWindowTitle("تأكيد")
                d.setStandardButtons(QMessageBox.StandardButton.Ok|QMessageBox.StandardButton.Cancel)
                d.setDetailedText("".join(Names))
                ret = d.exec()

                if ret == QMessageBox.StandardButton.Ok:

                    for idBenefit in self.beneditsIds:
                        cr.execute("SELECT end_date FROM end_date WHERE benefitId=?",[idBenefit])
                        value = cr.fetchone()
                        if value == None:
                            addBenefitToTripComplete(idBenefit)
                        else:
                            dateFromDb = str(value[0])
                            splitters = dateFromDb.split("/")
                            if date(int(CURRENT_TIME_TEMP[0]), int(CURRENT_TIME_TEMP[1]), int(CURRENT_TIME_TEMP[2])) >= date(int(splitters[0]), int(splitters[1]), int(splitters[2])):
                                addBenefitToTripComplete(idBenefit)

                            else:
                                cr.execute("SELECT name FROM benefits WHERE identy=?", [idBenefit])
                                d = QMessageBox(parent=self,text=f"المستفيد {cr.fetchone()[0]} لم يقض مهلة الانتظار بعد هل تريد اضافته؟")
                                d.setWindowTitle("تأكيد")
                                d.setIcon(QMessageBox.Icon.Information)
                                d.setStandardButtons(QMessageBox.StandardButton.Cancel|QMessageBox.StandardButton.Ok)
                                state = d.exec()
                                if state == QMessageBox.StandardButton.Ok:
                                    addBenefitToTripComplete(idBenefit)

                    con.commit()
                    if numbers > 0: 
                        d = QMessageBox(parent=self,text=f"تم اضافة {numbers} من المستفيدين الى الرحلة بنجاح")
                        d.setIcon(QMessageBox.Icon.Information)
                        d.setWindowTitle("نجاح")
                        d.setStandardButtons(QMessageBox.StandardButton.Ok)
                        d.exec()
                        self.destroyFrame(self.addBenefitToTripFrame)
                    self.showBenefits()
            else:
                message = QMessageBox(parent=self,text="تجاوزت الحد الأقصى للرحلة")
                message.setIcon(QMessageBox.Icon.Critical)
                message.setWindowTitle("فشل")
                message.exec()
        else:
            message = QMessageBox(parent=self,text="يرجى اختيار رحلة")
            message.setIcon(QMessageBox.Icon.Critical)
            message.setWindowTitle("فشل")
            message.exec()
    
    def addEntrySearchTripsBenefits(self):
        if self.comboBoxChoices.currentText() == "الكل":
            try:
                self.searchEntryTripsBenefits.destroy()
                self.searchEntryTripsBenefits.hide()
            except:
                pass
            self.loadTripsToAddToBenefit()
        else:
            try:
                self.searchEntryTripsBenefits.destroy()
                self.searchEntryTripsBenefits.hide()
            except:
                pass
            if self.comboBoxChoices.currentText() == "التاريخ":
                self.searchEntryTripsBenefits = CustomDateEdit(self.addBenefitToTripFrame)
                self.searchEntryTripsBenefits.setCalendarPopup(True)
                self.searchEntryTripsBenefits.setDisplayFormat("yyyy/MM/dd")
                arabic_locale = QLocale(QLocale.Language.Arabic, QLocale.Country.SaudiArabia)
                self.searchEntryTripsBenefits.setLocale(arabic_locale)
                self.searchEntryTripsBenefits.setFont(QFont("Arial",12))
                self.searchEntryTripsBenefits.setStyleSheet("background-color:white;color:black")
                self.todayButton = QPushButton("اليوم",clicked=lambda:self.searchEntryTripsBenefits.calendarWidget().setSelectedDate(QDate().currentDate()))
                self.todayButton.setStyleSheet("background-color:green;")
                self.searchEntryTripsBenefits.calendarWidget().layout().addWidget(self.todayButton)
                self.searchEntryTripsBenefits.dateChanged.connect(self.searchTripsBenefitsFun)
                self.searchEntryTripsBenefits.calendarWidget().setSelectedDate(QDate().currentDate())
                self.searchTripsBenefitsFun()
            else:
                self.searchEntryTripsBenefits = QLineEdit(self.addBenefitToTripFrame)
                self.searchEntryTripsBenefits.textChanged.connect(self.searchTripsBenefitsFun)

            self.searchEntryTripsBenefits.setStyleSheet("background-color:white")
            self.searchEntryTripsBenefits.setGeometry(100,100,150,20)
            self.searchEntryTripsBenefits.show()
    def searchTripsBenefitsFun(self):
        if len(self.searchEntryTripsBenefits.text())==0:
            self.loadTripsToAddToBenefit()
        else:
            self.TripsTableToChoice.setRowCount(0)
            tempThing = [] 
            posiple = []
            if self.comboBoxChoices.currentText()=="اسم الرحلة":
                cr.execute("SELECT name FROM trips")
                choices = cr.fetchall()
                for o in choices:
                    for n,i in enumerate(o):
                        x = (str(o[0])).split()
                        y = self.searchEntryTripsBenefits.text().split()
                        try:
                            l = len(y)
                            for n in (x):
                                for Q in (y):
                                    if Q == n[0:len(Q)]:
                                        l-=1
                            if l==0:
                                if i not in posiple:
                                    posiple.append(i)
                        except:
                            pass
                
                for p in posiple:
                    cr.execute("SELECT id, name, begin_date FROM trips WHERE name = ?", [p])
                    for i in cr.fetchall():
                        i = list(i)
                        i.insert(1,"")
                        tempThing.append(i)
            else:
                cr.execute("SELECT id, name, begin_date FROM trips WHERE begin_date = ?", [str(self.searchEntryTripsBenefits.text())])
                for i in cr.fetchall():
                    i = list(i)
                    i.insert(1,"")
                    tempThing.append(i)
            for row in range(len(tempThing)):
                self.TripsTableToChoice.insertRow(self.TripsTableToChoice.rowCount())
                for col in range(self.TripsTableToChoice.columnCount()):
                    self.TripsTableToChoice.setItem(row,col,QTableWidgetItem(str(tempThing[row][col])))
                    if col==1:
                        button = QRadioButton()
                        button.clicked.connect(lambda ch,tripId=tempThing[row][0]:self.changeCurrentId(tripId))
                        if tempThing[row][0] == self.currentIdTrip:
                            button.setChecked(True)
                        self.TripsTableToChoice.setIndexWidget(self.TripsTableToChoice.model().index(row,1),button)
    def loadTripsToAddToBenefit(self):
        self.TripsTableToChoice.setRowCount(0)
        cr.execute("SELECT id, name, begin_date FROM trips")
        
        tempThing = []
        for i in cr.fetchall():
            i = list(i)
            i.insert(1,"")
            tempThing.append(i)

        for row in range(len(tempThing)):
            self.TripsTableToChoice.insertRow(self.TripsTableToChoice.rowCount())
            for col in range(self.TripsTableToChoice.columnCount()):
                self.TripsTableToChoice.setItem(row,col,QTableWidgetItem(str(tempThing[row][col])))
                if col==1:
                    button = QRadioButton()
                    button.clicked.connect(lambda ch,tripId=tempThing[row][0]:self.changeCurrentId(tripId))
                    if tempThing[row][0] == self.currentIdTrip:
                        button.setChecked(True)
                    self.TripsTableToChoice.setIndexWidget(self.TripsTableToChoice.model().index(row,1),button)
    def changeCurrentId(self,tripId):

        self.currentIdTrip = tripId
    def addSearchEntry(self):
        if self.comboBoxChoices.currentText() == "الكل":
            try:
                self.choiceEntry.destroy()
                self.choiceEntry.close()
            except:
                pass
        else:
            try:
                self.choiceEntry.destroy()
                self.choiceEntry.close()
            except:
                pass
            self.choiceEntry = QLineEdit(parent=self.addBenefitToTripFrame)
            self.choiceEntry.setStyleSheet("background-color:white;")
            self.choiceEntry.setGeometry(120,100,121,20)
            self.choiceEntry.show()
    def addBenefitsFromExternal(self):
        try:
            self.destroyFrame(self.addBenefitsFromExternalFrame)
        except:
            pass
        self.addBenefitsFromExternalFrame = QFrame(self.mainFrame)
        self.addBenefitsFromExternalFrame.setGeometry((self.mainFrame.width()-206)//2,(self.mainFrame.height()-122)//2,206,122)
        self.addBenefitsFromExternalFrame.setStyleSheet("background-color:white;border:2px solid black")

        closeButton = QPushButton(self.addBenefitsFromExternalFrame)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:12px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 26px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(170,10,31,31)
        closeButton.clicked.connect(lambda x, frame=self.addBenefitsFromExternalFrame:self.destroyFrame(frame))

        label = QLabel(parent=self.addBenefitsFromExternalFrame,text="قم بتحميل الملف من ")
        label.setGeometry(30,50,161,20)
        label.setStyleSheet("font: 14pt 'Arial';border:none")
        
        downloadButton = QPushButton(self.addBenefitsFromExternalFrame,text="هنا")
        downloadButton.setStyleSheet("text-decoration:underline;font: 14pt 'Arial';color:rgb(2, 128, 186);border:none;")
        downloadButton.setGeometry(40,50,31,23)
        downloadButton.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        downloadButton.clicked.connect(self.downloadExetrnalFile)

        label = QLabel(parent=self.addBenefitsFromExternalFrame,text="قم بتعبئته ثم ارفعه من ")
        label.setGeometry(30,80,161,20)
        label.setStyleSheet("font: 14pt 'Arial';border:none")
        
        uploadButton = QPushButton(self.addBenefitsFromExternalFrame,text="هنا")
        uploadButton.setStyleSheet("text-decoration:underline;font: 14pt 'Arial';color:rgb(2, 128, 186);border:none;")
        uploadButton.setGeometry(30,80,31,23)
        uploadButton.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        uploadButton.clicked.connect(self.uploadExetrnalFile)

        self.addBenefitsFromExternalFrame.show()
    
    def downloadExetrnalFile(self):
        filePath = QFileDialog.getExistingDirectory(self,"Select a Directory")
        if len(filePath) > 0:
            try:
                shutil.copyfile("assests/examble.xlsx",f"{filePath}/نموذج تعبئة.xlsx")
                message = QMessageBox(parent=self,text="تم التحميل بنجاح")
                message.setIcon(QMessageBox.Icon.Information)
                message.setWindowTitle("نجاح")
                message.exec()
            except:
                message = QMessageBox(parent=self,text="فشل التحميل")
                message.setIcon(QMessageBox.Icon.Critical)
                message.setWindowTitle("فشل")
                message.exec()
    def uploadExetrnalFile(self):
        filePath = QFileDialog.getOpenFileName(self,'Select file','./','Excel Files (*.xlsx)')
        if len(filePath[0])!=0:
            cr.execute("SELECT identy FROM benefits")
            identiys = cr.fetchall()
            dataBaseIdentiys = []
            for i in identiys:
                for j in i:
                    dataBaseIdentiys.append(j)

            wk = load_workbook(filePath[0])
            ws = wk.active
            names = []
            idnent = []
            phones = []
            cities = []
            births = []
            gender = []
            ages = []
            nations = []
            for i in range(2,61):
                
                if ws["C"+str(i)].value is not None and str(ws["C"+str(i)].value) not in idnent and str(ws["C"+str(i)].value) not in dataBaseIdentiys:
                    idnent.append(str(ws["C"+str(i)].value))
                    if str(ws["A"+str(i)].value) is not None:
                        names.append(str(ws["A"+str(i)].value))

                    if str(ws["E"+str(i)].value) is not None:
                        phones.append(str(ws["E"+str(i)].value))

                    if str(ws["G"+str(i)].value) is not None:
                        cities.append(str(ws["G"+str(i)].value))

                    if str(ws["I"+str(i)].value) is not None:
                        births.append(str(ws["I"+str(i)].value))

                    if str(ws["K"+str(i)].value) is not None:
                        gender.append(str(ws["K"+str(i)].value))

                    if str(ws["M"+str(i)].value) is not None:
                        ages.append(str(ws["M"+str(i)].value))

                    if str(ws["O"+str(i)].value) is not None:
                        nations.append(str(ws["O"+str(i)].value))

            for r in range(len(names)):
                cr.execute(f"INSERT INTO benefits (name,identy,phone,city,birth,gender,age,nationality) values (?,?,?,?,?,?,?,?)",(names[r],idnent[r],phones[r],cities[r],births[r],gender[r],ages[r],nations[r]))
            con.commit()
            message = QMessageBox(parent=self,text="تمت الاضافة بنجاح")
            message.setIcon(QMessageBox.Icon.Information)
            message.setWindowTitle("نجاح")
            message.exec()
            self.loadBenefits()
    def exportAllBenefits(self):
        try:
            self.destroyFrame(self.exportBenefitFrame)
        except:
            pass
        self.exportBenefitFrame = QFrame(parent=self.mainFrame)
        self.exportBenefitFrame.setGeometry((self.mainFrame.width()-160)//2,(self.mainFrame.height()-174)//2,160,147)
        self.exportBenefitFrame.setStyleSheet("background-color:white;border:2px solid black")

        closeButton = QPushButton(self.exportBenefitFrame)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:8px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 18px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(120,10,31,21)
        closeButton.clicked.connect(lambda x, frame=self.exportBenefitFrame:self.destroyFrame(frame))

        label = QLabel(parent=self.exportBenefitFrame,text="الصيغة")
        label.setStyleSheet('font: 14pt "Arial";border:none')
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setGeometry(20,50,121,20)

        self.formatComboBox = QComboBox(self.exportBenefitFrame)
        self.formatComboBox.addItems(["Word","Pdf","Excel"])
        self.formatComboBox.setGeometry(8,80,141,22)
        
        exportButton = QPushButton(parent=self.exportBenefitFrame,text="تصدير")
        exportButton.setGeometry(30,110,101,31)
        exportButton.setStyleSheet("QPushButton:hover {background-color:#c8c8c8;}")
        exportButton.clicked.connect(self.complateExportAllBenefits)

        self.exportBenefitFrame.show()
    def complateExportAllBenefits(self):
        filePath = QFileDialog.getExistingDirectory(self,"Select a Directory")
        if len(filePath) > 0:
            cr.execute("SELECT * FROM benefits")
            benefits = cr.fetchall()
            if self.formatComboBox.currentText() == "Excel":
                headers = ["الاسم","الهوية","الجوال","المدينة","الميلاد","الجنس","العمر","الجنسية"]

                wb = Workbook()
                sheet = wb.active
                sheet.title = "sheet1"
                col = 1
                #write the headers
                for header in headers:
                    cell = sheet.cell(row=1, column=col)
                    cell.value = header
                    col+=2
                #End headers

                for row,benefit in enumerate(benefits):
                    col = 1
                    for j in benefit:
                        cell = sheet.cell(row=row+2, column=col)
                        cell.value = j
                        col+=2
                wb.save(f"{filePath}/جميع المستفيدين.xlsx")    
            elif self.formatComboBox.currentText() == "Word" or self.formatComboBox.currentText() == "Pdf":

                doc = docx.Document()

                sections = doc.sections
                sections.page_height = 11.69
                sections.page_width = 8.27
                sections = sections[-1]
                sections.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE


                new_width,new_height = sections.page_height,sections.page_width
                sections.page_width = new_width
                sections.page_height = new_height

                sections = doc.sections

                for section in sections:
                    section.top_margin = docx.shared.Cm(0.3)
                    section.bottom_margin = docx.shared.Cm(0.3)
                    section.left_margin = docx.shared.Cm(0.3)
                    section.right_margin = docx.shared.Cm(0.3)

                benefits_table = doc.add_table(rows=1,cols=9)
                benefits_table.style = "Table Grid"
                hdr_Cells = benefits_table.rows[0].cells
                hdr_Cells[8].text = "م"
                hdr_Cells[7].text = "الاسم"
                hdr_Cells[6].text = "الهوية"
                hdr_Cells[5].text = "الجوال"
                hdr_Cells[4].text = "المدينة"
                hdr_Cells[3].text = "الميلاد"
                hdr_Cells[2].text = "الجنس"
                hdr_Cells[1].text = "العمر"
                hdr_Cells[0].text = "الجنسية"

                for cell in hdr_Cells:
                    self.set_arabic_format(cell)

                b = 0
                
                for i in benefits:
                    b+=1   
                    row_Cells = benefits_table.add_row().cells
                    row_Cells[0].size = docx.shared.Pt(15)
                    row_Cells[1].size = docx.shared.Pt(15)
                    row_Cells[2].size = docx.shared.Pt(15)
                    row_Cells[3].size = docx.shared.Pt(15)
                    row_Cells[4].size = docx.shared.Pt(15)
                    row_Cells[5].size = docx.shared.Pt(15)
                    row_Cells[6].size = docx.shared.Pt(15)
                    row_Cells[7].size = docx.shared.Pt(15)
                    row_Cells[8].size = docx.shared.Pt(15)

                    row_Cells[8].text = str(b)
                    row_Cells[7].text = str(i[0])
                    row_Cells[6].text = str(i[1])
                    row_Cells[5].text = str(i[2])
                    row_Cells[4].text = str(i[3])
                    row_Cells[3].text = str(i[4])
                    row_Cells[2].text = str(i[5])
                    row_Cells[1].text = str(i[6])
                    row_Cells[0].text = str(i[7])
                    for cell in row_Cells:
                        self.set_arabic_format(cell)

                widths = (docx.shared.Inches(4), docx.shared.Inches(4), docx.shared.Inches(4), docx.shared.Inches(4), docx.shared.Inches(4), docx.shared.Inches(4),docx.shared.Inches(3),docx.shared.Inches(7),docx.shared.Inches(0.5))
                for row in benefits_table.rows:
                    for idx, width in enumerate(widths):
                        row.cells[idx].width = width
                for row in benefits_table.rows:
                    for cell in row.cells:
                        paragraphs = cell.paragraphs
                        for paragraph in paragraphs:
                            for run in paragraph.runs:
                                font = run.font
                                font.size= docx.shared.Pt(17)

                doc.save(f"{filePath}\جميع المستفيدين.docx")

            if self.formatComboBox.currentText() == "Pdf":
                with suppress_output():
                    convert(f"{filePath}\جميع المستفيدين.docx",f"{filePath}\جميع المستفيدين.pdf")

                os.remove(f"{filePath}\جميع المستفيدين.docx")

            message = QMessageBox(parent=self,text="تم التصدير بنجاح")
            message.setIcon(QMessageBox.Icon.Information)
            message.setWindowTitle("نجاح")
            message.exec()
    def loadBenefits(self):
        self.BenefitsTable.setColumnCount(8)
        self.BenefitsTable.setRowCount(0)
        self.BenefitsTable.setHorizontalHeaderLabels(["اسم المستفيد","رقم الهوية","رقم الجوال","المدينة","الميلاد","الجنس","العمر","الجنسية"])        
        
        cr.execute("SELECT * FROM benefits")
        tempThing = [] 
        for i in cr.fetchall():
            tempThing.append(i)
        if self.showCheckBoxes:
            self.BenefitsTable.insertColumn(0)
            self.BenefitsTable.setColumnWidth(0,40)
            self.BenefitsTable.setHorizontalHeaderLabels(["","اسم المستفيد","رقم الهوية","رقم الجوال","المدينة","الميلاد","الجنس","العمر","الجنسية"])
            for row,i in enumerate(tempThing):
                self.BenefitsTable.insertRow(self.BenefitsTable.rowCount())
                i = list(i)
                i.insert(0,"")
                for col,val in enumerate(i):
                    if col==0:
                        button = QCheckBox()
                        button.clicked.connect(lambda ch,BenefitId=i[2]:self.addBenefitIdToBenefitsList(BenefitId))
                        self.BenefitsTable.setIndexWidget(self.BenefitsTable.model().index(row,0),button)
                        if i[2] in self.beneditsIds:
                            button.setCheckState(Qt.CheckState.Checked)
                    else:
                        self.BenefitsTable.setItem(row,col,QTableWidgetItem(str(val)))
        else:
            for row,i in enumerate(tempThing):
                self.BenefitsTable.insertRow(self.BenefitsTable.rowCount())
                for col,val in enumerate(i):
                    self.BenefitsTable.setItem(row,col,QTableWidgetItem(str(val)))
    def addBenefitIdToBenefitsList(self,BenefitId):
        if (BenefitId in self.beneditsIds):
            self.beneditsIds.remove(BenefitId)
        else:
            self.beneditsIds.append(BenefitId)

    def destroyFrame(self,frame):

        for i in frame.children():
            i.deleteLater()

        frame.destroy()
        frame.deleteLater()

