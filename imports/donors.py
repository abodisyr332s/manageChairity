from imports.stuff import *

class Donors(QWidget):
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
    def showDonars(self):
        try:
            self.DonarsFrame.deleteLater()
        except:
            pass
        self.DonarsFrame = QFrame(self.mainFrame)
        self.DonarsFrame.setGeometry((self.mainFrame.width()-847)//2,(self.mainFrame.height()-610)//2,847,610)
        

        self.comboSearchDonarsBox = QComboBox(self.DonarsFrame)
        self.comboSearchDonarsBox.setGeometry(500,30,150,20)
        self.comboSearchDonarsBox.addItems(["الكل", "الاسم", "رقم الجوال"])
        self.comboSearchDonarsBox.activated.connect(self.addEntrySearchDonars)

        closeButton = QPushButton(self.DonarsFrame)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:12px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 26px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(800,10,41,31)
        closeButton.clicked.connect(lambda x, frame=self.DonarsFrame:self.destroyFrame(frame))

        self.DonarsFrame.setStyleSheet("background-color:white")

        self.DonorsTable = QTableWidget(self.DonarsFrame)
        self.DonorsTable.setGeometry(10,60,641,541)

        self.DonorsTable.setColumnCount(5)
        self.DonorsTable.setHorizontalHeaderLabels(["id","اسم المتبرع","رقم الجوال","مبلغ التبرع","عدد الرحل"])
        self.DonorsTable.setColumnHidden(0,True)
        
        self.contextMenuDonarsTable = QMenu(self.DonorsTable)
        self.contextMenuDonarsTable.setStyleSheet("background-color:grey")
        self.createButtonDonarsTable()

        self.DonorsTable.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.DonorsTable.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.DonorsTable.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.DonorsTable.customContextMenuRequested.connect(self.showMenuDonarsTable)
        
        self.DonorsTable.setColumnWidth(1,170)
        self.DonorsTable.setColumnWidth(2,150)
        self.DonorsTable.setColumnWidth(3,150)
        self.DonorsTable.setColumnWidth(4,150)

        #Start minMenu Buttons style
        mainMenuFrame = QFrame(self.DonarsFrame)
        mainMenuFrame.setGeometry(660,60,181,261)
        mainMenuFrame.setStyleSheet("background-color:white;border:2px solid black")
        
        label = QLabel(mainMenuFrame,text="القائمة الرئيسيه")
        label.setStyleSheet("background-color:white;border-bottom:none;font: 14pt 'Arial';")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setGeometry(0,0,181,31)

        addDonarButton = QPushButton(mainMenuFrame,text="اضافة متبرع")
        addDonarButton.setGeometry(0,50,181,31)
        addDonarButton.setStyleSheet("QPushButton {background-color:white;border:2px solid black;font: 14pt 'Arial';} QPushButton:hover {background-color:#c8c8c8;}")
        addDonarButton.clicked.connect(self.addDonar)


        exportDonars = QPushButton(mainMenuFrame,text="تصدير معلومات المتبرعين")
        exportDonars.setGeometry(0,90,181,31)
        exportDonars.setStyleSheet("QPushButton {background-color:white;border:2px solid black;font: 14pt 'Arial';} QPushButton:hover {background-color:#c8c8c8;}")
        exportDonars.clicked.connect(self.exportAllDonors)

        #End minMenu Buttons style
        self.loadDonars()
        self.DonarsFrame.show()
    def addEntrySearchDonars(self):
        if self.comboSearchDonarsBox.currentText() == "الكل":
            try:
                self.searchEntryDonars.destroy()
                self.searchEntryDonars.hide()
            except:
                pass
            self.loadDonars()
        else:
            try:
                self.searchEntryDonars.destroy()
                self.searchEntryDonars.hide()
            except:
                pass
            self.searchEntryDonars = QLineEdit(self.DonarsFrame)
            self.searchEntryDonars.textChanged.connect(self.searchDonarsFun)
            self.searchEntryDonars.setGeometry(340,30,150,20)
            self.searchEntryDonars.show()
    def searchDonarsFun(self):
        if len(self.searchEntryDonars.text())==0:
            self.loadDonars()
        else:
            self.DonorsTable.setRowCount(0)
            tempThing = [] 

            if self.comboSearchDonarsBox.currentText()=="الاسم":
                cr.execute("SELECT name FROM donars")
            else:
                cr.execute("SELECT phone FROM donars")
            choices = cr.fetchall()
            posiple = []
            if self.comboSearchDonarsBox.currentText()=="الاسم":
                for o in choices:
                    for n,i in enumerate(o):
                        x = (str(o[0])).split()
                        y = self.searchEntryDonars.text().split()
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
            else:
                for o in choices:
                    for n,i in enumerate(o):
                        try:
                            if o[n][:len(self.searchEntryDonars.text())]==self.searchEntryDonars.text():
                                    if i not in posiple:
                                        posiple.append(i)
                        except:
                            pass

            for p in posiple:
                if self.comboSearchDonarsBox.currentText()=="الاسم":
                    cr.execute("SELECT * FROM donars WHERE name = ? ORDER BY amount DESC", [p])
                else:
                    cr.execute("SELECT * FROM donars WHERE phone = ? ORDER BY amount DESC", [p])
                for i in cr.fetchall():
                    tempThing.append(i)

            for row,i in enumerate(tempThing):
                self.DonorsTable.insertRow(self.DonorsTable.rowCount())
                for col,val in enumerate(i):
                    self.DonorsTable.setItem(row,col,QTableWidgetItem(str(val)))

    def addDonar(self):
        try:
            self.addDonorFrame.deleteLater()
        except:
            pass
        self.addDonorFrame = QFrame(self.mainFrame)
        self.addDonorFrame.setGeometry((self.mainFrame.width()-350)//2,(self.mainFrame.height()-350)//2,350,350)
        self.addDonorFrame.setStyleSheet("background-color:white;border:2px solid black")

        closeButton = QPushButton(self.addDonorFrame)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:12px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 26px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(300,10,41,31)
        closeButton.clicked.connect(lambda x, frame=self.addDonorFrame:self.destroyFrame(frame))
        
        #Start scrolAria
        
        self.frame = QFrame()

        layout = QVBoxLayout()
        self.frame.setLayout(layout)


        label = QLabel("اسم المتبرع")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.nameEntry = QLineEdit()

        layout.addWidget(label)
        layout.addWidget(self.nameEntry)

        label = QLabel("رقم الجوال")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.phoneEntry = QLineEdit()

        layout.addWidget(label)
        layout.addWidget(self.phoneEntry)

        addButton = QPushButton(text="اضافة")
        addButton.clicked.connect(self.completeAddDonor)
        addButton.setStyleSheet("QPushButton:hover {background-color:#c8c8c8;}")        

        layout.addWidget(addButton)

        self.scroolAria = QScrollArea(self.addDonorFrame)
        self.scroolAria.setWidget(self.frame)
        self.scroolAria.setStyleSheet("border:1px solid gray")
        self.scroolAria.move(20,50)
        self.scroolAria.resize(321,200)

        self.scroolAria.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)


        self.scroolAria.setWidgetResizable(True)


        #End scrolAria
        self.addDonorFrame.show()
    def completeAddDonor(self):
        if (len(self.nameEntry.text()) > 0 and len(self.phoneEntry.text()) > 0):
            cr.execute("INSERT INTO donars (name,phone,amount,countOfTrips) values (?,?,?,?)",(self.nameEntry.text(),self.phoneEntry.text(),0,0))
            con.commit()
            message = QMessageBox(parent=self,text="تمت الاضافة بنجاح")
            message.setIcon(QMessageBox.Icon.Information)
            message.setWindowTitle("نجاح")
            message.exec()
            self.loadDonars()
        else:
            message = QMessageBox(parent=self,text="يرجى تعبئة جميع الحقول")
            message.setIcon(QMessageBox.Icon.Critical)
            message.setWindowTitle("فشل")
            message.exec()
    def loadDonars(self):
        self.DonorsTable.setRowCount(0)
        cr.execute("SELECT * FROM donars ORDER BY amount DESC")
        tempThing = [] 
        for i in cr.fetchall():
            tempThing.append(i)
        for row,i in enumerate(tempThing):
            self.DonorsTable.insertRow(self.DonorsTable.rowCount())
            for col,val in enumerate(i):
                self.DonorsTable.setItem(row,col,QTableWidgetItem(str(val)))
    def createButtonDonarsTable(self):
        self.deleteButton = QAction(self.DonorsTable)
        self.deleteButton.setIcon(QIcon("assests/deleteUser.png"))
        self.deleteButton.setText("حذف")
        self.deleteButton.setFont(QFont("Arial" , 12))
        self.deleteButton.triggered.connect(self.deleteDonar)

        self.editButton = QAction(self.DonorsTable)
        self.editButton.setIcon(QIcon("assests/edit.png"))
        self.editButton.setText("تعديل")
        self.editButton.setFont(QFont("Arial" , 12))
        self.editButton.triggered.connect(self.editDonars)
        
        self.addToTrip = QAction(self.DonorsTable)
        self.addToTrip.setText("اضافة رحلة")
        self.addToTrip.setFont(QFont("Arial" , 12))
        self.addToTrip.triggered.connect(self.addTripToDonar)

        self.showTrips = QAction(self.DonorsTable)
        self.showTrips.setText("اظهار الرحلات")
        self.showTrips.setFont(QFont("Arial" , 12))
        self.showTrips.triggered.connect(self.showTripsDonors)


        self.contextMenuDonarsTable.addAction(self.deleteButton)
        self.contextMenuDonarsTable.addAction(self.editButton)
        self.contextMenuDonarsTable.addAction(self.addToTrip)
        self.contextMenuDonarsTable.addAction(self.showTrips)
    def showMenuDonarsTable(self,position):
        indexes = self.DonorsTable.selectedIndexes()
        for index in indexes:
            self.contextMenuDonarsTable.exec(self.DonorsTable.viewport().mapToGlobal(position))
    def deleteDonar(self):
        donorIdDeleteDonor = self.DonorsTable.item(self.DonorsTable.selectedIndexes()[0].row(),0).text()
        d = QMessageBox(parent=self,text=f"تأكيد حذف {self.DonorsTable.item(self.DonorsTable.selectedIndexes()[0].row(),1).text()}")
        d.setIcon(QMessageBox.Icon.Information)
        d.setWindowTitle("تأكيد")
        d.setStyleSheet("background-color:white")
        d.setStandardButtons(QMessageBox.StandardButton.Cancel|QMessageBox.StandardButton.Ok)
        important = d.exec()
        if important == QMessageBox.StandardButton.Ok:
            cr.execute("DELETE FROM donars WHERE id=?",[donorIdDeleteDonor])
            cr.execute("SELECT id FROM trips WHERE donorId=?",[donorIdDeleteDonor])
            for i in cr.fetchall():
                for j in i:
                    cr.execute("DELETE FROM trips_benfit WHERE trip_id=?",[j])
                    cr.execute("DELETE FROM end_date WHERE trip_id=?",[j])
                    cr.execute("DELETE FROM trips WHERE id=?",[j])
            con.commit()
            self.loadDonars()
            d = QMessageBox(parent=self,text="تم الحذف بنجاح")
            d.setWindowTitle("نجاح")
            d.setIcon(QMessageBox.Icon.Information)
            d.setStyleSheet("background-color:white")
            ret = d.exec()
    def editDonars(self):
        try:
            self.editdonorFrame.deleteLater()
        except:
            pass
        self.editdonorFrame = QFrame(self.mainFrame)
        self.editdonorFrame.setGeometry((self.mainFrame.width()-350)//2,(self.mainFrame.height()-350)//2,350,350)
        self.editdonorFrame.setStyleSheet("background-color:white;border:2px solid black")

        closeButton = QPushButton(self.editdonorFrame)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:12px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 26px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(300,10,41,31)
        closeButton.clicked.connect(lambda x, frame=self.editdonorFrame:self.destroyFrame(frame))
        
        #Start scrolAria
        
        self.frame = QFrame()

        layout = QVBoxLayout()
        self.frame.setLayout(layout)


        label = QLabel("اسم المتبرع")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.nameEntry = QLineEdit()

        layout.addWidget(label)
        layout.addWidget(self.nameEntry)

        label = QLabel("رقم الجوال")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.phoneEntry = QLineEdit()

        layout.addWidget(label)
        layout.addWidget(self.phoneEntry)

        editButton = QPushButton(text="تعديل")
        editButton.clicked.connect(self.completeEditDonor)
        editButton.setStyleSheet("QPushButton:hover {background-color:#c8c8c8;}")        

        layout.addWidget(editButton)

        self.scroolAria = QScrollArea(self.editdonorFrame)
        self.scroolAria.setWidget(self.frame)
        self.scroolAria.setStyleSheet("border:1px solid gray")
        self.scroolAria.move(20,50)
        self.scroolAria.resize(321,200)

        self.scroolAria.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)


        self.scroolAria.setWidgetResizable(True)

        self.donorIdEditDonor = self.DonorsTable.item(self.DonorsTable.selectedIndexes()[0].row(),0).text()
        cr.execute("SELECT id,name,phone FROM donars WHERE id = ?",[self.donorIdEditDonor])
        values = cr.fetchall()[0]

        self.nameEntry.setText(values[1])
        self.phoneEntry.setText(values[2])

        #End scrolAria
        self.editdonorFrame.show()
    def completeEditDonor(self):
        cr.execute("UPDATE donars set name=?, phone=? WHERE id=?",(self.nameEntry.text(),self.phoneEntry.text(),self.donorIdEditDonor))
        con.commit()
        message = QMessageBox(parent=self,text="تم التعديل بنجاح")
        message.setIcon(QMessageBox.Icon.Information)
        message.setWindowTitle("نجاح")
        message.exec()
        self.loadDonars()
    def exportAllDonors(self):
        try:
            self.exportDnorsFrame.deleteLater()
        except:
            pass
        self.exportDnorsFrame = QFrame(parent=self.mainFrame)
        self.exportDnorsFrame.setGeometry((self.mainFrame.width()-160)//2,(self.mainFrame.height()-174)//2,160,147)
        self.exportDnorsFrame.setStyleSheet("background-color:white;border:2px solid black")

        closeButton = QPushButton(self.exportDnorsFrame)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:8px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 18px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(120,10,31,21)
        closeButton.clicked.connect(lambda x, frame=self.exportDnorsFrame:self.destroyFrame(frame))

        label = QLabel(parent=self.exportDnorsFrame,text="الصيغة")
        label.setStyleSheet('font: 14pt "Arial";border:none')
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setGeometry(20,50,121,20)

        self.formatComboBox = QComboBox(self.exportDnorsFrame)
        self.formatComboBox.addItems(["Word","Pdf","Excel"])
        self.formatComboBox.setGeometry(8,80,141,22)
        
        exportButton = QPushButton(parent=self.exportDnorsFrame,text="تصدير")
        exportButton.setGeometry(30,110,101,31)
        exportButton.setStyleSheet("QPushButton:hover {background-color:#c8c8c8;}")
        exportButton.clicked.connect(self.complateExportAllDonors)

        self.exportDnorsFrame.show()
    def complateExportAllDonors(self):
        filePath = QFileDialog.getExistingDirectory(self,"Select a Directory")
        if len(filePath) > 0:
            cr.execute("SELECT name,phone,amount,countOfTrips FROM donars")
            donors = cr.fetchall()
            if self.formatComboBox.currentText() == "Excel":
                headers = ["الاسم","الجوال","المبلغ","عدد الرحل"]

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

                for row,donor in enumerate(donors):
                    col = 1
                    for j in donor:
                        cell = sheet.cell(row=row+2, column=col)
                        cell.value = j
                        col+=2
                wb.save(f"{filePath}/جميع المتبرعين.xlsx")    
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

                donors_table = doc.add_table(rows=1,cols=5)
                donors_table.style = "Table Grid"
                hdr_Cells = donors_table.rows[0].cells
                hdr_Cells[4].text = "م"
                hdr_Cells[3].text = "الاسم"
                hdr_Cells[2].text = "الجوال"
                hdr_Cells[1].text = "المبلغ"
                hdr_Cells[0].text = "عدد الرحل"

                for cell in hdr_Cells:
                    self.set_arabic_format(cell)

                b = 0
                
                for i in donors:
                    b+=1   
                    row_Cells = donors_table.add_row().cells
                    row_Cells[0].size = docx.shared.Pt(15)
                    row_Cells[1].size = docx.shared.Pt(15)
                    row_Cells[2].size = docx.shared.Pt(15)
                    row_Cells[3].size = docx.shared.Pt(15)
                    row_Cells[4].size = docx.shared.Pt(15)


                    row_Cells[4].text = str(b)
                    row_Cells[3].text = str(i[0])
                    row_Cells[2].text = str(i[1])
                    row_Cells[1].text = str(i[2])
                    row_Cells[0].text = str(i[3])

                    for cell in row_Cells:
                        self.set_arabic_format(cell)

                widths = (docx.shared.Inches(4), docx.shared.Inches(4), docx.shared.Inches(4),docx.shared.Inches(4),docx.shared.Inches(0.5))
                for row in donors_table.rows:
                    for idx, width in enumerate(widths):
                        row.cells[idx].width = width
                        
                for row in donors_table.rows:
                    for cell in row.cells:
                        paragraphs = cell.paragraphs
                        for paragraph in paragraphs:
                            for run in paragraph.runs:
                                font = run.font
                                font.size= docx.shared.Pt(17)

                doc.save(f"{filePath}\جميع المتبرعين.docx")

            if self.formatComboBox.currentText() == "Pdf":
                with suppress_output():
                    convert(f"{filePath}\جميع المتبرعين.docx",f"{filePath}\جميع المتبرعين.pdf")
                os.remove(f"{filePath}\جميع المتبرعين.docx")

            message = QMessageBox(parent=self,text="تم التصدير بنجاح")
            message.setIcon(QMessageBox.Icon.Information)
            message.setWindowTitle("نجاح")
            message.exec()
    def addTripToDonar(self):
        try:
            self.addTripFrame.deleteLater()
        except:
            pass
        self.donorIdAddTrip = int(self.DonorsTable.item(self.DonorsTable.selectedIndexes()[0].row(),0).text())
        self.addTripFrame = QFrame(self.mainFrame)
        self.addTripFrame.setGeometry((self.mainFrame.width()-358)//2,(self.mainFrame.height()-350)//2,358,350)
        self.addTripFrame.setStyleSheet("background-color:white;border:2px solid black")

        closeButton = QPushButton(self.addTripFrame)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:12px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 26px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(310,10,31,31)
        closeButton.clicked.connect(lambda x, frame=self.addTripFrame:self.destroyFrame(frame))
        
        #Start scrolAria
        
        self.frame = QFrame()

        layout = QVBoxLayout()
        self.frame.setLayout(layout)


        label = QLabel("اسم الرحلة")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.tripName = QLineEdit()

        layout.addWidget(label)
        layout.addWidget(self.tripName)

        label = QLabel("العدد الأقصى للمعتمرين")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")

        self.BenefitsCount = QLineEdit()
        self.BenefitsCount.setValidator(QIntValidator(0,1000000000))
        self.BenefitsCount.setText('48')

        layout.addWidget(label)
        layout.addWidget(self.BenefitsCount)

        label = QLabel("تاريخ بداية الرحلة")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")

        self.dateHigryEntry = QLineEdit()
        self.dateHigryEntry.setDisabled(True)

        self.dateEntry = QDateEdit()
        self.dateEntry.setCalendarPopup(True)
        self.dateEntry.setDisplayFormat("yyyy/MM/dd")
        arabic_locale = QLocale(QLocale.Language.Arabic, QLocale.Country.SaudiArabia)
        self.dateEntry.setLocale(arabic_locale)
        self.dateEntry.setFont(QFont("Arial",12))
        self.dateEntry.setStyleSheet("background-color:white;color:black")
        self.todayButton = QPushButton("اليوم",clicked=lambda:self.dateEntry.calendarWidget().setSelectedDate(QDate().currentDate()))
        self.todayButton.setStyleSheet("background-color:green;")
        self.dateEntry.calendarWidget().layout().addWidget(self.todayButton)
        self.dateEntry.dateChanged.connect(self.changeHijryDate)
        self.dateEntry.calendarWidget().setSelectedDate(QDate().currentDate())

        dateHigryLabel = QLabel("التاريخ الهجري")
        dateHigryLabel.setStyleSheet("border:none;font: 15pt 'Arial';")
        dateHigryLabel.setAlignment(Qt.AlignmentFlag.AlignCenter)


        layout.addWidget(label)
        layout.addWidget(self.dateEntry)
        layout.addWidget(dateHigryLabel)
        layout.addWidget(self.dateHigryEntry)

        label = QLabel("تكلفة الرحلة")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("border:none;font: 15pt 'Arial';")
        self.tripAmount = QLineEdit()
        self.tripAmount.setValidator(QIntValidator(0,1000000000))

        layout.addWidget(label)
        layout.addWidget(self.tripAmount)

        addButton = QPushButton(text="اضافة")
        addButton.clicked.connect(self.completeAddTripToDonor)
        addButton.setStyleSheet("QPushButton:hover {background-color:#c8c8c8;}")        

        layout.addWidget(addButton)

        self.scroolAria = QScrollArea(self.addTripFrame)
        self.scroolAria.setWidget(self.frame)
        self.scroolAria.setStyleSheet("border:1px solid gray")
        self.scroolAria.move(20,50)
        self.scroolAria.resize(321,270)

        self.scroolAria.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)


        self.scroolAria.setWidgetResizable(True)

        #End scrolAria
        self.changeHijryDate()
        self.addTripFrame.show()
    def completeAddTripToDonor(self):
        if len(self.tripName.text()) > 0 and len(self.BenefitsCount.text()) > 0 and len(self.dateEntry.text()) and len(self.dateHigryEntry.text()) > 0 and len(self.tripAmount.text()) > 0:
            cr.execute("INSERT INTO trips (donorId, name, benefitsCount, lemit, begin_date,begin_date_hijri, amount) VALUES (?, ?, ?, ?, ?, ?,?)",(self.donorIdAddTrip,self.tripName.text(),self.BenefitsCount.text(),self.BenefitsCount.text(),self.dateEntry.text(),self.dateHigryEntry.text(),self.tripAmount.text()))
            cr.execute(f"UPDATE donars SET amount = amount + {self.tripAmount.text()} ,countOfTrips = countOfTrips + 1 WHERE id=?",[self.donorIdAddTrip])
            con.commit()
            message = QMessageBox(parent=self,text="تمت الاضافة بنجاح")
            message.setIcon(QMessageBox.Icon.Information)
            message.setWindowTitle("نجاح")
            message.exec()
            self.loadDonars()
    def changeHijryDate(self):
        currentDate = (str(self.dateEntry.text()))
        currentDateHijry = (currentDate).split("/")
        currentDateHijryFinesh = str(Gregorian(int(currentDateHijry[0]),int(currentDateHijry[1]),int(currentDateHijry[2])).to_hijri()).replace("-","/")
        self.dateHigryEntry.setText(currentDateHijryFinesh)
    def showTripsDonors(self):
        try:
            self.showTripsFrame.deleteLater()
        except:
            pass
        self.donorIdTrips = int(self.DonorsTable.item(self.DonorsTable.selectedIndexes()[0].row(),0).text())

        self.showTripsFrame = QFrame(self.mainFrame)
        self.showTripsFrame.setGeometry((self.mainFrame.width()-419)//2,(self.mainFrame.height()-319)//2,419,319)
        self.showTripsFrame.setStyleSheet("background-color:white;border:2px solid black")

        closeButton = QPushButton(self.showTripsFrame)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:12px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 26px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(380,10,31,31)
        closeButton.clicked.connect(lambda x, frame=self.showTripsFrame:self.destroyFrame(frame))
        
        self.showTripsTable = QTableWidget(self.showTripsFrame)
        self.showTripsTable.setStyleSheet("border:1px solid gray")
        self.showTripsTable.setColumnCount(7)
        self.showTripsTable.setColumnWidth(2,170)
        self.showTripsTable.setColumnWidth(3,170)
        self.showTripsTable.setColumnWidth(4,170)
        self.showTripsTable.setHorizontalHeaderLabels(["اسم الرحلة", "تاريخ البداية","تاريخ البداية هجري", "عدد المستفيدين المشاركين", "عدد المستفيدين المتبقي", "تكلفة الرحلة", "المتبرع"])
        self.showTripsTable.setGeometry(10,50,401,261)
        self.showTripsTable.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.showTripsTable.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)

        self.loadTripsToDonor()
        self.showTripsFrame.show()
    def loadTripsToDonor(self):
        self.showTripsTable.setRowCount(0)
        cr.execute("SELECT name, donorId, begin_date, begin_date_hijri, benefitsCount, lemit, amount FROM trips WHERE donorId=?",[self.donorIdTrips])
        
        tempThing = []
        for i in cr.fetchall():
            cr.execute("SELECT name FROM donars WHERE id = ?",[i[1]])
            i = list(i)
            i.remove(i[1])
            i.append(cr.fetchone()[0])
            tempThing.append(i)
        for row,i in enumerate(tempThing):
            self.showTripsTable.insertRow(self.showTripsTable.rowCount())
            for col,val in enumerate(i):
                self.showTripsTable.setItem(row,col,QTableWidgetItem(str(val)))
