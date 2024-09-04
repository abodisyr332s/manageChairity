
from imports.benefitsClass import *
from imports.donors import *

class MainWindow(BenefitsWindow, Donors):
    def __init__(self):
        super().__init__()
        uic.loadUi("assests/mainWindow.ui",self)
        self.setWindowTitle(title)
        self.setWindowIcon(QIcon(icon))
        self.benfitsButton.clicked.connect(self.showBenefits)
        self.donationsButton.clicked.connect(self.showDonars)
        self.showTripsButton.clicked.connect(self.showTrips)
        self.programmerButton.clicked.connect(self.showProgrammer)
        self.statsButton.clicked.connect(self.showStats)
        
        self.preWidth = 995
        self.preHeight = 680

        self.resizeEvent = self.resizeWindow
    def resizeWindow(self,resizeEvent):
        changedWidth = self.width() - 995
        changedHeight = self.height() - 680
        
        if self.width() >=995 and self.height() >=680:
            self.mainFrame.resize(self.width(), self.mainFrame.height() + self.height() - self.preHeight)
            self.buttonsFrame.resize(self.width(), 41)
            self.label.move((self.mainFrame.width() - self.label.width())//2, (self.mainFrame.height() - self.label.height())//2)

            
            self.preWidth = self.width()
            self.preHeight = self.height()

        try:
            self.showTripsAllFrame.move((self.width()-600)//2,(self.mainFrame.height()-419)//2)
        except:
            pass

        try:
            self.showProgrammerFrame.move((self.width()-273)//2,(self.mainFrame.height()-368)//2)
        except:
            pass

        try:
            self.showStatsFrame.move((self.width()-273)//2,(self.mainFrame.height()-354)//2)
        except:
            pass

        try:
            self.showBenefitsParticipateFrame.move((self.width()-419)//2,(self.mainFrame.height()-319)//2)
        except:
            pass

        try:
            self.exportTripReportFrame.move((self.width()-160)//2,(self.mainFrame.height()-174)//2)
        except:
            pass

        try:
            self.editTripFrame.move((self.width()-350)//2,(self.mainFrame.height()-350)//2)
        except:
            pass

        try:
            self.DonarsFrame.move((self.width()-847)//2,(self.mainFrame.height()-610)//2)
        except:
            pass

        try:
            self.addDonorFrame.move((self.width()-350)//2,(self.mainFrame.height()-350)//2)
        except:
            pass

        try:
            self.editdonorFrame.move((self.width()-350)//2,(self.mainFrame.height()-350)//2)
        except:
            pass

        try:
            self.exportDnorsFrame.move((self.width()-160)//2,(self.mainFrame.height()-174)//2)
        except:
            pass

        try:
            self.addTripFrame.move((self.width()-358)//2,(self.mainFrame.height()-350)//2)
        except:
            pass

        try:
            self.showTripsFrame.move((self.width()-419)//2,(self.mainFrame.height()-319)//2)
        except:
            pass

        try:
            if self.width()-60 <= 847 or self.mainFrame.height()-60 <= 610:
                self.BenefitsFrame.setGeometry((self.width()-847)//2,(self.mainFrame.height()-610)//2,847,610)
                self.BenefitsTable.setGeometry(10,60,641,541)
                self.mainMenuFrame.move(660,60)
                self.closeButtonBenefetsFrame.setGeometry(800,10,41,31)
            else:
                self.BenefitsFrame.move(30,30)
                self.BenefitsFrame.resize(self.width()-60, self.mainFrame.height()-60)

                self.BenefitsTable.resize(self.width()-self.mainMenuFrame.width() - 89, self.BenefitsFrame.height() - 69)
                self.mainMenuFrame.move(self.BenefitsTable.width() + 19, 60)
                self.closeButtonBenefetsFrame.move((self.mainMenuFrame.x() + self.mainMenuFrame.width()) - 41, 10)

            deserve = self.BenefitsTable.columnCount()
            for i in range(self.BenefitsTable.columnCount()):
                if self.showCheckBoxes and i == 0:
                    deserve -=1
                    continue
                self.BenefitsTable.setColumnWidth(i, (self.BenefitsTable.width() - 20) // deserve)
        except:
            pass

        try:
            self.showBenefitsTripsFrame.move((self.width()-419)//2,(self.mainFrame.height()-319)//2)
        except:
            pass

        try:
            self.editBenefitFrame.move((self.width()-350)//2,(self.mainFrame.height()-506)//2)
        except:
            pass
        
        try:
            self.addBenefitFrame.move((self.width()-350)//2,(self.mainFrame.height()-506)//2)
        except:
            pass

        try:
            self.addBenefitToTripFrame.move((self.width()-360)//2,(self.mainFrame.height()-506)//2)
        except:
            pass

        try:
            self.addBenefitsFromExternalFrame.move((self.width()-206)//2,(self.mainFrame.height()-122)//2)
        except:
            pass

        try:
            self.exportBenefitFrame.move((self.width()-160)//2,(self.mainFrame.height()-174)//2)
        except:
            pass

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
            
    def showTrips(self):
        try:
            self.showTripsAllFrame.deleteLater()
        except:
            pass
        self.showTripsAllFrame = QFrame(self.mainFrame)
        self.showTripsAllFrame.setGeometry((self.mainFrame.width()-600)//2,(self.mainFrame.height()-519)//2,600,519)
        self.showTripsAllFrame.setStyleSheet("background-color:white")

        closeButton = QPushButton(self.showTripsAllFrame)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:12px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 26px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(560,10,31,31)
        closeButton.clicked.connect(lambda x, frame=self.showTripsAllFrame:self.destroyFrame(frame))
        

        self.comboSearchTripsBox = QComboBox(self.showTripsAllFrame)
        self.comboSearchTripsBox.setGeometry(300,20,150,20)
        self.comboSearchTripsBox.addItems(["الكل", "الاسم", "التاريخ"])
        self.comboSearchTripsBox.activated.connect(self.addEntrySearchTrips)

        self.showAllTripsTable = QTableWidget(self.showTripsAllFrame)
        self.showAllTripsTable.setStyleSheet("border:1px solid gray")
        self.showAllTripsTable.setColumnCount(8)
        self.showAllTripsTable.setColumnHidden(0,True)
        self.showAllTripsTable.setColumnWidth(2,170)
        self.showAllTripsTable.setColumnWidth(3,170)
        self.showAllTripsTable.setColumnWidth(4,170)

        self.contextMenuTripsTableAll = QMenu(self.showAllTripsTable)
        self.contextMenuTripsTableAll.setStyleSheet("background-color:grey")
        self.createButtonsTripsTable()

        self.showAllTripsTable.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.showAllTripsTable.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.showAllTripsTable.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.showAllTripsTable.customContextMenuRequested.connect(self.showMenuTripsAllTable)

        self.showAllTripsTable.setHorizontalHeaderLabels(["id","اسم الرحلة", "تاريخ البداية","تاريخ البداية هجري", "عدد المستفيدين المشاركين", "عدد المستفيدين المتبقي", "تكلفة الرحلة", "المتبرع"])
        self.showAllTripsTable.setGeometry(10,50,581,361)
        self.showAllTripsTable.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.showAllTripsTable.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)

        exportAllButton = QPushButton(parent=self.showTripsAllFrame,text="تصدير الكل")
        
        exportAllButton.setStyleSheet("QPushButton {border:2px solid black;font:14pt 'Arial'}QPushButton:hover {background-color:#c8c8c8;}")        
        exportAllButton.setGeometry(220,440,180,31)
        exportAllButton.clicked.connect(self.exportAllTrips)

        self.loadTrips()
        self.showTripsAllFrame.show()
    def exportAllTrips(self):
        try:
            self.destroyFrame(self.exportAllTripsFrame)
        except:
            pass

        self.exportAllTripsFrame = QFrame(parent=self.mainFrame)
        self.exportAllTripsFrame.setGeometry((self.mainFrame.width()-160)//2,(self.mainFrame.height()-174)//2,160,147)
        self.exportAllTripsFrame.setStyleSheet("background-color:white;border:2px solid black")

        closeButton = QPushButton(self.exportAllTripsFrame)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:8px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 18px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(120,10,31,21)
        closeButton.clicked.connect(lambda x, frame=self.exportAllTripsFrame:self.destroyFrame(frame))

        label = QLabel(parent=self.exportAllTripsFrame,text="الصيغة")
        label.setStyleSheet('font: 14pt "Arial";border:none')
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setGeometry(20,50,121,20)

        self.formatComboBoxExportTrips = QComboBox(self.exportAllTripsFrame)
        self.formatComboBoxExportTrips.addItems(["Word","Pdf","Excel"])
        self.formatComboBoxExportTrips.setGeometry(8,80,141,22)
        
        exportButton = QPushButton(parent=self.exportAllTripsFrame,text="تصدير")
        exportButton.setGeometry(30,110,101,31)
        exportButton.setStyleSheet("QPushButton:hover {background-color:#c8c8c8;}")
        exportButton.clicked.connect(self.completeExportAllTrips)

        self.exportAllTripsFrame.show()
    def completeExportAllTrips(self):
        filePath = QFileDialog.getExistingDirectory(self,"Select a Directory")
        if len(filePath) > 0:
            cr.execute("SELECT name, benefitsCount, begin_date, begin_date_hijri, amount, donorId FROM trips ORDER BY amount DESC")
            trips = cr.fetchall()

            if self.formatComboBoxExportTrips.currentText() == "Excel":
                headers = ["اسم الرحلة", "عدد المستفيدين", "تاريخ البداية","تاريخ البداية هجري","المبلغ","اسم المتبرع"]

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

                for row,trip in enumerate(trips):
                    col = 1
                    for j in trip:
                        if col == 11:
                            cr.execute("SELECT name FROM donars WHERE id = ?",[j])
                            cell = sheet.cell(row=row+2, column=col)
                            cell.value = cr.fetchone()[0]
                            col+=2
                        else:
                            cell = sheet.cell(row=row+2, column=col)
                            cell.value = j
                            col+=2
                wb.save(f"{filePath}/جميع الرحل.xlsx")    
            elif self.formatComboBoxExportTrips.currentText() == "Word" or self.formatComboBoxExportTrips.currentText() == "Pdf":

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

                trips_table = doc.add_table(rows=1,cols=7)
                trips_table.style = "Table Grid"
                hdr_Cells = trips_table.rows[0].cells
                hdr_Cells[6].text = "م"
                hdr_Cells[5].text = "اسم الرحلة"
                hdr_Cells[4].text = "عدد المستفيدين"
                hdr_Cells[3].text = "تاريخ البداية"
                hdr_Cells[2].text = "تاريخ البداية هجري"
                hdr_Cells[1].text = "المبلغ"
                hdr_Cells[0].text = "اسم المتبرع"

                for cell in hdr_Cells:
                    self.set_arabic_format(cell)

                b = 0
                
                for i in trips:
                    b+=1   
                    row_Cells = trips_table.add_row().cells
                    row_Cells[0].size = docx.shared.Pt(15)
                    row_Cells[1].size = docx.shared.Pt(15)
                    row_Cells[2].size = docx.shared.Pt(15)
                    row_Cells[3].size = docx.shared.Pt(15)
                    row_Cells[4].size = docx.shared.Pt(15)
                    row_Cells[5].size = docx.shared.Pt(15)
                    row_Cells[6].size = docx.shared.Pt(15)


                    row_Cells[6].text = str(b)
                    row_Cells[5].text = str(i[0])
                    row_Cells[4].text = str(i[1])
                    row_Cells[3].text = str(i[2])
                    row_Cells[2].text = str(i[3])
                    row_Cells[1].text = str(i[4])
                    cr.execute("SELECT name FROM donars WHERE id = ?",[i[5]])
                    row_Cells[0].text = str(cr.fetchone()[0])
                    for cell in row_Cells:
                        self.set_arabic_format(cell)
                widths = (docx.shared.Inches(4), docx.shared.Inches(4), docx.shared.Inches(4), docx.shared.Inches(4), docx.shared.Inches(4), docx.shared.Inches(4),docx.shared.Inches(0.5))
                
                for row in trips_table.rows:
                    for idx, width in enumerate(widths):
                        row.cells[idx].width = width
                for row in trips_table.rows:
                    for cell in row.cells:
                        paragraphs = cell.paragraphs
                        for paragraph in paragraphs:
                            for run in paragraph.runs:
                                font = run.font
                                font.size= docx.shared.Pt(17)


                doc.save(f"{filePath}\جميع الرحل.docx")

            if self.formatComboBoxExportTrips.currentText() == "Pdf":
                with suppress_output():
                    convert(f"{filePath}\جميع الرحل.docx",f"{filePath}\جميع الرحل.pdf")

                os.remove(f"{filePath}\جميع الرحل.docx")

            message = QMessageBox(parent=self,text="تم التصدير بنجاح")
            message.setIcon(QMessageBox.Icon.Information)
            message.setWindowTitle("نجاح")
            message.exec()
    def addEntrySearchTrips(self):
        if self.comboSearchTripsBox.currentText() == "الكل":
            try:
                self.searchEntryTrips.destroy()
                self.searchEntryTrips.hide()
            except:
                pass
            self.loadTrips()
        else:
            try:
                self.searchEntryTrips.destroy()
                self.searchEntryTrips.hide()
            except:
                pass
            if self.comboSearchTripsBox.currentText() == "التاريخ":
                self.searchEntryTrips = CustomDateEdit(self.showTripsAllFrame)
                self.searchEntryTrips.setCalendarPopup(True)
                self.searchEntryTrips.setDisplayFormat("yyyy/MM/dd")
                arabic_locale = QLocale(QLocale.Language.Arabic, QLocale.Country.SaudiArabia)
                self.searchEntryTrips.setLocale(arabic_locale)
                self.searchEntryTrips.setFont(QFont("Arial",12))
                self.searchEntryTrips.setStyleSheet("background-color:white;color:black")
                self.todayButton = QPushButton("اليوم",clicked=lambda:self.searchEntryTrips.calendarWidget().setSelectedDate(QDate().currentDate()))
                self.todayButton.setStyleSheet("background-color:green;")
                self.searchEntryTrips.calendarWidget().layout().addWidget(self.todayButton)
                self.searchEntryTrips.dateChanged.connect(self.searchTripsFun)
                self.searchEntryTrips.calendarWidget().setSelectedDate(QDate().currentDate())
                self.searchTripsFun()
            else:
                self.searchEntryTrips = QLineEdit(self.showTripsAllFrame)
                self.searchEntryTrips.textChanged.connect(self.searchTripsFun)
            self.searchEntryTrips.setGeometry(140,20,150,20)
            self.searchEntryTrips.show()
    def searchTripsFun(self):
        if len(self.searchEntryTrips.text())==0:
            self.loadTrips()
        else:
            self.showAllTripsTable.setRowCount(0)
            tempThing = [] 
            posiple = []

            if self.comboSearchTripsBox.currentText()=="الاسم":
                cr.execute("SELECT name FROM trips")
                choices = cr.fetchall()
                for o in choices:
                    for n,i in enumerate(o):
                        x = (str(o[0])).split()
                        y = self.searchEntryTrips.text().split()
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
                    cr.execute("SELECT id, name, donorId, begin_date, begin_date_hijri, benefitsCount, lemit, amount FROM trips WHERE name = ?", [p])
                    for i in cr.fetchall():
                        cr.execute("SELECT name FROM donars WHERE id = ?",[i[2]])
                        i = list(i)
                        i.remove(i[2])
                        i.insert(4,(int(i[4]) - int(i[5])))
                        i.remove(i[5])
                        i.append(cr.fetchone()[0])
                        tempThing.append(i)
            else:
                cr.execute("SELECT id, name, donorId, begin_date, begin_date_hijri, benefitsCount, lemit, amount FROM trips WHERE begin_date = ?", [str(self.searchEntryTrips.text())])
                for i in cr.fetchall():
                    cr.execute("SELECT name FROM donars WHERE id = ?",[i[2]])
                    i = list(i)
                    i.remove(i[2])
                    i.insert(4,(int(i[4]) - int(i[5])))
                    i.remove(i[5])
                    i.append(cr.fetchone()[0])
                    tempThing.append(i)


            for row,i in enumerate(tempThing):
                self.showAllTripsTable.insertRow(self.showAllTripsTable.rowCount())
                for col,val in enumerate(i):
                    self.showAllTripsTable.setItem(row,col,QTableWidgetItem(str(val)))

    def loadTrips(self):
        self.showAllTripsTable.setRowCount(0)
        cr.execute("SELECT id, name, donorId, begin_date, begin_date_hijri, benefitsCount, lemit, amount FROM trips")

        # self.showAllTripsTable.setHorizontalHeaderLabels(["id","اسم الرحلة", "تاريخ البداية","تاريخ البداية هجري", "عدد المستفيدين المشاركين", "عدد المستفيدين المتبقي", "تكلفة الرحلة", "المتبرع"])
        

        tempThing = []
        for i in cr.fetchall():
            lis = []
            i = list(i)

            lis.insert(0, i[0])
            lis.insert(1, i[1])
            lis.insert(2, i[3])

            cr.execute("SELECT name FROM donars WHERE id = ?",[i[2]])
            
            lis.insert(3, i[4])
            lis.insert(4, int(i[5]) - int(i[6]))

            lis.insert(5, i[6])
            lis.insert(6, i[7])

            lis.insert(7, cr.fetchone()[0])

            tempThing.append(lis)
        for row,i in enumerate(tempThing):
            self.showAllTripsTable.insertRow(self.showAllTripsTable.rowCount())
            for col,val in enumerate(i):
                self.showAllTripsTable.setItem(row,col,QTableWidgetItem(str(val)))
    def createButtonsTripsTable(self):
        self.deleteTripButton = QAction(self.showAllTripsTable)
        self.deleteTripButton.setIcon(QIcon("assests/trash.png"))
        self.deleteTripButton.setText("حذف")
        self.deleteTripButton.setFont(QFont("Arial" , 12))
        self.deleteTripButton.triggered.connect(self.deleteTrip)

        self.editTripButton = QAction(self.showAllTripsTable)
        self.editTripButton.setIcon(QIcon("assests/edit.png"))
        self.editTripButton.setText("تعديل")
        self.editTripButton.setFont(QFont("Arial" , 12))
        self.editTripButton.triggered.connect(self.editTrip)

        self.showBenefitsParticipateButton = QAction(self.showAllTripsTable)
        self.showBenefitsParticipateButton.setText("اظهار المستفيدين")
        self.showBenefitsParticipateButton.setFont(QFont("Arial" , 12))
        self.showBenefitsParticipateButton.triggered.connect(self.showBenefitsParticipate)

        self.exportReportButton = QAction(self.showAllTripsTable)
        self.exportReportButton.setText("تصدير كشف رحلة")
        self.exportReportButton.setFont(QFont("Arial" , 12))
        self.exportReportButton.triggered.connect(self.exportReport)

        self.contextMenuTripsTableAll.addAction(self.deleteTripButton)
        self.contextMenuTripsTableAll.addAction(self.editTripButton)
        self.contextMenuTripsTableAll.addAction(self.showBenefitsParticipateButton)
        self.contextMenuTripsTableAll.addAction(self.exportReportButton)
    def showBenefitsParticipate(self):
        try:
            self.destroyFrame(self.showBenefitsParticipateFrame)
        except:
            pass
        self.TripIdToShowParticipateBenefits = self.showAllTripsTable.item(self.showAllTripsTable.selectedIndexes()[0].row(),0).text()

        self.showBenefitsParticipateFrame = QFrame(self.mainFrame)
        self.showBenefitsParticipateFrame.setGeometry((self.mainFrame.width()-419)//2,(self.mainFrame.height()-319)//2,419,319)
        
        closeButton = QPushButton(self.showBenefitsParticipateFrame)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:12px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 26px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(380,10,31,31)
        closeButton.clicked.connect(lambda x, frame=self.showBenefitsParticipateFrame:self.destroyFrame(frame))

        self.showBenefitsParticipateFrame.setStyleSheet("background-color:white")

        self.showBenefitsParticipateTable = QTableWidget(self.showBenefitsParticipateFrame)
        self.showBenefitsParticipateTable.setGeometry(10,50,401,261)

        self.contextMenuShowBenefitsParticipateTable = QMenu(self.showBenefitsParticipateTable)
        self.contextMenuShowBenefitsParticipateTable.setStyleSheet("background-color:grey")
        self.createButtonShowBenefitsParticipateTable()

        self.showBenefitsParticipateTable.setColumnCount(3)
        self.showBenefitsParticipateTable.setHorizontalHeaderLabels(["اسم المستفيد", "رقم الهوية", "رقم الجوال"])
        self.showBenefitsParticipateTable.setColumnWidth(0, 190)
        self.showBenefitsParticipateTable.setColumnWidth(1, 100)
        self.showBenefitsParticipateTable.setColumnWidth(2, 100)



        self.showBenefitsParticipateTable.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.showBenefitsParticipateTable.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.showBenefitsParticipateTable.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.showBenefitsParticipateTable.customContextMenuRequested.connect(self.showMenuBenefitsParticipate)

        self.loadBenefitsParticipate()
        self.showBenefitsParticipateFrame.show()
    def loadBenefitsParticipate(self):
        self.showBenefitsParticipateTable.setRowCount(0)
        benefitsParticipatedIds = []
        values = []
        cr.execute("SELECT benfit_id FROM trips_benfit WHERE trip_id = ?",[self.TripIdToShowParticipateBenefits])
        for i in cr.fetchall():
            for j in i:
                benefitsParticipatedIds.append(j)
        
        for participatedBenfitId in benefitsParticipatedIds:
            cr.execute("SELECT name, identy, phone FROM benefits WHERE identy=?",[participatedBenfitId])
            for i in cr.fetchall():
                values.append(list(i))

        for row,i in enumerate(values):
            self.showBenefitsParticipateTable.insertRow(self.showBenefitsParticipateTable.rowCount())
            for col,val in enumerate(i):
                self.showBenefitsParticipateTable.setItem(row,col,QTableWidgetItem(str(val)))

    def deleteBenefitsFromTrip(self):
        BenefitId = self.showBenefitsParticipateTable.item(self.showBenefitsParticipateTable.selectedIndexes()[0].row(),1).text()
        d = QMessageBox(parent=self,text=f"تأكيد حذف {self.showBenefitsParticipateTable.item(self.showBenefitsParticipateTable.selectedIndexes()[0].row(),0).text()}")
        d.setIcon(QMessageBox.Icon.Information)
        d.setWindowTitle("تأكيد")
        d.setStyleSheet("background-color:white")
        d.setStandardButtons(QMessageBox.StandardButton.Cancel|QMessageBox.StandardButton.Ok)
        important = d.exec()
        if important == QMessageBox.StandardButton.Ok:
            cr.execute("DELETE FROM trips_benfit WHERE benfit_id=? and trip_id=?", (BenefitId, self.TripIdToShowParticipateBenefits))
            cr.execute("DELETE FROM end_date WHERE benefitId=? and trip_id=?", (BenefitId, self.TripIdToShowParticipateBenefits))
            cr.execute("UPDATE trips SET lemit = lemit + 1 WHERE id = ?",[self.TripIdToShowParticipateBenefits])
            con.commit()
            self.loadTrips()
            self.loadBenefitsParticipate()
            d = QMessageBox(parent=self,text="تم الحذف بنجاح")
            d.setWindowTitle("نجاح")
            d.setIcon(QMessageBox.Icon.Information)
            d.setStyleSheet("background-color:white")
            d.exec()

    def createButtonShowBenefitsParticipateTable(self):
        self.deleteButton = QAction(self.showBenefitsParticipateTable)
        self.deleteButton.setIcon(QIcon("assests/deleteUser.png"))
        self.deleteButton.setText("حذف")
        self.deleteButton.setFont(QFont("Arial" , 12))
        self.deleteButton.triggered.connect(self.deleteBenefitsFromTrip)

        self.contextMenuShowBenefitsParticipateTable.addAction(self.deleteButton)
    def showMenuBenefitsParticipate(self,position):
        indexes = self.showBenefitsParticipateTable.selectedIndexes()
        for index in indexes:
            self.contextMenuShowBenefitsParticipateTable.exec(self.showBenefitsParticipateTable.viewport().mapToGlobal(position))

    def exportReport(self):
        try:
            self.destroyFrame(self.exportTripReportFrame)
        except:
            pass

        self.tripId = self.showAllTripsTable.item(self.showAllTripsTable.selectedIndexes()[0].row(),0).text()
        self.tripName = self.showAllTripsTable.item(self.showAllTripsTable.selectedIndexes()[0].row(),1).text()
        cr.execute("SELECT donorId, benefitsCount, lemit FROM trips WHERE id=?",([self.tripId]))
        values = cr.fetchall()[0]
        self.BenefitsCountToExport = str(int(values[1]) - int(values[2]))
        cr.execute("SELECT name FROM donars WHERE id=?",([values[0]]))
        self.donorName = cr.fetchone()[0]

        self.exportTripReportFrame = QFrame(parent=self.mainFrame)
        self.exportTripReportFrame.setGeometry((self.mainFrame.width()-160)//2,(self.mainFrame.height()-174)//2,160,147)
        self.exportTripReportFrame.setStyleSheet("background-color:white;border:2px solid black")

        closeButton = QPushButton(self.exportTripReportFrame)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:8px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 18px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(120,10,31,21)
        closeButton.clicked.connect(lambda x, frame=self.exportTripReportFrame:self.destroyFrame(frame))

        label = QLabel(parent=self.exportTripReportFrame,text="الصيغة")
        label.setStyleSheet('font: 14pt "Arial";border:none')
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setGeometry(20,50,121,20)

        self.formatComboBox = QComboBox(self.exportTripReportFrame)
        self.formatComboBox.addItems(["Word","Pdf"])
        self.formatComboBox.setGeometry(8,80,141,22)
        
        exportButton = QPushButton(parent=self.exportTripReportFrame,text="تصدير")
        exportButton.setGeometry(30,110,101,31)
        exportButton.setStyleSheet("QPushButton:hover {background-color:#c8c8c8;}")
        exportButton.clicked.connect(self.completeExportReportTrip)

        self.exportTripReportFrame.show()
    def completeExportReportTrip(self):
        filePath = QFileDialog.getExistingDirectory(self,"Select a Directory")
        if len(filePath) > 0: 
            doc = docx.Document()

            sections = doc.sections
            sections.page_height = 11.69
            sections.page_width = 8.27
            mystyle = doc.styles.add_style('mystyle', docx.enum.style.WD_STYLE_TYPE.CHARACTER)


            sections = doc.sections

            for section in sections:
                section.top_margin = docx.shared.Cm(0.3)
                section.bottom_margin = docx.shared.Cm(0.3)
                section.left_margin = docx.shared.Cm(0.3)
                section.right_margin = docx.shared.Cm(0.3)

            
            doc.add_paragraph("بسم الله الرحمن الرحيم", style="Body Text")

            doc.add_paragraph("\nكشف رحلة", style="Body Text")

            doc.add_picture("assests/logo.png",width=docx.shared.Inches(5), height=docx.shared.Inches(5))

            doc.add_paragraph("\n\n\n\n", style="Body Text")

            doc.add_paragraph(f"اسم الرحلة: {self.tripName} ", style="Body Text")
            doc.add_paragraph(f"عدد المستفيدين: {self.BenefitsCountToExport}", style="Body Text")
            doc.add_paragraph(f"المتبرع: {self.donorName}", style="Body Text")

            for p in doc.paragraphs:
                for run in p.runs:
                    run.font.size = docx.shared.Pt(24)

                pPr = p._element.get_or_add_pPr()
                bidi = OxmlElement('w:bidi')
                bidi.set(qn('w:val'), '1')
                pPr.append(bidi)
                for run in p.runs:
                    run.font.name = 'Arial'
                lang = OxmlElement('w:lang')
                lang.set(qn('w:val'), 'ar-SA')
                pPr.append(lang)

                p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

            for p in doc.paragraphs[-3:]:
                p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT

            doc.add_page_break()

            benefits_table_word_To_export = doc.add_table(rows=1,cols=4)
            benefits_table_word_To_export.style = "Table Grid"
            hdr_Cells = benefits_table_word_To_export.rows[0].cells
            hdr_Cells[3].text = "م"
            hdr_Cells[2].text = "اسم المستفيد"
            hdr_Cells[1].text = "رقم الهوية"
            hdr_Cells[0].text = "رقم الجوال"

            for cell in hdr_Cells:
                self.set_arabic_format(cell)

            b = 0
            benefitsIds = []
            cr.execute("SELECT benfit_id FROM trips_benfit WHERE trip_id = ?", [self.tripId])
            for i in cr.fetchall():
                for benfit_id in i:
                    benefitsIds.append(benfit_id)

            for benfit_id in benefitsIds:
                cr.execute("SELECT name, phone FROM benefits WHERE identy=?",[benfit_id])

                values = cr.fetchall()[0]
                b+=1   
                row_Cells = benefits_table_word_To_export.add_row().cells
                row_Cells[0].size = docx.shared.Pt(15)
                row_Cells[1].size = docx.shared.Pt(15)
                row_Cells[2].size = docx.shared.Pt(15)
                row_Cells[3].size = docx.shared.Pt(15)

                row_Cells[3].text = str(b)
                row_Cells[2].text = str(values[0])
                row_Cells[1].text = str(benfit_id)
                row_Cells[0].text = str(values[1])

                for cell in row_Cells:
                    self.set_arabic_format(cell)


            widths = (docx.shared.Inches(4),docx.shared.Inches(3),docx.shared.Inches(3),docx.shared.Inches(0.5))

            for row in benefits_table_word_To_export.rows:
                for idx, width in enumerate(widths):
                    row.cells[idx].width = width


            for row in benefits_table_word_To_export.rows:
                for cell in row.cells:
                    paragraphs = cell.paragraphs
                    for paragraph in paragraphs:
                        for run in paragraph.runs:
                            font = run.font
                            font.size= docx.shared.Pt(17)


            doc.save(f"{filePath}\كشف رحلة {self.tripName}.docx")

        if self.formatComboBox.currentText() == "Pdf":
            with suppress_output():
                convert(f"{filePath}\كشف رحلة {self.tripName}.docx",f"{filePath}\كشف رحلة {self.tripName}.pdf")
            os.remove(f"{filePath}\كشف رحلة {self.tripName}.docx")

        message = QMessageBox(parent=self,text="تم التصدير بنجاح")
        message.setIcon(QMessageBox.Icon.Information)
        message.setWindowTitle("نجاح")
        message.exec()
    def showMenuTripsAllTable(self,position):
        indexes = self.showAllTripsTable.selectedIndexes()
        for index in indexes:
            self.contextMenuTripsTableAll.exec(self.showAllTripsTable.viewport().mapToGlobal(position))
    def deleteTrip(self):
        tripId = self.showAllTripsTable.item(self.showAllTripsTable.selectedIndexes()[0].row(),0).text()
        d = QMessageBox(parent=self,text=f"تأكيد حذف {self.showAllTripsTable.item(self.showAllTripsTable.selectedIndexes()[0].row(),1).text()}")
        d.setIcon(QMessageBox.Icon.Information)
        d.setWindowTitle("تأكيد")
        d.setStyleSheet("background-color:white")
        d.setStandardButtons(QMessageBox.StandardButton.Cancel|QMessageBox.StandardButton.Ok)
        important = d.exec()
        if important == QMessageBox.StandardButton.Ok:

            cr.execute("SELECT donorId, amount FROM trips WHERE id=? ",[tripId])
            val = cr.fetchall()[0]
            cr.execute(f"UPDATE donars SET amount = amount - {val[1]}, countOfTrips = countOfTrips - 1 WHERE id = ?",[val[0]])
            cr.execute("DELETE FROM end_date WHERE trip_id = ?",[tripId])
            cr.execute("DELETE FROM trips_benfit WHERE trip_id = ?",[tripId])
            cr.execute("DELETE FROM trips WHERE id = ?",[tripId])

            con.commit()

            self.loadTrips()
            d = QMessageBox(parent=self,text="تم الحذف بنجاح")
            d.setWindowTitle("نجاح")
            d.setIcon(QMessageBox.Icon.Information)
            d.setStyleSheet("background-color:white")
            d.exec()
    def editTrip(self):
        self.TripIdEdit = self.showAllTripsTable.item(self.showAllTripsTable.selectedIndexes()[0].row(),0).text()
        try:
            self.editTripFrame.deleteLater()
        except:
            pass
        self.editTripFrame = QFrame(self.mainFrame)
        self.editTripFrame.setGeometry((self.mainFrame.width()-350)//2,(self.mainFrame.height()-350)//2,350,350)
        self.editTripFrame.setStyleSheet("background-color:white;border:2px solid black")

        closeButton = QPushButton(self.editTripFrame)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:12px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 26px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(300,10,41,31)
        closeButton.clicked.connect(lambda x, frame=self.editTripFrame:self.destroyFrame(frame))
        
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

        self.scroolAria = QScrollArea(self.editTripFrame)
        self.scroolAria.setWidget(self.frame)
        self.scroolAria.setStyleSheet("border:1px solid gray")
        self.scroolAria.move(20,50)
        self.scroolAria.resize(321,270)

        self.scroolAria.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)


        self.scroolAria.setWidgetResizable(True)

        #End scrolAria

        editButton = QPushButton(text="تعديل")
        editButton.clicked.connect(self.completeEditTrip)
        editButton.setStyleSheet("QPushButton:hover {background-color:#c8c8c8;}")        

        layout.addWidget(editButton)
        #End scrolAria

        cr.execute("SELECT name, benefitsCount, begin_date, begin_date_hijri, amount FROM trips WHERE id=?", [self.TripIdEdit])
        values = cr.fetchall()[0]
        
        self.tripName.setText(values[0])
        self.BenefitsCount.setText(str(values[1]))
        self.dateHigryEntry.setText(values[3])
        tempDateVar = values[2].split("/")
        self.dateEntry.setDate(QDate(int(tempDateVar[0]), int(tempDateVar[1]), int(tempDateVar[2])))
        self.tripAmount.setText(str(values[4]))

        self.changeHijryDate()
        self.editTripFrame.show()
    def completeEditTrip(self):
        cr.execute("SELECT benefitsCount,lemit FROM trips WHERE id=?", [self.TripIdEdit])
        values = cr.fetchall()[0]
        joinedBenefits = int(values[0] - values[1])
        if(joinedBenefits > int(self.BenefitsCount.text())):
            message = QMessageBox(parent=self,text="عذرا المستفيدين المشاركين في الرحلة اكثر من رقم المستفيدين الجديد الذي ادخلته")
            message.setIcon(QMessageBox.Icon.Critical)
            message.setWindowTitle("فشل")
            message.exec()
        else:
            if len(self.tripName.text()) > 0 and len(self.BenefitsCount.text()) > 0 and len(self.dateEntry.text()) and len(self.dateHigryEntry.text()) > 0 and len(self.tripAmount.text()) > 0:
                cr.execute("SELECT donorId, amount FROM trips WHERE id=?", [self.TripIdEdit])
                values = cr.fetchall()[0]
                diffrence = int(self.tripAmount.text()) - values[1]
                cr.execute("UPDATE trips SET name=?, benefitsCount=?, lemit=?, begin_date=?, begin_date_hijri=?, amount=? WHERE id=? ", (self.tripName.text(), self.BenefitsCount.text(),int(self.BenefitsCount.text()) - joinedBenefits, self.dateEntry.text(), self.dateHigryEntry.text(), self.tripAmount.text(), self.TripIdEdit))
                cr.execute("UPDATE donars SET amount = amount + ? WHERE id = ?",(diffrence, values[0]))
                con.commit()
                message = QMessageBox(parent=self,text="تم تعديل معلومات الرحلة بنجاح")
                message.setIcon(QMessageBox.Icon.Information)
                message.setWindowTitle("نجاح")
                message.exec()
                self.loadTrips()
    def showProgrammer(self):
        try:
            self.destroyFrame(self.showProgrammerFrame)
        except:
            pass

        self.showProgrammerFrame = QFrame(parent=self.mainFrame)
        self.showProgrammerFrame.setGeometry((self.mainFrame.width()-273)//2,(self.mainFrame.height()-368)//2,273,368)
        self.showProgrammerFrame.setStyleSheet("background-color:white")


        
        logo = QLabel(parent=self.showProgrammerFrame,text="")

        closeButton = QPushButton(logo)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:8px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 18px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(360,0,31,31)
        closeButton.clicked.connect(lambda x, frame=self.showProgrammerFrame:self.destroyFrame(frame))

        logo.setGeometry(-139,10,571,281)
        logo.setPixmap(QPixmap("assests/MyLogo.png"))
        logo.setScaledContents(True)

        label = QLabel(parent=self.showProgrammerFrame,text="برمجة وتطوير: م.عبدالله ماهر الشامي")
        label.setStyleSheet('font: 12pt "Arial";color:rgb(255, 0, 0);')
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setGeometry(30,250,201,20)

        label = QLabel(parent=self.showProgrammerFrame,text="للتواصل: 966558967920+")
        label.setStyleSheet('font: 12pt "Arial";color:rgb(255, 0, 0);')
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setGeometry(50,280,181,20)

        label = QLabel(parent=self.showProgrammerFrame,text="جميع الحقوق محفوظة")
        label.setStyleSheet('font: 12pt "Arial";color:rgb(255, 0, 0);')
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setGeometry(10,340,261,20)

        self.showProgrammerFrame.show()
    def showStats(self):
        try:
            self.destroyFrame(self.showStatsFrame)
        except:
            pass

        self.showStatsFrame = QFrame(parent=self.mainFrame)
        self.showStatsFrame.setGeometry((self.mainFrame.width()-273)//2,(self.mainFrame.height()-354)//2,273,354)
        self.showStatsFrame.setStyleSheet("background-color:white")


        logo = QLabel(parent=self.showStatsFrame,text="")

        closeButton = QPushButton(logo)
        closeButton.setStyleSheet("QPushButton {border:2px solid black;font-size:8px;qproperty-icon: url('assests/close.png');qproperty-iconSize: 18px}QPushButton:hover {background-color:#c8c8c8;}")        
        closeButton.setGeometry(250,30,31,31)
        closeButton.clicked.connect(lambda x, frame=self.showStatsFrame:self.destroyFrame(frame))

        logo.setGeometry(-20,-10,305,281)
        logo.setPixmap(QPixmap("assests/logo.png"))
        logo.setScaledContents(True)
        
        cr.execute("SELECT name FROM donars")
        donarsCount = len(cr.fetchall())

        cr.execute("SELECT name FROM benefits")
        benefitsCount = len(cr.fetchall())

        donationsMoney = 0
        cr.execute("SELECT amount FROM donars")
        for i in cr.fetchall():
            for j in i:
                donationsMoney+=j

        cr.execute("SELECT name FROM trips")
        tripsCount = len(cr.fetchall())

        label = QLabel(parent=self.showStatsFrame,text=f"عدد المتبرعين: {donarsCount}")
        label.setStyleSheet('font: 12pt "Arial";color:rgb(255, 0, 0);')
        label.setGeometry(30,250,201,20)

        label = QLabel(parent=self.showStatsFrame,text=f"عدد المستفيدين: {benefitsCount}")
        label.setStyleSheet('font: 12pt "Arial";color:rgb(255, 0, 0);')
        label.setGeometry(50,274,181,21)

        label = QLabel(parent=self.showStatsFrame,text=f"اجمالي مبلغ التبرع: {donationsMoney}")
        label.setStyleSheet('font: 12pt "Arial";color:rgb(255, 0, 0);')
        label.setGeometry(50,299,181,20)

        label = QLabel(parent=self.showStatsFrame,text=f"اجمالي عدد الرحل: {tripsCount}")
        label.setStyleSheet('font: 12pt "Arial";color:rgb(255, 0, 0);')
        label.setGeometry(50,320,181,20)

        self.showStatsFrame.show()