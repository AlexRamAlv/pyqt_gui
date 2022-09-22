"""
    To make an installer we have to install pyinstaller library.
    Then we run on terminal in the exact location of the application you want to make the installer
    the following commands depending on what you want or need.

    pyinstaller --windowed --onefile --icon=./path Appname 

"""

from datetime import datetime
import sys
from threading import Timer
from PyQt5.QtWidgets import QApplication, QDialog, QDesktopWidget, QWidget
from PyQt5.QtCore import QPoint

from greenProgressBar import *
from findFPWindow import *
import win32com.client as win32
from pathlib import Path


class OutlookError(Exception):
    pass


style_able = (
    "QLineEdit{\n"
    "    padding: 4px 10px;\n"
    "    font-size: 12px;\n"
    "    font-weight: 400;\n"
    "    line-height: 1.5px;\n"
    "    color: #212529;\n"
    "    background-color: #fff;\n"
    "    background-clip: padding-box;\n"
    "    border: 1px solid #ced4da;\n"
    "    border-radius: 6px;\n"
    "}"
)

style_disable = (
    "QLineEdit{\n"
    "    padding: 4px 10px;\n"
    "    font-size: 12px;\n"
    "    font-weight: 400;\n"
    "    line-height: 1.5px;\n"
    "    color: #adadad;\n"
    "    background-color: #fff;\n"
    "    background-clip: padding-box;\n"
    "    border: 1px solid #ced4da;\n"
    "    border-radius: 6px;\n"
    "}"
)


class ProgressBar(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.ui = Ui_Form()
        self.ui.setupUi(self)


class TheForm(QDialog):
    def __init__(self) -> None:
        # cwd = Path.cwd()
        # file_path = cwd / "files/path.txt"
        # file_path.touch(exist_ok=True)

        # file_mails = cwd / "files/mails.txt"
        # file_mails.touch(exist_ok=True)

        super().__init__()
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
        self.ui.pushButton_2.clicked.connect(lambda: self.close())
        self.ui.pushButton_3.clicked.connect(lambda: self.showMinimized())

        # QLineEdit 2
        with open(file="Find FP\\files\\path.txt", mode="r", encoding="utf-8") as f:
            self.path = f.read().strip()

        self.ui.lineEdit_2.setText(self.path)
        self.ui.lineEdit_2.setReadOnly(True)
        self.ui.lineEdit_2.setStyleSheet(style_disable)
        self.ui.checkBox.clicked.connect(self.modify_path)

        # QLineEdit 3
        with open(file="Find FP\\files\\mails.txt", mode="r", encoding="utf-8") as f:
            self.mails = f.read().strip()

        self.ui.lineEdit_3.setText(self.mails)
        self.ui.lineEdit_3.setReadOnly(True)
        self.ui.lineEdit_3.setStyleSheet(style_disable)
        self.ui.checkBox_2.stateChanged.connect(self.modify_mails)

        # Find button
        self.ui.pushButton.clicked.connect(self.find_fotopolimer_item)

    # Draggable window
    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def mousePressEvent(self, event):
        self.oldPos = event.globalPos()

    def mouseMoveEvent(self, event):
        delta = QPoint(event.globalPos() - self.oldPos)
        self.move(self.x() + delta.x(), self.y() + delta.y())
        self.oldPos = event.globalPos()

    def modify_path(self):
        if self.ui.checkBox.isChecked():
            self.ui.lineEdit_2.setStyleSheet(style_able)
            self.ui.lineEdit_2.setReadOnly(False)
        else:
            self.ui.lineEdit_2.setReadOnly(True)
            self.ui.lineEdit_2.setStyleSheet(style_disable)

    def modify_mails(self):
        if self.ui.checkBox_2.isChecked():
            self.ui.lineEdit_3.setStyleSheet(style_able)
            self.ui.lineEdit_3.setReadOnly(False)
        else:
            self.ui.lineEdit_3.setReadOnly(True)
            self.ui.lineEdit_3.setStyleSheet(style_disable)

    def find_fotopolimer_item(self):
        if (
            self.ui.lineEdit.text() == ""
            or self.ui.lineEdit_2.text() == ""
            or self.ui.lineEdit_3.text() == ""
        ):
            msg = Timer(
                0.1,
                self.ui.label_5.setText,
                ("Asegúrate que no haya campos vacíos",),
            )
            blank = Timer(5.0, self.ui.label_5.setText, ("",))
            msg.start()
            blank.start()
        else:
            outlook = win32.Dispatch("outlook.application")
            bar = ProgressBar()
            items = self.ui.lineEdit.text().strip().upper()
            if "," in items:
                fp_to_be_found = items.split(",")

            else:
                fp_to_be_found = items.split()

            try:
                mail = outlook.CreateItem(0)

            except Exception as e:
                raise OutlookError(f"No se pudo crear instancia de Outlook: {str(e)}")

            mail.To = self.ui.lineEdit_3.text()
            mail.CC = "ancruz@rowe.com.do"
            mail.Subject = f"Solicitud de cotización Fotopolímeros {datetime.now().strftime('%d %b %Y')}"
            mail.HTMLBody = """
                <p style="font-family:Tahoma; font-size:16px">
                    Estimados, gusto en saludarles.
                    Favor colaborarnos con una cotización para el(los) siguiente(s) fotopolímero(s) que dejo en adjunto.
                    <br/><br/>
                    Agradezco de antemano la acostumbrada colaboración.<br/>
                    Quedo atento.<br/>
                    Cordial saludo.<br/>
                </p>
            """
            path_for_finding = Path(self.ui.lineEdit_2.text())

            if self.ui.checkBox.isChecked() and self.ui.lineEdit_2.text() != self.path:
                with open(
                    file="Find FP\\files\\path.txt", mode="w", encoding="utf-8"
                ) as f:
                    f.write(self.ui.lineEdit_2.text())

            if (
                self.ui.checkBox_2.isChecked()
                and self.ui.lineEdit_3.text() != self.mails
            ):
                with open(
                    file="Find FP\\files\\mails.txt", mode="w", encoding="utf-8"
                ) as f:
                    f.write(self.ui.lineEdit_3.text())

            # FP00225, FP00226CM, FP01122, FP01239
            errors = []
            counter = 0
            error = 0
            bar.show()
            for i in fp_to_be_found:
                try:
                    error += 1
                    attachment_path = list(
                        path_for_finding.rglob(f"{ i.strip() }*.ai")
                    )[0]

                    print(attachment_path)
                    mail.Attachments.Add(Source=str(attachment_path))
                    counter += 1
                    bar.ui.progressBar.setValue(
                        int(counter / len(fp_to_be_found) * 100)
                    )
                except IndexError:
                    error = error
                    errors.append(error)

            if errors:
                msg_1 = Timer(
                    0.1,
                    self.ui.label_5.setText,
                    (f"Hay errores en el(los) código(s) número(s) { errors }!",),
                )
                blank_1 = Timer(5.0, self.ui.label_5.setText, ("",))
                msg_1.start()
                blank_1.start()
                print(f"Hay errores en el(los) código(s) número(s) { errors }!")
            if not counter < 1:
                mail.Display()

            msg_2 = Timer(
                0.1,
                self.ui.label_6.setText,
                (
                    f"{ counter } código(s) de { len(fp_to_be_found) } fue(ron) encontrado(s)!",
                ),
            )
            blank_2 = Timer(5.0, self.ui.label_6.setText, ("",))
            msg_2.start()
            blank_2.start()
            print(
                f"{ counter } código(s) de { len(fp_to_be_found) } fue(ron) encontrado(s)!"
            )


if __name__ == "__main__":
    app = QApplication(sys.argv)

    form = TheForm()
    form.show()

    sys.exit(app.exec_())
