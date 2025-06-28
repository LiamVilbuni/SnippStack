import sys, json, os, keyboard
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QListWidget, QLineEdit, QMessageBox, QFileDialog
from PyQt6.QtGui import QIcon

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STORAGE = os.path.join(BASE_DIR, "snipps.json")
ICON = os.path.join(BASE_DIR, "icon.ico")

try:
    from ctypes import windll
    myappid = 'InviBull.NovaObscura.SnippStack.1'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass

class SnippStack(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SnippStack")
        self.setWindowIcon(QIcon(ICON))
        self.data = self.loadSnippets(STORAGE)

        self.keyList = QListWidget()
        self.valueList = QListWidget()

        self.keyInput = QLineEdit()
        self.valueInput = QLineEdit()
        self.keyInput.setPlaceholderText("Enter snippet")
        self.valueInput.setPlaceholderText("Enter content")

        self.addButton = QPushButton("Add")
        self.delButton = QPushButton("Delete")
        self.editButton = QPushButton("Edit")
        self.backupButton = QPushButton("Backup")

        mainLayout = QVBoxLayout()
        listLayout = QHBoxLayout()
        inputLayout = QHBoxLayout()
        buttonLayout = QHBoxLayout()

        listLayout.addWidget(self.keyList)
        listLayout.addWidget(self.valueList)

        inputLayout.addWidget(self.keyInput)
        inputLayout.addWidget(self.valueInput)

        buttonLayout.addWidget(self.addButton)
        buttonLayout.addWidget(self.editButton)
        buttonLayout.addWidget(self.delButton)
        buttonLayout.addWidget(self.backupButton)

        mainLayout.addLayout(listLayout)
        mainLayout.addLayout(inputLayout)
        mainLayout.addLayout(buttonLayout)
        self.setLayout(mainLayout)

        self.addButton.clicked.connect(self.addSnippet)
        self.delButton.clicked.connect(self.delSnippet)
        self.editButton.clicked.connect(self.editSnippet)
        self.backupButton.clicked.connect(self.backupSnippets)
        self.keyList.currentRowChanged.connect(self.syncSelection)

        self.refreshLists()
        self.clearInputs()

    def loadSnippets(self, filename):
        if os.path.exists(filename):
            try:
                with open(filename, 'r') as f:
                    return json.load(f)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load {filename}:\n{e}")
        return {}

    def dumpSnippets(self):
        try:
            with open(STORAGE, 'w') as f:
                json.dump(self.data, f, indent=4)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save to {STORAGE}:\n{e}")

    def clearInputs(self):
        self.keyInput.clear()
        self.valueInput.clear()
        self.keyList.clearSelection()
        self.valueList.clearSelection()

    def syncSelection(self, row):
        if row >= 0:
            self.valueList.setCurrentRow(row)
            key = self.keyList.item(row).text()
            self.keyInput.setText(key)
            self.valueInput.setText(self.data.get(key, ""))

    def refreshLists(self):
        self.keyList.clear()
        self.valueList.clear()

        for key, value in self.data.items():
            self.keyList.addItem(key)
            self.valueList.addItem(value)

        self.keyList.clearSelection()
        self.valueList.clearSelection()

        self.refreshSnippets()

    def refreshSnippets(self):
        keyboard.unhook_all()

        for key, value in self.data.items():
            keyboard.add_abbreviation(key, value)

    def addSnippet(self):
        key = self.keyInput.text().strip()
        value = self.valueInput.text().strip()

        if not key or not value:
            QMessageBox.warning(self, "Input Error", "Snippet or content cannot be empty.")
            return
        if key in self.data:
            QMessageBox.warning(self, "Duplicate Snippet", f"The snippet '{key}' already exists.")
            return

        self.data[key] = value
        self.dumpSnippets()
        self.clearInputs()
        self.refreshLists()

    def delSnippet(self):
        row = self.keyList.currentRow()
        if row >= 0:
            key = self.keyList.item(row).text()
            del self.data[key]
            self.dumpSnippets()
            self.clearInputs()
            self.refreshLists()

    def editSnippet(self):
        row = self.keyList.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Selection Error", "Please select a snippet to edit.")
            return

        oldKey = self.keyList.item(row).text()
        newKey = self.keyInput.text().strip()
        newVal = self.valueInput.text().strip()

        if not newKey or not newVal:
            QMessageBox.warning(self, "Input Error", "Snippet or content cannot be empty.")
            return

        if newKey != oldKey and newKey in self.data:
            QMessageBox.warning(self, "Duplicate Key", f"The snippet '{newKey}' already exists.")
            return

        if oldKey != newKey:
            del self.data[oldKey]
        self.data[newKey] = newVal
        self.dumpSnippets()
        self.clearInputs()
        self.refreshLists()

    def backupSnippets(self):
        filename, _ = QFileDialog.getSaveFileName(self, "Save Backup", "", "Text Files (*.txt)")
        if filename:
            try:
                with open(filename, "w") as f:
                    json.dump(self.data, f, indent=4)
                QMessageBox.information(self, "Backup Saved", f"Backup saved to {filename}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save backup:\n{e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    WN = SnippStack()
    WN.resize(600, 400)
    WN.show()
    sys.exit(app.exec())
