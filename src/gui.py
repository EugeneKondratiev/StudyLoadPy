from flexx import flx, event, ui
import tkinter as tk
from tkinter import filedialog

from src.FileViewer import FileViewer


class StudyLoadPy(flx.Widget):
    CSS = """
    .flx-Button {
        margin-left: 25px;
        width: 150px;
        background: #9d9;
        border-radius: 10px;
    }
    .flx-LineEdit {
        margin-left: 25px;
        width: 150px;
        border: 2px solid #9d9;
        border-radius: 10px;
    }
    .flx-ComboBox {
        margin-left: 25px;
        width: 150px;
        border-radius: 10px;
    }
    .flx-Label {
        margin-top: 30px;
    }
    .flx-title {
        margin-left: 25px;
    }
    """

    def init(self):
        with flx.HFix():
            with flx.FormLayout() as self.form:
                self.b1 = flx.LineEdit(title='Name:', text='')
                self.button_open = flx.Button(text='Open')
                #filedialog.askopenfile()
                types_of_parse = ['Денна 1 сем', 'Денна 2 сем', 'Заочна 1 сем', 'Заочна 2 сем']
                self.combo = flx.ComboBox(editable=False, options=types_of_parse)
                self.button_submit = flx.Button(text='Submit')
            with flx.FormLayout() as self.form:
                self.label = flx.Label()
                self.label2 = flx.Label()
                flx.Widget(flex=1)  # Add a spacer

    @event.reaction
    def update_label(self):
        text = self.combo.text
        self.label2.set_text(text)

    @flx.reaction('button_open.pointer_click')
    def _button_clicked(self, *events):
        ev = events[-1]
        f = FileViewer(ui.FileBrowserWidget)
        text = self.label.text + ev.source.text
        self.label.set_text(text)


if __name__ == '__main__':
    app = flx.App(StudyLoadPy)
    app.serve('StudyLoadPy')
    flx.run()
