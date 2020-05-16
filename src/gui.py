from flexx import flx, event


class ThemedForm(flx.Widget):
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
    """

    def init(self):
        with flx.HFix():
            with flx.FormLayout() as self.form:
                self.b1 = flx.LineEdit(title='Name:', text='')
                self.button_open = flx.Button(text='Open')
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
        #if self.combo.selected_index is not None:
            #text += ' (index %i)' % self.combo.selected_index
        self.label2.set_text(text)

    @flx.reaction('button_open.pointer_click')
    def _button_clicked(self, *events):
        ev = events[-1]
        text = self.label.text + ev.source.text
        self.label.set_text(text)

if __name__ == '__main__':
    m = flx.launch(ThemedForm, 'app')
    flx.run()
