# imports
from flexx import app, event, ui, flx, config


## These two lines are  UNCOMMENTED when I run
# config.hostname = 'myhost'            # publish flexx app on hostname not just localhost
# config.port = 12345                   # serve flexx app on this port

class FileViewer(ui.FileBrowserWidget):
    print(f'DEBUG: Instantiated FileViewer instance')

    @flx.reaction('selected')
    def fileSelected(self, *events):
        print(f'DEBUG: We are in fileSelected reaction')
        ev = events[-1]  # only care about last event
        print(f'DEBUG: File {ev.filename} selected in FileViewer')
        try:
            f = open(ev.filename, 'r')
            if f.mode == 'r':
                contents = f.read()
                print(f'DEBUG: file content is: {contents}')
        except:
            print(f'ERROR: unable to open file {ev.filename} for reading')


if __name__ == '__main__':
    app = flx.App(FileViewer)
    app.serve('FileViewer')
    flx.start()
