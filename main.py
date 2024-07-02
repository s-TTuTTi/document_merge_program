import src.gui_document.handle_gui as handle_gui


class App:
    def __init__(self):
        self.gui = handle_gui.Gui()

    def run(self):
        self.gui.initialize_ui()
        self.gui.root.mainloop()

if __name__ == '__main__':
    app = App()
    app.run()


