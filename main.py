import tkinter as tk
from main_window import Application
import logging
from logging_settings import logging_start

logging_start()
logger = logging.getLogger('logging_settings').getChild('main')

def main():
    logger.info("Program Start")
    window = tk.Tk()
    app = Application(master=window)
    app.mainloop()

if __name__ == "__main__":
    main()