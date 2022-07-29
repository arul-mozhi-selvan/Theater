import openpyxl
import tkinter
from functools import partial

class get_from_user():

    def __init__(self):
        # INSTRUCTIONS FOR FRONT END:
        # call get_name function
        self.get_name()
        self.selected_screen = None
        self.show = None
        self.getscreen_and_time()
        # enter  screen button code here, button command should be
        # select_screen1 for screen 1 button and so on

    def getscreen_and_time(self):
        def prt():
            self.selected_screen = screen.get()
            self.show = show_time.get()
            y.destroy()

        y = tkinter.Tk()
        y.geometry('760x560')
        sc_opt = ['Screen1_', 'Screen2_']
        screen = tkinter.StringVar()
        sc_drop = tkinter.OptionMenu(y, screen, *sc_opt)
        sc_drop.pack()

        sh_opt = ['10am', '7pm']
        show_time = tkinter.StringVar()
        sh_drop = tkinter.OptionMenu(y, show_time, *sh_opt)
        sh_drop.pack()

        b = tkinter.Button(y, text='See seat availability', command=prt)
        b.pack()
        y.mainloop()


    def get_name(self):
        self.username = None  # default till front end is filled
        y = tkinter.Tk()
        l = tkinter.Label(text = "Enter username")
        l.pack()
        def get_text():
            self.username = e.get()
            y.destroy()
        e = tkinter.Entry()
        e.pack()
        b = tkinter.Button(text='submit', command=get_text)
        b.pack()
        y.mainloop()

    def generate_ID(self):
        self.bookingID = 1
        # add expression to get it from the workbook and increment 1

    def print_ticket(self):
        pass

    def get_seats(self, filled):
    # call front end function that selects seats
    # should return a list of selected seats
        sh1 = tkinter.Tk()
        sh1.geometry('700x650')
        b = ['A', 'B', 'C', 'D', 'E']
        yc = 100


        self.l = []

        def book(str1):
            self.l.append(str1)
            print(self.l)

        for i in range(5):
            for j in range(10):
                if j < 5:
                    if b[i] + str(j + 1) in filled:
                        a = tkinter.Button(sh1, text=b[i] + str(j + 1), bg='red', state = tkinter.DISABLED)
                        a.place(x=100 + (j * 50), y=yc)
                    else:
                        a = tkinter.Button(sh1, text=b[i] + str(j + 1), bg='green',
                                            command=partial(book, (b[i] + str(j + 1))))
                        a.place(x=100 + (j * 50), y=yc)
                else:
                    if b[i] + str(j + 1) in filled:
                        a = tkinter.Button(sh1, text=b[i] + str(j + 1), bg='red',state = tkinter.DISABLED)
                        a.place(x=180 + (j * 50), y=yc)
                    else:
                        a = tkinter.Button(sh1, text=b[i] + str(j + 1), bg='green',
                                            command=partial(book, (b[i] + str(j + 1))))
                        a.place(x=180 + (j * 50), y=yc)
            yc = yc + 50
        confirm = tkinter.Button(sh1, text='Book', command = sh1.destroy)
        confirm.place(x = 200, y = yc+100)
        eyes = '''____________________________________________________________
                    SCREEN'''
        la = tkinter.Label(sh1, text=eyes)
        la.place(x=240, y=400)
        sh1.mainloop()
        return self.l

    def save_data(self):
        pass


ticket = get_from_user()
#defining the path of the XL file to access
screen_path = "C:\\Users\\arulo\\Documents\\Python\\Theatre_project\\Theater_Project\\" + ticket.selected_screen + ticket.show + ".xlsx"

#get seat list from front end
seats = []

filename = ticket.selected_screen + ticket.show + ".xlsx"

print(screen_path)

#opening the required XL file
workbook = openpyxl.load_workbook(screen_path)
worksheet = workbook["Sheet1"]

#list of filled seats
filled = []
for row in worksheet:
    for cell in row:
        if cell.value == True:
            filled.append(cell.coordinate)

print (filled)
#seats = ticket.get_seats(filled)
seats = ticket.get_seats(filled)
#add new seats to the data base
for seat in seats:
    worksheet[seat] = True

customer_data = openpyxl.load_workbook("C:\\Users\\arulo\\Documents\\Python\\Theatre_project\\Theater_Project\\Customer_database.xlsx")

print(ticket.username)
workbook.save(filename)

