import tkinter as tk
from tkinter import ttk
from requests import Session
from bs4 import BeautifulSoup
from time import sleep
import xlsxwriter


class App(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.geometry('400x200')
        self.resizable(False, False)
        self.title('Parser of marks')
        self.set_ui()

    def set_ui(self):
        self.login_frame = ttk.LabelFrame(self)
        self.login_frame.pack(fill=tk.X)
        ttk.Label(self.login_frame, text='Login').pack(side=tk.TOP, anchor='w')
        self.login_entry = ttk.Entry(
            self.login_frame, justify=tk.LEFT)
        self.login_entry.pack(side=tk.LEFT)

        self.password_frame = ttk.LabelFrame(self)
        self.password_frame.pack(fill=tk.X)
        ttk.Label(self.password_frame, text='Password').pack(
            side=tk.TOP, anchor='w')
        self.password_entry = ttk.Entry(
            self.password_frame, justify=tk.LEFT)
        self.password_entry.pack(side=tk.LEFT)
        ttk.Button(self, text='Parse', command=self.go).pack()

    def go(self):
        log = self.login_entry.get()
        passwd = self.password_entry.get()
        self.conn(login=log, password=passwd)
        self.writer()

    def conn(self, login, password):
        header = {
            'UserAgent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 \
                (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36'}

        work = Session()
        work.get('https://www.google.com/search?q=triton+knu&rlz=1C1CHBD_ruUA1002UA1002&oq=tr&aqs=chrome.0.69i59l2j69i57j0i131i433i512l2j69i60l3.518j0j7&sourceid=chrome&ie=UTF-8',
                 headers=header)

        response = work.get('https://student.triton.knu.ua/', headers=header)
        soup = BeautifulSoup(response.text, 'lxml')

        token = soup.find('div', class_='well bs-component').find(
            'form').find_all('input')[-2].get('value')

        data = {'Login': f'{login}', 'Password': f'{password}',
                '__RequestVerificationToken': f'{token}'}

        logined = work.post('https://student.triton.knu.ua/',
                            headers=header, data=data, allow_redirects=True)

        marks = work.get(
            'https://student.triton.knu.ua/Study/Marks', headers=header)
        sp = BeautifulSoup(marks.text, 'lxml')

        self.mark_table = sp.find_all('tr', class_='success')

    def make_array(self):
        for row in self.mark_table:
            mark = row.text
            temp_arr = mark.split('\n')
            temp_arr.pop(-1)
            temp_arr.pop(0)
            yield temp_arr

    def writer(self):
        book = xlsxwriter.Workbook(r'.\marks.xlsx')
        page = book.add_worksheet('marks')

        row = 0
        column = 0

        page.set_column('A:A', 20)
        page.set_column('B:B', 20)
        page.set_column('C:C', 20)
        page.set_column('D:D', 20)
        page.set_column('E:E', 20)
        page.set_column('F:F', 20)
        page.set_column('G:G', 20)
        page.set_column('H:H', 20)

        for item in self.make_array():
            page.write(row, column, item[0])
            page.write(row, column+1, item[1])
            page.write(row, column+2, item[2])
            page.write(row, column+3, item[3])
            page.write(row, column+4, item[4])
            page.write(row, column+5, item[5])
            page.write(row, column+6, item[6])
            page.write(row, column+7, item[7])
            row += 1
        book.close()


if __name__ == '__main__':
    root = App()
    root.mainloop()
