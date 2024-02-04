import pyodbc, os, shutil, sys, threading
from pyodbc import pooling
from threading import Thread
from multiprocessing import Process
from datetime import *
from tkinter import *
from tkinter import scrolledtext, filedialog, Tk
from tkinter.ttk import Combobox
from tkcalendar import Calendar, DateEntry
import datetime as dt, csv
import paramiko, time, logging
import babel
from babel import numbers

SQL = 'C:\\Python\\SQL\\'
RES = 'C:\\Python\\Output\\'
LOG = 'C:\\Python\\Output\\logs\\'
datasource_timeout = 300
d = dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
d2 = dt.datetime.now().strftime('%H:%M:%S')
d3 = dt.datetime.now().strftime('%Y_%m_%d_%H_%M_%S')
d4 = dt.datetime.now().strftime('%Y%m%d%H%M%S')
version = os.path.basename(__file__)

#-1
def Datasource(dsn_id, dsn_user, dsn_pwd):

    start_time = time.time()
    #conn = None
    while True:
        try:
            conn = pyodbc.connect(dsn=dsn_id, uid=dsn_user, pwd=dsn_pwd)
            return conn
        except pyodbc.Error as pye:
            pye_error = "Error connecting to database, DSN license unavailable, retrying "
            screen_logger(pye_error)
            print("Error connecting to database, DSN license unavailable")
            pass
        if time.time() - start_time > datasource_timeout:
            print(' DSN Timeout set at ' + datasource_timeout + ' seconds ' )
            break
        else:
            time.sleep(10)
    if conn:
        print("Connected to DSN")

def main_account_screen():
    global main_screen
    main_screen = Tk()
    main_screen_width = 450
    main_screen_height = 300
    computer_screen_width = main_screen.winfo_screenwidth() #1600
    computer_screen_height = main_screen.winfo_screenheight() #900
    x = (computer_screen_width/2) -(main_screen_width/2)
    y = (computer_screen_height/2) - (main_screen_height/2)
    main_screen.geometry(f'{main_screen_width}x{main_screen_height}+{int(x)}+{int(y)}')
    main_screen.iconbitmap('c:\python\sql\hiport.ico')
    main_screen.title(" StatPro Reports " + "Version : " + version)
    Label(text=" HiPortfolio DRO Reports ", width="250", height="2", font=("Calibri", 12)).pack()
    Label(text="").pack()
    ### call login()
    Button(text="Login", height="2", width="30", bg ='grey' , command=login).pack()
    Label(text="").pack()
    ### call delete_main_screen
    Button(text="Exit", height="2", width="30", bg='red', command=delete_main_screen).pack()
    Label(text="").pack()
    dte = dt.datetime.now()
    fte = f"{dte:%a, %b %d, %Y}"
    Label(text=fte).pack()
    main_screen.mainloop()

#1
def login():

    global login_screen

    login_screen = Tk() #Toplevel(main_screen)
    login_screen_width = 550
    login_screen_height = 550
    computer_screen_width = main_screen.winfo_screenwidth()  # 1600
    computer_screen_height = main_screen.winfo_screenheight()  # 900

    x = (computer_screen_width / 2) - (login_screen_width / 2)
    y = (computer_screen_height / 2) - (login_screen_height / 2)
    login_screen.geometry(f'{login_screen_width}x{login_screen_height}+{int(x)}+{int(y)}')
    login_screen.iconbitmap('c:\python\sql\hiport.ico')
    login_screen.title("Login")

    global username_verify
    global password_verify
    global username_login_entry
    global password_login_entry
    global dsn_username_verify
    global dsn_password_verify

    def limitusername_verify(*args):
        value = username_verify.get()
        if len(value) > 2: username_verify.set(value[:6])

    def limit_password_verify(*args):
        value = password_verify.get()
        if len(value) > 2: password_verify.set(value[:6])

    username_verify = StringVar(login_screen,value='ADMIN')
    username_verify.trace('w', limitusername_verify)

    password_verify = StringVar(login_screen,value='ADMIN')
    password_verify.trace('w', limit_password_verify)


    ###############################################################################
    def dsn_limitusername_verify(*args):
        value = dsn_username_verify.get()
        if len(value) > 2: dsn_username_verify.set(value[:6])

    dsn_username_verify = StringVar(login_screen, value='ADMIN')
    dsn_username_verify.trace('w', dsn_limitusername_verify)

    def dsn_limit_password_verify(*args):
        value = dsn_password_verify.get()
        if len(value) > 2: dsn_password_verify.set(value[:8])

    dsn_password_verify = StringVar(login_screen, value='')
    dsn_password_verify.trace('w', dsn_limit_password_verify)
    Label(login_screen, text="").pack()

    ##############################################################################################################

    Label(login_screen, text="").pack()

    Label(login_screen, text="Username * ").pack()
    username_login_entry = Entry(login_screen, textvariable=username_verify)
    username_login_entry.pack()
    Label(login_screen, text="").pack()

    Label(login_screen, text="Password * ").pack()
    password_login_entry = Entry(login_screen, textvariable=password_verify, show='*')
    password_login_entry.pack()
    Label(login_screen, text="").pack()

    global dsn, dsns, dsn_list

    label = Label(login_screen, text="Data Source", height="2", width="30")
    label.pack()
    sources = pyodbc.dataSources()
    dsns = list(sources.keys())
    dsn = StringVar()
    dsn.set('PIC22')
    dsn_list = Combobox(login_screen, textvariable=dsn)
    dsn_list['values'] = dsns
    dsn_list['state'] = 'normal'  # readonly
    dsn_list.pack()
    Label(login_screen, text="").pack()

    Label(login_screen, text="DSN Username * ").pack()
    Label(login_screen, text="").pack()
    dsn_login_entry = Entry(login_screen, textvariable=dsn_username_verify)
    dsn_login_entry.pack()
    Label(login_screen, text="").pack()

    Label(login_screen, text="DSN Password * ").pack()
    Label(login_screen, text="").pack()
    dsn_password_login_entry = Entry(login_screen, textvariable=dsn_password_verify, show='*')
    dsn_password_login_entry.pack()
    Label(login_screen, text="").pack()

    #print(dsn)
    Label(login_screen, text="").pack()
    Button(login_screen, text="Login", width=10, bg='green', height=1, command=login_verify).pack()
    return dsn

#2
def user_not_found():
    global user_not_found_screen
    user_not_found_screen = Toplevel(login_screen)
    user_not_found_screen.title("Error")
    user_not_found_screen.iconbitmap('c:\python\sql\hiport.ico')
    user_not_found_screen.geometry("200x100")
    Label(user_not_found_screen, text="User or Password incorrect").pack()
    Label(user_not_found_screen, text="").pack()
    Button(user_not_found_screen, text="OK", bg='grey' ,command=delete_user_not_found_screen).pack()

#3
def delete_login_screen():
    login_screen.destroy()

#4
def delete_user_not_found_screen():
    user_not_found_screen.destroy()

#5
def delete_main_screen():
    main_screen.destroy()

#6
def delete_app_screen():
    app.destroy()

#7
def login_error_out():
    login_screen.destroy()

#8
def archive_files():

    source_dir = 'C:\\Python\\Output\\'
    target_dir = 'C:\\Python\\ARCHIVE\\'
    os.chmod(source_dir, 0o777)
    os.chmod(target_dir, 0o777)
    file_names = os.listdir(source_dir)
    print('source dir files :' , file_names)

    for file_name in file_names:
        shutil.move(source_dir + file_name, target_dir + file_name + '_' + d4)

        '''os.rename(target_dir + file_name , target_dir + file_name + '_' + d3)'''
    archived_files = os.listdir(target_dir)
    T.insert(END,archived_files)
    T.insert(END,'')
    T.insert(END,'  Archive Completed')

#9
def login_verify():

    global dsn_id, dsn_pwd, dsn_user

    #hip_user = username_verify.get()
    #hip_pwd = password_verify.get()
    dsn_user = dsn_username_verify.get()
    dsn_pwd = dsn_password_verify.get()
    dsn_id = dsn_list.get()
    print(dsn_id,'',dsn_user,'',dsn_pwd )
    username1 = username_verify.get()
    password1 = password_verify.get()
    username_login_entry.delete(0, END)
    password_login_entry.delete(0, END)

    #conn = pyodbc.connect(dsn=dsn_id, uid=dsn_user, pwd=dsn_pwd)
    a = Datasource(dsn_id,dsn_user,dsn_pwd)
    ver = a.cursor()

    ver.execute("SELECT CODE FROM OAUSER.USER")
    res = ver.fetchall()
    res = [i[0] for i in res]
    clean_res = []
    for ele in res:
        j = ele.replace(' ', '')
        clean_res.append(j)
    if username1 in clean_res and password1 in clean_res:
        x = 1
        mygui.gui(x)
    else:
        user_not_found()
        Button(login_screen, text="Exit", width=10, bg='red', height=1, command=login_error_out).pack(pady=20)
    ver.close()

#10
class mygui:

    def gui(x):
        delete_login_screen()
        main_screen.destroy()

        global app
        app = Tk()

        #app.geometry('550x750')
        app_screen_width = 800
        app_screen_height = 850
        computer_screen_width = app.winfo_screenwidth()  # 1600
        computer_screen_height = app.winfo_screenheight()  # 900
        x = (computer_screen_width / 2) - (app_screen_width / 2)
        y = (computer_screen_height / 2) - (app_screen_height / 2)

        #app.geometry(f'{app_screen_width}x{app_screen_height}+{int(x)}+{int(y)}')
        app.geometry(f'{app_screen_width}x{app_screen_height}+{0}+{0}')
        #app.geometry(f'{app_screen_width}x{app_screen_height}')

        app.title("StatPro Reports" + '              Version : ' + version)
        app.iconbitmap('c:\python\sql\hiport.ico')

        ## menu bar
        def donothing():
            x = 0
        menubar = Menu(app)
        filemenu = Menu(menubar, tearoff=0)

        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=app.quit)
        menubar.add_cascade(label="File", menu=filemenu)
        helpmenu = Menu(menubar, tearoff=0)
        '''helpmenu.add_command(label="Help Index", command=webbrowser.open('http://wisley-hipsupp:18080/DRO18.1'))'''
        helpmenu.add_command(label="About...", command=donothing)
        menubar.add_cascade(label="Help", menu=helpmenu)
        app.config(menu=menubar)

        ##reports drop down
        Label(text="").pack(pady=10)
        label = Label(app, text="Report Name:", height="2", width="30")
        label.pack()
        report = StringVar()
        report.set('Totals')
        report_list = Combobox(app, textvariable=report)
        report_list['values'] = ('Cash', 'Transactions', 'Totals', 'Risk', 'Holdings', 'Securities', 'Portfolios', 'Issuer','AllReports')
        report_list['state'] = 'readonly'
        report_list.pack()

        ##portfolio drop down
        label = Label(app, text="Portfolio Code:", height="2", width="30")
        label.pack()
        portfolio = StringVar()
        #portfolio.set('Ai')
        portfolio_list = Combobox(app, textvariable=portfolio)
        c = Datasource(dsn_id, dsn_user, dsn_pwd).cursor()
        c.execute("SELECT CODE FROM OAUSER.PORTFOLIO WHERE ActiveFlag=1 AND PfolioType in ('P','G')")
        res = c.fetchall()
        ress = [i[0] for i in res]
        portfolio_list['values'] = ress
        portfolio_list['state'] = 'readonly'  # normal
        portfolio_list.pack()

        def get_parameters():

            global v_date , y_date, report_name, portfolio_code, data_source, eday_int
            v_date = s_date.get()
            y_date = e_date.get()
            report_name = report_list.get()
            portfolio_code= portfolio_list.get()
            data_source = dsn.get()
            smonth = v_date[5:-3]
            syear = y_date[0:4]
            sday = v_date[-2:]
            emonth = y_date[5:-3]
            eyear = y_date[0:4]
            eday = y_date[-2:]
            sday_int = int(sday)
            eday_int = int(eday)
            smonth_int = int(smonth)
            emonth_int = int(emonth)
            syear_int = int(syear)
            eyear_int = int(eyear)
            dateCheck(smonth_int,sday_int,eday_int,emonth_int,syear_int, eyear_int)

        Label(app, text="Start Date:", height="2", width="30").pack()
        s_date = DateEntry(app, width="20", year=2021, month=1, day=1, date_pattern='yyyy-mm-dd')
        #run_report['state'] = DISABLED
        s_date.pack()
        Label(text="").pack()
        Label(app, text="End Date:", height="2", width="30").pack()
        #run_report['state'] = DISABLED
        e_date = DateEntry(app, width="20", year=2021, month=1, day=1, date_pattern='yyyy-mm-dd')
        e_date.pack()
        Label(text="").pack()


        Button(app, text='Submit Parameters',  height="1", width="30", bg ='grey', command=get_parameters).pack(pady=20)
        #Label(text="").pack()
        #Label(app,text="").pack()

        global run_report
        run_report = Button(app,text="Run Report", height="1", width="30", bg ='grey', state=DISABLED, command=runapp)
        run_report.pack()
        Label(app, text="").pack()

        def browse_button():
            filename = filedialog.askopenfilename(initialdir=os.path.dirname(RES), title="Select file",
                                                  filetypes=[("All Files",".*")])

        Button(app, text='View Files', height="1", width="30", bg='grey', command=browse_button).pack()
        Label(text="").pack()

        Button(app, text='Archive Old Files', height="1", width="30", bg='grey', command=archive_files).pack()
        Label(text="").pack()

        Button(app, text="Exit", height="1", width="30", bg='red', command=delete_app_screen).pack()
        Label(app, text="").pack()

        ################################################################canvas#########################################################

        frame = Frame(app, width=50, height=50)
        frame.pack(padx=50, expand=True, fill=BOTH)  # .grid(row=0,column=0)

        global T
        T = scrolledtext.ScrolledText(frame, state='normal', width =200 , height=6)
        l = Label(frame, text = "")
        l.config(font = ("Courier", 10))
        Label(text="").pack()
        Button(text="Clear Logs", height="1", width="10", bg ='grey' , command=lambda: T.delete(1.0,END)).pack()
        l.pack()
        T.pack()

        ################################################################################################################################

        Label(app, text="").pack()
        dte = dt.datetime.now()
        fte = f"{dte:%a, %b %d, %Y}"
        Label(app, text='System Date: ' + fte).pack()
        Label(app, text="").pack()
        Label(app, text="").pack()
        app.mainloop()

#11
def screen_logger(message):
    time.sleep(1)
    T.insert(END, message)
    T.insert(END, '\n')
    T.see(END)
    T.insert(END, '\n')
    return

#12
def dateCheck(smonth_int, sday_int, eday_int, emonth_int, syear_int, eyear_int):
    run_report['state'] = DISABLED
    if (smonth_int > emonth_int) or (smonth_int < emonth_int):
        message=('Please run for the same Month!')
        screen_logger(message)

    elif sday_int > eday_int:
        message = ('Start Date cannot be greater that End Date!')
        screen_logger(message)

    elif (syear_int < eyear_int) or (syear_int > eyear_int):
        message = (' Please run for the same Year!')
        screen_logger(message)

    else:
        Label(app, text="").pack()
        message = ('StartDate=' + v_date + ' EndDate=' + y_date + ' ReportName=' + report_name + ' PorfolioCode=' + portfolio_code + ' DataSource=' + dsn_id + '\n')
        screen_logger(message)
        run_report['state'] = NORMAL
        T.see(END)

#13
def multi_thread(sday_int, eday_int,smonth_int,emonth_int,syear_int,eyear_int, syear, smonth, a, report_name):

    threads = []

    while (smonth_int == emonth_int) and (syear_int == eyear_int) and (sday_int <= eday_int):

        if sday_int < 10:
            pstartdate = (syear + '-' + smonth + '-' + '0' + str(sday_int))
        else:
            pstartdate = (syear + '-' + smonth + '-' + str(sday_int))

        if eday_int < 10:
            penddate = (syear + '-' + smonth + '-' + '0' + str(eday_int))
        else:
            penddate = (syear + '-' + smonth + '-' + str(eday_int))

        t = Thread(target=getdata, args=(pstartdate , penddate , portfolio_code,  sqlfile, report_name, dsn_id, dsn_user, dsn_pwd,a))

        t.start()
        threads.append(t)
        Label(app, text="").pack()
        #logging.info('Started ' + report_name + ' for ' + pstartdate + ' at ' + d)
        message = (report_name + ' report Started for Startdate ' + pstartdate + ' and Enddate ' + penddate + ' at ' + d)
        screen_logger(message)
        sday_int = sday_int + 1
        print(sday_int -1)
        T.see(END)

#447
def runapp():
    global sqlfile

    if report_name == 'AllReports':
        sqlfile = SQL + 'Holdings' + '_sql.txt'
        get_report_data(v_date, y_date, 'Holdings', portfolio_code, sqlfile, RES)

        sqlfile = SQL + 'Transactions' + '_sql.txt'
        get_report_data(v_date, y_date, 'Transactions', portfolio_code, sqlfile, RES)

        sqlfile = SQL + 'Risk' + '_sql.txt'
        get_report_data(v_date, y_date, 'Risk', portfolio_code, sqlfile, RES)

        sqlfile = SQL + 'Totals' + '_sql.txt'
        get_report_data(v_date, y_date, 'Totals', portfolio_code, sqlfile, RES)

        sqlfile = SQL + 'Portfolio' + '_sql.txt'
        get_report_data(v_date, y_date, 'Portfolio', portfolio_code, sqlfile, RES)

        sqlfile = SQL + 'Securities' + '_sql.txt'
        get_report_data(v_date, y_date, 'Securities', portfolio_code, sqlfile, RES)

        sqlfile = SQL + 'Issuer' + '_sql.txt'
        get_report_data(v_date, y_date, 'Issuer', portfolio_code, sqlfile, RES)
    else:
        sqlfile = SQL + report_name + '_sql.txt'
        get_report_data(v_date, y_date, report_name, portfolio_code, sqlfile, RES)
        T.see(END)

#454
def get_report_data(v_date, y_date, report_name, portfolio_code, sqlfile, RES):
    smonth = v_date[5:-3]
    syear = y_date[0:4]
    sday = v_date[-2:]
    emonth = y_date[5:-3]
    eyear = y_date[0:4]
    eday = y_date[-2:]

    ## totals report
    def tot():
        sday_int = int(sday)
        smonth_int = int(smonth)
        syear_int = int(syear)
        eday_int = int(eday)
        emonth_int = int(emonth)
        eyear_int = int(eyear)
        print(sday, smonth, syear)
        print(sday_int, smonth_int, syear_int)
        a = 0
        multi_thread(sday_int, eday_int,smonth_int,emonth_int,syear_int,eyear_int, syear, smonth, a, report_name)

    ## risk report
    def rsk():
        T.insert(END, 'Started Risk Report at ' + d2)
        sday_int = int(sday)
        eday_int = int(eday)
        smonth_int = int(smonth)
        emonth_int = int(emonth)
        syear_int = int(syear)
        eyear_int = int(eyear)
        a = 1
        # multi_thread_rsk(sday_int, eday_int, smonth_int, emonth_int, syear_int, eyear_int, syear, smonth, a)
        multi_thread(sday_int, eday_int, smonth_int, emonth_int, syear_int, eyear_int, syear, smonth, a, report_name)

    ## transaction report
    def trn():
        sday_int = int(sday)
        eday_int = int(eday)
        smonth_int = int(smonth)
        emonth_int = int(emonth)
        syear_int = int(syear)
        eyear_int = int(eyear)
        a = 0
        multi_thread(sday_int, eday_int, smonth_int, emonth_int, syear_int, eyear_int, syear, smonth, a, report_name)

    ## cash report
    def csh():
        sday_int = int(sday)
        eday_int = int(eday)
        smonth_int = int(smonth)
        emonth_int = int(emonth)
        syear_int = int(syear)
        eyear_int = int(eyear)
        a = 0
        multi_thread(sday_int, eday_int, smonth_int, emonth_int, syear_int, eyear_int, syear, smonth, a, report_name)

    ## holdings report
    def hld():
        sday_int = int(sday)
        eday_int = int(eday)
        smonth_int = int(smonth)
        emonth_int = int(emonth)
        syear_int = int(syear)
        eyear_int = int(eyear)
        a = 0
        multi_thread(sday_int, eday_int, smonth_int, emonth_int, syear_int, eyear_int, syear, smonth, a, report_name)

    ## issuer report
    def issuer():
        Label(app, text="").pack()
        threads = []
        t = Thread(target=staticdata, args=(sqlfile, report_name))
        t.start()
        threads.append(t)

    ## security report
    def sec():
        threads = []
        t = Thread(target=staticdata, args=(sqlfile, report_name))
        t.start()
        threads.append(t)

    ## portfolio report
    def por():
        threads = []
        t = Thread(target=staticdata, args=(sqlfile, report_name))
        t.start()
        threads.append(t)

    def all():

        screen_logger('running All Reports')
        hld()
        csh()
        por()
        rsk()
        trn()
        tot()
        por()
        issuer()
        message = 'running All Reports'

    ### fetch report name to run
    if report_name == 'Totals':
        tot()
    elif report_name == 'Risk':
        rsk()
    elif report_name == 'Issuer':
        issuer()
    elif report_name == 'Portfolios':
        por()
    elif report_name == 'Cash':
        csh()
    elif report_name == 'Transactions':
        trn()
    elif report_name == 'Securities':
        sec()
    elif report_name == 'Holdings':
        hld()
    elif report_name == 'AllReports':
        all()

#573
def getdata(v_date, ydate, portfolio_code, sqlfile, report_name, dsn_id, dsn_user, dsn_pwd, a):
    time.sleep(120)
    e = Datasource(dsn_id, dsn_user, dsn_pwd).cursor()
    in_file = open(sqlfile, 'r')
    sql = in_file.read()
    if a == 0:
        portfolio_code = portfolio_code.strip()
        if (report_name == 'Totals'):
            e.execute(sql, y_date, portfolio_code)
        elif (report_name == 'Cash'):
            e.execute(sql,v_date,ydate,portfolio_code)
        elif (report_name == 'Transactions'):
            e.execute(sql,y_date,portfolio_code)
        elif (report_name == 'Holdings'):
            e.execute(sql, y_date, portfolio_code)
        elif (report_name == 'Securities'):
            e.execute(sql, y_date, portfolio_code)
        elif (report_name == 'Portfolios'):
            e.execute(sql, y_date, portfolio_code)
        elif (report_name == 'Issuers'):
            e.execute(sql, y_date, portfolio_code)
    else:

        pfc = portfolio_code.strip()
        e.execute(sql, (pfc, v_date, v_date))
        print(sql, portfolio_code, v_date, v_date)

    res = e.fetchall()
    print(res)
    time.sleep(1)
    os.system('cls')
    writedata(res, RES, report_name)
    e.close()

#606
def staticdata(sqlfile, report_name):
    sd = Datasource(dsn_id,dsn_user,dsn_pwd).cursor()
    #logging.info('Started ' + report_name + ' at ' + d)
    message = (report_name + ' Started running at ' + d)
    screen_logger(message)
    #sd = Datasource.cursor()
    in_file = open(sqlfile, 'r')
    sql = in_file.read()
    sd.execute(sql,v_date)
    res = sd.fetchall()
    writedata(res, RES, report_name)
    sd.close()
    T.see(END)

#621
def writedata(res, RES, report_name):

    if (report_name == 'Risk'):
        extension = 'rsk'
    elif (report_name == 'Cash'):
        extension = 'csh'
    elif (report_name == 'Transactions'):
        extension = 'trn'
    elif (report_name == 'Portfolios'):
        extension = 'por'
    elif (report_name == 'Holdings'):
        extension = 'hld'
    elif (report_name == 'Totals'):
        extension = 'tot'
    elif (report_name == 'Securities'):
        extension = 'sec'
    else:
        extension = 'sta'

    with open(RES + 'SSNC_ALL_EXCL_GEPF.' + extension, 'a') as fp:
        a = csv.writer(fp, delimiter=',', doublequote=True, lineterminator='\n')
        i = 1
        while i <= 1 : # len(res):
            a.writerows(res)
            print(res)
            #screen_logger(res)
            i = i + 1
        writedata_msg =(report_name + " report SSNC_ALL_EXCL_GEPF." + extension + ' completed at ' + dt.datetime.now().strftime('%H:%M:%S'))
    screen_logger(writedata_msg)

if __name__ == "__main__":
    print("Logon time:" + datetime.now().strftime('%H %M %S'))
    main_account_screen()
    print("Logout time:" + datetime.now().strftime('%H %M %S'))
