import csv
import pyautogui
from pynput.mouse import Listener
from pynput.mouse import Button, Controller
from autoit import mouse_click
from pynput import keyboard
from timeit import default_timer as timer
import time
import pyautogui
import tkinter as tk
from tkinter import ttk

def timer():
    now = time.localtime(time.time())
    return now[5]

def popupmsg(msg):
    NORM_FONT = ("Helvetica", 10)
    popup = tk.Tk()
    popup.wm_title("!")
    label = ttk.Label(popup, text=msg, font=NORM_FONT)
    label.pack(side="top", fill="x", pady=10)
    B1 = ttk.Button(popup, text="Okay", command = exit)
    B1.pack()
    #popup.mainloop()

def usingscript():
    test = tk.Tk()
    test.wm_title("Andrew's Paper Machine")
    tk.Label(test, text="Project Name").grid(row=0)
    project_name = tk.Entry(test)
    project_name.grid(row=0, column=1)
    #projectname = f"{projectname}.txt"
    def usingscript2():
        projectname = project_name.get()
        projectname = f"{projectname}.txt"
        csv_file = open(projectname, mode="r", encoding="latin-1")
        for line in csv_file:
            line = line.strip()
            line = line.split(',')
            list2 = []
            list3 = []
            buttonpress = []
            buttonrelease = []
            if len(line) == 3:
                for d in line:
                    if d.isdigit() == True:
                        d = int(d)
                        list2.append(d)
                    elif d.isdigit() == False:
                        if "Button." in d:
                            d = d.replace("Button.", "")
                            list2.append(d)
            if len(line) == 1:
                for d in line:
                    #print(d)
                    #d = d.isdigit()
                    if d.isdigit() == True:
                        d = d
                        list3.append(d)
                    elif d.isdigit() == False:
                        if ".time" in d:
                            d = d.replace(".time", "")
                            d = int(float(d))
                            list3.append(d)
                            continue
                        if "Key." in d:
                            d = d.replace("Key.", "")
                        if ".press" in d:
                            d = d.replace(".press", "")
                            buttonpress.append(d)
                        if ".release" in d:
                            d = d.replace(".release", "")
                            buttonrelease.append(d)
                        list2.append(d)

            #print(list2)
            #Controller.move(f"{line[0]}",f"{line[1]}")
            #print(f"{list2[2]}", f"{list2[0]}", f"{list2[1]}", 1)
            #read items in list and if == 3 then do mouse.
            if len(list2) == 3:
                mouse_click(list2[2], (list2[0]), (list2[1]), 1)
                pyautogui.click()
                #time.sleep(1)
            if len(list3) == 1:
                time.sleep(list3[0])
            if len(buttonpress) == 1:
                #line = line.replace(".press", "")
                #print("line")
                print(buttonpress[0])
                pyautogui.keyDown(buttonpress[0])
                #time.sleep(1)
            if len(buttonrelease) == 1:
                #line = line.replace(".release", "")
                #print(line)
                print(buttonrelease[0])
                pyautogui.keyUp(buttonrelease[0])
        csv_file.close()
    #projectname = f"{projectname}.txt"
    tk.Button(test,
              text='Playback Movements',
              command=usingscript2).grid(row=4,
                                      column=0,
                                      sticky=tk.W,
                                      pady=4)
timearray2 = []
def makingscript():
    test = tk.Tk()
    test.wm_title("Andrew's Paper Machine")
    tk.Label(test, text="Project Name").grid(row=0)
    project_name = tk.Entry(test)
    project_name.grid(row=0, column=1)
    #csv_file = open('test2.txt', mode="w", encoding="latin-1")
    def guiscript():
        #Takes whatever current time is an minuses it to make it = 0 so we can have a clean slate and track time.
        waiting2 = time.time()
        waiting = time.time() - time.time()
        going = 0
        projectname = project_name.get()
        projectname = f"{projectname}.txt"
        #popupmsg("Recording... Press ESC when done")
        def on_click(x, y, button, pressed):
            global going
            if pressed:
                going = time.time() - waiting2
                #timearray2.append(going)
                timearray2.append(going)
                print(timearray2)
                if len(timearray2) == 1:
                    going2 = going - timearray2[0]
                if len(timearray2) >= 2:
                    going2 = going - timearray2[-2]
                csv_file = open(projectname, mode="a", encoding="latin-1")
                print('{0},{1},{2}'.format(x, y, button))
                csv_file.write(f"{going2}.time\n")
                csv_file.write('{0},{1},{2}\n'.format(x, y, button))
                print(going2)
        #with Listener(on_click=on_click) as listener:
        #    listener.join()
        def on_press(key):
            global going
            try:
                going = time.time() - waiting2
                # timearray2.append(going)
                timearray2.append(going)
                print(timearray2)
                if len(timearray2) == 1:
                    going2 = going - timearray2[0]
                if len(timearray2) >= 2:
                    going2 = going - timearray2[-2]
                csv_file = open(projectname, mode="a", encoding="latin-1")
                csv_file.write(f"{going2}.time\n")
                csv_file.write('{0}.press\n'.format(key.char))
                print(going2)
            except AttributeError:
                going = time.time() - waiting2
                # timearray2.append(going)
                timearray2.append(going)
                print(timearray2)
                if len(timearray2) == 1:
                    going2 = going - timearray2[0]
                if len(timearray2) >= 2:
                    going2 = going - timearray2[-2]
                csv_file = open(projectname, mode="a", encoding="latin-1")
                csv_file.write(f"{going2}.time\n")
                csv_file.write('{0}.press\n'.format(key))
                print(going2)
        def on_release(key):
            try:
                going = time.time() - waiting2
                # timearray2.append(going)
                timearray2.append(going)
                print(timearray2)
                if len(timearray2) == 1:
                    going2 = going - timearray2[0]
                if len(timearray2) >= 2:
                    going2 = going - timearray2[-2]
                print('{0} released'.format(
                    key))
                csv_file = open(projectname, mode="a", encoding="latin-1")
                csv_file.write(f"{going2}.time\n")
                csv_file.write('{0}.release\n'.format(key.char))
                print(going2)
            except AttributeError:
                going = time.time() - waiting2
                # timearray2.append(going)
                timearray2.append(going)
                print(timearray2)
                if len(timearray2) == 1:
                    going2 = going - timearray2[0]
                if len(timearray2) >= 2:
                    going2 = going - timearray2[-2]
                csv_file = open(projectname, mode="a", encoding="latin-1")
                csv_file.write(f"{going2}.time\n")
                csv_file.write('{0}.release\n'.format(key))
                print(going2)
            if key == keyboard.Key.esc:
                # Stop listener
                return False

        #Grabs current time while listening
        #waiting = time.perf_counter()
        # Collect events until released
        with keyboard.Listener(on_press=on_press, on_release=on_release) as listener:
            with Listener(on_click=on_click) as listener:
                #waiting = time.perf_counter()
                #print("test")
                listener.join()
    tk.Button(test,
              text='Start Recording',
              command=guiscript).grid(row=4,
                                     column=0,
                                     sticky=tk.W,
                                     pady=4)
#usingscript()
#makingscript()

parent = tk.Tk()
parent.wm_title("Andrew's Automation")

frame = tk.Frame(parent)
frame.pack()

text_disp= tk.Button(frame,
                   text="Create Automation Segment",
                   command=makingscript
                   )

text_disp.pack(side=tk.LEFT)

exit_button = tk.Button(frame,
                   text="Use Automation Segment",
                   fg="green",
                   command=usingscript)
exit_button.pack(side=tk.RIGHT)

parent.mainloop()