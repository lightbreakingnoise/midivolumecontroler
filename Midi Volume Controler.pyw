import mido
import mido.backends.pygame
import time
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume, IAudioEndpointVolume
import tkinter as tk
import tkextrafont as tkfont
import json
import os

global win
win = tk.Tk()
global msize
msize = 24
global mainfont
mainfont = tkfont.Font(file="anpro.ttf", family="Anonymous Pro", size=msize, weight="bold")
win.resizable(0,0)
win.title("Volumecontroller")

global allentrys
allentrys = []
try:
    f = open("midivolchconfig.json", "r")
    config = json.loads(f.read())
    f.close()
    for entry in config:
        allentrys.append(entry)
except:
    allentrys = []

def saveconfig():
    global allentrys

    config = []
    for entry in allentrys:
        config.append({"proc": entry["proc"], "control": entry["control"], "perc": entry["perc"], "init": entry["init"]})
    
    f = open("midivolchconfig.json", "w")
    f.write(json.dumps(config))
    f.close()

global tochangeentry
tochangeentry = {}

def startconfigure(in_entry):
    global tochangeentry
    tochangeentry = in_entry
    in_entry["label"].config(fg = "#ed7712")

def removeentry(in_entry):
    global allentrys

    try:
        allentrys.remove(in_entry)
        in_entry["label"].destroy()
        saveconfig()
    except:
        pass

def showvolume(in_proc, in_perc):
    global win
    global mainfont
    global allentrys
    global tochangeentry

    n = 0
    matched = False
    for i in range(len(allentrys)):
        if allentrys[i]["proc"] == in_proc:
            n = i
            matched = True
            ctrl = allentrys[i]["control"]

    if matched == False:
        ctrl = ""

    val = in_perc // 3
    cent = int(float(in_perc) / 1.27)
    txt = f"{in_proc}"
    while len(txt) < 24:
        txt += " "
    cent = "%3d"%cent
    txt += f" {cent} ["
    
    for i in range(42):
        if i < val:
            txt += "#"
        else:
            txt += "_"
    txt += "]"

    color = "#f1f1f1"
    if matched:
        entry = allentrys[n]
    else:
        labello = tk.Label(win, text=txt, font=mainfont, bg="#070707")
        entry = {"proc": in_proc, "txt": txt, "label": labello, "control": -1, "perc": in_perc, "init": True}
        labello.bind("<Button-1>", lambda x: startconfigure(entry))
        labello.bind("<ButtonRelease-3>", lambda x: removeentry(entry))
        labello.pack(fill=tk.X)
        allentrys.append(entry)
    if entry["init"] == False:
        color = "#10f210"

    entry["txt"] = txt
    try:
        entry["label"].config(text = txt)
        entry["label"].config(fg = color)
    except:
        labello = tk.Label(win, text=txt, font=mainfont, bg="#070707", fg=color)
        labello.bind("<Button-1>", lambda x: startconfigure(entry))
        labello.bind("<ButtonRelease-3>", lambda x: removeentry(entry))
        labello.pack(fill=tk.X)
        entry["label"] = labello
    win.title(f"{in_proc} {cent}")

def changevol(proc, perc, changeit=True):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if session.Process and session.Process.name() == proc:
            if changeit:
                volume.SetMasterVolume(int(perc / 1.27) / 100, None)
                showvolume(proc, perc)
            else:
                showvolume(proc, 0)

    showvolume(proc, perc)

def changespkvol(perc, changeit=True):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = interface.QueryInterface(IAudioEndpointVolume)
    if changeit:
        volume.SetMasterVolumeLevelScalar(int(perc / 1.27) / 100, None)
        showvolume("Speakers", perc)
    else:
        cent = int(volume.GetMasterVolumeLevelScalar() * 1.27)
        showvolume("Speakers", cent)

def changemicvol(perc, changeit=True):
    devices = AudioUtilities.GetMicrophone()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = interface.QueryInterface(IAudioEndpointVolume)
    if changeit:
        volume.SetMasterVolumeLevelScalar(int(perc / 1.27) / 100, None)
        showvolume("Microphone", perc)
    else:
        cent = int(volume.GetMasterVolumeLevelScalar() * 1.27)
        showvolume("Microphone", cent)

def listnewdevs():
    global win
    global allentrys
    
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if session.Process:
            matched = False
            for n in range(len(allentrys)):
                if allentrys[n]["proc"] == session.Process.name():
                    matched = True
            if matched == False:
                showvolume(session.Process.name(), 0)
    
    win.after(500, listnewdevs)

for entry in allentrys:
    if entry["proc"] == "Speakers":
        changespkvol(entry["perc"], changeit=False)
    elif entry["proc"] == "Microphone":
        changemicvol(entry["perc"], changeit=False)
    else:
        changevol(entry["proc"], entry["perc"], changeit=False)

pygame = mido.Backend('mido.backends.pygame')
pygame.get_input_names()
global inport
inport = pygame.open_input()

changespkvol(0, changeit=False)
changemicvol(0, changeit=False)

def checkmidi():
    global allentrys
    global tochangeentry
    
    for msg in inport.iter_pending():
        if msg.is_cc():
            for entry in allentrys:
                if entry["proc"] == "Speakers":
                    if tochangeentry == entry:
                        tochangeentry = {}
                        entry["control"] = msg.control
                        entry["init"] = False
                        saveconfig()
                    if msg.control == entry["control"]:
                        changespkvol(msg.value)
                elif entry["proc"] == "Microphone":
                    if tochangeentry == entry:
                        tochangeentry = {}
                        entry["control"] = msg.control
                        entry["init"] = False
                        saveconfig()
                    if msg.control == entry["control"]:
                        changemicvol(msg.value)
                else:
                    if tochangeentry == entry:
                        tochangeentry = {}
                        entry["control"] = msg.control
                        entry["init"] = False
                        saveconfig()
                    if msg.control == entry["control"]:
                        changevol(entry["proc"], msg.value)

    win.after(20, checkmidi)

def zoomit(event):
    global win
    global mainfont
    global allentrys
    global msize

    if msize > 8 and event.delta < 0:
        msize -= 2
    if msize < 32 and event.delta > 0:
        msize += 2
    
    mainfont = tkfont.Font(family="Anonymous Pro", size=msize, weight="bold")

    for entry in allentrys:
        entry["label"].config(font = mainfont)
    
win.bind("<MouseWheel>", zoomit)
win.after(100, checkmidi)
win.after(100, listnewdevs)
win.mainloop()
