import mido
import mido.backends.pygame
import time
import comtypes
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume, IAudioEndpointVolume, IMMDeviceEnumerator, EDataFlow, DEVICE_STATE
from pycaw.constants import CLSID_MMDeviceEnumerator
import policyconfig as pc
import tkinter as tk
import tkinter.font as tkfont
#import tkextrafont as tkfont
import json
import os
import subprocess as sp
import shlex

global win
win = tk.Tk()
win.config(bg = "#204050")
global msize
msize = 14
global mainfont
mainfont = tkfont.Font(family="Terminal", size=msize, weight="bold")
win.resizable(0,0)
win.title("Midi Volume Controller")
win.iconbitmap("controler.ico")
global spkframe
spkframe = tk.Frame(win)
spkframe.pack(fill=tk.X, pady=1)
global micframe
micframe = tk.Frame(win)
micframe.pack(fill=tk.X, pady=1)
global appframe
appframe = tk.Frame(win)
appframe.pack(fill=tk.X, pady=1)
global runframe
runframe = tk.Frame(win)
runframe.pack(fill=tk.X, pady=1)

global allentrys
allentrys = []
try:
    f = open("midivolchconfig.json", "r")
    config = json.loads(f.read())
    f.close()
    for entry in config:
        if not "note" in entry:
            entry["note"] = -1
        if not "standard" in entry:
            entry["standard"] = False
        if not "smic" in entry:
            entry["smic"] = False
        if not "ctrl2" in entry:
            entry["ctrl2"] = -1
        allentrys.append(entry)
except:
    allentrys = []

def saveconfig():
    global allentrys

    config = []
    for entry in allentrys:
        config.append({"proc": entry["proc"], "control": entry["control"], "ctrl2": entry["ctrl2"], "note": entry["note"], "perc": entry["perc"], "init": entry["init"]})

    f = open("midivolchconfig.json", "w")
    f.write(json.dumps(config))
    f.close()

global tochangeentry
tochangeentry = {}

def pre_startconfigure(in_entry):
    global win
    global mainfont
    
    subwin = tk.Toplevel(win)
    subwin.resizable(0,0)
    subwin.attributes("-topmost", 1)
    btn1 = tk.Button(subwin, text="make standard", font=mainfont, command=lambda : startconfigure(in_entry, in_mode=2, rm_win=subwin))
    btn1.grid(column=0, row=0)
    btn2 = tk.Button(subwin, text="change volume", font=mainfont, command=lambda : startconfigure(in_entry, in_mode=1, rm_win=subwin))
    btn2.grid(column=1, row=0)

def startconfigure(in_entry, in_mode=1, rm_win=None):
    global tochangeentry
    tochangeentry = in_entry
    in_entry["label"].config(fg = "#ed7712")
    in_entry["mode"] = in_mode

    if rm_win is not None:
        rm_win.destroy()
        rm_win.update()

def removeentry(in_entry):
    global allentrys

    try:
        allentrys.remove(in_entry)
        in_entry["label"].destroy()
        saveconfig()
    except:
        pass

def showapp(in_proc):
    global win
    global mainfont
    global allentrys
    global runframe

    n = 0
    matched = False
    for i in range(len(allentrys)):
        if allentrys[i]["proc"] == in_proc:
            n = i
            matched = True

    end_proc = "> " + in_proc[1:][-90:]
    color = "#cecece"
    if matched:
        entry = allentrys[n]
    else:
        labello = tk.Label(runframe, text=end_proc, font=mainfont, bg="#070707", anchor="w")
        entry = {"proc": in_proc, "txt": end_proc, "label": labello, "control": -1, "ctrl2": -1, "note": -1, "perc": 0, "init": True, "standard": False, "smic": False, "mode": 0}
        labello.bind("<Button-1>", lambda x: startconfigure(entry))
        labello.bind("<ButtonRelease-3>", lambda x: removeentry(entry))
        labello.pack(fill=tk.X)
        allentrys.append(entry)

    try:
        entry["label"].config(text = end_proc)
        entry["label"].config(fg = color)
    except:
        labello = tk.Label(runframe, text=end_proc, font=mainfont, bg="#070707", fg=color, anchor="w")
        labello.bind("<Button-1>", lambda x: startconfigure(entry))
        labello.bind("<ButtonRelease-3>", lambda x: removeentry(entry))
        labello.pack(fill=tk.X)
        entry["label"] = labello

def showvolume(in_proc, in_perc):
    global win
    global mainfont
    global allentrys
    global tochangeentry
    global spkframe
    global micframe
    global appframe

    n = 0
    matched = False
    for i in range(len(allentrys)):
        if allentrys[i]["proc"] == in_proc:
            n = i
            matched = True
            ctrl = allentrys[i]["control"]

    if matched == False:
        ctrl = ""

    end_proc = in_proc[1:48]

    val = in_perc // 3
    cent = int(float(in_perc) / 1.27)
    ln = max(1, 50 - len(end_proc))
    space = " " * ln
    cent = "%3d"%cent
    raut = "#" * val
    unraut = " " * (42 - val)
    txt = f"{end_proc} {space} {cent} [{raut}{unraut}]"
    
    color = "#f1f1f1"
    if matched:
        entry = allentrys[n]
    else:
        if in_proc.startswith("+"):
            labello = tk.Label(spkframe, text=txt, font=mainfont, bg="#070707", anchor="w")
        if in_proc.startswith("-"):
            labello = tk.Label(micframe, text=txt, font=mainfont, bg="#070707", anchor="w")
        if in_proc.startswith("#"):
            labello = tk.Label(appframe, text=txt, font=mainfont, bg="#070707", anchor="w")
        entry = {"proc": in_proc, "txt": txt, "label": labello, "control": -1, "ctrl2": -1, "note": -1, "perc": in_perc, "init": True, "standard": False, "smic": False, "mode": 0}
        if in_proc.startswith("#"):
            labello.bind("<Button-1>", lambda x: startconfigure(entry))
        else:
            labello.bind("<Button-1>", lambda x: pre_startconfigure(entry))
        labello.bind("<ButtonRelease-3>", lambda x: removeentry(entry))
        labello.pack(fill=tk.X)
        allentrys.append(entry)
    if entry["init"] == False:
        color = "#10f210"
    if entry["standard"] or entry["smic"]:
        color = "#1097f2"

    entry["txt"] = txt
    try:
        entry["label"].config(text = txt)
        entry["label"].config(fg = color)
    except:
        if in_proc.startswith("+"):
            labello = tk.Label(spkframe, text=txt, font=mainfont, bg="#070707", fg=color, anchor="w")
        if in_proc.startswith("-"):
            labello = tk.Label(micframe, text=txt, font=mainfont, bg="#070707", fg=color, anchor="w")
        if in_proc.startswith("#"):
            labello = tk.Label(appframe, text=txt, font=mainfont, bg="#070707", fg=color, anchor="w")
        if in_proc.startswith("#"):
            labello.bind("<Button-1>", lambda x: startconfigure(entry))
        else:
            labello.bind("<Button-1>", lambda x: pre_startconfigure(entry))
        labello.bind("<ButtonRelease-3>", lambda x: removeentry(entry))
        labello.pack(fill=tk.X)
        entry["label"] = labello
    #win.title(f"{cent} {end_proc}")

comtypes.CoInitialize()
global deviceEnumerator
deviceEnumerator = comtypes.CoCreateInstance(CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, comtypes.CLSCTX_INPROC_SERVER)
State = DEVICE_STATE.ACTIVE.value
global collection
inpcollection = deviceEnumerator.EnumAudioEndpoints(EDataFlow.eCapture.value, State)
outcollection = deviceEnumerator.EnumAudioEndpoints(EDataFlow.eRender.value, State)
global alloutdevices
alloutdevices = []
for i in range(outcollection.GetCount()):
    dev = outcollection.Item(i)
    alloutdevices.append(AudioUtilities.CreateDevice(dev))
global alldevices
allinpdevices = []
for i in range(inpcollection.GetCount()):
    dev = inpcollection.Item(i)
    allinpdevices.append(AudioUtilities.CreateDevice(dev))

global policy_config
policy_config = comtypes.CoCreateInstance(pc.CLSID_PolicyConfigClient, pc.IPolicyConfig, comtypes.CLSCTX_ALL)

def changevol(proc, perc, changeit=True):
    outperc = perc
    
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if session.Process and session.Process.name() == proc[1:]:
            if changeit:
                volume.SetMasterVolume(int(perc / 1.27) / 100, None)
            else:
                outperc = int(volume.GetMasterVolume() * 127.0)

    showvolume(proc, outperc)

def changedevvol(proc, perc, changeit=True, State = DEVICE_STATE.ACTIVE.value):
    global alloutdevices
    global allinpdevices

    outperc = perc
    for device in alloutdevices:
        name = "+" + device.FriendlyName
        if proc == name:
            matched = False
            for n in range(len(allentrys)):
                if allentrys[n]["proc"] == name:
                    volume = device.EndpointVolume
                    if changeit:
                        volume.SetMasterVolumeLevelScalar(int(perc / 1.27) / 100, None)
                    else:
                        outperc = int(volume.GetMasterVolumeLevelScalar() * 127.0) + 1

    for device in allinpdevices:
        name = "-" + device.FriendlyName
        if proc == name:
            matched = False
            for n in range(len(allentrys)):
                if allentrys[n]["proc"] == name:
                    volume = device.EndpointVolume
                    if changeit:
                        volume.SetMasterVolumeLevelScalar(int(perc / 1.27) / 100, None)
                    else:
                        outperc = int(volume.GetMasterVolumeLevelScalar() * 127.0) + 1

    showvolume(proc, outperc)


def listnewdevs(State = DEVICE_STATE.ACTIVE.value):
    global win
    global allentrys
    global alldevices

    for device in alloutdevices:
        name = "+" + device.FriendlyName
        matched = False
        for n in range(len(allentrys)):
            if allentrys[n]["proc"] == name:
                matched = True
        changedevvol(name, 0, changeit=False)

    for device in allinpdevices:
        name = "-" + device.FriendlyName
        matched = False
        for n in range(len(allentrys)):
            if allentrys[n]["proc"] == name:
                matched = True
        changedevvol(name, 0, changeit=False)

    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if session.Process:
            name = "#" + session.Process.name()
            matched = False
            for n in range(len(allentrys)):
                if allentrys[n]["proc"] == name:
                    matched = True
            changevol(name, 0, changeit=False)
    
    win.after(500, listnewdevs)

def changedefault(proc):
    global allentrys
    global alldevices

    for device in alloutdevices:
        name = "+" + device.FriendlyName
        if proc == name:
            matched = False
            for n in range(len(allentrys)):
                if allentrys[n]["proc"] == name:
                    policy_config.SetDefaultEndpoint(device.id, 0)
                    policy_config.SetDefaultEndpoint(device.id, 2)

    for device in allinpdevices:
        name = "-" + device.FriendlyName
        if proc == name:
            matched = False
            for n in range(len(allentrys)):
                if allentrys[n]["proc"] == name:
                    policy_config.SetDefaultEndpoint(device.id, 0)
                    policy_config.SetDefaultEndpoint(device.id, 2)

def addapp(ntry):
    showapp("="+ntry.get())
    ntry.delete(0, tk.END)
    ntry.insert(0, "")

for entry in allentrys:
    if entry["proc"].startswith("="):
        showapp(entry["proc"])
    if entry["proc"].startswith("+") or entry["proc"].startswith("-"):
        changedevvol(entry["proc"], entry["perc"], changeit=False)
    if entry["proc"].startswith("#"):
        changevol(entry["proc"], entry["perc"], changeit=False)

mkentry = tk.Entry(win, font=mainfont, bg="#204050", fg="#cecece")
mkentry.bind("<Return>", lambda event, e=mkentry : addapp(e))
mkentry.pack(fill=tk.X)

pygame = mido.Backend('mido.backends.pygame')
pygame.get_input_names()
global inport
inport = pygame.open_input()

global rstcount
rstcount = 0

def resettitle():
    global win
    global rstcount

    if rstcount <= 0:
        rstcount = 50
        win.title("Midi Volume Controller")
    else:
        rstcount -= 1

    win.after(100, resettitle)

def checkmidi():
    global allentrys
    global tochangeentry
    global win
    global rstcount
    
    for msg in inport.iter_pending():
        if msg.is_cc():
            wperc = int(msg.value / 1.27)
            for entry in allentrys:
                    if entry["proc"].startswith("+") or entry["proc"].startswith("-"):
                        proc = entry["proc"][1:]
                        if tochangeentry == entry and entry["mode"] == 1:
                            tochangeentry = {}
                            entry["control"] = msg.control
                            entry["init"] = False
                            saveconfig()
                        if msg.control == entry["control"]:
                            changedevvol(entry["proc"], msg.value)
                            entry["perc"] = msg.value
                            win.title(f"{wperc} {proc}")
                            rstcount = 50
                            saveconfig()
                    elif entry["proc"].startswith("#"):
                        proc = entry["proc"][1:]
                        if tochangeentry == entry:
                            tochangeentry = {}
                            entry["control"] = msg.control
                            entry["init"] = False
                            saveconfig()
                        if msg.control == entry["control"]:
                            changevol(entry["proc"], msg.value)
                            entry["perc"] = msg.value
                            win.title(f"{wperc} {proc}")
                            rstcount = 50
                            saveconfig()
                    if entry["proc"].startswith("+"):
                        proc = entry["proc"][1:]
                        if tochangeentry == entry and entry["mode"] == 2:
                            tochangeentry = {}
                            entry["ctrl2"] = msg.control
                            entry["init"] = False
                            saveconfig()
                        if msg.control == entry["ctrl2"]:
                            changedefault(entry["proc"])
                            for ntry in allentrys:
                                ntry["standard"] = False
                                if ntry["init"]:
                                    if ntry["smic"]:
                                        ntry["label"].config(fg = "#1097f2")
                                    else:
                                        ntry["label"].config(fg = "#f1f1f1")
                                else:
                                    ntry["label"].config(fg = "#10f210")
                            entry["standard"] = True
                            entry["label"].config(fg = "#1097f2")
                            win.title(f"set std spk {proc}")
                            rstcount = 50
                            saveconfig()
                    if entry["proc"].startswith("-"):
                        proc = entry["proc"][1:]
                        if tochangeentry == entry and entry["mode"] == 2:
                            tochangeentry = {}
                            entry["ctrl2"] = msg.control
                            entry["init"] = False
                            saveconfig()
                        if msg.control == entry["ctrl2"]:
                            changedefault(entry["proc"])
                            for ntry in allentrys:
                                ntry["smic"] = False
                                if ntry["init"]:
                                    if ntry["standard"]:
                                        ntry["label"].config(fg = "#1097f2")
                                    else:
                                        ntry["label"].config(fg = "#f1f1f1")
                                else:
                                    ntry["label"].config(fg = "#10f210")
                            entry["smic"] = True
                            entry["label"].config(fg = "#1097f2")
                            win.title(f"set std mic {proc}")
                            rstcount = 50
                            saveconfig()
                    if entry["proc"].startswith("=") and msg.value > 63:
                        proc = entry["proc"][1:]
                        if tochangeentry == entry:
                            tochangeentry = {}
                            entry["control"] = msg.control
                            entry["init"] = False
                            entry["label"].config(fg = "#cecece")
                            saveconfig()
                        if msg.control == entry["control"] and msg.value > 63:
                            sp.Popen(shlex.split(proc))
                            entry["label"].config(fg = "#cecece")
                            win.title(f"run {proc}")
                            rstcount = 50
        if msg.type == "note_on":
            for entry in allentrys:
                if entry["proc"].startswith("="):
                    proc = entry["proc"][1:]
                    if tochangeentry == entry:
                        tochangeentry = {}
                        entry["note"] = msg.note
                        entry["init"] = False
                        entry["label"].config(fg = "#cecece")
                        saveconfig()
                    if msg.note == entry["note"]:
                        sp.Popen(shlex.split(entry["proc"][1:]))
                        entry["label"].config(fg = "#cecece")
                        win.title(f"run {proc}")
                        rstcount = 50
                if entry["proc"].startswith("+"):
                    proc = entry["proc"][1:]
                    if tochangeentry == entry:
                        tochangeentry = {}
                        entry["note"] = msg.note
                        saveconfig()
                    if msg.note == entry["note"]:
                        changedefault(entry["proc"])
                        for ntry in allentrys:
                            ntry["standard"] = False
                            if ntry["init"]:
                                if ntry["smic"]:
                                    ntry["label"].config(fg = "#1097f2")
                                else:
                                    ntry["label"].config(fg = "#f1f1f1")
                            else:
                                ntry["label"].config(fg = "#10f210")
                        entry["standard"] = True
                        entry["label"].config(fg = "#1097f2")
                        win.title(f"set std spk {proc}")
                        rstcount = 50
                        saveconfig()
                if entry["proc"].startswith("-"):
                    proc = entry["proc"][1:]
                    if tochangeentry == entry:
                        tochangeentry = {}
                        entry["note"] = msg.note
                        saveconfig()
                    if msg.note == entry["note"]:
                        changedefault(entry["proc"])
                        for ntry in allentrys:
                            ntry["smic"] = False
                            if ntry["init"]:
                                if ntry["standard"]:
                                    ntry["label"].config(fg = "#1097f2")
                                else:
                                    ntry["label"].config(fg = "#f1f1f1")
                            else:
                                ntry["label"].config(fg = "#10f210")
                        entry["smic"] = True
                        entry["label"].config(fg = "#1097f2")
                        win.title(f"set std mic {proc}")
                        rstcount = 50
                        saveconfig()

    win.after(20, checkmidi)

def zoomit(event):
    global win
    global mainfont
    global allentrys
    global msize

    if msize > 8 and event.delta < 0:
        msize -= 2
    if msize < 18 and event.delta > 0:
        msize += 2
    
    mainfont = tkfont.Font(family="Terminal", size=msize, weight="bold")

    for entry in allentrys:
        entry["label"].config(font = mainfont)

def updatedeviceslist(State = DEVICE_STATE.ACTIVE.value):
    global collection
    global alldevices
    global deviceEnumerator
    
    for device in alldevices:
        device._dev.Release()
    alldevices = []

    collection = deviceEnumerator.EnumAudioEndpoints(EDataFlow.eAll.value, State)

    for i in range(collection.GetCount()):
        dev = collection.Item(i)
        if dev is not None:
            alldevices.append(AudioUtilities.CreateDevice(dev))

    win.after(10000, updatedeviceslist)

win.bind("<MouseWheel>", zoomit)
win.after(100, checkmidi)
win.after(100, listnewdevs)
win.after(100, resettitle)
#win.after(10000, updatedeviceslist)
win.mainloop()
