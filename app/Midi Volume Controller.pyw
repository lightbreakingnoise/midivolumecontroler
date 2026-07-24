# library for handling midi controller inputs
import mido, mido.backends.pygame

# some other important libraries
import time, json, os, subprocess as sp, shlex

# libraries for controlling windows sound stuff
import comtypes, pycaw.pycaw as caw
import pycaw.constants as cawstants
import policyconfig as pc

# libraries for building GUI
import tkinter as tk
import tkextrafont as tkfont
import guicolors as gc

class Controller:
    def __init__(self):
        self.pygame = mido.Backend('mido.backends.pygame')
        self.portlist = self.pygame.get_input_names()
        self.selected = False
        if len(self.portlist) == 0:
            self.errmsg = "no midi"
            return
        if len(self.portlist) == 1:
            name = self.portlist[0]
            self.inport = self.pygame.open_input()
            self.errmsg =  f"midi {name} selected"
            self.selected = True
            return

    def select_midi(self, name):
        self.inport = self.pygame.open_input(name)

    # interval called function that checks midi controller inputs
    def checkmidi(self, gui, sound):
        for msg in self.inport.iter_pending():
            pre = ""
            if msg.is_cc():
                pre = "c"
            if msg.type == "note_on":
                pre = "n"

            if pre == "c" or pre == "n":
                for entry in gui.entrys:
                    ntype = entry["type"]
                    ncont = entry["content"]

                    # check if entry is waiting for assignment
                    if gui.chentry is not None and gui.chentry["type"] == ntype and gui.chentry["content"] == ncont:
                        gui.unconfigure(entry)
                        if pre == "c":
                            if gui.chmode == "vol":
                                entry["trigger"] = f"c{msg.control}"
                            if gui.chmode == "std":
                                entry["sectrig"] = f"c{msg.control}"
                            if gui.chmode == "mut":
                                entry["tritrig"] = f"c{msg.control}"
                            entry["init"] = True
                            gui.save()
                        if pre == "n":
                            if gui.chmode == "std":
                                entry["sectrig"] = f"n{msg.note}"
                            if gui.chmode == "mut":
                                entry["tritrig"] = f"n{msg.note}"
                            entry["init"] = True
                            gui.save()

                    # check if this entry was triggered
                    ntrig = entry["trigger"]
                    if ntrig.startswith("c") and msg.is_cc():
                        if int(ntrig[1:]) == msg.control:
                            entry["value"] = msg.value
                            self.triggercmd(entry, sound, gui, "vol")

                    # check if this entry has a secondary trigger
                    if "sectrig" in entry:
                        # check if its triggered
                        strig = entry["sectrig"]
                        if strig.startswith("c") and msg.is_cc():
                            if int(strig[1:]) == msg.control and msg.value > 63:
                                self.triggercmd(entry, sound, gui, "std")
                        if strig.startswith("n") and msg.type == "note_on":
                            if int(strig[1:]) == msg.note:
                                if ntype != "app":
                                    self.triggercmd(entry, sound, gui, "std")

                    if "tritrig" in entry:
                        ttrig = entry["tritrig"]
                        if ttrig.startswith("c") and msg.is_cc():
                            if int(ttrig[1:]) == msg.control and msg.value > 63:
                                self.triggercmd(entry, sound, gui, "mut")
                        if ttrig.startswith("n") and msg.type == "note_on":
                            if int(ttrig[1:]) == msg.note:
                                if ntype != "app":
                                    self.triggercmd(entry, sound, gui, "mut")
                    
        # gv means given -> given gui and given sound
        gui.win.after(50, lambda gvgui=gui, gvsound=sound : self.checkmidi(gvgui, gvsound))

    # function to call when controller input triggers a command
    def triggercmd(self, ntry, snd, gui, mod):
        typ = ntry["type"]
        con = ntry["content"]
        val = ntry["value"]
        ival = int(val / 1.27)
        if typ == "microphone" and mod == "vol":
            snd.change_mic_vol(con, val)
            gui.updateentry(ntry)
            gui.win.title(f"{ival} {con}")
            gui.rstcount = 50
        if typ == "speaker" and mod == "vol":
            snd.change_spk_vol(con, val)
            gui.updateentry(ntry)
            gui.win.title(f"{ival} {con}")
            gui.rstcount = 50
        if typ == "microphone" and mod == "std":
            snd.change_default_mic(con)
            for e in gui.entrys:
                e["std"] = False
                gui.updateentry(e)
            ntry["std"] = True
            gui.updateentry(ntry)
            gui.win.title(f"standard now: {con}")
            gui.rstcount = 50
        if typ == "speaker" and mod == "std":
            snd.change_default_spk(con)
            for e in gui.entrys:
                e["std"] = False
                gui.updateentry(e)
            ntry["std"] = True
            gui.updateentry(ntry)
            gui.win.title(f"standard now: {con}")
            gui.rstcount = 50
        if typ == "microphone" and mod == "mut":
            ismut = snd.toggle_mute_mic(con)
            ntry["mut"] = ismut
            xmut = "muted" if ismut == 1 else "unmuted"
            gui.updateentry(ntry)
            gui.win.title(f"{xmut} {con}")
            gui.rstcount = 50
        if typ == "speaker" and mod == "mut":
            ismut = snd.toggle_mute_spk(con)
            ntry["mut"] = ismut
            xmut = "muted" if ismut == 1 else "unmuted"
            gui.updateentry(ntry)
            gui.win.title(f"{xmut} {con}")
            gui.rstcount = 50
        if typ == "app":
            snd.change_app_vol(con, val)
            #xival = max(0, ival-1)
            gui.updateentry(ntry)
            gui.win.title(f"{ival} {con}")
            gui.rstcount = 50
        if typ == "script" and val > 63:
            sp.Popen(shlex.split(con))
            gui.updateentry(ntry)
            gui.win.title(f"running {con}")
            gui.rstcount = 50

class WinSound:
    def __init__(self):
        comtypes.CoInitialize()

        # the device enumerator
        self.devEnum = comtypes.CoCreateInstance(cawstants.CLSID_MMDeviceEnumerator,
            caw.IMMDeviceEnumerator, comtypes.CLSCTX_INPROC_SERVER)

        self.micdevs = []
        self.spkdevs = []
        self.enumdevices()
        self.app_sessions = {}
        
        # the policy config needed for changing standard devices
        self.policy_config = comtypes.CoCreateInstance(pc.CLSID_PolicyConfigClient,
            pc.IPolicyConfig, comtypes.CLSCTX_ALL)

    def enumdevices(self, gui=None):
        # important values
        micval = caw.EDataFlow.eCapture.value
        spkval = caw.EDataFlow.eRender.value
        state = caw.DEVICE_STATE.ACTIVE.value

        # collections of devices
        micdevs = self.devEnum.EnumAudioEndpoints(micval, state)
        spkdevs = self.devEnum.EnumAudioEndpoints(spkval, state)
        
        # copy collections to lists
        self.copydevs(micdevs, self.micdevs)
        self.copydevs(spkdevs, self.spkdevs)

        if gui is not None:
            gui.win.after(1000, lambda : self.enumdevices(gui))

    # copy a device collection to destination list
    def copydevs(self, srcdevs, dstdevs):
        checklist = []
        for i in range(srcdevs.GetCount()):
            dev = srcdevs.Item(i)
            ndev = caw.AudioUtilities.CreateDevice(dev)
            checklist.append(ndev)
            if not ndev.FriendlyName in [x.FriendlyName for x in dstdevs]:
                dstdevs.append(ndev)
        todels = []
        for dev in dstdevs:
            if not dev.FriendlyName in [x.FriendlyName for x in checklist]:
                todels.append(dev)
        for todel in todels:
            dstdevs.remove(todel)

    # call this to change volume of an app
    def change_app_vol(self, appname, value, changeit=True):
        try:
            volume = self.app_sessions[appname]
            if volume:
                if changeit:
                    volume.SetMasterVolume(value / 127.0, None)
                else:
                    out = volume.GetMasterVolume()
                    return int(round(out * 127.0))
        except:
            pass
        
    # do not call this directly
    def change_dev_vol(self, devname, value, devlist, changeit=True):
        for device in devlist:
            if devname == device.FriendlyName:
                volume = device.EndpointVolume
                if changeit:
                    volume.SetMasterVolumeLevelScalar(int(value / 1.27) / 100, None)
                else:
                    return int(volume.GetMasterVolumeLevelScalar() * 127.0) + 1

    # call this to change volume of a microphone device
    def change_mic_vol(self, micname, value, changeit=True):
        self.change_dev_vol(micname, value, self.micdevs, changeit=changeit)

    def toggle_mute_mic(self, devname, changeit=True):
        return self.toggle_mute(devname, self.micdevs, changeit=changeit)

    def toggle_mute_spk(self, devname, changeit=True):
        return self.toggle_mute(devname, self.spkdevs, changeit=changeit)
    
    def toggle_mute(self, devname, devlist, changeit=True):
        for device in devlist:
            if devname == device.FriendlyName:
                volume = device.EndpointVolume
                if volume.GetMute() == 0:
                    mute = 1
                else:
                    mute = 0
                volume.SetMute(mute, None)
                return mute
        return 2

    # call this to change volume of a speaker device
    def change_spk_vol(self, spkname, value, changeit=True):
        self.change_dev_vol(spkname, value, self.spkdevs, changeit=changeit)

    def change_default_dev(self, devname, devlist):
        for dev in devlist:
            if devname == dev.FriendlyName:
                self.policy_config.SetDefaultEndpoint(dev.id, 0)
                self.policy_config.SetDefaultEndpoint(dev.id, 2)

    def change_default_mic(self, micname):
        self.change_default_dev(micname, self.micdevs)

    def change_default_spk(self, spkname):
        self.change_default_dev(spkname, self.spkdevs)
    
    def list_devs(self, gui, devlist, dsttype):
        for dev in devlist:
            devname = dev.FriendlyName
            value = self.change_dev_vol(devname, 0, devlist, changeit=False)
            volume = dev.EndpointVolume
            ismut = volume.GetMute()
            ln = len(gui.entrys)
            matched = False
            for n in range(ln):
                gtype = gui.entrys[n]["type"]
                gcont = gui.entrys[n]["content"]
                if gtype == dsttype and gcont == devname:
                    gui.entrys[n]["value"] = value
                    gui.entrys[n]["mut"] = ismut
                    gui.updateentry(gui.entrys[n])
                    matched = True
            if matched == False:
                d = {"type": dsttype, "content": devname,
                     "trigger": "", "value": value, "std": False, "init": False}
                gui.addentry(d)

        todels = []
        for e in gui.entrys:
            if e["type"] == dsttype and e["content"] not in [x.FriendlyName for x in devlist]:
                todels.append(e)
        for todel in todels:
            gui.removeentry(todel)

    def listdevices(self, gui):
        self.list_devs(gui, self.micdevs, "microphone")
        self.list_devs(gui, self.spkdevs, "speaker")
        gui.win.after(1000, lambda : self.listdevices(gui))

    def listapps(self, gui):
        self.app_sessions.clear()

        sessions = caw.AudioUtilities.GetAllSessions()
        for session in sessions:
            volume = session._ctl.QueryInterface(caw.ISimpleAudioVolume)
            if session.Process:
                self.app_sessions[session.Process.name()] = volume
                appname = session.Process.name()
                value = self.change_app_vol(appname, 0, changeit=False)
                matched = False
                ln = len(gui.entrys)
                for n in range(ln):
                    gtype = gui.entrys[n]["type"]
                    gcont = gui.entrys[n]["content"]
                    if gtype == "app" and gcont == appname:
                        #print("listapps", id(gui.entrys[n]), id(gui.entrys[n]["label"]), gui.entrys[n]["content"])
                        gui.entrys[n]["value"] = value
                        gui.updateentry(gui.entrys[n])
                        matched = True
                if matched == False:
                    d = {"type": "app", "content": appname,
                         "trigger": "", "value": value, "std": False, "init": False}
                    gui.addentry(d)

        todels = []
        for ntry in gui.entrys:
            if ntry["type"] == "app" and ntry["content"] not in self.app_sessions.keys():
                todels.append(ntry)

        for t in todels:
            gui.removeentry(t)

        gui.win.after(1000, lambda : self.listapps(gui))

class GUI:
    def __init__(self):
        self.win = tk.Tk()
        self.win.config(bg=gc.mainspace)
        self.mainfont = tkfont.Font(file="sspro.ttf",
            family="Source Code Pro", size=11, weight="normal")
        self.bigfont = tkfont.Font(family="Source Code Pro",
            size=24, weight="bold")
        self.win.resizable(0,0)
        self.win.title("Midi Volume Controller")
        self.rstcount = 50
        self.win.iconbitmap("controller.ico")
        self.win.attributes("-disabled", True)

        self.micpic = tk.PhotoImage(file="lmicrophone.png")
        self.spkpic = tk.PhotoImage(file="lspeaker.png")
        self.apppic = tk.PhotoImage(file="lapp.png")
        self.runpic = tk.PhotoImage(file="lscript.png")

        self.chmode = "vol"
        self.speakers = tk.Frame(self.win)
        self.speakers.pack(fill=tk.X)
        self.microphones = tk.Frame(self.win)
        self.microphones.pack(fill=tk.X)
        self.apps = tk.Frame(self.win)
        self.apps.pack(fill=tk.X)
        self.scripts = tk.Frame(self.win)
        self.scripts.pack(fill=tk.X)
        self.scripts.bind("<Expose>", self.onexpose)
        self.mkscript = tk.Entry(self.win, font=self.mainfont,
            bg=gc.entrycol, fg=gc.fieldback, border=8, relief=tk.FLAT)
        self.mkscript.bind("<Return>", lambda event : self.addscript())
        self.mkscript.pack(fill=tk.X)

        # type = speaker, microphone, app, script
        # content = name of speaker, microphone, app OR the script to run
        # trigger = trigger for volume -> note or control -> n or c followed by ID
        # sectrig = trigger for standard device
        # value = 0 - 127 of control
        # std = True if this is set as standard
        # init = True if this is initialized
        self.entrys = []
        self.chentry = None

        self.savedir = os.getenv("HOMEDRIVE") + os.getenv("HOMEPATH")
        savedir = os.path.join(self.savedir, "midivolctrl")
        if not os.path.exists(savedir):
            os.makedirs(savedir)
        self.savedir = savedir
        self.load()

    def onexpose(self, e):
        w = e.widget
        if not w.children:
            w.configure(height=1)

    def resetonzero(self):
        self.rstcount -= 1
        if self.rstcount <= 0:
            self.rstcount = 50
            self.win.title("Midi Volume Controller")
        self.win.after(100, self.resetonzero)

    def setstdmode(self):
        self.win.title("set standard button")
        self.rstcount = 50
        self.chmode = "std"

    def setvolmode(self):
        self.win.title("set volume fader")
        self.rstcount = 50
        self.chmode = "vol"

    def setmutmode(self):
        self.win.title("set mute button")
        self.rstcount = 50
        self.chmode = "mut"

    def unconfigure(self, ntry):
        if self.chentry is not None:
            self.chentry["label"].config(fg = self.oldfg)
        self.chentry = None

    def startconfigure(self, ntry, evt):
        if ntry["type"] == "microphone" or ntry["type"] == "speaker":
            if evt.x <= 16:
                self.setmutmode()
            elif evt.x > 16 and evt.x < 460:
                self.setstdmode()
            else:
                self.setvolmode()
        else:
            self.setvolmode()
        
        if self.chentry is not None:
            self.chentry["label"].config(fg = self.oldfg)
        if self.chentry == ntry:
            self.chentry = None
            return
        
        self.chentry = ntry
        self.oldfg = ntry["label"]["fg"]
        ntry["label"].config(fg = gc.mkassi)

    def removeentry(self, ntry):
        try:
            self.entrys.remove(ntry)
            ntry["label"].destroy()
            self.save()
        except:
            pass

    def checkpos(self, event):
        self.win.title(f"X={event.x}")
        self.rstcount = 50

    def updateentry(self, ntry):
        nmut = 0
        if "mut" in ntry:
            nmut = ntry["mut"]
        ntype = ntry["type"]
        ncont = ntry["content"][:46]
        nval = int(ntry["value"] / 1.27)
        numr = ntry["value"] // 3
        numu = 42 - numr
        ln = 48 - len(ncont)
        empty = ""
        if nmut == 0:
            txt = f"{ncont}{empty: <{ln}} {nval:3} [{empty:#<{numr}}{empty:_<{numu}}]"
        else:
            txt = f"{ncont}{empty: <{ln}} --- [{empty:#<{numr}}{empty:_<{numu}}]"
        if ntype == "script":
            txt = ntry["content"][-80:]
            icn = self.runpic
            dstframe = self.scripts
        if ntype == "microphone":
            icn = self.micpic
            dstframe = self.microphones
        if ntype == "speaker":
            icn = self.spkpic
            dstframe = self.speakers
        if ntype == "app":
            icn = self.apppic
            dstframe = self.apps

        if "label" not in ntry:
            ntry["label"] = tk.Label(dstframe, justify=tk.LEFT, compound=tk.LEFT, text=txt, image=icn,
                font=self.mainfont, bg=gc.fieldback, fg=gc.noassi,
                anchor="w")
            ntry["label"].bind("<Button-1>", lambda event, lntry=ntry : self.startconfigure(lntry, event))
            ntry["label"].bind("<Button-3>", lambda event, lntry=ntry : self.removeentry(lntry))
            #ntry["label"].bind("<Button-2>", self.checkpos)
            ntry["label"].pack(fill=tk.X)
        

        #print("updateentry", id(ntry), id(ntry["label"]), ntry["content"], txt)
        ntry["label"].config(text=txt)
        color = gc.noassi
        if ntry["init"]:
            color = gc.isassi
            if ntype == "script":
                color = gc.isscriptassi
        if ntry["std"]:
            color = gc.isstd
        if ntry["label"]["fg"] != gc.mkassi:
            ntry["label"].config(fg=color)
    
    def addentry(self, ntry):
        self.entrys.append(ntry)
        ntype = ntry["type"]
        ncont = ntry["content"][:46]
        nval = int(ntry["value"] / 1.27)
        numr = ntry["value"] // 3
        numu = 42 - numr
        ln = 48 - len(ncont)
        empty = ""
        txt = f"{ncont}{empty: <{ln}} {nval:3} [{empty:#<{numr}}{empty:_<{numu}}]"
        if ntype == "script":
            txt = ntry["content"][-80:]
            icn = self.runpic
            dstframe = self.scripts
        if ntype == "microphone":
            icn = self.micpic
            dstframe = self.microphones
        if ntype == "speaker":
            icn = self.spkpic
            dstframe = self.speakers
        if ntype == "app":
            icn = self.apppic
            dstframe = self.apps

        ntry["label"] = tk.Label(dstframe, justify=tk.LEFT, compound=tk.LEFT, text=txt, image=icn,
            font=self.mainfont, bg=gc.fieldback, fg=gc.noassi,
            anchor="w")
        ntry["label"].bind("<Button-1>", lambda event, lntry=ntry : self.startconfigure(lntry, event))
        ntry["label"].bind("<Button-3>", lambda event, lntry=ntry : self.removeentry(lntry))
        #ntry["label"].bind("<Button-2>", self.checkpos)
        ntry["label"].pack(fill=tk.X)
        
    def addscript(self):
        entry = self.mkscript.get()
        d = {"type": "script", "content": entry,
             "trigger": "", "value": 0, "std": False, "init": False}
        self.addentry(d)
        self.mkscript.delete(0, tk.END)
        self.save()

    def load(self):
        try:
            fname = os.path.join(self.savedir, "midivolctrl.json")
            f = open(fname, "r")
            self.entrys = json.loads(f.read())
            f.close()
            for entry in self.entrys:
                self.updateentry(entry)
        except:
            self.entrys = []

    def save(self):
        try:
            entrys = []
            for ntry in self.entrys:
                tmpntry = ntry.copy()
                tmpntry.pop("label", None)
                tmpntry["value"] = min(127, tmpntry["value"])
                if tmpntry["init"]:
                    entrys.append(tmpntry)
            
            fname = os.path.join(self.savedir, "midivolctrl.json")
            f = open(fname, "w")
            f.write(json.dumps(entrys))
            f.close()
        except:
            pass

    def mainloop(self):
        self.win.mainloop()

def main():
    gui = GUI()
    sound = WinSound()
    ctrl = Controller()

    def onselect(event):
        widget = event.widget
        index = int(widget.curselection()[0])
        val = widget.get(index)
        ctrl.select_midi(val)
        gui.selector.destroy()
        gui.win.attributes("-disabled", False)
        gui.win.attributes("-topmost", 1)
        gui.win.attributes("-topmost", 0)
        gui.win.after(50, lambda : ctrl.checkmidi(gui, sound))
        gui.win.after(100, lambda : sound.listdevices(gui))
        gui.win.after(500, lambda : sound.listapps(gui))
        gui.win.after(700, gui.resetonzero)
        gui.win.after(2000, lambda : sound.enumdevices(gui))

    def disable_event():
        pass

    if ctrl.selected:
        gui.win.attributes("-disabled", False)
        gui.win.attributes("-topmost", 1)
        gui.win.attributes("-topmost", 0)
        gui.win.after(50, lambda : ctrl.checkmidi(gui, sound))
        gui.win.after(100, lambda : sound.listdevices(gui))
        gui.win.after(500, lambda : sound.listapps(gui))
        gui.win.after(700, gui.resetonzero)
        gui.win.after(2000, lambda : sound.enumdevices(gui))

    else:
        gui.selector = tk.Toplevel(gui.win)
        gui.selector.resizable(0,0)
        gui.selector.title("Select Midi Controller")
        gui.selector.iconbitmap("controller.ico")
        gui.selector.protocol("WM_DELETE_WINDOW", disable_event)
        gui.lbox = tk.Listbox(gui.selector, selectmode=tk.SINGLE, height=len(ctrl.portlist),
            width=30, font=gui.bigfont, bg=gc.fieldback, fg=gc.noassi)
        for x in ctrl.portlist:
            gui.lbox.insert(tk.END, x)
        gui.lbox.bind("<<ListboxSelect>>", onselect)
        gui.lbox.pack()

    gui.mainloop()


if __name__ == "__main__":
    main()
