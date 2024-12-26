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
        if len(self.portlist) == 0:
            self.errmsg = "no midi"
            return
        if len(self.portlist) == 4:
            name = self.portlist[0]
            self.inport = self.pygame.open_input()
            self.errmsg =  f"midi {name} selected"
            return
        if len(self.portlist) == 1:
            self.errmsg = "select midi"
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
                            entry["init"] = True
                            gui.save()
                        if pre == "n":
                            if gui.chmode == "vol":
                                entry["trigger"] = f"n{msg.note}"
                            if gui.chmode == "std":
                                entry["sectrig"] = f"n{msg.note}"
                            entry["init"] = True
                            gui.save()

                    # check if this entry was triggered
                    ntrig = entry["trigger"]
                    if ntrig.startswith("c") and msg.is_cc():
                        if int(ntrig[1:]) == msg.control:
                            entry["value"] = msg.value
                            self.triggercmd(entry, sound, gui, "vol")
                    if ntrig.startswith("n") and msg.type == "note_on":
                        if int(ntrig[1:]) == msg.note:
                            if ntype != "app":
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

        # gv means given -> given gui and given sound
        gui.win.after(50, lambda gvgui=gui, gvsound=sound : self.checkmidi(gvgui, gvsound))

    # function to call when controller input triggers a command
    def triggercmd(self, ntry, snd, gui, mod):
        typ = ntry["type"]
        con = ntry["content"]
        val = ntry["value"]
        if typ == "microphone" and mod == "vol":
            snd.change_mic_vol(con, val)
            gui.updateentry(ntry)
        if typ == "speaker" and mod == "vol":
            snd.change_spk_vol(con, val)
            gui.updateentry(ntry)
        if typ == "microphone" and mod == "std":
            snd.change_default_mic(con)
            for e in gui.entrys:
                e["std"] = False
                gui.updateentry(e)
            ntry["std"] = True
            gui.updateentry(ntry)
        if typ == "speaker" and mod == "std":
            snd.change_default_spk(con)
            for e in gui.entrys:
                e["std"] = False
                gui.updateentry(e)
            ntry["std"] = True
            gui.updateentry(ntry)
        if typ == "app":
            snd.change_app_vol(con, val)
            gui.updateentry(ntry)
        if typ == "script":
            sp.Popen(shlex.split(con))

class WinSound:
    def __init__(self):
        comtypes.CoInitialize()

        # the device enumerator
        self.devEnum = comtypes.CoCreateInstance(cawstants.CLSID_MMDeviceEnumerator,
            caw.IMMDeviceEnumerator, comtypes.CLSCTX_INPROC_SERVER)

        # important values
        micval = caw.EDataFlow.eCapture.value
        spkval = caw.EDataFlow.eRender.value
        state = caw.DEVICE_STATE.ACTIVE.value

        # collections of devices
        self.micdevs = []
        micdevs = self.devEnum.EnumAudioEndpoints(micval, state)
        self.spkdevs = []
        spkdevs = self.devEnum.EnumAudioEndpoints(spkval, state)
        
        # copy collections to lists
        self.copydevs(micdevs, self.micdevs)
        self.copydevs(spkdevs, self.spkdevs)

        # the policy config needed for changing standard devices
        self.policy_config = comtypes.CoCreateInstance(pc.CLSID_PolicyConfigClient,
            pc.IPolicyConfig, comtypes.CLSCTX_ALL)

    # copy a device collection to destination list
    def copydevs(self, srcdevs, dstdevs):
        for i in range(srcdevs.GetCount()):
            dev = srcdevs.Item(i)
            dstdevs.append(caw.AudioUtilities.CreateDevice(dev))

    # call this to change volume of an app
    def change_app_vol(self, appname, value, changeit=True):
        sessions = caw.AudioUtilities.GetAllSessions()
        for session in sessions:
            volume = session._ctl.QueryInterface(caw.ISimpleAudioVolume)
            if session.Process and session.Process.name() == appname:
                if changeit:
                    volume.SetMasterVolume(int(value / 1.27) / 100, None)
                else:
                    return int(volume.GetMasterVolume() * 127.0)

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
            ln = len(gui.entrys)
            matched = False
            for n in range(ln):
                gtype = gui.entrys[n]["type"]
                gcont = gui.entrys[n]["content"]
                if gtype == dsttype and gcont == devname:
                    gui.entrys[n]["value"] = value
                    gui.updateentry(gui.entrys[n])
                    matched = True
            if matched == False:
                d = {"type": dsttype, "content": devname,
                     "trigger": "", "value": value, "std": False, "init": False}
                gui.addentry(d)
    
    def listdevices(self, gui):
        self.list_devs(gui, self.micdevs, "microphone")
        self.list_devs(gui, self.spkdevs, "speaker")
        gui.win.after(1000, lambda : self.listdevices(gui))

    def listapps(self, gui):
        sessions = caw.AudioUtilities.GetAllSessions()
        for session in sessions:
            volume = session._ctl.QueryInterface(caw.ISimpleAudioVolume)
            if session.Process:
                appname = session.Process.name()
                value = self.change_app_vol(appname, 0, changeit=False)
                matched = False
                ln = len(gui.entrys)
                for n in range(ln):
                    gtype = gui.entrys[n]["type"]
                    gcont = gui.entrys[n]["content"]
                    if gtype == "app" and gcont == appname:
                        gui.entrys[n]["value"] = value
                        gui.updateentry(gui.entrys[n])
                        matched = True
                if matched == False:
                    d = {"type": "app", "content": appname,
                         "trigger": "", "value": value, "std": False, "init": False}
                    gui.addentry(d)
        gui.win.after(500, lambda : self.listapps(gui))

class GUI:
    def __init__(self):
        self.win = tk.Tk()
        self.win.config(bg=gc.mainspace)
        self.mainfont = tkfont.Font(file="space.ttf",
            family="Space Mono", size=11, weight="normal")
        self.bigfont = tkfont.Font(family="Space Mono",
            size=24, weight="bold")
        self.win.resizable(0,0)
        self.win.title("Midi Volume Controller")
        self.win.iconbitmap("controller.ico")
        self.win.attributes("-disabled", True)

        self.chmode = "vol"
        self.question = tk.Frame(self.win, bg=gc.fieldback)
        self.question.pack(fill=tk.X, pady=8)
        self.question.grid_columnconfigure((0,1), weight=1)
        # grid_columnconfigure makes content centered
        # in grid sticky=ew makes content stretched left and right
        self.answerstd = tk.Button(self.question, text="make standard", command=self.setstdmode,
            font=self.mainfont, border=4)
        self.answerstd.grid(column=0, row=0, sticky="ew")
        self.answervol = tk.Button(self.question, text="✓ change volume ✓", command=self.setvolmode,
            font=self.mainfont, border=4)
        self.answervol.grid(column=1, row=0, sticky="ew")
        self.speakers = tk.Frame(self.win)
        self.speakers.pack(fill=tk.X, pady=1)
        self.microphones = tk.Frame(self.win)
        self.microphones.pack(fill=tk.X, pady=6)
        self.apps = tk.Frame(self.win)
        self.apps.pack(fill=tk.X, pady=1)
        self.scripts = tk.Frame(self.win)
        self.scripts.pack(fill=tk.X, pady=6)
        self.mkscript = tk.Entry(self.win, font=self.mainfont,
            bg=gc.fieldback, fg=gc.scripts)
        self.mkscript.bind("<Return>", self.addscript)
        self.mkscript.pack(fill=tk.X, pady=1)

        # type = speaker, microphone, app, script
        # content = name of speaker, microphone, app OR the script to run
        # trigger = trigger for volume -> note or control -> n or c followed by ID
        # sectrig = trigger for standard device
        # value = 0 - 127 of control
        # std = True if this is set as standard
        # init = True if this is initialized
        self.entrys = []
        self.chentry = None
        self.load()

    def setstdmode(self):
        self.chmode = "std"
        self.answerstd.config(text="✓ make standard ✓")
        self.answervol.config(text="change volume")

    def setvolmode(self):
        self.chmode = "vol"
        self.answerstd.config(text="make standard")
        self.answervol.config(text="✓ change volume ✓")

    def unconfigure(self, ntry):
        if self.chentry is not None:
            self.chentry["label"].config(fg = self.oldfg)
        self.chentry = None

    def startconfigure(self, ntry):
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

    def updateentry(self, ntry):
        ntype = ntry["type"]
        ncont = ntry["content"][:48]
        nval = int(ntry["value"] / 1.27)
        numr = ntry["value"] // 3
        numu = 42 - numr
        ln = 50 - len(ncont)
        empty = ""
        txt = f"{ncont}{empty: <{ln}} {nval:3} [{empty:#<{numr}}{empty:_<{numu}}]"
        if ntype == "script":
            txt = "> " + ncont
            dstframe = self.apps
        if ntype == "microphone":
            dstframe = self.microphones
        if ntype == "speaker":
            dstframe = self.speakers
        if ntype == "app":
            dstframe = self.apps

        if "label" not in ntry:
            ntry["label"] = tk.Label(dstframe, text=txt,
                font=self.mainfont, bg=gc.fieldback, fg=gc.noassi,
                anchor="w")
            ntry["label"].bind("<Button-1>", lambda event, lntry=ntry : self.startconfigure(lntry))
            ntry["label"].bind("<Button-3>", lambda event, lntry=ntry : self.removeentry(lntry))
            ntry["label"].pack(fill=tk.X)
        

        ntry["label"].config(text=txt)
        color = gc.noassi
        if ntry["init"]:
            color = gc.isassi
        if ntry["std"]:
            color = gc.isstd
        if ntry["label"]["fg"] != gc.mkassi:
            ntry["label"].config(fg=color)
    
    def addentry(self, ntry):
        self.entrys.append(ntry)
        ntype = ntry["type"]
        ncont = ntry["content"][:48]
        nval = int(ntry["value"] / 1.27)
        numr = ntry["value"] // 3
        numu = 42 - numr
        ln = 50 - len(ncont)
        empty = ""
        txt = f"{ncont}{empty: <{ln}} {nval:3} [{empty:#<{numr}}{empty:_<{numu}}]"
        if ntype == "script":
            txt = "> " + ncont
            dstframe = self.apps
        if ntype == "microphone":
            dstframe = self.microphones
        if ntype == "speaker":
            dstframe = self.speakers
        if ntype == "app":
            dstframe = self.apps

        ntry["label"] = tk.Label(dstframe, text=txt,
            font=self.mainfont, bg=gc.fieldback, fg=gc.noassi,
            anchor="w")
        ntry["label"].bind("<Button-1>", lambda event, lntry=ntry : self.startconfigure(lntry))
        ntry["label"].bind("<Button-3>", lambda event, lntry=ntry : self.removeentry(lntry))
        ntry["label"].pack(fill=tk.X)
        
    def addscript(self, event):
        entry = self.mkscript.get()
        d = {"type": "script", "content": entry,
             "trigger": "", "value": -1, "std": False, "init": False}
        self.entrys.append(d)
        self.mkscript.delete(0, tk.END)
        # do i really need to insert nothing?
        self.mkscript.insert(0, "")
        d["label"] = tk.Label(self.scripts, text="> " + entry[:80],
            font=self.mainfont, bg=gc.fieldback, fg=gc.scripts)
        d["label"].bind("<Button-1>", lambda event, ntry=d : self.startconfigure(ntry))
        d["label"].bind("<Button-3>", lambda event, ntry=d : self.removeentry(ntry))
        
    def load(self):
        try:
            f = open("mivoco.json", "r")
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
            
            f = open("mivoco.json", "w")
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
        gui.win.after(50, lambda : ctrl.checkmidi(gui, sound))
        gui.win.after(100, lambda : sound.listdevices(gui))
        gui.win.after(500, lambda : sound.listapps(gui))

    def disable_event():
        pass
    
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
