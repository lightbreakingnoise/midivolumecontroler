# midivolumecontroler
### change volume levels of apps, speaker and microphone with your midi controler

> [!NOTE]
> actually it **only** works on windows

**it needs python!**

> you can install it with [chocolatey](https://chocolatey.org/install)<br>
> ```Shell
> choco install python -y
> ```
> enter this in administrator terminal

**or just go to [python.org](https://www.python.org/),<br>
then download and install the setup (tick _"add python.exe to path"_)**

> after setup, my midivolumecontroler needs a few python libraries<br>
> ```Shell
> pip install tk tkextrafont mido pygame comtypes pycaw
> ```
> enter this in administrator terminal

## use it
> as soon as it has started, choose your midi controler, then you should see output/input devices and sound producing apps.<br>
> Just click in one of these lines, it should go red, then turn a knob/fader and it should go green.<br>
> you can now change it's volume with this knob/fader. just right click in one line to remove a link.

> update:<br>
> now, thanks to KillerBOSS2019 you can assign a note or control on your midi controler after clicking on a device when "make standard" is selected.<br>
> this makes the device standard **AND** standard communication everytime you tap this note/control.

> update:<br>
> there is a text entry field where you can enter app/command to start (with parameters if you want).<br>
> after adding app/command you can assign this with a note/button.

## credits
> [KillerBOSS2019](https://github.com/KillerBOSS2019) and<br>
> [kdschlosserfor](https://github.com/kdschlosser)<br>
> for the [policyconfig.py](/app/policyconfig.py) file.<br>
when you give [RetroBar](https://github.com/dremin/RetroBar) a try, you can see last volume change in your taskbar via window title.

## screenshot
![screenshot](/screenshot_005.png?raw=true)
