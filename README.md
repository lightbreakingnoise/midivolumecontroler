# midivolumecontroler
## change volume levels of apps, speaker and microphone with your midi controler

actually it only works on windows

it needs python!<br>
you can install it with [chocolatey](https://chocolatey.org/install)

```Shell
choco install python -y
```
enter this in administrator terminal

or just go to [python.org](https://www.python.org/), then download and install the setup (tick "add python.exe to path")

after setup, my midivolumecontroler needs a few python libraries

```Shell
pip install tk tkextrafont mido pygame comtypes pycaw
```
enter this in administrator terminal

# use it
as soon as it has started, with your default midi controler, you should see output/input devices and sound producing apps. Just click in one of these lines, it should go red, then turn a knob/fader and it should go green. you can now change it's volume with this knob/fader. just right click in one line to remove a link.

update:
now, thanks to KillerBOSS2019 you can assign a note or control on your midi controler after clicking on a device when "make standard" is selected. this makes the device standard AND standard communication everytime you tap this note/control.

update:
there is a text entry field where you can enter app to start (with parameters if you want). after adding app you can link this app with a note.

# credits
[KillerBOSS2019](https://github.com/KillerBOSS2019) and [kdschlosserfor](https://github.com/kdschlosser) for the [policyconfig.py](/app/policyconfig.py) file.

when you give [RetroBar](https://github.com/dremin/RetroBar) a try, you can see last volume change in your taskbar via window title.

# screenshot
![screenshot](/screenshot_005.png?raw=true)
