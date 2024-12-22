# midivolumecontroler
change volume levels of apps, speaker and microphone with your midi controler

actually it only works on windows

it needs python!
you can install it with

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
as soon as it has started, with your default midi controler, you should see Speakers, Microphone and sound producing apps. Just click in one of these lines, it should go red, then turn a knob/fader and it should go green. you can now change it's volume with this knob/fader. just right click in one line to remove a link. now you can tap a note on your midi controler after clicking on a device. this makes the device standard AND standard communication everytime you tap this note. there is a text entry field where you can enter app (with parameters if you want) to start. after adding app you can link this app with a note. with mousewheel you can change font size and that changes window size.

when you give RetroBar a try, you can see last volume change in your taskbar via window title.

![screenshot](/screenshot.png?raw=true&ver=0)
