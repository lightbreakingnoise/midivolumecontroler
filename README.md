# midivolumecontroller
### change volume levels of apps, speakers and microphones with your midi controller<br>even running commands by pressing button on your midi controller is possible

> [!WARNING]
> actually it **only** works on windows

<br><br>
> [!NOTE]
> **it needs python!**

you can install it with [chocolatey](https://chocolatey.org/install)<br>
> ```Shell
> choco install python -y
> ```
> enter this in administrator terminal

or you can install it regular
> just go to [python.org](https://www.python.org/),<br>
> then download and install the setup (tick _"add python.exe to path"_)

<br><br>
after python is installed correctly, my midivolumecontroller needs a few python libraries<br>
> ```Shell
> pip install tk tkextrafont mido pygame comtypes pycaw
> ```
> enter this in administrator terminal

<br><br>
## use it
> as soon as it has started, choose your midi controler, then you should see output/input devices and sound producing apps.<br>
> Just click in one of these lines, it should go red, then turn a knob/fader and it should go green.<br>
> you can now change it's volume with this knob/fader. just right click in one line to remove an assignment.

> update:<br>
> now, thanks to KillerBOSS2019 you can assign a note or control on your midi controler after clicking on a device when "make standard" is selected.<br>
> this makes the device standard **AND** standard communication everytime you tap this note/button.

> update:<br>
> there is a text entry field where you can enter app/command to start (with parameters if you want).<br>
> after adding app/command you can assign this with a note/button.

<br><br>
## credits
> [KillerBOSS2019](https://github.com/KillerBOSS2019)<br>
> [kdschlosserfor](https://github.com/kdschlosser)<br>
> for the [policyconfig.py](/app/policyconfig.py) file.

when you give [RetroBar](https://github.com/dremin/RetroBar) a try, you can see last volume change in your taskbar via window title.

<br>
## screenshot

![screenshot](/screenshot_005.png?raw=true)
