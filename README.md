# dragonfly-commands

My commands for Dragonfly, an extension to Dragon NaturallySpeaking.

These command modules are intended as a reference supplement to my blog,
[handsfreecoding.org](http://handsfreecoding.org). Currently, most of the
commands are stuffed into `_repeat.py`. They are a work in progress, so expect
breaking changes. Many commands are specific to my setup, so I recommend using
this as a source of ideas and examples instead of directly. If you wish to make
it easier for others to use directly, I am open to pull requests, but please do
so in a way that will not slow down development (i.e. follow the Don't Repeat
Yourself principle).

Several of these modules are based on examples from the [official
dragonfly-modules repository](https://github.com/t4ngo/dragonfly-modules).

## Installation

If you do choose to use this code directly, here are the basic steps:

1. Install Dragon, NatLink, and Dragonfly. I generally recommend `pip install
   dragonfly2` to install Dragonfly, but occasionally this grammar may use
   features only available in my [development
   branch](https://github.com/wolfmanstout/dragonfly/tree/develop) that have not
   yet been integrated upstream.
2. Install dependencies: `pip install -r requirements.txt`
3. (Optional) Update Chrome to listen on port 9222
   ([instructions](http://handsfreecoding.org/2015/02/21/custom-web-commands-with-webdriver/)).
4. Copy the contents of this repository into your macros directory (typically
   the MacroSystem directory).
5. Rename ```_dragonfly_local.py.template``` to ```_dragonfly_local.py```.
6. Restart Dragon.

Those are the basic steps needed to get the code to run without errors. Some
interesting functionality will still be missing (e.g. eye tracking, WebDriver
integration). Here is how to integrate eye tracking:

1. Run `pip install pythonnet`.
2. Download the [latest
   Tobii.Interaction](https://www.nuget.org/packages/Tobii.Interaction/) package
   from NuGet (these instructions have been tested on 0.7.3).
3. Rename the file extension to .zip and expand the contents.
4. Copy these 3 DLLs to a directory of your choice:
   build/AnyCPU/Tobii.EyeX.Client.dll, lib/net45/Tobii.Interaction.Model.dll,
   lib/net45/Tobii.Interaction.Net.dll.
5. Ensure that the files are not blocked (right-click Properties, and if there
   is a "Security" section at the bottom, check the "Unblock" box.)
6. Set `DLL_DIRECTORY` in `_dragonfly_local.py` to point to the directory used
   in the previous step.

Please check out [my blog](http://handsfreecoding.org) for instructions integrating
other optional features.
