# dragonfly-commands

My commands for Dragonfly, an extension to Dragon NaturallySpeaking.

These command modules are intended as a reference supplement to my blog,
[handsfreecoding.org](http://handsfreecoding.org). They are a work in progress,
so expect breaking changes. Currently, many commands are specific to my setup,
so I recommend using this as a source of ideas and examples instead of
directly. If you wish to make it easier for others to use directly, I am open to
a pull request, but please do so in a way that will not slow down development
(i.e. follow the Don't Repeat Yourself principle).

Several of these modules are based on examples from the [official
dragonfly-modules repository](https://github.com/t4ngo/dragonfly-modules).

## Installation

If you do choose to use this code directly, here are the basic steps:

1. Install Dragon, NatLink, and dragonfly.
2. [Modify dragonfly to export RuleWrap](https://github.com/wolfmanstout/dragonfly/blob/2a3161b59c76fdda7e009fefeda9503c4b4631f2/dragonfly/__init__.py#L37).
3. Install [Python bindings for WebDriver](https://pypi.python.org/pypi/selenium).
4. (Optional) Update Chrome to listen on port 9222 ([instructions](http://handsfreecoding.org/2015/02/21/custom-web-commands-with-webdriver/)).
5. Copy the contents of this repository into your macros directory (typically the MacroSystem directory).
6. Rename ```_dragonfly_local.py.template``` to ```_dragonfly_local.py```.
7. Restart Dragon.

Those are the basic steps needed to get the code to run without errors. Some
interesting functionality will still be missing (e.g. eye tracking, WebDriver
integration). Please check out my blog for instructions on integrating these
features (hint: you will need to update ```_dragonfly_local.py``` if you wish to
use my built-in grammars).
