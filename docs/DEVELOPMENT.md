# Development

Please use an editor that offers linting information and
recommends pep8 and identifies possible errors.

Don't fall in the temptation to change indentation or things that
don't sum up.

Typos and fixes are welcome, improvement on wording and clarity are
good for everyone.

As submodule is aihordeclient to avoid installing packages. It handles
all the communication with aihorde.

## Getting the sources for the first time

```
git clone --recurse-submodules https://github.com/ikks/libreoffice-stable-diffusion.git
```

[uv](https://docs.astral.sh/uv/) is used to prepare urllib3, which is
vendored in the .oxt file.

## Help and resources

## Testing distribution, LibreOffice and flatpak

It should be possible to run the script to show the dialog, this
allows iterating a lot faster when building the interface.

```
python src/StableHordeForLibreOffice.py 
```

To create the oxt, type `make`, to prepare the languages,
`make langs` is needed.  It's not included in make, because
it tends to change due to line numbers changes.

To build, install and launch libreoffice issue `make run`

### Linux

Look in Makefile to test flatpak, usually from CLI can be
invoked via

 flatpak run org.libreoffice.LibreOffice

LibreOffice distributes packages that do not clash with distro
packages. Are installed in /opt/libreofficeYY.M/program/

Linux distribution binaries are usually in /usr/bin

Envvars when running make can be changed with

PREFIX : By default point to /usr/bin , can be changed to /opt/libre...
UNOPKG : The name of the extension installer
LOBIN : The name of LibreOffice executable


### Interaction

* [💬 irc #libreoffice-dev](https://web.libera.chat)
* [🐛 reporting bugs to libreoffice](https://bugs.documentfoundation.org/)
* [🙋 forum](https://ask.libreoffice.org/)

### Local documentation

* file:///usr/share/doc/libreoffice-dev-doc/api/
* file:///usr/lib/libreoffice/share/Scripts/python/
* file:///usr/lib/libreoffice/sdk/examples/html


## Helpers

Install apso and use the python console that comes with it, easier
to get help from there.

### Some recipes

* [WIP wiki](https://wiki.documentfoundation.org/User:Ikks/RawPython)

#### Ways to get context

```
doc = XSCRIPTCONTEXT.getDocument()
doc.get_current_controller().ComponentWindow.StyleSettings.LightColor.is_dark()


# Inside APSO
import uno
ctx = uno.getComponentContext()
doc = uno.getComponentContext().getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", ctx).getCurrentComponent()
```

