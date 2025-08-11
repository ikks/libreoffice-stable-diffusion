# How to contribute

Filling an issue is one way, sending a pull request, translating
or suggesting a feature. Also if you would like, please contact
the author to buy him a coffee.

## Translating

If you wish to help with a translation, please submit an
[issue](https://github.com/ikks/libreoffice-stable-diffusion/issues)
indicating it.

Steps are usually:

1. [File an issue](https://github.com/ikks/libreoffice-stable-diffusion/issues)
  telling it so, to avoid reprocess.
1. Clone the repository
1. Edit the files
1. Make the pull request

There are four files to translate in this repository and
the [description of the extension](https://extensions.libreoffice.org/en/extensions/show/99431).
In the issue, please include the description of the extension.  Can be
plain text or markdown.  Ideally attach two or three screenshots to show
how it looks in your language.

The files from the repository to translate are:

1. oxt/description/description_en.txt : The description when installing
1. oxt/description.xml: Name of the extension
1. oxt/Accelerators.xcu : Shortcut Ctrl+Shift+H
1. oxt/Addons.xcu : The menu and the toolbar tip
1. src/po/messages.pot : Dialog, messages and errors

For the rest of this section we will use as sample, Spanish (es),
please adjust to your language, using the
[ISO 639 Code](https://en.wikipedia.org/wiki/List_of_ISO_639_language_codes),
it's also possible to use
[ISO639-2](https://www.loc.gov/standards/iso639-2/php/English_list.php)
if your language does not have a two letter code.


### Explanation of each step

Please add your translation in alphabetical order, to
allow others find a consistent sample.

#### Description: text file

Copy `oxt/description/description_en.txt` to
`oxt/description/description_es.txt` and put your translation there.

#### Name: XML file

Edit `oxt/description.xml` and look for the tag `display-name` add
inside it a tag `name` with the property `lang` with the value `es`
like:

```xml
<name lang="es">Cliente de Stable Horde para LibreOffice</name>
```

Additionally we need to link the description, look for the tag
`extension-description` and add inside it a `src` tag with
the property `lang` with the value `es` and the property
`xlink:href` with the value `description/description_es.txt` like:

```xml
<src xlink:href="description/description_es.txt" lang="es"/>
```

#### Menu and Toolbar: XML file

Edit `oxt/Addons.xcu` and look for lines that contain something like
`<value xml:lang="en-US"...`, add below each of them a line with
your corresponding language. There should be added two entries,
one for the menu and the other for the toolbar tooltip.

#### The messages: PO file

When you open the issue, you can copy the file `src/po/messages.pot`
to `oxt/locale/es.po` and use a tool like [poedit](https://poedit.net/) or the one you
prefer to make the translation.

### PR

Before PR, please use an spell checker to make sure there isn't a typo
or something that can lower your personal signature.

Thanks for helping with it.

