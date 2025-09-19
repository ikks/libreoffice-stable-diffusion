# libreoffice-stable-diffusion-horde

Creating images from text. That is what this extension for
[LibreOffice](https://libreoffice.org) does. Making use of
[AIHorde](https://aihorde.net).

AIHorde is a cluster of stable-diffusion servers run by
volunteers. You can create stable-diffusion images for free without
running a colab notebook or a local server. Please check the section
[Limitations](README#Limitations) to better understand where the
limits are.

Please check [CHANGELOG.md](../CHANGELOG.md) for the latest
changes.

## Installation

### If you are on Linux Debian based Distribution

Please first install `libreoffice-script-provider-python`
with your preferred package manager, alternatively you can issue:

```
sudo apt install -y libreoffice-script-provider-python
```

### Download the extension 

Download [loshd.oxt](https://raw.githubusercontent.com/ikks/libreoffice-stable-diffusion/refs/heads/main/loshd.oxt).

### LibreOffice

This Extension is known to work from LibreOffice 7.4 and upwards

1. Once you [Downloaded](https://raw.githubusercontent.com/ikks/libreoffice-stable-diffusion/refs/heads/main/src/loshd.oxt)
  the extension.
2. Open LibreOffice, go to Tools > Extensions...
  in the Dialog push the button Add , browse and select
  the Downloaded file `loshd.oxt`. Accept the License, Close and Restart LibreOffice.
<img width="1366" height="768" alt="annotated01" src="https://github.com/user-attachments/assets/8b8a5996-3dc8-48eb-bea2-4522afe584ab" />
3. Reopen LibreOffice. You should see a menu entry under `Insert`
  Menu that reads `Image from Text...` and a new colorful button
  in the toolbar to trigger the action. Additionally, you
  will have the shortcut `Ctrl+Shift+h` (Cmd+Shift+h on Mac)
  configured to launch the extension.
<img width="1366" height="768" alt="annotated03" src="https://github.com/user-attachments/assets/b7d83df2-a129-460a-86c6-d8bdd7041454" />
You are ready to

## Generate images

1. Launch the extension with `Ctrl+Shift+h`, the menu or
the button toolbar, a dialog will appear allowing you
to write a prompt and generate the image, if you don't have a
Writer document opened, the extension will open it for you.
<img width="1920" height="1080" alt="do1" src="https://github.com/user-attachments/assets/f4a772ff-890b-4a51-9897-7ee8b103d13d" />

You can select a text, and when the macro is launched, the text
will be used as a prompt.  The progress bar will guide you on what
is happening; at the end of the process, an image will be inserted
before your selected text or in the cursor position.

<details>
  <summary>Details on using the plugin</summary>

- **Prompt:** How the generated image should look like.

- **Max Wait:** The maximum time in minutes you want to wait until
image generation is finished. When the max time is reached, a timeout
happens and the generation request is stopped.

- **Strength:** How much the AI should follow the prompt. The
higher the value, the more the AI will generate an image which looks
like your prompt. 7.5 is a good value to use.

- **Steps:** How many steps the AI should use to generate the
image. The higher the value, the more the AI will work on details. But
it also means, the longer the generation takes and the more the GPU
is used. 50 is a good value to use.

- **NSFW:** If you want to send a prompt, which is explicitly NSFW
(Not Safe For Work).
    - If you flag your request as NSFW, only servers, which accept
    NSFW prompts, work on the request. It's very likely, that it
    takes then longer than usual to generate the image. If you don't
    flag the prompt, but it is NSFW, you will receive a black image.
    - If you didn't flag your request as NSFW and don't prompt NSFW,
    you will receive in some cases a black image, although it's not
    NSFW (false positive). Just rerun the generation in that case.

- **Seed:** This parameter is optional. If it is empty, a random seed
will be generated on the server. If you use a seed, the same image is
generated again in the case the same parameters for init strength,
steps, etc. are used. A slightly different image will be generated,
if the parameters are modified. You find the seed in an additional
layer at the top left.

- **API key:** This parameter is optional. If you don't enter an
API key, you run the image generation as anonymous. The downside
is, that you will have then the lowest priority in the generation
queue. For that reason it is recommended registering for free on
[AIHorde](https://aihorde.net) and getting an API key.

2. Click on the OK button. The values you inserted into the dialog
will be transmitted to the server, which dispatches the request now to
one of the stable-diffusion servers in the cluster. Your generation
request is added to queue. You can see the queue position in the
status bar. When the image has been generated successfully, it will
be shown as a new image in LibreOffice inside the current document.

## Limitations
- **Generation speed:** AIHorde is a cluster of stable-diffusion
servers run by volunteers. The generation speed depends on how many
servers are in the cluster, which hardware they use and how many others
want to generate with AIHorde. The upside is, that AIHorde is
free to use, the downside that the generation speed is unpredictable.

- **Privacy:** The privacy AIHorde offers is similar to generating
in a public discord channel. So, please assume, that neither your
prompts nor your generated images are private.

- **Features:** Currently text2img if you are looking for img2img and
inpainting consider the
[Gimp Stable Horde plugin](https://github.com/ikks/gimp-stable-diffusion/)

</details>
  
## Troubleshooting

### LibreOffice

##### Linux

##### Scriptforge

If you hit errors related to scriptforge, please install scriptforge on
Debian, Ubuntu, Mint and derivatives it should suffice to do it from
your package manager or from a shell do:

```
sudo apt install python3-scriptforge
```

Errors look usually like

<img width="656" height="512" alt="da" src="https://github.com/user-attachments/assets/9092e297-2516-42a7-a4ec-65d127b94600" />

The dependency was removed starting from 0.8.0

##### macOS

##### macOS/Linux

## FAQ

**How do I report an error or request a new feature?** Please open
a new issue [here](https://github.com/ikks/libreoffice-stable-diffusion/issues).
If you have questions, head to [Discord](https://discord.gg/PmxqTjUB).

**How do I get the latest extension?** Sometimes we are ahead of the
[stable extension](https://extensions.libreoffice.org/en/extensions/show/99431)
with the latest features
[here](https://github.com/ikks/libreoffice-stable-diffusion/blob/main/loshd.oxt).

## Internals

**How do I troubleshot myself?** Open LibreOffice from a terminal and look
at the output it gives.  You can turn on DEBUG editing the plugin file
`StableHordeForLibreOffice.py` and changing `DEBUG = False` to
`DEBUG = True` (case matters).

## References and other options

* [LibreOffice](https://libreoffice.org): The Document Foundation
* [AIHorde](https://aihorde.net): A collaborative network to share resources
* [Stable diffusion for Gimp](https://github.com/ikks/gimp-stable-diffusion/)
