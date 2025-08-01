# libreoffice-stable-diffusion-horde

This repository includes a
[LibreOffice](https://libreoffice.org) plugin to make use of
[StableHorde](https://stablehorde.net). Stablehorde is a cluster
of stable-diffusion servers run by volunteers. You can create
stable-diffusion images for free without running a colab notebook
or a local server. Please check the section "Limitations" to better
understand where the limits are.

Please check [CHANGELOG.md](../CHANGELOG.md) for the latest
changes.

## Installation
### Download the macro

Download [StableHordeForLibreOffice.py](https://raw.githubusercontent.com/ikks/libreoffice-stable-diffusion/refs/heads/main/src/StableHordeForLibreOffice.py).

### LibreOffice

This Macro is known to work from LibreOffice 7.4 and upwards

1. [Download](https://raw.githubusercontent.com/ikks/libreoffice-stable-diffusion/refs/heads/main/src/StableHordeForLibreOffice.py)
  the macro
2. Copy the downloaded file `StableHordeLibreOffice.py` to
  your LibreOffice .config directory
  `libreoffice/4/user/Scripts/python/stablediffusion`.  If you
  are in doubt about where is your .config directory go to
  Tools -> Options... | LibreOffice > Paths to find yours.
  For each Operating System is different:
    * C:\Users\username... for Windows
    * /home/username... for Linux
    * /Users/username... for Mac.
<img width="1920" height="1080" alt="determinepath" src="https://github.com/user-attachments/assets/08c5ce95-4171-4d4c-9a5b-491d4874f92b" />

3. Reopen LibreOffice.
4. Invoke the macro going to Tools -> Macros -> Run Macro... |
  My Macros > stablediffusion > StableHordeForLibreOffice ->
  create_image and push Run button.
<img width="1920" height="1080" alt="runmacro" src="https://github.com/user-attachments/assets/5e344742-eae3-4647-b0d6-c1876e911c24" />
5. Optionally add a shortcut going to Tools -> Customize... |
  1. Tab Keyboard,
  2. on Category > Application Macros > My Macros >
    stablediffusion > StableHordeForLibreOffice
  3. Select Function | create_image
  4. On Shortcuts Keys select your preferred combination, for
    example Ctrl+Shift+h and push the Button `Assign` and
    finally `OK`.
<img width="1920" height="1080" alt="shortcut" src="https://github.com/user-attachments/assets/518ccb75-81a3-4698-8dc4-1fb5611d5e58" />


## Generate images

Now we are ready for generating images.

1. Launch the macro and run it, a dialog will appear allowing you
to write a prompt and generate the image, if you don't have a
Writer document opened, the macro will open it for you.

You can select a text, and when the macro is launched, the text
will be used as a prompt.  The progress bar will guide you on what
is happening; at the end of the process, an image will be inserted
above your selected text or in the cursor position.

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
[StableHorde](https://stablehorde.net) and getting an API key.

2. Click on the OK button. The values you inserted into the dialog
will be transmitted to the server, which dispatches the request now to
one of the stable-diffusion servers in the cluster. Your generation
request is added to queue. You can see the queue position in the
status bar. When the image has been generated successfully, it will
be shown as a new image in LibreOffice inside the current document.

## Limitations
- **Generation speed:** StableHorde is a cluster of stable-diffusion
servers run by volunteers. The generation speed depends on how many
servers are in the cluster, which hardware they use and how many others
want to generate with StableHorde. The upside is, that StableHorde is
free to use, the downside that the generation speed is unpredictable.

- **Privacy:** The privacy StableHorde offers is similar to generating
in a public discord channel. So, please assume, that neither your
prompts nor your generated images are private.

- **Features:** Currently text2img if you are looking for img2img and
inpainting consider the
[Gimp Stable Horde plugin](https://github.com/ikks/gimp-stable-diffusion/)

</details>
  
## Troubleshooting

### LibreOffice
##### Linux

##### macOS

##### macOS/Linux

## FAQ

**How do I report an error or request a new feature?** Please open
a new issue in this repository.

## Internals

**How do I troubleshot myself?** Open LibreOffice from a terminal and look
at the output it gives.  You can turn on DEBUG editing the plugin file
`StableHordeForLibreOffice.py` and changing `DEBUG = False` to
`DEBUG = True` (case matters).

## References and other options

* [LibreOffice](https://libreoffice.org): The Document Foundation
* [StableHorde](https://stablehorde.net): A collaborative network to share resources
* [Stable diffusion for Gimp](https://github.com/ikks/gimp-stable-diffusion/)
