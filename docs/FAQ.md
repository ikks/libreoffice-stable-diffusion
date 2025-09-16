## Why I am not getting good images?

Maybe the prompt needs some polishing, consider visiting [artbot](https://artbot.site/)
to view samples of prompts and models used to generate some interesting images.

Models are trained with millions of images and descriptions of them to generate
new compositions, models are trained with specific use cases, for example, targeted
to photograph quality, cartoons and different styles.  The number of different variations
is bounded by the computing capacity.

## I created an image and I want to reuse it, where I can find it?

Starting from Version 0.8, the downloaded images are stored in the `store/aihorde.net_images`
directory. It's possible to free some space removing the images that you dont want to
have anymore.


## Why the images I'm getting are not as good as other services out there?

With the same prompt on different providers and different runs on the same provider
with the same parameters only changing the seed can make subtle or big differences.

AiHorde incorporate new models as they are produced in the huge ecosystem and the
computing capacity of any worker in aihorde maybe is not enough for a given model.

## Where can I see some sample of models in action?

[artbot](https://artbot.site/) has descriptions and samples for the models present
in Aihorde [here](https://artbot.site/info/models).


## How can I get an api_key?

[Here](https://aihorde.net), it's free.


## Can I use my api_key in artbot?

Yes, libreoffice AiHorde plugin use [AiHorde](https://aihorde.net).

## I've heard that there is inpainting and image to image, libreoffice plugin will support them?

Not in the plan, for that kind of work you can use [GIMP](https://gimp.org)
with the [gimp AiHorde plugin](https://github.com/ikks/gimp-stable-diffusion), you can
use the same api_key.  Gimp is targeted to have way more advanced options, it can
for example enlarge an image.

## I want to continue using the plugin, How can I help to sustain this?

The plugin works thanks to [AiHorde](https://aihorde.net) and you can help the funding
[via Patreon](https://www.patreon.com/db0), [adding a worker](https://github.com/Haidra-Org/AI-Horde/blob/main/README_StableHorde.md#joining-the-horde)
if you have a good GPU(advanced graphics card), spreading the word.  For the plugin, you
can help [translating to a new language](https://github.com/Haidra-Org/AI-Horde/blob/main/README_StableHorde.md#joining-the-horde).

In the plugin side it's also possible to making pull requests with new features, it's
recommended to first [open an issue](https://github.com/ikks/libreoffice-stable-diffusion/issues),
letting know what is the purpose and attaching the [PR](https://github.com/ikks/libreoffice-stable-diffusion/pulls)
to the given issue.

## Which services does the plugin connect to?

Once you installed the plugin

* Generating images require https://aihorde.net
* Model information and new versions information require https://github.com
* Translations depend on https://igortamara-opus-translate.hf.space/ which is part of [huggingface](https://huggingface.co/)

