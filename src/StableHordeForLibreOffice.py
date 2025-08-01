# StableHorde client
# Igor Támara 2025
# No Warranties, use on your own risk
#
# This file is part of the LibreOffice project.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# This file incorporates work covered by the following license notice:
#
#   Licensed to the Apache Software Foundation (ASF) under one or more
#   contributor license agreements. See the NOTICE file distributed
#   with this work for additional information regarding copyright
#   ownership. The ASF licenses this file to you under the Apache
#   License, Version 2.0 (the "License"); you may not use this file
#   except in compliance with the License. You may obtain a copy of
#   the License at http://www.apache.org/licenses/LICENSE-2.0 .
#
import base64
import json
import locale
import os
import sched
import tempfile
import time
import traceback
import uno
from com.sun.star.text.TextContentAnchorType import AS_CHARACTER
from com.sun.star.beans.PropertyAttribute import TRANSIENT
from com.sun.star.beans import PropertyExistException
from com.sun.star.beans import UnknownPropertyException
from datetime import datetime
from scriptforge import CreateScriptService
from urllib.error import HTTPError, URLError
from urllib.request import urlopen, Request

VERSION = "0.1"
API_ROOT = "https://stablehorde.net/api/v2/"
REGISTER_URL = "https://aihorde.net/register"
URL_VERSION_UPDATE = "https://raw.githubusercontent.com/ikks/libreoffice-stable-diffusion/main/version.json"
DEBUG = False
ANONYMOUS = "0000000000"
# check between 5 and 15 seconds
CHECK_WAIT = 5
MAX_TIME_REFRESH = 15
HELP_URL = "https://aihorde.net/faq"
STATUS_BAR = ["|", "/", "-", "\\"]
MIN_WIDTH = 384
MIN_HEIGHT = 384
MAX_WIDTH = 1024
MAX_HEIGHT = 1024
HORDE_CLIENT_NAME = "StableHordeForLibreOffice"


onaction = "vnd.sun.star.script:stablediffusion|StableHordeForLibreOffice.py$validate_form?language=Python&location=user"
onhelp = "vnd.sun.star.script:stablediffusion|StableHordeForLibreOffice.py$get_help?language=Python&location=user"
# https://wiki.documentfoundation.org/Documentation/DevGuide/Scripting_Framework#Python_script When migrating to extension, change this one

######## Fix this
import gettext  # noqa: E402

gettext.bindtextdomain("StableHordeForLibreOffice", "/path/to/my/language/directory")
gettext.textdomain("StableHordeForLibreOffice")
_ = gettext.gettext
########


MODELS = [
    "AbsoluteReality",
    "AlbedoBase XL (SDXL)",
    "AlbedoBase XL 3.1",
    "AMPonyXL",
    "Analog Madness",
    "Anything Diffusion",
    "Babes",
    "BB95 Furry Mix",
    "BB95 Furry Mix v14",
    "BlenderMix Pony",
    "Counterfeit",
    "CyberRealistic Pony",
    "Deliberate",
    "Deliberate 3.0",
    "Dreamshaper",
    "DreamShaper XL",
    "DucHaiten GameArt (Unreal) Pony",
    "Flux.1-Schnell fp8 (Compact)",
    "Fustercluck",
    "Hassaku XL",
    "HolyMix ILXL",
    "ICBINP - I Can't Believe It's Not Photography",
    "ICBINP XL",
    "Juggernaut XL",
    "KaynegIllustriousXL",
    "majicMIX realistic",
    "NatViS",
    "noobEvo",
    "Nova Anime XL",
    "Nova Furry Pony",
    "NTR MIX IL-Noob XL",
    "Pony Diffusion XL",
    "Pony Realism",
    "Prefect Pony",
    "Realistic Vision",
    "SDXL 1.0",
    "Stable Cascade 1.0",
    "stable_diffusion",
    "SwamPonyXL",
    "TUNIX Pony",
    "Unstable Diffusers XL",
    "WAI-ANI-NSFW-PONYXL",
    "WAI-CUTE Pony",
    "waifu_diffusion",
    "White Pony Diffusion 4",
    "Yiffy",
    "ZavyChromaXL",
]


class StableHordeClient:
    def __init__(self, api_key: str = ANONYMOUS, platform: str = HORDE_CLIENT_NAME):
        self.api_key: str = api_key
        self.settings: json = {"api_key": api_key}
        self.client_name: str = platform
        self.headers = {
            "Content-Type": "application/json",
            "Accept": "application/json",
            "apikey": self.api_key,
            "Client-Agent": self.client_name,
        }
        self.informer: LibreOfficeInteraction = None
        self.progress = 0
        self.progress_text = "Starting"
        self.scheduler = sched.scheduler(time.time, time.sleep)
        show_debugging_data(self.headers)

    def url_open(self, url, timeout=10):
        def __url_open__():
            with urlopen(url, timeout=timeout) as response:
                self.response_data = json.loads(response.read().decode("utf-8"))

        for i in range(3):
            self.scheduler.enter(i, 1, self.inform_progress, ())
        self.scheduler.enter(0.1, 2, __url_open__, ())
        self.scheduler.run()

    def check_update(self) -> str:
        """
        Inform the user regarding a plugin update
        """
        already_asked = self.informer.get_property_value("stable_horde_checked_update")
        message = ""
        if already_asked:
            return message

        try:
            # Check for updates by fetching version information from a URL
            url = URL_VERSION_UPDATE
            with urlopen(url) as response:
                data = json.loads(response.read())
            self.informer.set_property("stable_horde_checked_update", True)

            local_version = (*(int(i) for i in VERSION.split(".")),)
            incoming_version = (*(int(i) for i in data["version"].split(".")),)

            if local_version < incoming_version:
                lang = locale.getlocale()[0][:2]
                message = data["message"].get(lang, data["message"]["en"])
        except (HTTPError, URLError):
            message = _(
                "Tried to check for most recent version, check your Internet connection"
            )
        return message

    def get_balance(self) -> str:
        """
        Given an Stable Horde token, returns the balance for the account
        """
        if self.api_key == ANONYMOUS:
            return f"Register at { REGISTER_URL }"
        url = API_ROOT + "find_user"
        request = Request(url, headers=self.headers)
        try:
            self.url_open(request, timeout=15)
            data = self.response_data
            show_debugging_data(data)
        except HTTPError as ex:
            raise (ex)

        return f"\n\nYou have { data['kudos'] } kudos"

    def generate_image(self, options: json, informer) -> str:
        self.informer = informer
        self.check_counter = 1
        self.check_max = (options["max_wait_minutes"] * 60) / CHECK_WAIT
        # Id assigned when requesting the generation of an image
        self.id = ""

        # Used for the progressbar.  We depend on the max time the user indicated
        self.max_time = (
            int(datetime.now().timestamp()) + options["max_wait_minutes"] * 60
        )
        self.factor = 5 / (
            3.0 * options["max_wait_minutes"]
        )  # Percentage and minutes 100*ellapsed/(max_wait*60)

        self.progress_text = _("Contacting the Horde")
        try:
            params = {
                "cfg_scale": float(options["prompt_strength"]),
                "steps": int(options["steps"]),
                "seed": options["seed"],
            }

            data_to_send = {
                "params": params,
                "prompt": options["prompt"],
                "nsfw": options["nsfw"],
                "censor_nsfw": options["censor_nsfw"],
                "r2": True,
            }

            if options["image_width"] % 64 != 0:
                width = int(options["simage_width"] / 64) * 64
            else:
                width = options["image_width"]

            if options["image_height"] % 64 != 0:
                height = int(options["image_height"] / 64) * 64
            else:
                height = options["image_height"]

            params.update({"width": int(width)})
            params.update({"height": int(height)})

            data_to_send.update({"models": [options["model"]]})

            data_to_send = json.dumps(data_to_send)
            post_data = data_to_send.encode("utf-8")

            url = f"{ API_ROOT }generate/async"

            request = Request(url, headers=self.headers, data=post_data)
            try:
                show_debugging_data(data_to_send)
                self.inform_progress()
                self.url_open(request, timeout=15)
                data = self.response_data
                show_debugging_data(data)
                text = "Horde Contacted"
                show_debugging_data(text + f" {self.check_counter} { self.progress }")
                self.progress_text = text
                self.inform_progress()
                self.id = data["id"]
            except HTTPError as ex:
                try:
                    data = ex.read().decode("utf-8")
                    data = json.loads(data)
                    message = data.get("message", str(ex))
                    if data.get("rc", "") == "KudosUpfront":
                        if options["api_key"] == ANONYMOUS:
                            message = (
                                _(
                                    f"Register at { REGISTER_URL } and use your key to improve your rate success. Detail:"
                                )
                                + f" { message }"
                            )
                        else:
                            message = (
                                f"{ HELP_URL } "
                                + _("to learn to earn kudos. Detail:")
                                + " { message }"
                            )
                except Exception as ex2:
                    show_debugging_data(ex2, "No way to recover error msg")
                    message = str(ex)
                show_debugging_data(message, data)
                informer.show_error(f"'{ message }'.")
                return ""
            except URLError as ex:
                show_debugging_data(ex, data)
                informer.show_error(
                    _("Internet required, chek your connection: ") + f"'{ ex }'."
                )
                return ""
            except Exception as ex:
                show_debugging_data(ex, data)
                informer.show_error(f"{ ex }")
                return ""

            self.check_status()
            images = self.get_images()
            image_name = self.get_image_filename(images, options["model"])

        except HTTPError as ex:
            try:
                data = ex.read().decode("utf-8")
                data = json.loads(data)
                message = data.get("message", str(ex))
                show_debugging_data(ex)
            except Exception as ex3:
                show_debugging_data(ex3)
                message = str(ex)
            show_debugging_data(ex, data)
            informer.show_error(_(f"Stablehorde said: '{ message }'."))
            return ""
        except URLError as ex:
            show_debugging_data(ex, data)
            informer.show_error(
                _(f"Internet required, check your connection: '{ ex }'.")
            )
            return ""
        except Exception as ex:
            show_debugging_data(ex)
            informer.show_error(_(f"Service failed with: '{ ex }'."))
            return ""
        finally:
            pass
            message = self.check_update()
            if message:
                informer.show_error(message)

        return image_name

    def inform_progress(self):
        progress = 100 - int((self.max_time - datetime.now().timestamp()) * self.factor)

        show_debugging_data(f"{progress} {self.progress_text}")

        if self.informer and progress != self.progress:
            self.informer.update_status(self.progress_text, progress)
            self.progress = progress

    def check_status(self):
        url = f"{ API_ROOT }generate/check/{ self.id }"

        self.inform_progress()
        self.url_open(url)
        data = self.response_data

        show_debugging_data(data)

        self.check_counter = self.check_counter + 1

        if data["processing"] == 0:
            text = _("Queue position: ") + str(data["queue_position"])
            show_debugging_data(text, self.progress)
        elif data["processing"] > 0:
            text = _("Generating...")
            show_debugging_data(text + f" {self.check_counter} { self.progress }")
        self.progress_text = text

        if self.check_counter < self.check_max and data["done"] is False:
            if data["is_possible"] is True:
                wait_time = min(
                    max(CHECK_WAIT, int(data["wait_time"] / 2)), MAX_TIME_REFRESH
                )
                for i in range(1, wait_time + 3):
                    self.scheduler.enter(i, 1, self.inform_progress, ())
                self.scheduler.enter(wait_time, 2, self.check_status, ())
                self.scheduler.run()
            else:
                show_debugging_data(data)
                raise Exception(
                    _(
                        "Currently no worker available to generate your image. Please try again later."
                    )
                )
        elif self.check_counter >= self.check_max:
            minutes = (self.check_max * CHECK_WAIT) / 60
            show_debugging_data(data)
            raise Exception(
                _(f"Image generation timed out after { minutes } minutes.")
                + _("Please try again later.")
            )
        elif data["done"]:
            self.progress_text = _("Downloading generated image")
            self.inform_progress()
            return

    def get_images(self):
        url = f"{ API_ROOT }generate/status/{ self.id }"
        self.progress_text = _("fetching images")
        self.inform_progress()
        self.url_open(url)
        data = self.response_data
        show_debugging_data(data)

        return data["generations"]

    def get_image_filename(self, images, model):
        show_debugging_data("Start to download generated images")
        with tempfile.NamedTemporaryFile(
            "wb+", delete=False, suffix=".webp"
        ) as generated_file:
            for image in images:
                if image["img"].startswith("https"):
                    self.progress_text = _("Downloading image")
                    self.inform_progress()
                    with urlopen(image["img"]) as response:
                        bytes = response.read()
                else:
                    bytes = base64.b64decode(image["img"])

                show_debugging_data(f"dumping to { generated_file.name }")

                generated_file.write(bytes)
            return generated_file.name
        return ""

    def get_settings(self) -> json:
        return {}

    def set_settings(self, settings: json):
        self.settings = settings


def validate_form(event: uno):
    if event is None:
        return
    ctrl = CreateScriptService("DialogEvent", event)
    dialog = ctrl.Parent
    btn_ok = dialog.Controls("btn_ok")
    btn_ok.Enabled = len(dialog.Controls("txt_prompt").Value) > 10


def get_help(event: uno):
    if event is None:
        return
    session = CreateScriptService("Session")
    session.OpenURLInBrowser(HELP_URL)
    session.Dispose()


class LibreOfficeInteraction:
    def get_type_doc(self, doc):
        TYPE_DOC = {
            "writer": "com.sun.star.text.TextDocument",
            "impress": "com.sun.star.presentation.PresentationDocument",
            # "calc": "com.sun.star.sheet.SpreadsheetDocument",
            # "draw": "com.sun.star.drawing.DrawingDocument",
        }
        for k, v in TYPE_DOC.items():
            if doc.supportsService(v):
                return k
        return "new-writer"

    def __init__(self):
        self.desktop = XSCRIPTCONTEXT.getDesktop()
        self.image_insert_to = {
            "impress": self.__insert_image_in_presentation__,
            "writer": self.__insert_image_in_text__,
            "calc": self.__insert_image_in_calc__,
            "draw": self.__insert_image_in_draw__,
        }
        self.model = self.desktop.getCurrentComponent()

        self.bas = CreateScriptService("Basic")
        self.ui = CreateScriptService("UI")
        self.platform = CreateScriptService("Platform")
        self.base_info = "-_".join(
            [
                HORDE_CLIENT_NAME,
                VERSION,
                self.platform.OSPlatform,
                self.platform.PythonVersion,
                self.platform.OfficeVersion,
                self.platform.Machine,
            ]
        )

        # For now we only add images to Writer
        self.inside = self.get_type_doc(self.model)
        if self.inside == "new-writer":
            self.model = self.desktop.loadComponentFromURL(
                "private:factory/swriter", "_blank", 0, ()
            )
            self.inside = "writer"

        try:
            # Get selected text
            if self.inside == "writer":
                self.selected = self.model.CurrentSelection.getByIndex(0).getString()
            else:
                self.selected = self.model.CurrentSelection.getString()
        except AttributeError:
            # If the selection is not text, let's wait for the user to write down
            self.selected = ""

        self.doc = self.model

        self.dlg = self.__create_dialog__()

    def __create_dialog__(self):
        dlg = CreateScriptService(
            "NewDialog", "StableHordeOptionsDialog", (47, 10, 265, 206)
        )
        dlg.Caption = "Stable Horde Configuration"
        dlg.CreateGroupBox("framebox", (16, 11, 236, 163))
        # Labels
        lbl = dlg.CreateFixedText("label_prompt", (29, 31, 39, 13))
        lbl.Caption = _("Prompt")
        lbl = dlg.CreateFixedText("label_height", (155, 65, 39, 13))
        lbl.Caption = _("Height")
        lbl = dlg.CreateFixedText("label_width", (29, 65, 39, 13))
        lbl.Caption = _("Width")
        lbl = dlg.CreateFixedText("label_model", (29, 82, 39, 13))
        lbl.Caption = _("Model")
        lbl = dlg.CreateFixedText("label_max_wait", (155, 82, 39, 13))
        lbl.Caption = _("Max Wait")
        lbl = dlg.CreateFixedText("label_strength", (29, 99, 39, 13))
        lbl.Caption = _("Strength")
        lbl = dlg.CreateFixedText("label_steps", (155, 99, 39, 13))
        lbl.Caption = _("Steps")
        lbl = dlg.CreateFixedText("label_seed", (94, 130, 49, 13))
        lbl.Caption = _("Seed (Optional)")
        lbl = dlg.CreateFixedText("label_token", (93, 148, 49, 13))
        lbl.Caption = _("Token (Optional)")

        # Buttons
        button_ok = dlg.CreateButton("btn_ok", (145, 182, 45, 13), push="OK")
        button_ok.Caption = "Process"
        button_ok.TabIndex = 12
        button_cancel = dlg.CreateButton("btn_cancel", (78, 182, 49, 12), push="CANCEL")
        button_cancel.Caption = "Cancel"
        button_ok.TabIndex = 13
        button_help = dlg.CreateButton("CommandButton1", (23, 15, 13, 10))
        button_help.Caption = "?"
        button_help.TipText = _("About Horde")
        button_help.OnMouseReleased = onhelp
        button_ok.TabIndex = 14

        # Controls
        ctrl = dlg.CreateComboBox(
            "lst_model",
            (60, 80, 79, 15),
        )
        ctrl.TabIndex = 4
        ctrl = dlg.CreateTextField(
            "txt_prompt",
            (60, 16, 188, 42),
            multiline=True,
        )
        ctrl.TabIndex = 1
        ctrl.TipText = _(
            "Let your imagination run wild or put a proper description of your desired output."
        )
        ctrl.OnTextChanged = onaction
        ctrl = dlg.CreateTextField("txt_token", (155, 147, 90, 13))
        ctrl.TabIndex = 11
        ctrl.TipText = _("Get yours at https://stablehorde.net/ for free")

        ctrl = dlg.CreateTextField("txt_seed", (155, 128, 90, 13))
        ctrl.TabIndex = 10
        ctrl.TipText = _(
            "If you want the process repeatable, put something here, otherwise, enthropy will win"
        )

        ctrl = dlg.CreateNumericField(
            "int_width",
            (87, 63, 52, 13),
            accuracy=0,
            minvalue=384,
            maxvalue=1024,
            increment=64,
            spinbutton=True,
        )
        ctrl.Value = 384
        ctrl.TabIndex = 2
        ctrl = dlg.CreateNumericField(
            "int_height",
            (196, 63, 52, 13),
            accuracy=0,
            minvalue=384,
            maxvalue=1024,
            increment=64,
            spinbutton=True,
        )
        ctrl.Value = 384
        ctrl.TabIndex = 3
        ctrl = dlg.CreateNumericField(
            "int_strength",
            (87, 100, 52, 13),
            minvalue=0,
            maxvalue=20,
            increment=0.5,
            accuracy=2,
            spinbutton=True,
        )
        ctrl.Value = 15
        ctrl.TabIndex = 5
        ctrl.TipText = _(
            "How much the AI will follow the prompt, the higher, the more obedient"
        )
        ctrl = dlg.CreateNumericField(
            "int_steps",
            (196, 97, 52, 13),
            minvalue=1,
            maxvalue=150,
            spinbutton=True,
            increment=10,
            accuracy=0,
        )
        ctrl.Value = 25
        ctrl.TabIndex = 6
        ctrl.TipText = _("More steps mean more details, affects time and GPU usage")
        ctrl = dlg.CreateNumericField(
            "int_waiting",
            (196, 80, 52, 13),
            minvalue=1,
            maxvalue=5,
            spinbutton=True,
            accuracy=0,
        )
        ctrl.Value = 3
        ctrl.TabIndex = 9
        ctrl.TipText = _(
            "In minutes. Depends on your patience and your kudos.  You'll get a complain message if timeout is reached"
        )
        ctrl = dlg.CreateCheckBox("bool_nsfw", (29, 130, 30, 10))
        ctrl.Caption = "NSFW"
        ctrl.TabIndex = 7
        ctrl.TipText = _(
            "If not marked, it's faster, when marked you are on the edge..."
        )

        ctrl = dlg.CreateCheckBox("bool_censure", (29, 145, 30, 10))
        ctrl.Caption = "Censor NSFW"
        ctrl.TipText = _("Allow if you want to avoid unexpected images...")
        ctrl.TabIndex = 8

        return dlg

    def prepare_options(self, options: json = None) -> json:
        dlg = self.dlg
        dlg.Controls("txt_prompt").Value = self.selected
        api_key = options.get("api_key", "")
        if api_key == ANONYMOUS:
            api_key = ""
        dlg.Controls("txt_token").Value = api_key
        dlg.Controls("lst_model").RowSource = MODELS
        dlg.Controls("lst_model").Value = "stable_diffusion"
        dlg.Controls("btn_ok").Enabled = len(self.selected) > 10

        dlg.Controls("int_width").Value = options.get("image_width", MIN_WIDTH)
        dlg.Controls("int_height").Value = options.get("image_height", MIN_HEIGHT)
        dlg.Controls("lst_model").Value = options.get("model", "stable_diffusion")
        dlg.Controls("int_strength").Value = options.get("prompt_strength", 0.3)
        dlg.Controls("int_steps").Value = options.get("steps", 25)
        dlg.Controls("bool_nsfw").Value = options.get("nsfw", 0)
        dlg.Controls("bool_censure").Value = options.get("censor_nsfw", 1)
        dlg.Controls("int_waiting").Value = options.get("max_wait_minutes", 3)
        dlg.Controls("txt_seed").Value = options.get("seed", "")
        rc = dlg.Execute()

        if rc != dlg.OKBUTTON:
            show_debugging_data("User scaped, nothing to do")
            dlg.Terminate()
            dlg.Dispose()
            return None

        show_debugging_data("good")
        options = {
            "prompt": dlg.Controls("txt_prompt").Value,
            "image_width": dlg.Controls("int_width").Value,
            "image_height": dlg.Controls("int_height").Value,
            "model": dlg.Controls("lst_model").Value,
            "prompt_strength": dlg.Controls("int_strength").Value,
            "steps": dlg.Controls("int_steps").Value,
            "nsfw": dlg.Controls("bool_nsfw").Value == 1,
            "censor_nsfw": dlg.Controls("bool_censure").Value == 1,
            "api_key": dlg.Controls("txt_token").Value or ANONYMOUS,
            "max_wait_minutes": dlg.Controls("int_waiting").Value,
            "seed": dlg.Controls("txt_seed").Value,
        }
        dlg.Terminate()
        dlg.Dispose()
        self.options = options
        return options

    def free(self):
        self.bas.Dispose()
        self.ui.Dispose()
        self.platform.Dispose()
        self.set_finished()

    def update_status(self, text: str, progress: int = None):
        if progress:
            self.progress = progress
        self.ui.SetStatusbar(text, self.progress)

    def set_finished(self):
        self.ui.SetStatusbar("")

    def show_error(self, message):
        self.bas.MsgBox(message)

    def insert_image(self, img_path: str, width: int, height: int):
        self.image_insert_to[self.inside](img_path, width, height)

    def __insert_image_in_presentation__(self, img_path: str, width: int, height: int):
        from com.sun.star.awt import Size
        from com.sun.star.awt import Point

        size = Size(width * 10, height * 10)
        image = self.doc.createInstance("com.sun.star.presentation.GraphicObjectShape")
        image.GraphicURL = uno.systemPathToFileUrl(img_path)

        ctrllr = self.model.CurrentController
        draw_page = ctrllr.CurrentPage

        draw_page.addTop(image)
        added_image = draw_page[-1]
        added_image.setSize(size)
        position = Point(
            ((added_image.Parent.Width - (width * 10)) / 2),
            ((added_image.Parent.Height - (height * 10)) / 2),
        )
        added_image.setPosition(position)
        added_image.setPropertyValue("ZOrder", draw_page.Count)

        # The placeholder does not update
        added_image.PlaceholderText = ""
        added_image.setPropertyValue("PlaceholderText", "")

        added_image.setPropertyValue("Title", "Stable Horde Generated Image")
        added_image.setPropertyValue(
            "Description", f"{ self.options["prompt"] } by { self.options["model"] }"
        )
        added_image.Visible = True
        self.model.Modified = True
        os.unlink(img_path)

    def __insert_image_in_calc__(self, img_path: str, width: int, height: int):
        self.bas.MsgBox("TBD")

    def __insert_image_in_draw__(self, img_path: str, width: int, height: int):
        self.bas.MsgBox("TBD")

    def __insert_image_in_text__(self, img_path: str, width: int, height: int):
        image = self.doc.createInstance("com.sun.star.text.GraphicObject")
        image.GraphicURL = uno.systemPathToFileUrl(img_path)
        image.AnchorType = AS_CHARACTER
        image.Width = width * 10
        image.Height = height * 10

        ctrllr = self.model.CurrentController
        curview = ctrllr.ViewCursor
        curview.String = f'{ self.options["prompt"] }  by  { self.options["model"] }'
        self.doc.Text.insertTextContent(curview, image, False)
        os.unlink(img_path)

    def get_property_value(self, property_name: str) -> str | bool | None:
        value = None
        oDocProps = self.model.getDocumentProperties()
        self.userProps = oDocProps.getUserDefinedProperties()
        try:
            value = self.userProps.getPropertyValue(property_name)
        except UnknownPropertyException:
            # The property was not present
            return None
        return value

    def set_property(self, property_name: str, value: str | bool):
        oDocProps = self.model.getDocumentProperties()
        self.userProps = oDocProps.getUserDefinedProperties()
        if value is None:
            str_value = ""
        else:
            str_value = str(value)

        try:
            self.userProps.addProperty(property_name, TRANSIENT, str_value)
        except PropertyExistException:
            self.userProps.setPropertyValue(property_name, str_value)


class Settings:
    def __init__(self):
        self.base_directory = "/tmp"
        self.settingsfile = "stablehordesettings.json"
        self.file = self.base_directory + "/" + self.settingsfile

    def load(self) -> json:
        if not os.path.exists(self.file):
            return {}
        with open(self.file) as myfile:
            return json.loads(myfile.read())

    def save(self, settings: json):
        with open(self.file, "w") as myfile:
            myfile.write(json.dumps(settings))


def show_debugging_data(information, additional=""):
    if not DEBUG:
        return

    dnow = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(information, Exception):
        ln = information.__traceback__.tb_lineno
        print(f"[{ dnow }]{ln}: { information }")
        print(
            "".join(
                traceback.format_exception(None, information, information.__traceback__)
            )
        )
    else:
        print(f"[{ dnow }] { information }")
    if additional:
        print(additional)


def create_image():
    """Creates an image from a prompt provided by the user, making use
    of Stable Horde"""

    st_manager = Settings()
    saved_options = st_manager.load()
    lo_manager = LibreOfficeInteraction()
    sh_client = StableHordeClient(
        saved_options.get("api_key", ANONYMOUS), lo_manager.base_info
    )

    show_debugging_data(lo_manager.base_info)

    options = lo_manager.prepare_options(saved_options)

    if options is None:
        # User cancelled, nothing to be done
        lo_manager.free()
        return

    show_debugging_data(options)
    image_path = sh_client.generate_image(options, lo_manager)

    if image_path:
        lo_manager.insert_image(
            image_path, options["image_width"], options["image_height"]
        )
        st_manager.save(options)

    lo_manager.free()


g_exportedScripts = (create_image,)

if __name__ == "__main__":
    create_image()

# TODO:
#    - Determine the correct place to store the options saved
#    - Store the file with proper permission 700
#    - Recommend to use a shared key
# * [ ] Go to github
# * [ ] Issue bugs for Impress with placeholdertext
#    - Wayland transparent png
#    - Wishlist to have right alignment for numeric
# * [ ] Add to menu
# * [ ] Add a popup context menu
# * [ ] Internationalization
# * [ ] Use styles support
#    -  Show Styles and Advanced View
#    -  Download and cache Styles
# * [ ] Add to Calc
# * [ ] Add to Draw
# * [ ] Publish the extension
#   - https://extensions.libreoffice.org/en/home/using-this-site-as-an-extension-maintainer
#   - Create extension
#   - Create account
#   - Upload extension
# * [ ] Announce in stablehorde
#
# Local documentation
# file:///usr/share/doc/libreoffice-dev-doc/api/
# /usr/lib/libreoffice/share/Scripts/python/
# /usr/lib/libreoffice/sdk/examples/html
