# AIHorde client for LibreOffice
# Igor Támara 2025
# No Warranties, use on your own risk
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

import gettext
import json
import logging
import os
import sys
import tempfile
import uno
import unohelper
from com.sun.star.awt import Point
from com.sun.star.awt import Size
from com.sun.star.beans import PropertyExistException
from com.sun.star.beans import UnknownPropertyException
from com.sun.star.beans.PropertyAttribute import TRANSIENT
from com.sun.star.document import XEventListener
from com.sun.star.task import XJobExecutor
from com.sun.star.text.TextContentAnchorType import AS_CHARACTER
from pathlib import Path
from scriptforge import CreateScriptService
from typing import Union

import_message_error = None

logger = logging.getLogger(__name__)
DEBUG = False
LOGGING_LEVEL = logging.ERROR

VERSION = "0.6"
LIBREOFFICE_EXTENSION_ID = "org.fectp.StableHordeForLibreOffice"
GETTEXT_DOMAIN = "stablehordeforlibreoffice"

log_file = os.path.join(tempfile.gettempdir(), "libreoffice_shotd.log")
if DEBUG:
    print(f"your log is at {log_file}")
    LOGGING_LEVEL = logging.DEBUG
logging.basicConfig(
    filename=log_file,
    level=LOGGING_LEVEL,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

file_path = os.path.dirname(os.path.realpath(__file__))
submodule_path = os.path.join(file_path, "module")
sys.path.append(str(submodule_path))

from aihordeclient import (  # noqa: E402
    ANONYMOUS_KEY,  # noqa: E402
    DEFAULT_MODEL,
    MIN_PROMPT_LENGTH,
    MIN_HEIGHT,
    MIN_WIDTH,
    MODELS,
)
from aihordeclient import (  # noqa: E402
    AiHordeClient,
    InformerFrontend,
    HordeClientSettings,
)

HELP_URL = "https://aihorde.net/faq"
"""
Help url for the extension
"""

URL_VERSION_UPDATE = "https://raw.githubusercontent.com/ikks/libreoffice-stable-diffusion/main/version.json"
"""
Latest version for the extension
"""

PROPERTY_CURRENT_SESSION = "ai_horde_checked_update"

URL_DOWNLOAD = "https://extensions.libreoffice.org/en/extensions/show/99431"
"""
Download URL for libreoffice-stable-diffusion
"""

HORDE_CLIENT_NAME = "StableHordeForLibreOffice"
"""
Name of the client sent to API
"""

# onaction = "service:org.fectp.StableHordeForLibreOffice$validate_form?language=Python"
# onhelp = "service:org.fectp.StableHordeForLibreOffice$get_help?language=Python&location=application"
# onmenupopup = "vnd.sun.star.script:stablediffusion|StableHordeForLibreOffice.py$popup_click?language=Python&location=user"
# https://wiki.documentfoundation.org/Documentation/DevGuide/Scripting_Framework#Python_script When migrating to extension, change this one


# gettext usual alias for i18n
_ = gettext.gettext
gettext.textdomain(GETTEXT_DOMAIN)


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


def popup_click(poEvent: uno = None):
    """
    Intended for popup
    """
    if poEvent is None:
        return
    my_popup = CreateScriptService("SFWidgets.PopupMenu", poEvent)

    my_popup.AddItem(_("Generate Image"))
    # Populate popupmenu with items
    response = my_popup.Execute()
    logger.debug(response)
    my_popup.Dispose()


class LibreOfficeInteraction(InformerFrontend):
    def get_type_doc(self, doc):
        TYPE_DOC = {
            "writer": "com.sun.star.text.TextDocument",
            "impress": "com.sun.star.presentation.PresentationDocument",
            "calc": "com.sun.star.sheet.SpreadsheetDocument",
            "draw": "com.sun.star.drawing.DrawingDocument",
        }
        for k, v in TYPE_DOC.items():
            if doc.supportsService(v):
                return k
        return "new-writer"

    def __init__(self, desktop, context):
        self.desktop = desktop
        self.context = context
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
        self.session = CreateScriptService("Session")
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
            "NewDialog", "AIHordeOptionsDialog", (47, 10, 265, 206)
        )
        dlg.Caption = _("AI Horde for LibreOffice - ") + VERSION
        dlg.CreateGroupBox("framebox", (16, 9, 236, 165))
        # Labels
        lbl = dlg.CreateFixedText("label_prompt", (29, 31, 45, 13))
        lbl.Caption = _("Prompt")
        lbl = dlg.CreateFixedText("label_height", (155, 65, 45, 13))
        lbl.Caption = _("Height")
        lbl = dlg.CreateFixedText("label_width", (29, 65, 45, 13))
        lbl.Caption = _("Width")
        lbl = dlg.CreateFixedText("label_model", (29, 82, 45, 13))
        lbl.Caption = _("Model")
        lbl = dlg.CreateFixedText("label_max_wait", (155, 82, 45, 13))
        lbl.Caption = _("Max Wait")
        lbl = dlg.CreateFixedText("label_strength", (29, 99, 45, 13))
        lbl.Caption = _("Strength")
        lbl = dlg.CreateFixedText("label_steps", (155, 99, 45, 13))
        lbl.Caption = _("Steps")
        lbl = dlg.CreateFixedText("label_seed", (96, 130, 49, 13))
        lbl.Caption = _("Seed (Optional)")
        lbl = dlg.CreateFixedText("label_token", (96, 149, 49, 13))
        lbl.Caption = _("ApiKey (Optional)")

        # Buttons
        button_ok = dlg.CreateButton("btn_ok", (73, 182, 49, 13), push="OK")
        button_ok.Caption = _("Process")
        button_ok.TabIndex = 4
        button_cancel = dlg.CreateButton(
            "btn_cancel", (145, 182, 49, 13), push="CANCEL"
        )
        button_cancel.Caption = _("Cancel")
        button_cancel.TabIndex = 13
        # button_help = dlg.CreateButton("CommandButton1", (23, 15, 13, 10))
        # button_help.Caption = "?"
        # button_help.TipText = _("About Horde")
        # button_help.OnMouseReleased = onhelp
        # button_ok.TabIndex = 14

        # Controls
        ctrl = dlg.CreateComboBox(
            "lst_model",
            (60, 80, 79, 15),
            linecount=10,
        )
        ctrl.TabIndex = 3
        ctrl = dlg.CreateTextField(
            "txt_prompt",
            (60, 16, 188, 42),
            multiline=True,
        )
        ctrl.TabIndex = 1
        ctrl.TipText = _("""
        Let your imagination run wild or put a proper description of your
        desired output. Use full grammar for Flux, use tag-like language
        for sd15, use short phrases for sdxl.

        Write at least 5 words or 10 characters.
        """)
        # ctrl.OnTextChanged = onaction
        ctrl = dlg.CreateTextField("txt_token", (155, 147, 92, 13))
        ctrl.TabIndex = 11
        ctrl.TipText = _("""
        Get yours at https://aihorde.net/ for free. Recommended:
        Anonymous users are last in the queue.
        """)

        ctrl = dlg.CreateTextField("txt_seed", (155, 128, 92, 13))
        ctrl.TabIndex = 2
        ctrl.TipText = _(
            "Set a seed to regenerate (reproducible), or it'll be chosen at random by the worker."
        )

        ctrl = dlg.CreateNumericField(
            "int_width",
            (91, 63, 48, 13),
            accuracy=0,
            minvalue=384,
            maxvalue=1024,
            increment=64,
            spinbutton=True,
        )
        ctrl.Value = 384
        ctrl.TabIndex = 5
        ctrl.TipText = _(
            "Height and Width together at most can be 2048x2048=4194304 pixels"
        )
        ctrl = dlg.CreateNumericField(
            "int_strength",
            (91, 100, 48, 13),
            minvalue=0,
            maxvalue=20,
            increment=0.5,
            accuracy=2,
            spinbutton=True,
        )
        ctrl.Value = 15
        ctrl.TabIndex = 7
        ctrl.TipText = _("""
         How strongly the AI follows the prompt vs how much creativity to allow it.
        Set to 1 for Flux, use 2-4 for LCM and lightning, 5-7 is common for SDXL
        models, 6-9 is common for sd15.
        """)
        ctrl = dlg.CreateNumericField(
            "int_height",
            (200, 63, 48, 13),
            accuracy=0,
            minvalue=384,
            maxvalue=1024,
            increment=64,
            spinbutton=True,
        )
        ctrl.Value = 384
        ctrl.TabIndex = 6
        ctrl.TipText = _(
            "Height and Width together at most can be 2048x2048=4194304 pixels"
        )
        ctrl = dlg.CreateNumericField(
            "int_waiting",
            (200, 80, 48, 13),
            minvalue=1,
            maxvalue=45,
            spinbutton=True,
            accuracy=0,
        )
        ctrl.Value = 3
        ctrl.TabIndex = 8
        ctrl.TipText = _("""
        How long to wait for your generation to complete.
        Depends on number of workers and user priority (more
        kudos = more priority. Anonymous users are last)
        """)
        ctrl = dlg.CreateNumericField(
            "int_steps",
            (200, 97, 48, 13),
            minvalue=1,
            maxvalue=150,
            spinbutton=True,
            increment=10,
            accuracy=0,
        )
        ctrl.Value = 25
        ctrl.TabIndex = 7
        ctrl.TipText = _("""
        How many sampling steps to perform for generation. Should
        generally be at least double the CFG unless using a second-order
        or higher sampler (anything with dpmpp is second order)
        """)
        ctrl = dlg.CreateCheckBox("bool_nsfw", (29, 130, 55, 10))
        ctrl.Caption = _("NSFW")
        ctrl.TabIndex = 9
        ctrl.TipText = _("""
        Whether or not your image is intended to be NSFW. May
        reduce generation speed (workers can choose if they wish
        to take nsfw requests)
        """)

        ctrl = dlg.CreateCheckBox("bool_censure", (29, 145, 55, 10))
        ctrl.Caption = _("Censor NSFW")
        ctrl.TipText = _("""
        Separate from the NSFW flag, should workers
        return nsfw images. Censorship is implemented to be safe
        and overcensor rather than risk returning unwanted NSFW.
        """)
        ctrl.TabIndex = 10
        if DEBUG:
            ctrl = dlg.CreateCheckBox("bool_debug", (19, 162, 50, 10))
            ctrl.Caption = _("📜")
            ctrl.TabIndex = 12
            ctrl.TipText = _("""
            See what's happening when something does not go as expected.
            Recommended to run libreoffice from terminal or cmd.
            """)

        return dlg

    def prepare_options(self, options: json = None) -> json:
        dlg = self.dlg
        dlg.Controls("txt_prompt").Value = self.selected
        api_key = options.get("api_key", ANONYMOUS_KEY)
        ctrl_token = dlg.Controls("txt_token")
        ctrl_token.Value = api_key
        if api_key == ANONYMOUS_KEY:
            ctrl_token.Value = ""
            ctrl_token.TabIndex = 1
        choices = options.get("local_settings", {"models": MODELS}).get(
            "models", MODELS
        )
        choices = choices or MODELS
        dlg.Controls("lst_model").RowSource = choices
        dlg.Controls("lst_model").Value = DEFAULT_MODEL
        # dlg.Controls("btn_ok").Enabled = len(self.selected) > MIN_PROMPT_LENGTH

        dlg.Controls("int_width").Value = options.get("image_width", MIN_WIDTH)
        dlg.Controls("int_height").Value = options.get("image_height", MIN_HEIGHT)
        dlg.Controls("lst_model").Value = options.get("model", DEFAULT_MODEL)
        dlg.Controls("int_strength").Value = options.get("prompt_strength", 6.3)
        dlg.Controls("int_steps").Value = options.get("steps", 25)
        dlg.Controls("bool_nsfw").Value = options.get("nsfw", 0)
        dlg.Controls("bool_censure").Value = options.get("censor_nsfw", 1)
        dlg.Controls("int_waiting").Value = options.get("max_wait_minutes", 15)
        dlg.Controls("txt_seed").Value = options.get("seed", "")
        if DEBUG:
            dlg.Controls("bool_debug").Value = options.get("debug", 1 if DEBUG else 0)
        rc = dlg.Execute()

        if rc != dlg.OKBUTTON:
            logger.debug("User scaped, nothing to do")
            dlg.Terminate()
            dlg.Dispose()
            return None
        if DEBUG:
            if dlg.Controls("bool_debug").Value == 1:
                print(f"your log is at {log_file}. do:\ntailf {log_file}")
                LOGGING_LEVEL = logging.DEBUG
            else:
                LOGGING_LEVEL = logging.ERROR
            logger.setLevel(LOGGING_LEVEL)
        logger.debug("good")

        options.update(
            {
                "prompt": dlg.Controls("txt_prompt").Value,
                "image_width": dlg.Controls("int_width").Value,
                "image_height": dlg.Controls("int_height").Value,
                "model": dlg.Controls("lst_model").Value,
                "prompt_strength": dlg.Controls("int_strength").Value,
                "steps": dlg.Controls("int_steps").Value,
                "nsfw": dlg.Controls("bool_nsfw").Value == 1,
                "censor_nsfw": dlg.Controls("bool_censure").Value == 1,
                "api_key": dlg.Controls("txt_token").Value or ANONYMOUS_KEY,
                "max_wait_minutes": dlg.Controls("int_waiting").Value,
                "seed": dlg.Controls("txt_seed").Value,
            }
        )
        dlg.Terminate()
        dlg.Dispose()
        self.options = options
        return options

    def free(self):
        self.bas.Dispose()
        self.ui.Dispose()
        self.platform.Dispose()
        self.session.Dispose()
        self.set_finished()

    def update_status(self, text: str, progress: float = 0.0):
        """
        Updates the status to the frontend and the progress from 0 to 100
        """
        if progress:
            self.progress = progress
        self.ui.SetStatusbar(text.rjust(32), self.progress)

    def set_finished(self):
        """
        Tells the frontend that the process has finished successfully
        """
        self.ui.SetStatusbar("")

    def __msg_usr__(self, message, buttons=0, title="", url=""):
        """
        Shows a message dialog, if url is given, shows
        OK, Cancel, when the user presses OK, opens the URL in the
        browser
        """
        if url:
            buttons = self.bas.MB_OKCANCEL | buttons
            res = self.bas.MsgBox(
                message,
                buttons=buttons,
                title=title,
            )
            if res == self.bas.IDOK:
                self.session.OpenURLInBrowser(url)
            return

        self.bas.MsgBox(message, buttons=buttons, title=title)

    def show_error(self, message, url="", title="", buttons=0):
        """
        Shows an error message dialog
        if url is given, shows OK, Cancel, when the user presses OK, opens the URL in the
        browser
        title is the title of the dialog to be shown
        buttons are the options that the user can have
        """
        if title == "":
            title = _("Watch out!")
        buttons = buttons | self.bas.MB_ICONSTOP
        self.__msg_usr__(message, buttons=buttons, title=title, url=url)
        self.set_finished()

    def show_message(self, message, url="", title="", buttons=0):
        """
        Shows an informative message dialog
        if url is given, shows OK, Cancel, when the user presses OK, opens the URL in the
        browser
        title is the title of the dialog to be shown
        buttons are the options that the user can have
        """
        if title == "":
            title = _("Good")
        self.__msg_usr__(
            message,
            buttons=buttons + self.bas.MB_ICONINFORMATION,
            title=title,
            url=url,
        )

    def insert_image(
        self, img_path: str, width: int, height: int, sh_client: AiHordeClient
    ):
        """
        Inserts the image with width*height from the path in the document adding
        the accessibility data from sh_client.

        Depending on self.inside type document, inserts directly in drawing,
        spreadsheet, presentation and text document, when in other type of document,
        creates a new text document and inserts the image in there.
        """
        # Normalizing pixel to cm
        self.image_insert_to[self.inside](
            img_path, width * 25.4, height * 25.4, sh_client
        )

    def __insert_image_in_presentation__(
        self, img_path: str, width: int, height: int, sh_client: AiHordeClient
    ):
        """
        Inserts the image with width*height from the path in the document adding
        the accessibility data from sh_client in a presentation document centered
        and above all other elements in the current page.
        """

        size = Size(width, height)
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1presentation_1_1GraphicObjectShape.html
        image = self.doc.createInstance("com.sun.star.presentation.GraphicObjectShape")
        image.GraphicURL = uno.systemPathToFileUrl(img_path)

        ctrllr = self.model.CurrentController
        draw_page = ctrllr.CurrentPage

        draw_page.addTop(image)
        added_image = draw_page[-1]
        added_image.setSize(size)
        position = Point(
            ((added_image.Parent.Width - width) / 2),
            ((added_image.Parent.Height - height) / 2),
        )
        added_image.setPosition(position)
        added_image.setPropertyValue("ZOrder", draw_page.Count)

        # The placeholder does not update
        # https://bugs.documentfoundation.org/show_bug.cgi?id=167809
        added_image.PlaceholderText = ""
        added_image.setPropertyValue("PlaceholderText", "")

        added_image.setPropertyValue("Title", sh_client.get_title())
        added_image.setPropertyValue("Name", sh_client.get_imagename())
        added_image.setPropertyValue("Description", sh_client.get_full_description())
        added_image.Visible = True
        self.model.Modified = True
        os.unlink(img_path)

    def __insert_image_in_calc__(
        self, img_path: str, width: int, height: int, sh_client: AiHordeClient
    ):
        """
        Inserts the image with width*height from the path in the document adding
        the accessibility data from sh_client in a calc spreadsheet left topmost
        and above all other elements in the current page.
        """

        size = Size(width, height)
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1GraphicObjectShape.html
        # https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1sheet.html
        image = self.doc.createInstance("com.sun.star.drawing.GraphicObjectShape")
        image.GraphicURL = uno.systemPathToFileUrl(img_path)

        ctrllr = self.model.CurrentController
        draw_page = ctrllr.ActiveSheet.DrawPage

        draw_page.addTop(image)
        added_image = draw_page[-1]
        added_image.setSize(size)
        added_image.setPropertyValue("ZOrder", draw_page.Count)

        added_image.Title = sh_client.get_title()
        added_image.Name = sh_client.get_imagename()
        added_image.Description = sh_client.get_full_description()
        added_image.Visible = True
        self.model.Modified = True
        os.unlink(img_path)

    def __insert_image_in_draw__(
        self, img_path: str, width: int, height: int, sh_client: AiHordeClient
    ):
        """
        Inserts the image with width*height from the path in the document adding
        the accessibility data from sh_client in a draw document centered
        and above all other elements in the current page.
        """

        size = Size(width, height)
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1GraphicObjectShape.html
        image = self.doc.createInstance("com.sun.star.drawing.GraphicObjectShape")
        image.GraphicURL = uno.systemPathToFileUrl(img_path)

        ctrllr = self.model.CurrentController
        draw_page = ctrllr.CurrentPage

        draw_page.addTop(image)
        added_image = draw_page[-1]
        added_image.setSize(size)
        position = Point(
            ((added_image.Parent.Width - width) / 2),
            ((added_image.Parent.Height - height) / 2),
        )
        added_image.setPosition(position)
        added_image.setPropertyValue("ZOrder", draw_page.Count)

        added_image.setPropertyValue("Title", sh_client.get_title())
        added_image.setPropertyValue("Name", sh_client.get_imagename())
        added_image.setPropertyValue("Description", sh_client.get_full_description())
        added_image.Visible = True
        self.model.Modified = True
        os.unlink(img_path)

    def __insert_image_in_text__(
        self, img_path: str, width: int, height: int, sh_client: AiHordeClient
    ):
        """
        Inserts the image with width*height from the path in the document adding
        the accessibility data from sh_client in the current document
        at cursor position with the same text next to it.
        """
        logger.debug(f"Inserting {img_path} in writer")
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextGraphicObject.html
        image = self.doc.createInstance("com.sun.star.text.GraphicObject")
        image.GraphicURL = uno.systemPathToFileUrl(img_path)
        image.AnchorType = AS_CHARACTER
        image.Width = width
        image.Height = height
        image.Tooltip = sh_client.get_tooltip()
        image.Name = sh_client.get_imagename()
        image.Description = sh_client.get_full_description()
        image.Title = sh_client.get_title()

        ctrllr = self.model.CurrentController
        curview = ctrllr.ViewCursor
        curview.String = self.options["prompt"] + _(" by ") + self.options["model"]
        self.doc.Text.insertTextContent(curview, image, False)
        os.unlink(img_path)

    def get_frontend_property(self, property_name: str) -> Union[str, bool, None]:
        """
        Gets a property from the frontend application, used to retrieved stored
        information during this session.  Used when checking for update
        """
        value = None
        oDocProps = self.model.getDocumentProperties()
        self.userProps = oDocProps.getUserDefinedProperties()
        try:
            value = self.userProps.getPropertyValue(property_name)
        except UnknownPropertyException:
            # Expected when the property was not present
            return False
        return value

    def has_asked_for_update(self) -> bool:
        return (
            False if not self.get_frontend_property(PROPERTY_CURRENT_SESSION) else True
        )

    def just_asked_for_update(self) -> None:
        self.set_frontend_property(PROPERTY_CURRENT_SESSION, True)

    def set_frontend_property(self, property_name: str, value: Union[str, bool]):
        """
        Sets a property in the frontend application, used to retrieved stored
        information during this session.  Used when checking for update.
        """
        oDocProps = self.model.getDocumentProperties()
        self.userProps = oDocProps.getUserDefinedProperties()
        if value is None:
            str_value = ""
        else:
            str_value = str(value)

        try:
            self.userProps.addProperty(property_name, TRANSIENT, str_value)
        except PropertyExistException:
            # It's ok, if the property existed, we update it
            self.userProps.setPropertyValue(property_name, str_value)

    def path_store_directory(self) -> str:
        """
        Returns the basepath for the directory offered by the frontend
        to store data for the plugin, cache and user settings
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1util_1_1PathSubstitution.html

        create_service = self.context.ServiceManager.createInstance
        path_finder = create_service("com.sun.star.util.PathSubstitution")

        config_url = path_finder.substituteVariables("$(user)/config", True)
        config_path = uno.fileUrlToSystemPath(config_url)
        return Path(config_path)


def get_locale_dir(extid):
    ctx = uno.getComponentContext()
    pip = ctx.getByName(
        "/singletons/com.sun.star.deployment.PackageInformationProvider"
    )
    extpath = pip.getPackageLocation(extid)
    locdir = os.path.join(uno.fileUrlToSystemPath(extpath), "locale")
    logger.debug(f"Locales folder: {locdir}")
    return locdir


def generate_image(desktop=None, context=None):
    """Creates an image from a prompt provided by the user, making use
    of AI Horde"""

    gettext.bindtextdomain(GETTEXT_DOMAIN, get_locale_dir(LIBREOFFICE_EXTENSION_ID))

    lo_manager = LibreOfficeInteraction(desktop, context)
    st_manager = HordeClientSettings(lo_manager.path_store_directory())
    saved_options = st_manager.load()
    sh_client = AiHordeClient(
        VERSION,
        URL_VERSION_UPDATE,
        HELP_URL,
        URL_DOWNLOAD,
        saved_options,
        lo_manager.base_info,
        lo_manager,
    )

    logger.debug(lo_manager.base_info)

    options = lo_manager.prepare_options(saved_options)

    if options is None:
        # User cancelled, nothing to be done
        lo_manager.free()
        return
    elif len(options["prompt"]) < MIN_PROMPT_LENGTH:
        lo_manager.show_error(_("Please provide a prompt with at least 5 words"))
        lo_manager.free()
        return

    logger.debug(options)

    def __real_work_with_api__():
        """
        Once we have collected everything, we make calls to the API and
        maintain LibreOffice interface responsive, show a message and
        insert the image in the cursor position.  It will need some
        refactor to have the dialog visible and the progressbar inside
        it.we have collected everything, we make calls to the API and
        maintain LibreOffice interface responsive, show a message and
        insert the image in the cursor position.  It will need some
        refactor to have the dialog visible and the progressbar inside
        it."""
        images_paths = sh_client.generate_image(options)

        logger.debug(images_paths)
        if images_paths:
            bas = CreateScriptService("Basic")
            bas.MsgBox(_("Your image was generated"), title=_("AIHorde has good news"))
            bas.Dispose()

            lo_manager.insert_image(
                images_paths[0],
                options["image_width"],
                options["image_height"],
                sh_client,
            )
            st_manager.save(sh_client.get_settings())

        lo_manager.free()

    from threading import Thread

    t = Thread(target=__real_work_with_api__)
    t.start()


class AiHordeForLibreOffice(unohelper.Base, XJobExecutor, XEventListener):
    """Service that creates images from text. The url to be invoked is:
    service:org.fectp.AIHordeForLibreOffice
    """

    def trigger(self, args):
        if args == "create_image":
            generate_image(self.desktop, self.context)
        # if args == "validate_form":
        #     print(dir(self))
        # if args == "get_help":
        #     print(dir(self))

    def __init__(self, context):
        self.context = context
        # see https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1frame_1_1Desktop.html
        self.desktop = self.createUnoService("com.sun.star.frame.Desktop")
        # see https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1frame_1_1DispatchHelper.html
        self.dispatchhelper = self.createUnoService("com.sun.star.frame.DispatchHelper")

    def createUnoService(self, name):
        """little helper function to create services in our context"""
        # see https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1lang_1_1ServiceManager.html
        # see https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XMultiComponentFactory.html#a77f975d2f28df6d1e136819f78a57353
        return self.context.ServiceManager.createInstanceWithContext(name, self.context)

    def disposing(self, args):
        pass

    def notifyEvent(self, args):
        pass


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    AiHordeForLibreOffice,
    LIBREOFFICE_EXTENSION_ID,
    ("com.sun.star.task.JobExecutor",),
)

# TODO:
# * [X] Integrate changes from gimp work
# * [X] Issue bugs for Impress with placeholdertext bug 167809
# * [X] Get back support fot python 3.8
#   and showing the dialog with prefilled with the same
#   telling the user that it was NSFW
# * [X] Add information to the image https://discord.com/channels/781145214752129095/1401005281332433057/1406114848810467469
# * [X] Add to Impress and also to Sheets and Drawing
# * [ ] Add option on the Dialog to show debug
# * [ ] Add tips to show. Localized messages. Inpainting, Gimp.
# * [ ] Use singleton path for the config path
#       https://ask.libreoffice.org/t/what-is-the-proper-place-to-store-settings-for-an-extension-python/125134/6
# * [X] Add to Calc
# * [X] Add to Draw
# * [X] Repo for client and use it as a submodule
# * [X] Handle Warnings.  For each model the restrictions are present in
# * [X] Internationalization
# * [ ] Wishlist to have right alignment for numeric control option
# * [ ] Recommend to use a shared key to users
# * [X] Make release in Github
# * [X] Modify Makefile to upload to github with gh
# * [ ] Automate version propagation when publishing
# * [X] Fix Makefile for patterns on gettext languages
# * [ ] Add a popup context menu: Generate Image... [programming] https://wiki.documentfoundation.org/Macros/ScriptForge/PopupMenuExample
# * [ ] Use styles support from Horde
#    -  Show Styles and Advanced View
#    -  Download and cache Styles
#
# Local documentation
# file:///usr/share/doc/libreoffice-dev-doc/api/
# /usr/lib/libreoffice/share/Scripts/python/
# /usr/lib/libreoffice/sdk/examples/html
#
# doc = XSCRIPTCONTEXT.getDocument()
# doc.get_current_controller().ComponentWindow.StyleSettings.LightColor.is_dark()
# ctx = uno.getComponentContext()
# uno.getComponentContext().getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
