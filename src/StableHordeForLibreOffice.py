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
import logging
import os
import sys
import tempfile
import uno
import unohelper

from com.sun.star.awt import Point
from com.sun.star.awt import PosSize
from com.sun.star.awt import Size
from com.sun.star.awt import XActionListener
from com.sun.star.awt import XFocusListener
from com.sun.star.awt import ActionEvent
from com.sun.star.awt import FocusEvent
from com.sun.star.awt import KeyEvent
from com.sun.star.awt import SpinEvent
from com.sun.star.awt import TextEvent
from com.sun.star.awt import XKeyListener
from com.sun.star.awt import XSpinListener
from com.sun.star.awt import XTextListener
from com.sun.star.awt.Key import ESCAPE
from com.sun.star.beans import PropertyExistException
from com.sun.star.beans import UnknownPropertyException
from com.sun.star.beans.PropertyAttribute import TRANSIENT
from com.sun.star.document import XEventListener
from com.sun.star.task import XJobExecutor
from com.sun.star.text.TextContentAnchorType import AS_CHARACTER

from math import sqrt
from pathlib import Path
from scriptforge import CreateScriptService
from typing import Any, Dict, List, Union

# Change the next line replacing False to True if you need to debug. Case matters
DEBUG = False

VERSION = "0.6"

import_message_error = None

logger = logging.getLogger(__name__)
LOGGING_LEVEL = logging.ERROR

LIBREOFFICE_EXTENSION_ID = "org.fectp.StableHordeForLibreOffice"
GETTEXT_DOMAIN = "stablehordeforlibreoffice"

log_file = os.path.join(tempfile.gettempdir(), "libreoffice_shotd.log")
if DEBUG:
    LOGGING_LEVEL = logging.DEBUG
logging.basicConfig(
    filename=log_file,
    level=LOGGING_LEVEL,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

script_path = os.path.realpath(__file__)
file_path = os.path.dirname(script_path)
submodule_path = os.path.join(file_path, "module")
sys.path.append(str(submodule_path))

from aihordeclient import (  # noqa: E402
    ANONYMOUS_KEY,  # noqa: E402
    DEFAULT_MODEL,
    MIN_PROMPT_LENGTH,
    MAX_MP,
    MIN_WIDTH,
    MAX_WIDTH,
    MIN_HEIGHT,
    MAX_HEIGHT,
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

HORDE_CLIENT_NAME = "AiHordeForLibreOffice"
"""
Name of the client sent to API
"""
DEFAULT_HEIGHT = 384
DEFAULT_WIDTH = 384


# gettext usual alias for i18n
_ = gettext.gettext
gettext.textdomain(GETTEXT_DOMAIN)


def popup_click(poEvent: uno = None) -> None:
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


class LibreOfficeInteraction(
    unohelper.Base,
    InformerFrontend,
    XActionListener,
    XKeyListener,
    XTextListener,
    XSpinListener,
    XFocusListener,
):
    def get_type_doc(self, doc):
        TYPE_DOC = {
            "calc": "com.sun.star.sheet.SpreadsheetDocument",
            "draw": "com.sun.star.drawing.DrawingDocument",
            "impress": "com.sun.star.presentation.PresentationDocument",
            "web": "com.sun.star.text.WebDocument",
            "writer": "com.sun.star.text.TextDocument",
        }
        for k, v in TYPE_DOC.items():
            if doc.supportsService(v):
                return k
        return "new-writer"

    def __init__(self, desktop, context):
        self.options = {}
        self.desktop = desktop
        self.context = context
        self.model = self.desktop.getCurrentComponent()
        self.toolkit = self.context.ServiceManager.createInstanceWithContext(
            "com.sun.star.awt.ExtToolkit", self.context
        )

        self.in_progress = False

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

        self.DEFAULT_DLG_HEIGHT = 216
        self.displacement = 180
        self.ok_btn = None
        self.dlg = self.__create_dialog__()

    def __create_dialog__(self):
        def create_widget(
            dlg, typename: str, identifier: str, x: int, y: int, width: int, height: int
        ):
            """
            Adds to the dlg a control Model, with the identifier, positioned with
            widthxheight. For typename see UnoControl* at
            https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt.html
            """
            cmpt_type = f"com.sun.star.awt.UnoControl{typename}Model"
            cmpt = dlg.createInstance(cmpt_type)
            cmpt.Name = identifier
            cmpt.PositionX = str(x)
            cmpt.PositionY = str(y)
            cmpt.Width = width
            cmpt.Height = height
            dlg.insertByName(identifier, cmpt)
            return cmpt

        dc = self.context.ServiceManager.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialog", self.context
        )
        dm = self.context.ServiceManager.createInstance(
            "com.sun.star.awt.UnoControlDialogModel"
        )
        dc.setModel(dm)
        dc.addKeyListener(self)
        dm.Name = "stablehordeoptions"
        dm.PositionX = "47"
        dm.PositionY = "10"
        dm.Width = 265
        dm.Height = self.DEFAULT_DLG_HEIGHT
        dm.Closeable = True
        dm.Moveable = True
        dm.Title = _("AI Horde for LibreOffice - ") + VERSION

        create_widget(dm, "GroupBox", "framebox", 16, 9, 236, 165)
        lbl = create_widget(dm, "FixedText", "label_prompt", 29, 31, 45, 13)
        lbl.Label = _("Prompt")
        lbl = create_widget(dm, "FixedText", "label_height", 155, 65, 45, 13)
        lbl.Label = _("Height")
        lbl = create_widget(dm, "FixedText", "label_width", 29, 65, 45, 13)
        lbl.Label = _("Width")
        lbl = create_widget(dm, "FixedText", "label_model", 29, 82, 45, 13)
        lbl.Label = _("Model")
        lbl = create_widget(dm, "FixedText", "label_max_wait", 155, 82, 45, 13)
        lbl.Label = _("Max Wait")
        lbl = create_widget(dm, "FixedText", "label_strength", 29, 99, 45, 13)
        lbl.Label = _("Strength")
        lbl = create_widget(dm, "FixedText", "label_steps", 155, 99, 45, 13)
        lbl.Label = _("Steps")
        lbl = create_widget(dm, "FixedText", "label_seed", 96, 130, 49, 13)
        lbl.Label = _("Seed (Optional)")
        lbl = create_widget(dm, "FixedText", "label_token", 96, 149, 49, 13)
        lbl.Label = _("ApiKey (Optional)")
        if DEBUG:
            ctrl = create_widget(dm, "FixedText", "lbl_debug", 19, 162, 50, 10)
            ctrl.Label = f"📜 {log_file}"
            ctrl.HelpText = (
                _(
                    "You are debugging, make sure opening LibreOffice from the command line. Consider using"
                )
                + f"\n\n   tailf { log_file }"
            )

        # Buttons
        button_ok = create_widget(dm, "Button", "btn_ok", 73, 182, 49, 13)
        button_ok.Label = _("Process")
        button_ok.TabIndex = 4
        dc.getControl("btn_ok").addActionListener(self)
        dc.getControl("btn_ok").setActionCommand("btn_ok_OnClick")
        self.ok_btn = button_ok

        button_cancel = create_widget(dm, "Button", "btn_cancel", 145, 182, 49, 13)
        button_cancel.Label = _("Cancel")
        button_cancel.TabIndex = 13
        dc.getControl("btn_cancel").addActionListener(self)
        dc.getControl("btn_cancel").setActionCommand("btn_cancel_OnClick")

        button_help = create_widget(dm, "Button", "btn_help", 23, 15, 13, 10)
        button_help.Label = "?"
        button_help.HelpText = _("About Horde")
        button_help.TabIndex = 14
        dc.getControl("btn_help").addActionListener(self)
        dc.getControl("btn_help").setActionCommand("btn_help_OnClick")

        button_toggle = create_widget(dm, "Button", "btn_toggle", 2, 204, 12, 10)
        button_toggle.Label = "_"
        button_toggle.HelpText = _("Toggle")
        button_toggle.TabIndex = 15
        dc.getControl("btn_toggle").addActionListener(self)
        dc.getControl("btn_toggle").setActionCommand("btn_toggle_OnClick")

        # Controls
        ctrl = create_widget(
            dm,
            "ComboBox",
            "lst_model",
            60,
            80,
            79,
            15,
        )
        ctrl.TabIndex = 3
        ctrl.Dropdown = True
        ctrl.LineCount = 10

        ctrl = create_widget(
            dm,
            "Edit",
            "txt_prompt",
            60,
            16,
            188,
            42,
        )
        ctrl.MultiLine = True
        ctrl.TabIndex = 1
        ctrl.HelpText = _("""
        Let your imagination run wild or put a proper description of your
        desired output. Use full grammar for Flux, use tag-like language
        for sd15, use short phrases for sdxl.

        Write at least 5 words or 10 characters.
        """)
        dc.getControl("txt_prompt").addTextListener(self)
        dc.getControl("txt_prompt").addKeyListener(self)

        ctrl = create_widget(dm, "Edit", "txt_token", 155, 147, 92, 13)
        ctrl.TabIndex = 11
        ctrl.HelpText = _("""
        Get yours at https://aihorde.net/ for free. Recommended:
        Anonymous users are last in the queue.
        """)

        ctrl = create_widget(dm, "Edit", "txt_seed", 155, 128, 92, 13)
        ctrl.TabIndex = 2
        ctrl.HelpText = _(
            "Set a seed to regenerate (reproducible), or it'll be chosen at random by the worker."
        )

        ctrl = create_widget(dm, "NumericField", "int_width", 91, 63, 48, 13)
        ctrl.DecimalAccuracy = 0
        ctrl.ValueMin = MIN_WIDTH
        ctrl.ValueMax = MAX_WIDTH
        ctrl.ValueStep = 64
        ctrl.Spin = True
        ctrl.Value = DEFAULT_WIDTH
        ctrl.TabIndex = 5
        ctrl.HelpText = _(
            "Height and Width together at most can be 2048x2048=4194304 pixels"
        )
        dc.getControl("int_width").addTextListener(self)
        dc.getControl("int_width").addSpinListener(self)
        dc.getControl("int_width").addFocusListener(self)

        ctrl = create_widget(dm, "NumericField", "int_strength", 91, 100, 48, 13)
        ctrl.ValueMin = 0
        ctrl.ValueMax = 20
        ctrl.ValueStep = 0.5
        ctrl.DecimalAccuracy = 2
        ctrl.Spin = True
        ctrl.Value = 15
        ctrl.TabIndex = 7
        ctrl.HelpText = _("""
         How strongly the AI follows the prompt vs how much creativity to allow it.
        Set to 1 for Flux, use 2-4 for LCM and lightning, 5-7 is common for SDXL
        models, 6-9 is common for sd15.
        """)

        ctrl = create_widget(
            dm,
            "NumericField",
            "int_height",
            200,
            63,
            48,
            13,
        )
        ctrl.DecimalAccuracy = 0
        ctrl.ValueMin = MIN_HEIGHT
        ctrl.ValueMax = MAX_HEIGHT
        ctrl.ValueStep = 64
        ctrl.Spin = True
        ctrl.Value = DEFAULT_HEIGHT
        ctrl.TabIndex = 6
        ctrl.HelpText = _(
            "Height and Width together at most can be 2048x2048=4194304 pixels"
        )
        dc.getControl("int_height").addTextListener(self)
        dc.getControl("int_height").addSpinListener(self)
        dc.getControl("int_height").addFocusListener(self)

        ctrl = create_widget(
            dm,
            "NumericField",
            "int_waiting",
            200,
            80,
            48,
            13,
        )
        ctrl.ValueMin = 1
        ctrl.ValueMax = 15
        ctrl.Spin = True
        ctrl.DecimalAccuracy = 0
        ctrl.Value = 5
        ctrl.TabIndex = 8
        ctrl.HelpText = _("""
        How long to wait(minutes) for your generation to complete.
        Depends on number of workers and user priority (more
        kudos = more priority. Anonymous users are last)
        """)

        ctrl = create_widget(
            dm,
            "NumericField",
            "int_steps",
            200,
            97,
            48,
            13,
        )
        ctrl.ValueMin = 1
        ctrl.ValueMax = 150
        ctrl.Spin = True
        ctrl.ValueStep = 10
        ctrl.DecimalAccuracy = 0
        ctrl.Value = 25
        ctrl.TabIndex = 7
        ctrl.HelpText = _("""
        How many sampling steps to perform for generation. Should
        generally be at least double the CFG unless using a second-order
        or higher sampler (anything with dpmpp is second order)
        """)

        ctrl = create_widget(dm, "CheckBox", "bool_nsfw", 29, 130, 55, 10)
        ctrl.Label = _("NSFW")
        ctrl.TabIndex = 9
        ctrl.HelpText = _("""
        Whether or not your image is intended to be NSFW. May
        reduce generation speed (workers can choose if they wish
        to take nsfw requests)
        """)

        ctrl = create_widget(dm, "CheckBox", "bool_censure", 29, 145, 55, 10)
        ctrl.Label = _("Censor NSFW")
        ctrl.TabIndex = 10
        ctrl.HelpText = _("""
        Separate from the NSFW flag, should workers
        return nsfw images. Censorship is implemented to be safe
        and overcensor rather than risk returning unwanted NSFW.
        """)
        lbl = create_widget(dm, "FixedText", "label_progress", 20, 205, 150, 10)
        lbl.Label = ""
        ctrl = create_widget(
            dm,
            "ProgressBar",
            "prog_status",
            14,
            203,
            253,
            13,
        )
        self.progress_label = lbl
        self.progress_meter = ctrl

        return dc

    def show_ui(self):
        self.dlg.setVisible(True)
        self.dlg.createPeer(self.toolkit, None)
        self.dlg.getControl("btn_toggle").setVisible(False)
        size = self.dlg.getPosSize()
        self.DEFAULT_DLG_HEIGHT = size.Height
        self.displacement = self.dlg.getControl("label_progress").getPosSize().Y - 5

    def toggle_dialog(self):
        size = self.dlg.getPosSize()
        lbl = self.dlg.getControl("label_progress")
        btn = self.dlg.getControl("btn_toggle")
        prg = self.dlg.getControl("prog_status")
        frame = self.dlg.getControl("framebox")
        if size.Height == self.DEFAULT_DLG_HEIGHT:
            frame.setVisible(False)
            self.dlg.setPosSize(size.X, size.Y, size.Height, 30, PosSize.HEIGHT)
            size = lbl.getPosSize()
            lbl.setPosSize(
                size.X, size.Y - self.displacement, size.Height, size.Width, PosSize.Y
            )
            size = btn.getPosSize()
            btn.setPosSize(
                size.X, size.Y - self.displacement, size.Height, size.Width, PosSize.Y
            )
            size = prg.getPosSize()
            prg.setPosSize(
                size.X, size.Y - self.displacement, size.Height, size.Width, PosSize.Y
            )
        else:
            self.dlg.setPosSize(
                size.X, size.Y, size.Width, self.DEFAULT_DLG_HEIGHT, PosSize.HEIGHT
            )
            size = lbl.getPosSize()
            lbl.setPosSize(
                size.X, size.Y + self.displacement, size.Height, size.Width, PosSize.Y
            )
            size = btn.getPosSize()
            btn.setPosSize(
                size.X, size.Y + self.displacement, size.Height, size.Width, PosSize.Y
            )
            size = prg.getPosSize()
            prg.setPosSize(
                size.X, size.Y + self.displacement, size.Height, size.Width, PosSize.Y
            )
            frame.setVisible(True)

    def validate_fields(self) -> None:
        if self.in_progress:
            return
        enable_ok = (
            len(self.dlg.getControl("txt_prompt").Text) > MIN_PROMPT_LENGTH
            and self.dlg.getControl("int_width").Value
            * self.dlg.getControl("int_height").Value
            <= MAX_MP
        )
        self.ok_btn.Enabled = enable_ok

    def focusLost(self, oFocusEvent: FocusEvent) -> None:
        element = oFocusEvent.Source.getModel()
        if element.Name not in ["int_width", "int_height"]:
            return

        pixels = (
            self.dlg.getControl("int_width").Value
            * self.dlg.getControl("int_height").Value
        )
        if pixels >= MAX_MP:
            element.Value = 64 * int(sqrt((pixels - MAX_MP)) / 64)
        self.validate_fields()

    def down(self, oSpinActed: SpinEvent) -> None:
        if oSpinActed.Source.getModel().Name in ["int_width", "int_height"]:
            self.validate_fields()

    def up(self, oSpinActed: SpinEvent) -> None:
        if oSpinActed.Source.getModel().Name in ["int_width", "int_height"]:
            self.validate_fields()

    def keyReleased(self, oKeyReleased: KeyEvent) -> None:
        if oKeyReleased.KeyCode == ESCAPE:
            self.dlg.dispose()

    def textChanged(self, oTextChanged: TextEvent) -> None:
        if oTextChanged.Source.getModel().Name in [
            "txt_prompt",
            "int_width",
            "int_height",
        ]:
            self.validate_fields()

    def actionPerformed(self, oActionEvent: ActionEvent) -> None:
        """
        Function invoked when an event is fired from a widget
        """
        if oActionEvent.ActionCommand == "btn_ok_OnClick":
            if self.in_progress:
                # Fast clickers,need to learn to wait
                return
            self.in_progress = True
            self.start_processing()
        elif oActionEvent.ActionCommand == "btn_toggle_OnClick":
            self.toggle_dialog()
        elif oActionEvent.ActionCommand == "btn_cancel_OnClick":
            self.dlg.dispose()
        elif oActionEvent.ActionCommand == "btn_help_OnClick":
            self.session.OpenURLInBrowser(HELP_URL)

    def get_options_from_dialog(self) -> List[Dict[str, Any]]:
        """
        Updates the options from the dialog ready to be used
        """
        self.options.update(
            {
                "prompt": self.dlg.getControl("txt_prompt").Text,
                "image_width": self.dlg.getControl("int_width").Value,
                "image_height": self.dlg.getControl("int_height").Value,
                "model": self.dlg.getControl("lst_model").Text,
                "prompt_strength": self.dlg.getControl("int_strength").Value,
                "steps": self.dlg.getControl("int_steps").Value,
                "nsfw": self.dlg.getControl("bool_nsfw").State == 1,
                "censor_nsfw": self.dlg.getControl("bool_censure").State == 1,
                "api_key": self.dlg.getControl("txt_token").Text or ANONYMOUS_KEY,
                "max_wait_minutes": self.dlg.getControl("int_waiting").Value,
                "seed": self.dlg.getControl("txt_seed").Text,
            }
        )

        return self

    def start_processing(self) -> None:
        self.dlg.getControl("btn_ok").getModel().Enabled = False
        self.dlg.getControl("btn_toggle").setVisible(True)
        cancel_button = self.dlg.getControl("btn_cancel")
        cancel_button.setLabel(_("Close"))
        cancel_button.getModel().HelpText = _(
            "The image generation will continue while you are doing other tasks"
        )
        self.toggle_dialog()
        self.get_options_from_dialog()
        logger.debug(self.options)

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
            images_paths = self.sh_client.generate_image(self.options)

            logger.debug(images_paths)
            if images_paths:
                self.update_status("", 100)
                bas = CreateScriptService("Basic")
                bas.MsgBox(
                    _("Your image was generated"), title=_("AIHorde has good news")
                )
                bas.Dispose()

                self.insert_image(
                    images_paths[0],
                    self.options["image_width"],
                    self.options["image_height"],
                    self.sh_client,
                )
                self.st_manager.save(self.sh_client.get_settings())

            self.free()

        from threading import Thread

        self.worker = Thread(target=__real_work_with_api__)
        self.worker.start()

    def prepare_options(
        self,
        sh_client: AiHordeClient,
        st_manager: HordeClientSettings,
        options: List[Dict[str, Any]],
    ) -> None:
        self.options.update(options)
        self.sh_client = sh_client
        self.st_manager = st_manager
        dlg = self.dlg
        api_key = options.get("api_key", ANONYMOUS_KEY)
        ctrl_token = dlg.getControl("txt_token")
        ctrl_token.setText(api_key)
        if api_key == ANONYMOUS_KEY:
            ctrl_token.setText("")
            ctrl_token.getModel().TabIndex = 1

        choices = options.get("local_settings", {"models": MODELS}).get(
            "models", MODELS
        )
        choices = choices or MODELS
        lst_rep_model = dlg.getControl("lst_model").getModel()
        for i in range(len(choices)):
            lst_rep_model.insertItemText(i, choices[i])
        dlg.getControl("lst_model").Text = options.get("model", DEFAULT_MODEL)
        # dlg.getControl("btn_ok").Enabled = len(self.selected) > MIN_PROMPT_LENGTH

        dlg.getControl("txt_prompt").setText(options.get("prompt", ""))
        dlg.getControl("txt_seed").setText(options.get("seed", ""))
        dlg.getControl("int_width").getModel().Value = options.get(
            "image_width", DEFAULT_WIDTH
        )
        dlg.getControl("int_height").getModel().Value = options.get(
            "image_height", DEFAULT_HEIGHT
        )
        dlg.getControl("int_strength").getModel().Value = options.get(
            "prompt_strength", 6.3
        )
        dlg.getControl("int_steps").getModel().Value = options.get("steps", 25)
        dlg.getControl("int_waiting").getModel().Value = options.get(
            "max_wait_minutes", 15
        )
        dlg.getControl("bool_nsfw").State = options.get("nsfw", 0)
        dlg.getControl("bool_censure").State = options.get("censor_nsfw", 1)

    def free(self):
        self.dlg.dispose()
        self.bas.Dispose()
        self.ui.Dispose()
        self.platform.Dispose()
        self.session.Dispose()

    def update_status(self, text: str, progress: float = 0.0):
        """
        Updates the status to the frontend and the progress from 0 to 100
        """
        if progress:
            self.progress = progress
        self.progress_label.Label = text
        self.progress_meter.ProgressValue = self.progress

    def set_finished(self):
        """
        Tells the frontend that the process has finished successfully
        """
        self.progress_label.Label = ""
        self.progress_meter.ProgressValue = 100

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
        Inserts the image with width*height from img_path in the current document adding
        accessibility data from sh_client.

        Inserts directly in drawing, spreadsheet, presentation, web, xml form
        and text document, when in other type of document, creates a new text
        document and inserts the image in there.
        """

        # relative size from px
        width = width * 25.4
        height = height * 25.4

        def __insert_image_as_draw__():
            """
            Inserts the image with width*height from the path in the document adding
            the accessibility data from sh_client in a draw document centered
            and above all other elements in the current page.
            """

            size = Size(width, height)
            # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1GraphicObjectShape.html
            image = self.model.createInstance("com.sun.star.drawing.GraphicObjectShape")
            image.GraphicURL = uno.systemPathToFileUrl(img_path)

            ctrllr = self.model.CurrentController
            draw_page = ctrllr.CurrentPage

            draw_page.addTop(image)
            added_image = draw_page[-1]
            added_image.setPropertyValue("ZOrder", draw_page.Count)

            added_image.Title = sh_client.get_title()
            added_image.Name = sh_client.get_imagename()
            added_image.Description = sh_client.get_full_description()
            added_image.Visible = True
            self.model.Modified = True
            os.unlink(img_path)

            if self.inside == "calc":
                return

            added_image.setSize(size)
            position = Point(
                ((added_image.Parent.Width - width) / 2),
                ((added_image.Parent.Height - height) / 2),
            )
            added_image.setPosition(position)

        def __insert_image_in_text__():
            """
            Inserts the image with width*height from the path in the document adding
            the accessibility data from sh_client in the current document
            at cursor position with the same text next to it.
            """
            logger.debug(f"Inserting {img_path} in writer")
            # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextGraphicObject.html
            image = self.model.createInstance("com.sun.star.text.GraphicObject")
            image.GraphicURL = uno.systemPathToFileUrl(img_path)
            image.AnchorType = AS_CHARACTER
            image.Width = width
            image.Height = height
            image.Tooltip = sh_client.get_tooltip()
            image.Name = sh_client.get_imagename()
            image.Description = sh_client.get_full_description()
            image.Title = sh_client.get_title()

            curview = self.model.CurrentController.ViewCursor
            self.model.Text.insertTextContent(curview, image, False)
            os.unlink(img_path)

        image_insert_to = {
            "calc": __insert_image_as_draw__,
            "draw": __insert_image_as_draw__,
            "impress": __insert_image_as_draw__,
            "web": __insert_image_in_text__,
            "writer": __insert_image_in_text__,
        }
        # Normalizing pixel to cm
        image_insert_to[self.inside]()

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
        # https://api.libreoffice.org/docs/idl/ref/singletoncom_1_1sun_1_1star_1_1util_1_1thePathSettings.html

        pip = self.context.getByName("/singletons/com.sun.star.util.thePathSettings")
        config_path = uno.fileUrlToSystemPath(pip.BasePathUserLayer)

        return Path(config_path)


def generate_image(desktop=None, context=None):
    """Creates an image from a prompt provided by the user, making use
    of AI Horde"""

    def get_locale_dir():
        pip = context.getByName(
            "/singletons/com.sun.star.deployment.PackageInformationProvider"
        )
        extpath = pip.getPackageLocation(LIBREOFFICE_EXTENSION_ID)
        locdir = os.path.join(uno.fileUrlToSystemPath(extpath), "locale")

        return locdir

    gettext.bindtextdomain(GETTEXT_DOMAIN, get_locale_dir())

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

    lo_manager.prepare_options(sh_client, st_manager, saved_options)

    lo_manager.show_ui()


class AiHordeForLibreOffice(unohelper.Base, XJobExecutor, XEventListener):
    """Service that creates images from text. The url to be invoked is:
    service:org.fectp.AIHordeForLibreOffice
    """

    def trigger(self, args):
        if args == "create_image":
            generate_image(self.desktop, self.context)

    def __init__(self, context):
        self.context = context
        # see https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1frame_1_1Desktop.html
        self.desktop = self.createUnoService("com.sun.star.frame.Desktop")
        # see https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1frame_1_1DispatchHelper.html
        self.dispatchhelper = self.createUnoService("com.sun.star.frame.DispatchHelper")

        if DEBUG:
            print(f"your log is at {log_file}")
        else:
            message = (
                _("To view debugging messages, edit")
                + "\n\n   {}\n\n".format(script_path)
                + _("and set DEBUG to True (case matters)")
            )
            print(message)

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
    ("com.sun.star.task.theJobExecutor",),
)

# TODO:
# Great you are here looking at this, take a look at docs/CONTRIBUTING.md
# * [X] Recover form validation
#    -  Minimum length for prompt
#    -  Maximum size 4MP
# * [X] Add progress bar inside dialog
# * [ ] Cancel generation
#    - requires to have a flag to continue running or abort
#    - should_stop
#    - send delete
# * [ ] Add tips to show. Localized messages. Inpainting, Gimp.
#    - We need a panel to show the tip
#    - We need a button to expand and allow people to read
#    - Invite people to grow the knowledge
#    - They can be in github and refreshed each 10 days
#    -  url, title, description, image, visibility
# * [ ] Wishlist to have right alignment for numeric control option
# * [ ] Add translation using https://huggingface.co/spaces/Helsinki-NLP/opus-translate ,  understand how to avoid using gradio client, put the thing to be translated in the clipboard
# * [ ] Recommend to use a shared key to users
# * [ ] Automate version propagation when publishing: Wishlist for extensions
# * [ ] Add a popup context menu: Generate Image... [programming] https://wiki.documentfoundation.org/Macros/ScriptForge/PopupMenuExample
# * [ ] Use styles support from Horde
#    -  Show Styles and Advanced View
#    -  Download and cache Styles
#
