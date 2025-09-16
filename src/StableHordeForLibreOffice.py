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
import platform
import shutil
import sys
import tempfile
import uno
import unohelper

from collections import OrderedDict
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
from com.sun.star.beans import PropertyValue
from com.sun.star.beans import UnknownPropertyException
from com.sun.star.beans.PropertyAttribute import TRANSIENT
from com.sun.star.datatransfer import DataFlavor
from com.sun.star.datatransfer import XTransferable
from com.sun.star.document import XEventListener
from com.sun.star.uno import XComponentContext
from com.sun.star.task import XJobExecutor
from com.sun.star.text.TextContentAnchorType import AS_CHARACTER

from collections.abc import Iterable
from math import sqrt
from pathlib import Path
from scriptforge import CreateScriptService
from typing import Any, Dict, List, TYPE_CHECKING, Union

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlDialog
    from com.sun.star.awt import UnoControlButtonModel
    from com.sun.star.awt import UnoControlEditModel
    from com.sun.star.frame import Desktop
    from com.sun.star.frame import DispatchHelper
    from com.sun.star.awt import ExtToolkit
    from com.sun.star.awt import UnoControlDialogModel
    from com.sun.star.awt import UnoControlCheckBoxModel
    from com.sun.star.gallery import GalleryTheme
    from com.sun.star.gallery import GalleryThemeProvider

# Change the next line replacing False to True if you need to debug. Case matters
DEBUG = True

VERSION = "0.8"

import_message_error = None

logger = logging.getLogger(__name__)
LOGGING_LEVEL = logging.ERROR
GALLERY_NAME = "aihorde.net"

GALLERY_IMAGE_DIR = GALLERY_NAME + "_images"

LIBREOFFICE_EXTENSION_ID = "org.fectp.StableHordeForLibreOffice"
GETTEXT_DOMAIN = "stablehordeforlibreoffice"

log_file = os.path.realpath(Path(tempfile.gettempdir(), "libreoffice_shotd.log"))
if DEBUG:
    LOGGING_LEVEL = logging.DEBUG
logging.basicConfig(
    filename=log_file,
    level=LOGGING_LEVEL,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

script_path = os.path.realpath(__file__)
file_path = os.path.dirname(script_path)
submodule_path = os.path.join(file_path, "python_path")
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
    OPUSTM_SOURCE_LANGUAGES,
)
from aihordeclient import (  # noqa: E402
    AiHordeClient,
    InformerFrontend,
    HordeClientSettings,
    opustm_hf_translate,
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


CLIPBOARD_TEXT_FORMAT = "text/plain;charset=utf-16"

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


class DataTransferable(unohelper.Base, XTransferable):
    """Exchange data with Clipboard"""

    def __init__(self, text: str):
        dft = DataFlavor()
        dft.MimeType = CLIPBOARD_TEXT_FORMAT
        dft.HumanPresentableName = "Unicode-Text"
        self.data: Dict[str, str] = {}

        if isinstance(text, str):
            self.data[CLIPBOARD_TEXT_FORMAT] = text
        self.flavors: Iterable(DataFlavor) = (dft,)

    def getTransferData(self, flavor: DataFlavor):
        if not flavor:
            return
        return self.data.get(flavor.MimeType)

    def getTransferDataFlavors(self) -> None:
        return self.flavors

    def isDataFlavorSupported(self, flavor: DataFlavor) -> bool:
        if not flavor:
            return False
        return flavor.MimeType in self.data.keys()


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
        self.desktop: Desktop = desktop
        self.context = context
        self.cp = self.context.getServiceManager().createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", self.context
        )

        # Allows to store the URL of the image in case of an error
        self.generated_url: str = ""

        # Helps determine if on text, calc, draw, etc...
        self.model = self.desktop.getCurrentComponent()

        self.toolkit: ExtToolkit = (
            self.context.ServiceManager.createInstanceWithContext(
                "com.sun.star.awt.ExtToolkit", self.context
            )
        )

        # Used to control UI state
        self.in_progress: bool = False

        # Retrieves the previously selected text in LO
        self.selected: str = ""

        self.inside: str = "new-writer"
        self.bas = CreateScriptService("Basic")
        self.session = CreateScriptService("Session")
        self.key_debug_info = OrderedDict(
            [
                ("name", HORDE_CLIENT_NAME),
                ("version", VERSION),
                ("os", platform.system()),
                ("python", platform.python_version()),
                ("libreoffice", self.get_libreoffice_version()),
                ("arch", platform.machine()),
            ]
        )

        # Client identification to API
        self.base_info = "-_".join(self.key_debug_info.values())

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

        self.DEFAULT_DLG_HEIGHT: int = 216
        self.displacement: int = 180
        self.ok_btn: UnoControlButtonModel = None
        self.local_language: str = ""
        self.initial_prompt: str = ""
        self.dlg: UnoControlDialog = self.__create_dialog__()
        self.options: Dict[str, Any] = {}
        self.progress: float = 0.0

    def get_configuration_value(
        self,
        property_name: str,
        section: str,
        category: str = "Setup",
        prefix="org.openoffice",
    ):
        """
        To discover which properties are available as part of Setup,
        go to Tools > Options > Advanced > Open Expert configurations
        """
        node = PropertyValue()
        node.Name = "nodepath"
        node.Value = f"/{prefix}.{category}/{section}"
        prop = property_name
        try:
            cr = self.cp.createInstanceWithArguments(
                "com.sun.star.configuration.ConfigurationAccess", (node,)
            )
            if cr and (cr.hasByName(prop)):
                return cr.getPropertyValue(prop)
        except Exception as ex:
            logger.info(
                f"The property /{prefix}.{category}/{section} {property_name} is not present"
            )
            logger.debug(ex, stack_info=True)
            return ""

        logger.debug(
            f"The property /{prefix}.{category}/{section} {property_name} is not present"
        )
        return ""

    def get_libreoffice_version(self) -> str:
        return " ".join(
            [
                self.get_configuration_value("ooName", "Product"),
                self.get_configuration_value("ooSetupVersionAboutBox", "Product"),
                "(" + self.get_configuration_value("ooVendor", "Product") + ")",
            ]
        )

    def get_language(self) -> str:
        """
        Determines the UI current language
        Taken from MRI
        https://github.com/hanya/MRI
        """
        return self.get_configuration_value("ooLocale", "L10N")

    def __create_dialog__(self):
        def create_widget(
            dlg,
            typename: str,
            identifier: str,
            x: int,
            y: int,
            width: int,
            height: int,
            container=None,
        ):
            """
            Adds to the dlg a control Model, with the identifier, positioned with
            widthxheight, and put it inside container For typename see UnoControl* at
            https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt.html
            """
            if not container:
                container = dlg

            cmpt_type = f"com.sun.star.awt.UnoControl{typename}Model"
            cmpt = dlg.createInstance(cmpt_type)
            cmpt.Name = identifier
            cmpt.PositionX = str(x)
            cmpt.PositionY = str(y)
            cmpt.Width = width
            cmpt.Height = height
            container.insertByName(identifier, cmpt)
            return cmpt

        current_language = self.get_language()
        self.show_language = current_language in OPUSTM_SOURCE_LANGUAGES
        if self.show_language:
            self.local_language = current_language
            logger.info(f"locale is «{self.local_language}»")
        else:
            logger.warning("locale is «{current_locale}», non translatable")

        dc: UnoControlDialog = self.context.ServiceManager.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialog", self.context
        )
        dm: UnoControlDialogModel = self.context.ServiceManager.createInstance(
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
        button_ok: UnoControlButtonModel = create_widget(
            dm, "Button", "btn_ok", 73, 182, 49, 13
        )
        button_ok.Label = _("Process")
        button_ok.TabIndex = 4
        btn_ok = dc.getControl("btn_ok")
        btn_ok.addActionListener(self)
        btn_ok.setActionCommand("btn_ok_OnClick")
        self.ok_btn: UnoControlButtonModel = button_ok

        button_cancel = create_widget(dm, "Button", "btn_cancel", 145, 182, 49, 13)
        button_cancel.Label = _("Cancel")
        button_cancel.TabIndex = 13
        btn_cancel = dc.getControl("btn_cancel")
        btn_cancel.addActionListener(self)
        btn_cancel.setActionCommand("btn_cancel_OnClick")

        button_help = create_widget(dm, "Button", "btn_help", 250, 204, 13, 10)
        button_help.Label = "?"
        button_help.HelpText = _("About Horde")
        button_help.TabIndex = 14
        btn_help = dc.getControl("btn_help")
        btn_help.addActionListener(self)
        btn_help.setActionCommand("btn_help_OnClick")

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

        ctrl: UnoControlEditModel = create_widget(
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
        ctrl.HelpText = _("""        Let your imagination run wild or put a proper description of your
        desired output. Use full grammar for Flux, use tag-like language
        for sd15, use short phrases for sdxl.
        Write at least 5 words or 10 characters.""")
        dc.getControl("txt_prompt").addTextListener(self)
        dc.getControl("txt_prompt").addKeyListener(self)

        ctrl = create_widget(dm, "Edit", "txt_token", 155, 147, 92, 13)
        ctrl.TabIndex = 11
        ctrl.HelpText = _("""        Get yours at https://aihorde.net/ for free. Recommended:
        Anonymous users are last in the queue.""")

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

        ctrl = create_widget(dm, "NumericField", "int_strength", 91, 97, 48, 13)
        ctrl.ValueMin = 0
        ctrl.ValueMax = 20
        ctrl.ValueStep = 0.5
        ctrl.DecimalAccuracy = 2
        ctrl.Spin = True
        ctrl.Value = 15
        ctrl.TabIndex = 7
        ctrl.HelpText = _("""        How strongly the AI follows the prompt vs how much creativity to allow it.
        Set to 1 for Flux, use 2-4 for LCM and lightning, 5-7 is common for SDXL
        models, 6-9 is common for sd15.""")

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
        ctrl.HelpText = _("""        How long to wait(minutes) for your generation to complete.
        Depends on number of workers and user priority (more
        kudos = more priority. Anonymous users are last)""")

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
        ctrl.HelpText = _("""        How many sampling steps to perform for generation. Should
        generally be at least double the CFG unless using a second-order
        or higher sampler (anything with dpmpp is second order)""")

        ctrl: UnoControlCheckBoxModel = create_widget(
            dm, "CheckBox", "bool_trans", 29, 45, 30, 10
        )
        ctrl.Label = "🌏"
        ctrl.HelpText = _("""           Translate the prompt to English, wishing for the best.  If the result is not
        the expected, try toggling or changing the model""")

        if self.show_language:
            ctrl.TabIndex = 2

        ctrl = create_widget(dm, "CheckBox", "bool_nsfw", 29, 130, 55, 10)
        ctrl.Label = _("NSFW")
        ctrl.TabIndex = 9
        ctrl.HelpText = _("""        Whether or not your image is intended to be NSFW. May
        reduce generation speed (workers can choose if they wish
        to take nsfw requests)""")

        ctrl = create_widget(dm, "CheckBox", "bool_censure", 29, 145, 55, 10)
        ctrl.Label = _("Censor NSFW")
        ctrl.TabIndex = 10
        ctrl.HelpText = _("""        Separate from the NSFW flag, should workers
        return nsfw images. Censorship is implemented to be safe
        and overcensor rather than risk returning unwanted NSFW.""")

        lbl = create_widget(dm, "FixedText", "label_progress", 20, 205, 150, 10)
        lbl.Label = ""
        ctrl = create_widget(
            dm,
            "ProgressBar",
            "prog_status",
            14,
            203,
            235,
            13,
        )
        self.progress_label = lbl
        self.progress_meter = ctrl

        return dc

    def show_ui(self):
        self.dlg.setVisible(True)
        self.dlg.createPeer(self.toolkit, None)
        self.dlg.getControl("btn_toggle").setVisible(False)
        self.dlg.getControl("bool_trans").setVisible(self.show_language)
        size = self.dlg.getPosSize()
        self.DEFAULT_DLG_HEIGHT = size.Height
        self.displacement = self.dlg.getControl("label_progress").getPosSize().Y - 5

    def toggle_dialog(self):
        size = self.dlg.getPosSize()
        lbl = self.dlg.getControl("label_progress")
        btn = self.dlg.getControl("btn_toggle")
        hlp = self.dlg.getControl("btn_help")
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
            size = hlp.getPosSize()
            hlp.setPosSize(
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
            size = hlp.getPosSize()
            hlp.setPosSize(
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
        self.dlg.getControl("btn_trans").Enable = enable_ok

    def translate(self) -> None:
        if not self.show_language:
            return
        prompt_control = self.dlg.getControl("txt_prompt")
        text = prompt_control.Text
        self.initial_prompt = text
        try:
            self.update_status(_("Translating"), 1.0)
            logger.info(f"Translating {text} from {self.local_language}")
            translated_text = opustm_hf_translate(text, self.local_language)
        except Exception as ex:
            logger.info(ex, stack_info=True)
        finally:
            prompt_control.Text = translated_text or text

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
        if self.show_language and self.dlg.getControl("bool_trans").State == 1:
            self.translate()
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

        dlg.getControl("txt_prompt").setText(self.selected or options.get("prompt", ""))
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
        dlg.getControl("bool_trans").State = 1

    def free(self):
        self.dlg.dispose()
        self.bas.Dispose()
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
        if not url and self.generated_url:
            url = self.generated_url
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

        def locale_description(incoming_description):
            if self.show_language:
                full_description = (
                    f"original_prompt : {self.initial_prompt}\nsource_lang : {self.local_language}\n"
                    + incoming_description
                )
            else:
                full_description = incoming_description

            return full_description

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
            if self.inside == "calc":
                draw_page = ctrllr.ActiveSheet.DrawPage
            else:
                draw_page = ctrllr.CurrentPage

            draw_page.addTop(image)
            added_image = draw_page[-1]
            added_image.setSize(size)
            added_image.setPropertyValue("ZOrder", draw_page.Count)

            added_image.Title = sh_client.get_title()
            added_image.Name = sh_client.get_imagename()
            added_image.Description = locale_description(
                sh_client.get_full_description()
            )

            added_image.Visible = True
            self.model.Modified = True

            if self.inside == "calc":
                return

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
            image.Title = sh_client.get_title()
            image.Description = locale_description(sh_client.get_full_description())

            curview = self.model.CurrentController.ViewCursor
            self.model.Text.insertTextContent(curview, image, False)

        image_insert_to = {
            "calc": __insert_image_as_draw__,
            "draw": __insert_image_as_draw__,
            "impress": __insert_image_as_draw__,
            "web": __insert_image_in_text__,
            "writer": __insert_image_in_text__,
        }
        image_insert_to[self.inside]()

        # Add image to gallery
        self.add_image_to_gallery([img_path, sh_client.get_full_description()])

    def get_frontend_property(self, property_name: str) -> Union[str, bool, None]:
        """
        Returns the value stored for this session of the property_name, if not
        present, returns False.
        Used when checking for update.
        """
        value = None
        oDocProps = self.model.getDocumentProperties()
        userProps = oDocProps.getUserDefinedProperties()
        try:
            value = userProps.getPropertyValue(property_name)
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
        Sets property_name with value for the current session.
        Used when checking for update.
        """
        oDocProps = self.model.getDocumentProperties()
        userProps = oDocProps.getUserDefinedProperties()
        if value is None:
            str_value = ""
        else:
            str_value = str(value)

        try:
            userProps.addProperty(property_name, TRANSIENT, str_value)
        except PropertyExistException:
            # It's ok, if the property existed, we update it
            userProps.setPropertyValue(property_name, str_value)

    def add_image_to_gallery(self, image_info: List[str]) -> None:
        """
        Adds the image in image_path to the gallery theme, the image
        is moved from image_path to the store.
        """

        def path_store_images_directory() -> str:
            """
            Returns the basepath for the store objects directory
            to store images. Created if did not exist
            """
            # https://api.libreoffice.org/docs/idl/ref/singletoncom_1_1sun_1_1star_1_1util_1_1thePathSettings.html

            psettings = self.context.getByName(
                "/singletons/com.sun.star.util.thePathSettings"
            )
            images_local_path = (
                Path(uno.fileUrlToSystemPath(psettings.Storage_writable))
                / GALLERY_IMAGE_DIR
            )
            os.makedirs(images_local_path, exist_ok=True)

            return Path(images_local_path)

        def the_gallery():
            """
            Returns the default gallerytheme, creating it if if it did not exist
            https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1gallery_1_1GalleryTheme.html
            """
            themes_list: GalleryThemeProvider = (
                self.context.ServiceManager.createInstanceWithContext(
                    "com.sun.star.gallery.GalleryThemeProvider", self.context
                )
            )
            if themes_list.hasByName(GALLERY_NAME):
                logger.debug("Using existing theme gallery")
                aihorde_theme: GalleryTheme = themes_list.getByName(GALLERY_NAME)
            else:
                logger.debug("Creating theme gallery")
                aihorde_theme: GalleryTheme = themes_list.insertNewByName(GALLERY_NAME)
            return aihorde_theme

        images_dir = path_store_images_directory()
        image_filename = os.path.basename(image_info[0])
        shutil.move(image_info[0], images_dir)
        target_image = str(images_dir / image_filename)

        aihorde_theme: GalleryTheme = the_gallery()
        result = aihorde_theme.insertURLByIndex(
            uno.systemPathToFileUrl(target_image), -1
        )
        if result:
            logger.error(
                f"unable to set {target_image} in the gallery, reported result {result}"
            )
        image_ref = aihorde_theme.getByIndex(0)
        image_ref.Title = image_info[1]
        aihorde_theme.update()

    def path_store_directory(self) -> str:
        """
        Returns the basepath for the directory offered by the frontend
        to store data for the plugin, cache and user settings
        """
        # https://api.libreoffice.org/docs/idl/ref/singletoncom_1_1sun_1_1star_1_1util_1_1thePathSettings.html

        psettings = self.context.getByName(
            "/singletons/com.sun.star.util.thePathSettings"
        )
        config_path = uno.fileUrlToSystemPath(psettings.BasePathUserLayer)

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
        self.desktop: Desktop = self.createUnoService("com.sun.star.frame.Desktop")
        # see https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1frame_1_1DispatchHelper.html
        self.dispatchhelper: DispatchHelper = self.createUnoService(
            "com.sun.star.frame.DispatchHelper"
        )

        log_file = os.path.realpath(
            Path(tempfile.gettempdir(), "libreoffice_shotd.log")
        )
        if DEBUG:
            print(f"your log is at {log_file}")
        else:
            message = (
                _("To view debugging messages, edit")
                + "\n\n   {}\n\n".format(script_path)
                + _("and change DEBUG = False to DEBUG = True")
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

if __name__ == "__main__":
    """ Connect to LibreOffice proccess.
    1) Start the office in shell with command:
    libreoffice "--accept=socket,host=127.0.0.1,port=2002,tcpNoDelay=1;urp;StarOffice.ComponentContext" --norestore
    2) Run script
    """

    local_ctx: XComponentContext = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstance(
        "com.sun.star.bridge.UnoUrlResolver"
    )
    try:
        remote_ctx: XComponentContext = resolver.resolve(
            "uno:socket,"
            "host=127.0.0.1,"
            "port=2002,"
            "tcpNoDelay=1;"
            "urp;"
            "StarOffice.ComponentContext"
        )
    except Exception as err:
        print("""You are not running a libreoffice instance, to do so from a command line that has `libreoffice` in the path run:

       libreoffice "--accept=socket,host=127.0.0.1,port=2002,tcpNoDelay=1;urp;StarOffice.ComponentContext" --norestore  \nAnd try again
        """)
        err = err
        exit(1)
    desktop = remote_ctx.getServiceManager().createInstanceWithContext(
        "com.sun.star.frame.Desktop", remote_ctx
    )
    generate_image(desktop, remote_ctx)
    print("You will only see the dialog, for full testing issue: make run")


# TODO:
# Great you are here looking at this, take a look at docs/CONTRIBUTING.md
# * [X] Store retrieved images in the gallery
# * [ ] Make the extension appear in the Extension tabs
# * [ ] Add the Messagebox with path to gallery and copy system information
#       to send a bug report.
# * [ ] Show a dialog with info that will be copied to the clipboard to allow easier info sharing
# --- 0.8
# * [ ] When something fails in the middle, it's possible to show an URL to allow to recover the generated image by hand
# * [ ] Add an option to write the prompt with the image write_text
# * [ ] Add an option to use the main progressbar use_full_progress
# * [ ] Add an option to store in the gallery
# * [ ] Store and retrieve the UI options: translation, put text with the image, use full progress, add to gallery
# * [ ] Cancel generation
#    - requires to have a flag to continue running or abort
#    - should_stop
#    - send delete
# * [ ] Add tips to show. Localized messages. Inpainting, Gimp. Artbot https://artbot.site/
#    - We need a panel to show the tip
#    - Invite people to grow the knowledge
#    - They can be in github and refreshed each 10 days
#    -  url, title, description, image, visibility
# * [ ] Wishlist to have right alignment for numeric control option
# * [ ] Recommend to use a shared key to users
# * [ ] Automate version propagation when publishing: Wishlist for extensions
# * [ ] Add a popup context menu: Generate Image... [programming] https://wiki.documentfoundation.org/Macros/ScriptForge/PopupMenuExample
# * [ ] Use styles support from Horde
#    -  Show Styles and Advanced View
#    -  Download and cache Styles
#
