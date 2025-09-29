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
import time
import uno
import unohelper
import webbrowser

from com.sun.star.awt import ActionEvent
from com.sun.star.awt import FocusEvent
from com.sun.star.awt import KeyEvent
from com.sun.star.awt import MessageBoxButtons as mbb
from com.sun.star.awt import MessageBoxResults as mbr
from com.sun.star.awt import Point
from com.sun.star.awt import PosSize
from com.sun.star.awt import Size
from com.sun.star.awt import SpinEvent
from com.sun.star.awt import TextEvent
from com.sun.star.awt import XActionListener
from com.sun.star.awt import XFocusListener
from com.sun.star.awt import XKeyListener
from com.sun.star.awt import XSpinListener
from com.sun.star.awt import XTextListener
from com.sun.star.awt.Key import ESCAPE
from com.sun.star.awt.MessageBoxType import MESSAGEBOX
from com.sun.star.awt.MessageBoxType import WARNINGBOX
from com.sun.star.beans import PropertyExistException
from com.sun.star.beans import PropertyValue
from com.sun.star.beans import UnknownPropertyException
from com.sun.star.beans.PropertyAttribute import TRANSIENT
from com.sun.star.datatransfer import DataFlavor
from com.sun.star.datatransfer import XTransferable
from com.sun.star.document import XEventListener
from com.sun.star.task import XJobExecutor
from com.sun.star.text.TextContentAnchorType import AS_CHARACTER
from com.sun.star.text.TextContentAnchorType import AT_FRAME
from com.sun.star.uno import XComponentContext

from collections import OrderedDict
from collections.abc import Iterable
from math import sqrt
from pathlib import Path
from typing import Any, Dict, List, TYPE_CHECKING, Tuple, Union
from threading import Thread

if TYPE_CHECKING:
    from com.sun.star.awt import ExtToolkit
    from com.sun.star.awt import Toolkit
    from com.sun.star.awt import UnoControlButtonModel
    from com.sun.star.awt import UnoControlCheckBoxModel
    from com.sun.star.awt import UnoControlComboBoxModel
    from com.sun.star.awt import UnoControlDialog
    from com.sun.star.awt import UnoControlDialogModel
    from com.sun.star.awt import UnoControlEditModel
    from com.sun.star.awt import UnoControlFixedHyperlinkModel
    from com.sun.star.awt import UnoControlFixedTextModel
    from com.sun.star.awt import UnoControlModel
    from com.sun.star.awt import UnoControlNumericFieldModel
    from com.sun.star.awt.tab import UnoControlTabPageContainerModel
    from com.sun.star.awt.tab import UnoControlTabPageModel
    from com.sun.star.frame import Desktop
    from com.sun.star.frame import DispatchHelper
    from com.sun.star.gallery import GalleryTheme
    from com.sun.star.gallery import GalleryThemeProvider
    from com.sun.star.datatransfer.clipboard import SystemClipboard

# Change the next line replacing False to True if you need to debug. Case matters
DEBUG = False

VERSION = "0.9.1"

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
    REGISTER_AI_HORDE_URL,
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

URL_MODEL_SHOWCASE = "https://artbot.site/info/models"

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


def create_widget(
    container,
    typename: str,
    identifier: str,
    rect: Tuple[int, int, int, int],
    add_now: bool = True,
    additional_properties: List[Tuple[str, Any]] = None,
    insert_later: List[Tuple[Any, Any]] = None,
):
    """
    Adds to the dlg a control Model, with the identifier, positioned with
    widthxheight, and put it inside container For typename see UnoControl* at
    https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt.html
    """

    if typename.startswith("TabPage"):
        cmpt_type = f"com.sun.star.awt.tab.UnoControl{typename}Model"
    else:
        cmpt_type = f"com.sun.star.awt.UnoControl{typename}Model"

    cmpt: UnoControlModel = container.createInstance(cmpt_type)
    cmpt.setPropertyValues(
        [
            "Name",
            "PositionX",
            "PositionY",
            "Width",
            "Height",
        ],
        [
            identifier,
            rect[0],
            rect[1],
            rect[2],
            rect[3],
        ],
    )
    if add_now:
        container.insertByName(identifier, cmpt)
    else:
        insert_later.append((container, cmpt))

    if additional_properties:
        cmpt.setPropertyValues(*zip(*additional_properties))
    return cmpt


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
        self.clipboard: SystemClipboard = (
            self.context.getServiceManager().createInstanceWithContext(
                "com.sun.star.datatransfer.clipboard.SystemClipboard", self.context
            )
        )

        # Allows to store the URL of the image in case of an error
        self.generated_url: str = ""

        # Helps determine if on text, calc, draw, etc...
        self.model = self.desktop.getCurrentComponent()
        self.toolkit: Toolkit = self.context.ServiceManager.createInstanceWithContext(
            "com.sun.star.awt.Toolkit", self.context
        )
        self.extoolkit: ExtToolkit = (
            self.context.ServiceManager.createInstanceWithContext(
                "com.sun.star.awt.ExtToolkit", self.context
            )
        )

        # Used to control UI state
        self.in_progress: bool = False

        # Retrieves the previously selected text in LO
        self.selected: str = ""

        self.inside: str = "new-writer"
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

        self.DEFAULT_DLG_HEIGHT: int = 220
        self.PASSWORD_MASK = 42
        self.displacement: int = 180
        self.ok_btn: UnoControlButtonModel = None
        self.local_language: str = ""
        self.initial_prompt: str = ""
        self.dlg: UnoControlDialog = self.__create_dialog__()
        self.options: Dict[str, Any] = {"api_key": ANONYMOUS_KEY}
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
        def add_widget(
            container,
            typename: str,
            identifier: str,
            rect: Tuple[int, int, int, int],
            add_now: bool = True,
            additional_properties: List[Tuple[str, Any]] = None,
        ):
            """
            Adds to the container a control Model, with the identifier, positioned with
            widthxheight, and put it inside container if add_nos is True, the properties are added. For typename see UnoControl* at
            https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt.html
            """
            return create_widget(
                container,
                typename,
                identifier,
                rect,
                add_now,
                additional_properties,
                self.insert_in,
            )

        self.insert_in = []
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
        dc.createPeer(self.extoolkit, None)
        dc.addKeyListener(self)

        book: UnoControlTabPageContainerModel = add_widget(
            dm, "TabPageContainer", "tab_book", (18, 63, 233, 100)
        )
        page_ad: UnoControlTabPageModel = book.createTabPage(1)
        page_ad.Title = "✨ " + _("Generate")
        book.insertByIndex(0, page_ad)
        page_ux: UnoControlTabPageModel = book.createTabPage(2)
        page_ux.Title = "🛠️ " + _("Configure")
        book.insertByIndex(1, page_ux)
        page_in: UnoControlTabPageModel = book.createTabPage(3)
        page_in.Title = "💬 " + _("About")
        book.insertByIndex(1, page_in)
        self.book: UnoControlTabPageContainerModel = dc.getControl("tab_book")

        # Dialog placement
        dm.Name = "stablehordeoptions"
        dm.PositionX = "47"
        dm.PositionY = "10"
        dm.Width = 265
        dm.Height = self.DEFAULT_DLG_HEIGHT
        dm.Closeable = True
        dm.Moveable = True
        dm.Title = _("AI Horde for LibreOffice - ") + VERSION

        lbl = add_widget(dm, "FixedText", "label_prompt", (28, 18, 48, 10))
        lbl.Label = _("Prompt")
        self.bool_trans: UnoControlCheckBoxModel = add_widget(
            dm, "CheckBox", "bool_trans", (29, 33, 30, 10)
        )
        self.bool_trans.Label = "🌏"
        self.bool_trans.HelpText = _("""           Translate the prompt to English, wishing for the best.  If the result is not
        the expected, try toggling or changing the model""")

        if self.show_language:
            self.bool_trans.TabIndex = 2
        else:
            self.bool_trans.Tabstop = False
        self.bool_trans = dc.getControl(self.bool_trans.Name)
        self.bool_trans.setVisible(self.show_language)

        self.txt_prompt: UnoControlEditModel = add_widget(
            dm,
            "Edit",
            "txt_prompt",
            (
                60,
                16,
                188,
                42,
            ),
            add_now=False,
        )
        self.txt_prompt.MultiLine = True
        self.txt_prompt.TabIndex = 1
        self.txt_prompt.HelpText = _("""        Let your imagination run wild or put a proper description of your
        desired output. Use full grammar for Flux, use tag-like language
        for sd15, use short phrases for sdxl.
        Write at least 5 words or 10 characters.""")

        button_ok: UnoControlButtonModel = add_widget(
            dm, "Button", "btn_ok", (73, 186, 49, 13)
        )
        button_ok.Label = _("Process")
        button_ok.TabIndex = 4
        btn_ok = dc.getControl("btn_ok")
        btn_ok.addActionListener(self)
        btn_ok.setActionCommand("btn_ok_OnClick")
        self.ok_btn: UnoControlButtonModel = button_ok

        button_cancel = add_widget(dm, "Button", "btn_cancel", (145, 186, 49, 13))
        button_cancel.Label = _("Cancel")
        button_cancel.TabIndex = 13
        btn_cancel = dc.getControl("btn_cancel")
        btn_cancel.addActionListener(self)
        btn_cancel.setActionCommand("btn_cancel_OnClick")

        self.ctrl_token: UnoControlEditModel = add_widget(
            dm, "Edit", "txt_token", (170, 168, 80, 10), add_now=False
        )
        self.ctrl_token.TabIndex = 11
        self.ctrl_token.HelpText = _("""        Get yours at https://aihorde.net/ for free. Recommended:
        Anonymous users are last in the queue.""")
        self.ctrl_token.EchoChar = self.PASSWORD_MASK

        self.lbl_view_pass: UnoControlFixedHyperlinkModel = add_widget(
            dm, "FixedHyperlink", "lbl_view_pass", (241, 168, 10, 10)
        )
        self.lbl_view_pass.Label = "👀"
        self.lbl_view_pass.HelpText = _("""Click to view your AiHorde API Key""")
        dc.getControl("lbl_view_pass").addActionListener(self)

        self.btn_toggle: UnoControlButtonModel = add_widget(
            dm, "Button", "btn_toggle", (2, 208, 12, 10)
        )
        self.btn_toggle.Label = "_"
        self.btn_toggle.HelpText = _("Toggle")
        self.btn_toggle.TabIndex = 15

        self.btn_toggle = dc.getControl("btn_toggle")
        self.btn_toggle.addActionListener(self)
        self.btn_toggle.setActionCommand("btn_toggle_OnClick")

        lbl = add_widget(dm, "FixedText", "label_progress", (20, 208, 150, 10))
        lbl.Label = ""
        self.progress_label = lbl

        self.progress_meter = add_widget(
            dm,
            "ProgressBar",
            "prog_status",
            (
                14,
                206,
                250,
                13,
            ),
        )

        # Main page placement
        lbl = add_widget(
            page_ad, "FixedText", "label_height", (125, 8, 48, 13), add_now=False
        )
        lbl.Label = _("Height")
        lbl = add_widget(
            page_ad, "FixedText", "label_width", (5, 8, 48, 13), add_now=False
        )
        lbl.Label = _("Width")
        lbl = add_widget(
            page_ad, "FixedHyperlink", "label_model", (5, 27, 48, 13), add_now=False
        )
        lbl.Label = _("Model")
        lbl.URL = URL_MODEL_SHOWCASE
        lbl = add_widget(
            page_ad,
            "FixedText",
            "label_strength",
            (125, 46, 49, 13),
            add_now=False,
        )
        lbl.Label = _("Strength")
        lbl = add_widget(
            page_ad, "FixedText", "label_steps", (5, 46, 49, 13), add_now=False
        )
        lbl.Label = _("Steps")
        lbl = add_widget(
            page_ad, "FixedText", "label_seed", (125, 27, 49, 13), add_now=False
        )
        lbl.Label = _("Seed (Optional)")

        self.txt_seed: UnoControlEditModel = add_widget(
            page_ad, "Edit", "txt_seed", (175, 26, 48, 10), add_now=False
        )
        self.txt_seed.TabIndex = 2
        self.txt_seed.HelpText = _(
            "Set a seed to regenerate (reproducible), or it'll be chosen at random by the worker."
        )

        self.int_width: UnoControlNumericFieldModel = add_widget(
            page_ad, "NumericField", "int_width", (60, 5, 48, 13), add_now=False
        )
        self.int_width.DecimalAccuracy = 0
        self.int_width.ValueMin = MIN_WIDTH
        self.int_width.ValueMax = MAX_WIDTH
        self.int_width.ValueStep = 64
        self.int_width.Spin = True
        self.int_width.Value = DEFAULT_WIDTH
        self.int_width.TabIndex = 5
        self.int_width.HelpText = _(
            "Height and Width together at most can be 2048x2048=4194304 pixels"
        )

        self.int_strength: UnoControlNumericFieldModel = add_widget(
            page_ad, "NumericField", "int_strength", (175, 43, 48, 13), add_now=False
        )
        self.int_strength.ValueMin = 0
        self.int_strength.ValueMax = 20
        self.int_strength.ValueStep = 0.5
        self.int_strength.DecimalAccuracy = 2
        self.int_strength.Spin = True
        self.int_strength.Value = 15
        self.int_strength.TabIndex = 7
        self.int_strength.HelpText = _("""        How strongly the AI follows the prompt vs how much creativity to allow it.
        Set to 1 for Flux, use 2-4 for LCM and lightning, 5-7 is common for SDXL
        models, 6-9 is common for sd15.""")

        self.int_height: UnoControlNumericFieldModel = add_widget(
            page_ad,
            "NumericField",
            "int_height",
            (
                175,
                7,
                48,
                10,
            ),
            add_now=False,
        )
        self.int_height.DecimalAccuracy = 0
        self.int_height.ValueMin = MIN_HEIGHT
        self.int_height.ValueMax = MAX_HEIGHT
        self.int_height.ValueStep = 64
        self.int_height.Spin = True
        self.int_height.Value = DEFAULT_HEIGHT
        self.int_height.TabIndex = 6
        self.int_height.HelpText = _(
            "Height and Width together at most can be 2048x2048=4194304 pixels"
        )

        self.lst_model: UnoControlComboBoxModel = add_widget(
            page_ad,
            "ComboBox",
            "lst_model",
            (
                40,
                23,
                69,
                15,
            ),
            add_now=False,
        )
        self.lst_model.TabIndex = 3
        self.lst_model.Dropdown = True
        self.lst_model.LineCount = 10

        self.int_steps: UnoControlNumericFieldModel = add_widget(
            page_ad,
            "NumericField",
            "int_steps",
            (
                60,
                43,
                48,
                13,
            ),
            add_now=False,
        )
        self.int_steps.ValueMin = 1
        self.int_steps.ValueMax = 150
        self.int_steps.Spin = True
        self.int_steps.ValueStep = 10
        self.int_steps.DecimalAccuracy = 0
        self.int_steps.Value = 25
        self.int_steps.TabIndex = 7
        self.int_steps.HelpText = _("""        How many sampling steps to perform for generation. Should
        generally be at least double the CFG unless using a second-order
        or higher sampler (anything with dpmpp is second order)""")

        self.bool_nsfw: UnoControlCheckBoxModel = add_widget(
            page_ad, "CheckBox", "bool_nsfw", (110, 63, 55, 10), add_now=False
        )
        self.bool_nsfw.Label = _("NSFW")
        self.bool_nsfw.TabIndex = 9
        self.bool_nsfw.HelpText = _("""        Whether or not your image is intended to be NSFW. May
        reduce generation speed (workers can choose if they wish
        to take nsfw requests)""")

        self.bool_censure: UnoControlCheckBoxModel = add_widget(
            page_ad, "CheckBox", "bool_censure", (160, 63, 55, 10), add_now=False
        )
        self.bool_censure.Label = _("Censor NSFW")
        self.bool_censure.TabIndex = 10
        self.bool_censure.HelpText = _("""        Separate from the NSFW flag, should workers
        return nsfw images. Censorship is implemented to be safe
        and overcensor rather than risk returning unwanted NSFW.""")

        # Page UX placement
        lbl = add_widget(
            page_ux, "FixedText", "label_max_wait", (5, 7, 48, 13), add_now=False
        )
        lbl.Label = _("Max Wait")
        self.int_waiting: UnoControlNumericFieldModel = add_widget(
            page_ux,
            "NumericField",
            "int_waiting",
            (
                70,
                5,
                52,
                13,
            ),
            add_now=False,
        )
        self.int_waiting.ValueMin = 1
        self.int_waiting.ValueMax = 15
        self.int_waiting.Spin = True
        self.int_waiting.DecimalAccuracy = 0
        self.int_waiting.Value = 5
        self.int_waiting.TabIndex = 8
        self.int_waiting.HelpText = _("""        How long to wait(minutes) for your generation to complete.
        Depends on number of workers and user priority (more
        kudos = more priority. Anonymous users are last)""")
        self.bool_add_to_gallery: UnoControlCheckBoxModel = add_widget(
            page_ux, "CheckBox", "bool_add_to_gallery", (160, 10, 75, 10), add_now=False
        )
        self.bool_add_to_gallery.State = 1
        self.bool_add_to_gallery.Label = _("Add to Gallery")
        self.bool_add_to_gallery.TabIndex = 9
        self.bool_add_to_gallery.HelpText = _(
            """        Adds the generated image to the gallery        """
        )
        self.bool_add_frame: UnoControlCheckBoxModel = add_widget(
            page_ux, "CheckBox", "bool_add_frame", (160, 25, 75, 10), add_now=False
        )
        self.bool_add_frame.Label = _("Insert frame")
        self.bool_add_frame.TabIndex = 9
        self.bool_add_frame.HelpText = _(
            """        Adds a frame for the image and the text with the original prompt        """
        )

        if DEBUG:
            ctrl: UnoControlFixedHyperlinkModel = add_widget(
                page_ux, "FixedHyperlink", "lbl_debug", (175, 65, 50, 10), add_now=False
            )
            ctrl.Label = f"📜 {log_file}"
            ctrl.URL = uno.systemPathToFileUrl(log_file)
            ctrl.HelpText = (
                _(
                    "You are debugging, make sure opening LibreOffice from the command line. Consider using"
                )
                + f"\n\n   tailf { log_file }"
            )

        lbl: UnoControlFixedTextModel = add_widget(
            page_in, "FixedText", "label_prompt", (6, 8, 190, 50), add_now=False
        )
        lbl.Label = _("This is a horde client crafted with ") + "💗 @2025 - "

        lbl: UnoControlFixedHyperlinkModel = add_widget(
            page_in, "FixedHyperlink", "lbl_faq", (186, 6, 40, 10)
        )
        lbl.Label = "🤔 " + _("FAQ")
        lbl.URL = HELP_URL

        ctrl: UnoControlFixedHyperlinkModel = add_widget(
            page_in, "FixedHyperlink", "lbl_gallery", (183, 65, 50, 10), add_now=False
        )
        ctrl.Label = "🎨 " + _("Go to images")
        logger.info(
            "Download Image directory " + str(self.path_store_images_directory())
        )
        ctrl.URL = uno.systemPathToFileUrl(str(self.path_store_images_directory()))
        ctrl.HelpText = _("""       Click to browse your generated images        """)

        self.lbl_sysinfo: UnoControlFixedHyperlinkModel = add_widget(
            page_in, "FixedHyperlink", "lbl_sysinfo", (5, 65, 90, 10), add_now=False
        )
        self.lbl_sysinfo.Label = "🤖 " + _("System Information")
        self.lbl_sysinfo.HelpText = _(
            """Click to copy on your clipboard your system information to share when reporting something, you can paste it"""
        )

        return dc

    def show_ui(self):
        self.dlg.setVisible(True)

        # Post visibility setup
        self.book.ActiveTabPageID = 1
        for pair in self.insert_in:
            pair[0].insertByName(pair[1].Name, pair[1])

        self.lst_model = self.book.getTabPageByID(1).getControl(self.lst_model.Name)
        lst_rep_model = self.lst_model.getModel()
        for i in range(len(self.model_choices)):
            lst_rep_model.insertItemText(i, self.model_choices[i])
        self.lst_model.Text = self.default_model

        self.int_width = self.book.getTabPageByID(1).getControl(self.int_width.Name)
        self.int_width.addTextListener(self)
        self.int_width.addSpinListener(self)
        self.int_width.addFocusListener(self)
        self.int_height = self.book.getTabPageByID(1).getControl(self.int_height.Name)
        self.int_height.addTextListener(self)
        self.int_height.addSpinListener(self)
        self.int_height.addFocusListener(self)
        self.lbl_sysinfo = self.book.getTabPageByID(3).getControl(self.lbl_sysinfo.Name)
        self.lbl_sysinfo.addActionListener(self)
        self.txt_prompt = self.dlg.getControl("txt_prompt")
        self.txt_prompt.addTextListener(self)
        self.txt_prompt.addKeyListener(self)
        self.ctrl_token = self.dlg.getControl(self.ctrl_token.Name)

        # UI tweaks
        self.ctrl_token.getModel().EchoChar
        self.lbl_view_pass = self.dlg.getControl(self.lbl_view_pass.Name)
        if self.options.get("api_key", ANONYMOUS_KEY) == ANONYMOUS_KEY:
            self.ctrl_token.getModel().EchoChar = 0
            self.lbl_view_pass.setVisible(False)
        else:
            self.ctrl_token.getModel().EchoChar = self.PASSWORD_MASK
            self.lbl_view_pass.setVisible(True)

        self.btn_toggle.setVisible(False)

        size = self.dlg.getPosSize()
        self.DEFAULT_DLG_HEIGHT = size.Height
        self.displacement = self.dlg.getControl("label_progress").getPosSize().Y - 5

    def toggle_dialog(self):
        """
        Un/shrink the dialog to have the progressbar visible
        """
        size = self.dlg.getPosSize()
        lbl = self.dlg.getControl("label_progress")
        btn = self.dlg.getControl("btn_toggle")
        prg = self.dlg.getControl("prog_status")
        if size.Height == self.DEFAULT_DLG_HEIGHT:
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

    def validate_fields(self) -> None:
        if self.in_progress:
            return
        enable_ok = (
            len(self.txt_prompt.Text) > MIN_PROMPT_LENGTH
            and self.int_width.Value * self.int_height.Value <= MAX_MP
        )
        self.ok_btn.Enabled = enable_ok
        self.dlg.getControl("btn_trans").Enable = enable_ok

    def translate(self) -> None:
        self.continue_ticking = True

        def __emit_ticks__():
            i = 1.1
            logging.debug(1)
            while i < 15:
                time.sleep(0.5)
                if not self.continue_ticking:
                    return
                self.update_status(_("Translating"), i)
                logging.debug(i)
                i = i + 0.5

        if not self.show_language:
            return
        prompt_control = self.txt_prompt
        text = prompt_control.Text
        self.initial_prompt = text
        try:
            self.update_status(_("Translating"), 1.0)
            logger.info(f"Translating {text} from {self.local_language}")
            ticker = Thread(target=__emit_ticks__)
            ticker.start()
            logging.debug("starting")
            translated_text = opustm_hf_translate(text, self.local_language)
            logging.debug("Finished translating")
            self.continue_ticking = False
        except Exception as ex:
            logger.info(ex, stack_info=True)
        finally:
            prompt_control.Text = translated_text or text

    def focusLost(self, oFocusEvent: FocusEvent) -> None:
        element = oFocusEvent.Source.getModel()
        if element.Name not in ["int_width", "int_height"]:
            return

        pixels = self.int_width.Value * self.int_height.Value
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
        else:
            ctrl = oActionEvent.value.Source.getModel()
            if ctrl.Name == "lbl_view_pass":
                self.show_message(self.ctrl_token.Text, title=_("AIHorde API Key"))
            elif ctrl.Name == "lbl_sysinfo":
                self.export_system_information()

    def export_system_information(self):
        """
        Stores in the clipboard system information for issue reporting and showcase
        """
        system_info = "\n".join(
            [f"{key.capitalize()}: {val}" for key, val in self.key_debug_info.items()]
        )
        data = DataTransferable(system_info)
        self.clipboard.setContents(data, None)
        suffix = _("Above information is copied in your clipboard, paste it if needed")
        self.__msg_usr__(
            "\n\n" + system_info + "\n\n  " + suffix, title=_("System information")
        )

    def get_options_from_dialog(self) -> List[Dict[str, Any]]:
        """
        Updates the options from the dialog ready to be used
        """
        self.options.update(
            {
                "prompt": self.txt_prompt.Text,
                "image_width": self.int_width.Value,
                "image_height": self.int_height.Value,
                "model": self.lst_model.Text,
                "prompt_strength": self.int_strength.Value,
                "steps": self.int_steps.Value,
                "nsfw": self.bool_nsfw.State == 1,
                "censor_nsfw": self.bool_censure.State == 1,
                "api_key": self.ctrl_token.Text or ANONYMOUS_KEY,
                "max_wait_minutes": self.int_waiting.Value,
                "seed": self.txt_seed.Text,
            }
        )

        return self

    def start_processing(self) -> None:
        self.toggle_dialog()
        if self.inside in ["writer", "web"]:
            self.curview = self.model.CurrentController.ViewCursor
        self.ok_btn.Enabled = False
        self.btn_toggle.setVisible(True)
        cancel_button = self.dlg.getControl("btn_cancel")
        cancel_button.setLabel(_("Close"))
        cancel_button.getModel().HelpText = _(
            "The image generation will continue while you are doing other tasks"
        )
        if self.show_language and self.dlg.getControl("bool_trans").State == 1:
            self.translate()

        self.get_options_from_dialog()
        logger.debug(self.options)

        def __real_work_with_api__():
            images_paths = self.sh_client.generate_image(self.options)

            logger.debug(images_paths)
            if images_paths:
                self.update_status("", 100)
                self.show_message(
                    _("Your image was generated consuming {} kudos").format(
                        self.sh_client.kudos_cost
                    ),
                    title=_("AIHorde has good news"),
                )

                self.insert_image(
                    images_paths[0],
                    self.options["image_width"],
                    self.options["image_height"],
                    self.sh_client,
                    self.bool_add_to_gallery.State == 1,
                    self.bool_add_frame.State == 1,
                )
                settings_used = self.sh_client.get_settings()
                settings_used["translate"] = self.bool_trans.State
                settings_used["add_to_gallery"] = self.bool_add_to_gallery.State
                settings_used["add_text"] = self.bool_add_frame.State
                self.st_manager.save(settings_used)

            self.free()

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
        api_key = options.get("api_key", ANONYMOUS_KEY)
        self.ctrl_token.Text = api_key
        if api_key == ANONYMOUS_KEY:
            self.ctrl_token.Text = ""
            self.ctrl_token.TabIndex = 1
            lbl = create_widget(
                self.dlg.getModel(), "FixedHyperlink", "label_token", (110, 168, 48, 10)
            )
            lbl.Label = _("ApiKey (Optional)")
            lbl.URL = REGISTER_AI_HORDE_URL
        else:
            lbl = create_widget(
                self.dlg.getModel(), "FixedText", "label_token", (130, 168, 48, 10)
            )
            lbl.Label = _("ApiKey")

        self.model_choices = options.get("local_settings", {"models": MODELS}).get(
            "models", MODELS
        )
        self.model_choices = self.model_choices or MODELS
        self.default_model = options.get("model", DEFAULT_MODEL)

        self.txt_prompt.Text = self.selected or options.get("prompt", "")
        self.int_width.Value = options.get("image_width", DEFAULT_WIDTH)
        self.int_height.Value = options.get("image_height", DEFAULT_HEIGHT)
        self.int_strength.Value = options.get("prompt_strength", 6.3)
        self.int_steps.Value = options.get("steps", 25)
        self.int_waiting.Value = options.get("max_wait_minutes", 15)
        self.bool_nsfw.State = options.get("nsfw", 0)
        self.bool_censure.State = options.get("censor_nsfw", 1)
        self.bool_trans.State = options.get("translate", 1)
        self.bool_add_to_gallery.State = options.get("add_to_gallery", 1)
        self.bool_add_frame.State = options.get("add_text", 0)

    def free(self):
        self.dlg.dispose()

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

    def __msg_usr__(
        self, message, buttons=mbb.BUTTONS_OK, title="", url="", box_type=MESSAGEBOX
    ) -> int:
        """
        Shows a message dialog, if url is given, shows
        OK, Cancel, when the user presses OK, opens the URL in the
        browser.  Returns the status of the messagebox.
        """
        if url:
            buttons = mbb.BUTTONS_OK_CANCEL | buttons
            res = self.toolkit.createMessageBox(
                self.extoolkit, box_type, buttons, title, message
            ).execute()
            if res == mbr.OK:
                webbrowser.open(url, new=2)
            return res

        return self.toolkit.createMessageBox(
            self.extoolkit, box_type, buttons, title, message
        ).execute()

    def show_error(self, message, url="", title="", buttons=mbb.BUTTONS_OK):
        """
        Shows an error message dialog
        if url is given, shows OK, Cancel, when the user presses OK, opens the URL in the
        browser
        title is the title of the dialog to be shown
        buttons are the options that the user can have
        """
        if title == "":
            title = _("Watch out!")
        if not url and self.generated_url:
            url = self.generated_url
        self.__msg_usr__(
            message, buttons=buttons, title=title, url=url, box_type=WARNINGBOX
        )
        if self.options.get("api_key", ANONYMOUS_KEY) == ANONYMOUS_KEY:
            self.ctrl_token.setFocus()
        self.set_finished()

    def show_message(self, message, url="", title="", buttons=mbb.BUTTONS_OK):
        """
        Shows an informative message dialog
        if url is given, shows OK, Cancel, when the user presses OK, opens the URL in the
        browser
        title is the title of the dialog to be shown
        buttons are the options that the user can have
        """
        if title == "":
            title = _("Good")
        self.__msg_usr__(message, buttons=buttons, title=title, url=url)

    def insert_image(
        self,
        img_path: str,
        width: int,
        height: int,
        sh_client: AiHordeClient,
        add_to_gallery=True,
        add_frame=False,
    ):
        """
        Inserts the image with width*height from img_path in the current document adding
        accessibility data from sh_client, adding to the gallery.

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

        def __insert_image_in_text_doc__():
            """
            Inserts the image with width*height from the path in the document adding
            the accessibility data from sh_client in the current document
            at cursor position with the same text next to it.
            """

            def __insert_frame__(cursor, text, image):
                """Inserts the image and the text in the current position inside
                a frame, if it's not possible to place the image because it's inside
                another element that does not allow it, jumps to a start of page and
                inserts the frame with the image and the given text
                """
                text_frame = self.model.createInstance("com.sun.star.text.TextFrame")
                frame_size = Size()
                frame_size.Height = height + 150
                frame_size.Width = width + 150
                text_frame.setSize(frame_size)

                text_frame.setPropertyValue("AnchorType", AT_FRAME)
                try:
                    self.model.getText().insertTextContent(cursor, text_frame, False)
                except Exception:
                    # This happens if we are inside a frame.
                    cursor.jumpToStartOfPage()
                    self.model.getText().insertTextContent(cursor, text_frame, False)
                    logging.exception(
                        "Please try to not add the image inside other objects"
                    )

                frame_text = text_frame.getText()
                frame_cursor = frame_text.createTextCursor()
                text_frame.insertTextContent(frame_cursor, image, False)
                text_frame.insertString(frame_cursor, "\n" + text, False)

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

            if add_frame:
                __insert_frame__(self.curview, sh_client.get_title(), image)
            else:
                try:
                    self.model.Text.insertTextContent(self.curview, image, False)
                except Exception:
                    # This happens if we are inside a frame, or another element that
                    # does not allow to insert an image, then we jump and insert
                    logging.debug("Trying to insert the image without frame")
                    self.curview.jumpToStartOfPage()
                    self.model.getText().insertTextContent(self.curview, image, False)
                    logging.exception(
                        "Please try to not add the image inside other objects"
                    )

        image_insert_to = {
            "calc": __insert_image_as_draw__,
            "draw": __insert_image_as_draw__,
            "impress": __insert_image_as_draw__,
            "web": __insert_image_in_text_doc__,
            "writer": __insert_image_in_text_doc__,
        }
        image_insert_to[self.inside]()

        # Add image to gallery
        if add_to_gallery:
            self.add_image_to_gallery([img_path, sh_client.get_full_description()])
        else:
            # The downloaded image is removed, no gallery, no track of the image
            os.unlink.img_path

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

    def path_store_images_directory(self) -> str:
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

    def add_image_to_gallery(self, image_info: List[str]) -> None:
        """
        Adds the image in image_path to the gallery theme, the image
        is moved from image_path to the store.
        """

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

        images_dir = self.path_store_images_directory()
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
# Great!!! you are here looking at this, take a look at docs/CONTRIBUTING.md
# * [X] Store retrieved images in the gallery
# * [X] Show a message to open the directory that holds generated images with webbrowser
# * [X] Show a message to copy the system information
# * [X] Use TabPage for generate images, UX and information
# * [X] Add an option to store in the gallery
# * [X] Store consumed kudos
# * [X] Store and retrieve the UI options: translation, put text with the image, add to gallery
# * [X] Add an option to write the prompt with the image write_text
# --- 0.8
# * [ ] When something fails in the middle, it's possible to show an URL to allow to recover the generated image by hand
# * [ ] Add logo to page_in
# * [ ] Show kudos
# * [ ] Add tips to show. Localized messages. Inpainting, Gimp. Artbot https://artbot.site/
#    - We need a panel to show the tip
#    - Invite people to grow the knowledge
#    - They can be in github and refreshed each 10 days
#    - url, title, description, image, visibility
# * [ ] Wishlist to have right alignment for numeric control option
# * [ ] Recommend to use a shared key to users
# * [ ] Automate version propagation when publishing: Wishlist for extensions
# --- 1.0
# * [ ] Use styles support from Horde
#    -  Download and cache Styles
#
