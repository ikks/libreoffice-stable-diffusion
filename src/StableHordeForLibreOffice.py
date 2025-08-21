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

import abc
import asyncio
import base64
import contextvars
import gettext
import functools
import json
import locale
import logging
import os
import sys
import tempfile
import time
import traceback
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
from datetime import date
from datetime import datetime
from pathlib import Path
from scriptforge import CreateScriptService
from time import sleep
from typing import List
from typing import Union
from urllib.error import HTTPError, URLError
from urllib.request import urlopen, Request

DEBUG = True
VERSION = "0.5"
LIBREOFFICE_EXTENSION_ID = "org.fectp.StableHordeForLibreOffice"
GETTEXT_DOMAIN = "stablehordeforlibreoffice"

log_file = os.path.join(tempfile.gettempdir(), "libreoffice_shotd.log")
if DEBUG:
    print(f"your log is at {log_file}")
logging.basicConfig(
    filename=log_file,
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
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

URL_DOWNLOAD = "https://github.com/ikks/libreoffice-stable-diffusion/releases"
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


def show_debugging_data(information, additional="", important=False):
    if not DEBUG:
        return

    dnow = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(information, Exception):
        ln = information.__traceback__.tb_lineno
        logging.error(f"[{ dnow }]{ln}: { str(information) }")
        logging.error(
            "".join(
                traceback.format_exception(None, information, information.__traceback__)
            )
        )
    else:
        if important:
            logging.debug(f"[\033[1m{ dnow }\033[0m] { information }")
        else:
            logging.debug(f"[{ dnow }] { information }")
    if additional:
        logging.debug(f"[{ dnow }]{additional}")


# gettext usual alias for i18n
_ = gettext.gettext
gettext.textdomain(GETTEXT_DOMAIN)

API_ROOT = "https://aihorde.net/api/v2/"

REGISTER_AI_HORDE_URL = "https://aihorde.net/register"


class InformerFrontendInterface(metaclass=abc.ABCMeta):
    """
    Implementing this interface for an application frontend
    gives AIHordeClient a way to inform progress.  It's
    expected that AIHordeClient receives as parameter
    an instance of this Interface to be able to send messages
    and updates to the user.
    """

    @classmethod
    def __subclasshook__(cls, subclass):
        return (
            hasattr(subclass, "show_message")
            and callable(subclass.show_message)
            and hasattr(subclass, "show_error")
            and callable(subclass.show_error)
            and hasattr(subclass, "get_frontend_property")
            and callable(subclass.get_frontend_property)
            and hasattr(subclass, "set_frontend_property")
            and callable(subclass.set_frontend_property)
            and hasattr(subclass, "update_status")
            and callable(subclass.set_frontend_property)
            and hasattr(subclass, "set_finished")
            and callable(subclass.set_finished)
            and hasattr(subclass, "path_store_directory")
            and callable(subclass.path_store_directory)
            or NotImplemented
        )

    @abc.abstractclassmethod
    def show_message(
        self, message: str, url: str = "", title: str = "", buttons: int = 0
    ):
        """
        Shows an informative message dialog
        if url is given, shows OK, Cancel, when the user presses OK, opens the URL in the
        browser
        title is the title of the dialog to be shown
        buttons are the options that the user can have
        """
        raise NotImplementedError

    @abc.abstractclassmethod
    def show_error(self, message, url="", title="", buttons=0):
        """
        Shows an error message dialog
        if url is given, shows OK, Cancel, when the user presses OK, opens the URL in the
        browser
        title is the title of the dialog to be shown
        buttons are the options that the user can have
        """
        raise NotImplementedError

    @abc.abstractclassmethod
    def get_frontend_property(self, property_name: str) -> Union[str, bool, None]:
        """
        Gets a property from the frontend application, used to retrieved stored
        information during this session.  Used when checking for update.
        """
        raise NotImplementedError

    @abc.abstractclassmethod
    def set_frontend_property(self, property_name: str, value: Union[str, bool]):
        """
        Sets a property in the frontend application, used to retrieved stored
        information during this session.  Used when checking for update.
        """
        raise NotImplementedError

    @abc.abstractclassmethod
    def update_status(self, text: str, progress: float = 0.0):
        """
        Updates the status to the frontend and the progress from 0 to 100
        """
        raise NotImplementedError

    @abc.abstractclassmethod
    def set_finished(self):
        """
        Tells the frontend that the process has finished successfully
        """
        raise NotImplementedError

    @abc.abstractclassmethod
    def path_store_directory(self) -> str:
        """
        Returns the basepath for the directory offered by the frontend
        to store data for the plugin, cache and user settings
        """
        raise NotImplementedError


class IdentifiedError(Exception):
    """
    Exception raised for identified problems

    Attributes:
        message -- explanation of the error
        url -- Resource to understand and fix the problem
    """

    def __init__(self, message: str = "A custom error occurred", url: str = ""):
        self.message: str = message
        self.url: str = url
        super().__init__(self.message)

    def __str__(self):
        return self.message


ANONYMOUS = "0000000000"
"""
api_key for anonymous users
"""

DEFAULT_MODEL = "stable_diffusion"
"""
Model that is always present for image generation
"""

DEFAULT_INPAINT_MODEL = "stable_diffusion_inpainting"
"""
Model that is always present for inpainting
"""

MIN_WIDTH = 384
MAX_WIDTH = 1024
MIN_HEIGHT = 384
MAX_HEIGHT = 1024
MIN_PROMPT_LENGTH = 10
"""
It's  needed that the user writes down something to create an image from
"""

MODELS = [
    "Deliberate",
    "Dreamshaper",
    "NatViS",
    "noob_v_pencil XL",
    "Nova Anime XL",
    "Prefect Pony",
    "Realistic Vision",
    "stable_diffusion",
    "Ultraspice",
    "Unstable Diffusers XL",
    "WAI-ANI-NSFW-PONYXL",
]
"""
Initial list of models, new ones are downloaded from AIHorde API
"""

INPAINT_MODELS = [
    "A-Zovya RPG Inpainting",
    "Anything Diffusion Inpainting",
    "Epic Diffusion Inpainting",
    "iCoMix Inpainting",
    "Realistic Vision Inpainting",
    "stable_diffusion_inpainting",
]
"""
Initial list of inpainting models, new ones are downloaded from AIHorde API
"""


class AiHordeClient:
    """
    Interaction with AI Horde platform, currently supports:
    * Fetch the most used models in the month
    * Review the credits of an api_key
    * Request an image async and go all the way down until getting the image
    * Check if there is a newer version of the frontend client

    Attributes:
        settings -- configured in the constructor and later updated
    """

    # check model updates
    MAX_DAYS_MODEL_UPDATE = 5
    """
    We check at least this number of days for new models
    """

    MAX_MODELS_LIST = 50
    """
    Max Number of models to be presented to the user
    """

    CHECK_WAIT = 5
    """
    Number of seconds to wait before checking again if the image is generated
    """

    MAX_TIME_REFRESH = 15
    """
    If we are in a queue waiting, this is the max time in seconds before asking
    if we are still in queue
    """

    MODEL_REQUIREMENTS_URL = "https://raw.githubusercontent.com/Haidra-Org/AI-Horde-image-model-reference/refs/heads/main/stable_diffusion.json"

    def __init__(
        self,
        settings: json = None,
        platform: str = HORDE_CLIENT_NAME,
        informer: InformerFrontendInterface = None,
    ):
        """
        Creates a AI Horde client with the settings, if None, the API_KEY is
        set to ANONYMOUS, the name to identify the client to AI Horde and
        a reference of an obect that allows the client to send messages to the
        user.
        """
        if settings is None:
            self.settings = {"api_key": ANONYMOUS}
        else:
            self.settings: json = settings

        self.api_key: str = self.settings["api_key"]
        self.client_name: str = platform
        self.headers: json = {
            "Content-Type": "application/json",
            "Accept": "application/json",
            "apikey": self.api_key,
            "Client-Agent": self.client_name,
        }
        self.informer: InformerFrontendInterface = informer
        self.progress: float = 0.0
        self.progress_text: str = _("Starting...")
        self.warnings: List[json] = []

        # Sync informer and async request
        self.finished_task: bool = True
        self.censored: bool = False
        dt = self.headers.copy()
        del dt["apikey"]
        # Beware, not logging the api_key
        show_debugging_data(dt)

    def __url_open__(
        self, url: Union[str, Request], timeout: float = 10, refresh_each: float = 0.5
    ) -> None:
        """
        Opens a url request async with standard urllib, taking into account
        timeout informs `refresh_each` seconds.

        Requires Python 3.9

        Uses self.finished_task
        Invokes self.__inform_progress__()
        Stores the result in self.response_data
        """

        def real_url_open():
            show_debugging_data(f"starting request {url}")
            try:
                with urlopen(url, timeout=timeout) as response:
                    show_debugging_data("Data arrived")
                    self.response_data = json.loads(response.read().decode("utf-8"))
            except Exception as ex:
                show_debugging_data(ex)
                self.timeout = ex

            self.finished_task = True

        async def counter(until: int = 10) -> None:
            now = time.perf_counter()
            initial = now
            for i in range(0, until):
                if self.finished_task:
                    show_debugging_data(f"Request took {now - initial}")
                    break
                await asyncio.sleep(refresh_each)
                now = time.perf_counter()
                self.__inform_progress__()

        async def requester_with_counter() -> None:
            the_counter = asyncio.create_task(counter(int(timeout / refresh_each)))
            await asyncio.to_thread(real_url_open)
            await the_counter
            show_debugging_data("finished request")

        async def local_to_thread(func, /, *args, **kwargs):
            """
            python3.8 version do not have to_thread
            https://stackoverflow.com/a/69165563/107107
            """
            loop = asyncio.get_running_loop()
            ctx = contextvars.copy_context()
            func_call = functools.partial(ctx.run, func, *args, **kwargs)
            return await loop.run_in_executor(None, func_call)

        async def local_requester_with_counter():
            """
            Auxiliary function to add support for python3.8 missing
            asyncio.to_thread
            """
            task = asyncio.create_task(counter(30))
            await local_to_thread(real_url_open)
            self.finished_task = True
            await task

        self.finished_task = False
        running_python_version = [int(i) for i in sys.version.split()[0].split(".")]
        self.timeout = False
        if running_python_version >= [3, 9]:
            asyncio.run(requester_with_counter())
        elif running_python_version >= [3, 7]:
            ## python3.7 introduced create_task
            # https://docs.python.org/3/library/asyncio-task.html#asyncio.create_task
            asyncio.run(local_requester_with_counter())
        else:
            # Falling back to urllib, user experience will be uglier
            # when waiting...
            urlopen(url, timeout)
            self.finished_task = True

        if self.timeout:
            raise self.timeout

    def __update_models_requirements__(self) -> None:
        """
        Downloads model requirements.
        Usually it is a value to be updated, taking the lowest possible value.
        Add range when min and/or max are present as prefix of an attribute,
        the range is stored under the same name of the prefix attribute
        replaced.

        For example min_steps  and max_steps become range_steps
        max_cfg_scale becomes range_cfg_scale.

        Modifies self.settings["local_settings"]["requirements"]
        """
        # download json
        # filter the models that have requirements rules, store
        # the rules processed to be used later easily.
        # Store fixed value and range when possible
        # clip_skip
        # cfg_scale
        #
        # min_steps max_steps
        # min_cfg_scale max_cfg_scale
        # max_cfg_scale can be alone
        # [samplers]   -> can be single
        # [schedulers] -> can be single
        #

        if "local_settings" not in self.settings:
            return

        show_debugging_data("Getting requirements for models")
        url = self.MODEL_REQUIREMENTS_URL
        self.progress_text = _("Updating model requirements...")
        self.__url_open__(url)
        model_information = self.response_data
        req_info = {}

        for model, reqs in model_information.items():
            if "requirements" not in reqs:
                continue
            req_info[model] = {}
            # Model with requirement
            settings_range = {}
            for name, val in reqs["requirements"].items():
                # extract range where possible
                if name.startswith("max_"):
                    name_req = "range_" + name[4:]
                    if name_req in settings_range:
                        settings_range[name_req][1] = val
                    else:
                        settings_range[name_req] = [0, val]
                elif name.startswith("min_"):
                    name_req = "range_" + name[4:]
                    if name_req in settings_range:
                        settings_range[name_req][0] = val
                    else:
                        settings_range[name_req] = [val, val]
                else:
                    req_info[model][name] = val

            for name, range_vals in settings_range.items():
                if range_vals[0] == range_vals[1]:
                    req_info[model][name[6:]] = range_vals[0]
                else:
                    req_info[model][name] = range_vals

        show_debugging_data(f"We have requirements for {len(req_info)} models")

        if "requirements" not in self.settings["local_settings"]:
            show_debugging_data("Creating requirements in local_settings")
            self.settings["local_settings"]["requirements"] = req_info
        else:
            show_debugging_data("Updating requirements in local_settings")
            self.settings["local_settings"]["requirements"].update(req_info)

    def __get_model_requirements__(self, model: str) -> json:
        """
        Given the name of a model, fetch the requirements if any,
        to have the opportunity to mix the requirements for the
        model.

        Replaces values that must be fixed and if a value is out
        of range replaces by the min possible value of the range,
        if it was a list of possible values like schedulers, the
        key is replaced by scheduler_name and is enforced to have
        a valid value, if it resulted that was a wrong value,
        takes the first available option.

        Intended to set defaults for the model with the requirements
        present in self.MODEL_REQUIREMENTS_URL json

        The json return has keys with range_ or configuration requirements
        such as steps, cfg_scale, clip_skip, name of a sampler or a scheduler.
        """
        reqs = {}
        if not self.settings or "local_settings" not in self.settings:
            show_debugging_data("Too brand new... ")
            self.settings["local_settings"] = {}
        if "requirements" not in self.settings["local_settings"]:
            text_doing = self.progress_text
            self.__update_models_requirements__()
            self.progress_text = text_doing

        settings = self.settings["local_settings"]["requirements"].get(model, {})

        if not settings:
            show_debugging_data(f"No requirements for {model}")
            return reqs

        for key, val in settings.items():
            if key.startswith("range_") and (
                key[6:] not in settings
                or (settings[key[6:]] < val[0])
                or (val[1] < settings[key[6:]])
            ):
                reqs[key[6:]] = val[0]
            elif isinstance(val, list):
                key_name = key[:-1] + "_name"
                if key_name not in settings or settings[key_name] not in val:
                    reqs[key_name] = val[0]
            else:
                reqs[key] = val

        show_debugging_data(f"Requirements for { model } are { reqs }")
        return reqs

    def __get_model_restrictions__(self, model: str) -> json:
        """
        Returns a json that offers for each key a fixed value or
        a range for the requirements present in self.settings["local_settings"].
         * Fixed Value
         * Range

        Most commonly the result is an empty json.

        Intended for UI validation.

        Can offer range for initial min or max values, and also a
        list of strings or fixed values.
        """
        return self.settings.get("requirements", {model: {}}).get(model, {})

    def refresh_models(self):
        """
        Refreshes the model list with the 50 more used including always stable_diffusion
        we update self.settings to store the date when the models were refreshed.
        """
        default_models = MODELS
        self.staging = "Refresh models"
        previous_update = self.settings.get(
            "local_settings", {"date_refreshed_models": "2025-07-01"}
        ).get("date_refreshed_models", "2025-07-01")
        today = datetime.now().date()
        days_updated = (
            today - date(*[int(i) for i in previous_update.split("-")])
        ).days
        if days_updated < AiHordeClient.MAX_DAYS_MODEL_UPDATE:
            show_debugging_data(f"No need to update models {previous_update}")
            return

        show_debugging_data("time to update models")
        locals = self.settings.get("local_settings", {"models": MODELS})
        locals["date_refreshed_models"] = today.strftime("%Y-%m-%d")

        url = API_ROOT + "/stats/img/models?model_state=known"
        self.headers["X-Fields"] = "month"

        self.progress_text = _("Updating Models...")
        self.__inform_progress__()
        try:
            self.__url_open__(url)
            del self.headers["X-Fields"]
        except (HTTPError, URLError):
            message = _("Failed to get latest models, check your Internet connection")
            self.informer.show_error(message)
            return
        except TimeoutError:
            show_debugging_data("Failed updating models due to timeout")
            return

        # Select the most popular models
        popular_models = sorted(
            [(key, val) for key, val in self.response_data["month"].items()],
            key=lambda c: c[1],
            reverse=True,
        )
        show_debugging_data(f"Downloaded {len(popular_models)}")
        if self.settings.get("mode", "") == "MODE_INPAINTING":
            popular_models = [
                (key, val)
                for key, val in popular_models
                if key.lower().count("inpaint") > 0
            ][: AiHordeClient.MAX_MODELS_LIST]
            default_models = INPAINT_MODELS
        else:
            popular_models = [
                (key, val)
                for key, val in popular_models
                if key.lower().count("inpaint") == 0
            ][: AiHordeClient.MAX_MODELS_LIST]

        fetched_models = [model[0] for model in popular_models]
        default_model = self.settings.get("default_model", DEFAULT_MODEL)
        if default_model not in fetched_models:
            fetched_models.append(default_model)
        if len(fetched_models) > 3:
            compare = set(fetched_models)
            new_models = compare.difference(locals.get("models", default_models))
            if new_models:
                show_debugging_data(f"New models {len(new_models)}")
                locals["models"] = sorted(fetched_models, key=lambda c: c.upper())
                size_models = len(new_models)
                if size_models == 1:
                    message = _("We have a new model:\n\n * ") + next(iter(new_models))
                else:
                    if size_models > 10:
                        message = (
                            _("We have {} new models, including:").format(size_models)
                            + "\n * "
                            + "\n * ".join(list(new_models)[:10])
                        )
                    else:
                        message = (
                            _("We have {} new models:").format(size_models)
                            + "\n * "
                            + "\n * ".join(list(new_models)[:10])
                        )

                self.informer.show_message(message)

        self.settings["local_settings"] = locals

        self.__update_models_requirements__()

        if self.settings["model"] not in locals["models"]:
            self.settings["model"] = locals["models"][0]
        show_debugging_data(self.settings["local_settings"])

    def refresh_styles(self):
        """
        Refreshes the style list
        """
        # Fetch first 50 more used styles
        # We store the name of the styles and the date the last request was done
        # We fetch the style list if it haven't been updated during 5 days
        pass

    def check_update(self) -> str:
        """
        Inform the user regarding a plugin update. Returns "" if the
        installed is the latest one. Else the localized message,
        defaulting to english if there is no locale for the message.

        Uses PROPERTY_CURRENT_SESSION as the name of the property for
        checking only during this session.
        """
        message = ""
        current_local_session_key = PROPERTY_CURRENT_SESSION
        already_asked = self.informer.get_frontend_property(current_local_session_key)

        if already_asked:
            show_debugging_data(
                "We already checked for a new version during this session"
            )
            return ""
        show_debugging_data("Checking for update")

        try:
            # Check for updates by fetching version information from a URL
            url = URL_VERSION_UPDATE
            self.__url_open__(url, 15)
            data = self.response_data

            # During this session we will not check for update
            self.informer.set_frontend_property(current_local_session_key, True)
            local_version = (*(int(i) for i in str(VERSION).split(".")),)
            if isinstance(data["version"], int):
                # incoming_version has a deprecated format, local is newer
                return ""
            incoming_version = (*(int(i) for i in data["version"].split(".")),)

            if local_version < incoming_version:
                lang = locale.getlocale()[0][:2]
                message = data["message"].get(lang, data["message"]["en"])
        except (HTTPError, URLError):
            message = _(
                "Failed to check for most recent version, check your Internet connection"
            )
        return message

    def get_balance(self) -> str:
        """
        Given an AI Horde token, present in the attribute api_key,
        returns the balance for the account. If happens to be an
        anonymous account, invites to register
        """
        if self.api_key == ANONYMOUS:
            return _("Register at ") + REGISTER_AI_HORDE_URL
        url = API_ROOT + "find_user"
        request = Request(url, headers=self.headers)
        try:
            self.__url_open__(request, 15)
            data = self.response_data
            show_debugging_data(data)
        except HTTPError as ex:
            raise (ex)

        return f"\n\nYou have { data['kudos'] } kudos"

    def generate_image(self, options: json) -> [str]:
        """
        options have been prefilled for the selected model
        informer will be acknowledged on the process via show_progress
        Executes the flow to get an image from AI Horde

        1. Invokes endpoint to launch a work for image generation
        2. Reviews the status of the work
        3. Waits until the max_wait_minutes for the generation of
        the image
        4. Retrieves the resulting images and returns the local path of
        the downloaded images

        When no success, returns [].  raises exceptions, but tries to
        offer helpful messages
        """
        self.stage = "Nothing"
        self.settings.update(options)
        self.api_key = options["api_key"]
        self.headers["apikey"] = self.api_key
        self.check_counter = 1
        self.check_max = (options["max_wait_minutes"] * 60) / AiHordeClient.CHECK_WAIT
        # Id assigned when requesting the generation of an image
        self.id = ""

        # Used for the progressbar.  We depend on the max time the user indicated
        self.max_time = datetime.now().timestamp() + options["max_wait_minutes"] * 60
        self.factor = 5 / (
            3.0 * options["max_wait_minutes"]
        )  # Percentage and minutes 100*ellapsed/(max_wait*60)

        self.progress_text = _("Contacting the Horde...")
        try:
            params = {
                "cfg_scale": float(options["prompt_strength"]),
                "steps": int(options["steps"]),
                "seed": options["seed"],
            }

            restrictions = self.__get_model_requirements__(options["model"])
            params.update(restrictions)

            if options["image_width"] % 64 != 0:
                width = int(options["image_width"] / 64) * 64
            else:
                width = options["image_width"]

            if options["image_height"] % 64 != 0:
                height = int(options["image_height"] / 64) * 64
            else:
                height = options["image_height"]

            params.update({"width": int(width)})
            params.update({"height": int(height)})

            data_to_send = {
                "params": params,
                "prompt": options["prompt"],
                "nsfw": options["nsfw"],
                "censor_nsfw": options["censor_nsfw"],
                "r2": True,
            }

            data_to_send.update({"models": [options["model"]]})

            mode = options.get("mode", "")
            if mode == "MODE_IMG2IMG":
                data_to_send.update({"source_image": options["source_image"]})
                data_to_send.update({"source_processing": "img2img"})
                data_to_send["params"].update(
                    {"denoising_strength": (1 - float(options["init_strength"]))}
                )
                data_to_send["params"].update({"n": options["nimages"]})
            elif mode == "MODE_INPAINTING":
                data_to_send.update({"source_image": options["source_image"]})
                data_to_send.update({"source_processing": "inpainting"})
                data_to_send["params"].update({"n": options["nimages"]})

            dt = data_to_send.copy()
            if "source_image" in dt:
                del dt["source_image"]
                dt["source_image_size"] = len(data_to_send["source_image"])
            show_debugging_data(dt)

            data_to_send = json.dumps(data_to_send)
            post_data = data_to_send.encode("utf-8")

            url = f"{ API_ROOT }generate/async"

            request = Request(url, headers=self.headers, data=post_data)
            try:
                self.__inform_progress__()
                self.stage = "contacting"
                self.__url_open__(request, 15)
                data = self.response_data
                show_debugging_data(data)
                if "warnings" in data:
                    self.warnings = data["warnings"]
                text = _("Horde Contacted")
                show_debugging_data(text + f" {self.check_counter} { self.progress }")
                self.progress_text = text
                self.__inform_progress__()
                self.id = data["id"]
            except HTTPError as ex:
                try:
                    data = ex.read().decode("utf-8")
                    data = json.loads(data)
                    message = data.get("message", str(ex))
                    if data.get("rc", "") == "KudosUpfront":
                        if self.api_key == ANONYMOUS:
                            message = (
                                _(
                                    f"Register at { REGISTER_AI_HORDE_URL  } and use your key to improve your rate success. Detail:"
                                )
                                + f" { message }."
                            )
                        else:
                            message = (
                                f"{ HELP_URL } "
                                + _("to learn to earn kudos. Detail:")
                                + f" { message }."
                            )
                except Exception as ex2:
                    show_debugging_data(ex2, "No way to recover error msg")
                    message = str(ex)
                show_debugging_data(message, data)
                if self.api_key == ANONYMOUS and REGISTER_AI_HORDE_URL in message:
                    self.informer.show_error(f"{ message }", url=REGISTER_AI_HORDE_URL)
                else:
                    self.informer.show_error(f"{ message }")
                return ""
            except URLError as ex:
                show_debugging_data(ex, data)
                self.informer.show_error(
                    _("Internet required, chek your connection: ") + f"'{ ex }'."
                )
                return ""
            except Exception as ex:
                show_debugging_data(ex)
                self.informer.show_error(str(ex))
                return ""

            self.__check_if_ready__()
            images = self.__get_images__()
            images_names = self.__get_images_filenames__(images)

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
            self.informer.show_error(_("AIhorde response: ") + f"'{ message }'.")
            return ""
        except URLError as ex:
            show_debugging_data(ex, data)
            self.informer.show_error(_("Internet required, check your connection"))
            return ""
        except IdentifiedError as ex:
            if ex.url:
                self.informer.show_error(str(ex), url=ex.url)
            else:
                self.informer.show_error(str(ex))
            return ""
        except Exception as ex:
            show_debugging_data(ex)
            self.informer.show_error(_("Service failed with: ") + f"'{ ex }'.")
            return ""
        finally:
            self.informer.set_finished()
            message = self.check_update()
            if message:
                self.informer.show_message(message, url=URL_DOWNLOAD)

        return images_names

    def __inform_progress__(self):
        """
        Reports to informer the progress updating the attribute progress
        with the percentage elapsed time since the job started
        """
        progress = 100 - (int(self.max_time - datetime.now().timestamp()) * self.factor)

        show_debugging_data(f"{progress:.2f} {self.progress_text}")

        if self.informer and progress != self.progress:
            self.informer.update_status(self.progress_text, progress)
            self.progress = progress

    def __check_if_ready__(self) -> bool:
        """
        Queries AI horde API to check if the requested image has been generated,
        returns False if is not ready, otherwise True.
        When the time to get an image has been reached raises an Exception, also
        throws exceptions when there are network problems.

        Calls itself until max_time has been reached or the information from the API
        helps to conclude that the time will be longer than user configured.

        self.id holds the ID of the task that generates the image
        * Uses self.response_data
        * Uses self.check_counter
        * Uses self.max_time
        * Queries self.api_key
        """
        url = f"{ API_ROOT }generate/check/{ self.id }"

        self.__url_open__(url)
        data = self.response_data

        show_debugging_data(data)

        self.check_counter = self.check_counter + 1

        if data["done"]:
            self.progress_text = _("Downloading generated image...")
            self.__inform_progress__()
            return True

        if data["processing"] == 0:
            if data["queue_position"] == 0:
                text = _("You are first in the queue")
            else:
                text = _("Queue position: ") + str(data["queue_position"])
            show_debugging_data(f"Wait time {data['wait_time']}")
        elif data["processing"] > 0:
            text = _("Generating...")
            show_debugging_data(text + f" {self.check_counter} { self.progress }")
        self.progress_text = text

        if self.check_counter < self.check_max:
            if (
                data["processing"] == 0
                and data["wait_time"] + datetime.now().timestamp() > self.max_time
            ):
                # If we are in queue, we will not be served in time
                show_debugging_data(data)
                if self.api_key == ANONYMOUS:
                    message = (
                        _("Get a free API Key at ")
                        + REGISTER_AI_HORDE_URL
                        + _(
                            ".\n This model takes more time than your current configuration."
                        )
                    )
                    raise IdentifiedError(message, url=REGISTER_AI_HORDE_URL)
                else:
                    message = (
                        _("Please try another model,")
                        + f"{self.settings['model']} would take more time than you configured,"
                        + _(" or try again later.")
                    )
                    raise IdentifiedError(message)

            if data["is_possible"] is True:
                # We still have time to wait, given that the status is processing, we
                # wait between 5 secs and 15 secs to check again
                wait_time = min(
                    max(AiHordeClient.CHECK_WAIT, int(data["wait_time"] / 2)),
                    AiHordeClient.MAX_TIME_REFRESH,
                )
                for i in range(1, wait_time * 2):
                    sleep(0.5)
                    self.__inform_progress__()
                self.__check_if_ready__()
                return False
            else:
                show_debugging_data(data)
                raise IdentifiedError(
                    _(
                        "There are no workers available with these settings. Please try again later."
                    )
                )
        else:
            if self.api_key == ANONYMOUS:
                message = (
                    _("Get an Api key for free at ")
                    + REGISTER_AI_HORDE_URL
                    + _(
                        ".\n This model takes more time than your current configuration."
                    )
                )
                raise IdentifiedError(message, url=REGISTER_AI_HORDE_URL)
            else:
                minutes = (self.check_max * AiHordeClient.CHECK_WAIT) / 60
                show_debugging_data(data)
                if minutes == 1:
                    raise IdentifiedError(
                        _(f"Image generation timed out after { minutes } minute.")
                        + _("Please try again later.")
                    )
                else:
                    raise IdentifiedError(
                        _(f"Image generation timed out after { minutes } minutes.")
                        + _("Please try again later.")
                    )
        return False

    def __get_images__(self):
        """
        At this stage AI horde has generated the images and it's time
        to download them all.
        """
        self.stage = "Getting images"
        url = f"{ API_ROOT }generate/status/{ self.id }"
        self.progress_text = _("Fetching images...")
        self.__inform_progress__()
        self.__url_open__(url)
        data = self.response_data
        show_debugging_data(data)

        return data["generations"]

    def __get_images_filenames__(self, images: List[json]) -> List[str]:
        """
        Downloads the generated images and returns the full path of the
        downloaded images.
        """
        self.stage = "Downloading images"
        show_debugging_data("Start to download generated images")
        generated_filenames = []
        cont = 1
        nimages = len(images)
        for image in images:
            with tempfile.NamedTemporaryFile(
                "wb+", delete=False, suffix=".webp"
            ) as generated_file:
                if image["censored"]:
                    message = f'«{ self.settings["prompt"] }»' + _(
                        " is censored, try changing the prompt wording"
                    )
                    show_debugging_data(message)
                    self.informer.show_error(message, title="warning")
                    self.censored = True
                    break
                if image["img"].startswith("https"):
                    show_debugging_data(f"Downloading { image['img'] }")
                    if nimages == 1:
                        self.progress_text = _("Downloading result...")
                    else:
                        self.progress_text = _(
                            f"Downloading image { cont }/{ nimages }"
                        )
                    self.__inform_progress__()
                    with urlopen(image["img"]) as response:
                        bytes = response.read()
                else:
                    show_debugging_data(f"Storing embebed image { cont }")
                    bytes = base64.b64decode(image["img"])

                show_debugging_data(f"Dumping to { generated_file.name }")
                generated_file.write(bytes)
                generated_filenames.append(generated_file.name)
                cont += 1
        if self.warnings:
            message = _(
                "You may need to reduce your settings or choose another model, or you may have been censored. Horde message:\n * "
            ) + "\n * ".join([i["message"] for i in self.warnings])
            show_debugging_data(self.warnings)
            self.informer.show_error(message, title="warning")
            self.warnings = []
        self.refresh_models()
        return generated_filenames

    def get_imagename(self) -> str:
        """
        Returns a name for the image, intended to be used as identifier
        """
        if "prompt" not in self.settings:
            return "AIHorde will be invoked and this image will appear"
        return self.settings["prompt"] + " " + self.settings["model"]

    def get_title(self) -> str:
        """
        Intended to be used as the title to offer the user some information
        """
        if "prompt" not in self.settings:
            return "AIHorde will be invoked and this image will appear"
        return self.settings["prompt"] + _(" generated by ") + "AIHorde"

    def get_tooltip(self) -> str:
        """
        Intended for assistive technologies
        """
        if "prompt" not in self.settings:
            return "AIHorde will be invoked and this image will appear"
        return (
            self.settings["prompt"]
            + _(" with ")
            + self.settings["model"]
            + _(" generated by ")
            + "AIHorde"
        )

    def get_full_description(self) -> str:
        """
        Intended for reproducibility
        """
        if "prompt" not in self.settings:
            return "AIhorde shall be working sometime in the future"

        options = [
            "prompt",
            "model",
            "seed",
            "image_width",
            "image_height",
            "prompt_strength",
            "steps",
            "nsfw",
            "censor_nsfw",
        ]

        result = ["".join((op, " : ", str(self.settings[op]))) for op in options]

        return "\n".join(result)

    def get_settings(self) -> json:
        """
        Returns the stored settings
        """
        return self.settings

    def set_settings(self, settings: json):
        """
        Sets the settings, useful when fetching from a file or updating
        based on user selection.
        """
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
    show_debugging_data(response)
    my_popup.Dispose()


class LibreOfficeInteraction(InformerFrontendInterface):
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
        dlg.CreateGroupBox("framebox", (16, 11, 236, 163))
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
        button_ok = dlg.CreateButton("btn_ok", (78, 182, 45, 13), push="OK")
        button_ok.Caption = _("Process")
        button_ok.TabIndex = 12
        button_cancel = dlg.CreateButton(
            "btn_cancel", (145, 182, 49, 12), push="CANCEL"
        )
        button_cancel.Caption = _("Cancel")
        button_ok.TabIndex = 13
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
        ctrl.TabIndex = 4
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
        ctrl.TabIndex = 10
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
        ctrl.TabIndex = 2
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
        ctrl.TabIndex = 5
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
        ctrl.TabIndex = 3
        ctrl.TipText = _(
            "Height and Width together at most can be 2048x2048=4194304 pixels"
        )
        ctrl = dlg.CreateNumericField(
            "int_waiting",
            (200, 80, 48, 13),
            minvalue=1,
            maxvalue=5,
            spinbutton=True,
            accuracy=0,
        )
        ctrl.Value = 3
        ctrl.TabIndex = 9
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
        ctrl.TabIndex = 6
        ctrl.TipText = _("""
        How many sampling steps to perform for generation. Should
        generally be at least double the CFG unless using a second-order
        or higher sampler (anything with dpmpp is second order)
        """)
        ctrl = dlg.CreateCheckBox("bool_nsfw", (29, 130, 50, 10))
        ctrl.Caption = _("NSFW")
        ctrl.TabIndex = 7
        ctrl.TipText = _("""
        Whether or not your image is intended to be NSFW. May
        reduce generation speed (workers can choose if they wish
        to take nsfw requests)
        """)

        ctrl = dlg.CreateCheckBox("bool_censure", (29, 145, 50, 10))
        ctrl.Caption = _("Censor NSFW")
        ctrl.TipText = _("""
        Separate from the NSFW flag, should workers
        return nsfw images. Censorship is implemented to be safe
        and overcensor rather than risk returning unwanted NSFW.
        """)
        ctrl.TabIndex = 8

        return dlg

    def prepare_options(self, options: json = None) -> json:
        dlg = self.dlg
        dlg.Controls("txt_prompt").Value = self.selected
        api_key = options.get("api_key", "")
        dlg.Controls("txt_token").Value = "" if api_key == ANONYMOUS else api_key
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
        dlg.Controls("int_waiting").Value = options.get("max_wait_minutes", 3)
        dlg.Controls("txt_seed").Value = options.get("seed", "")
        rc = dlg.Execute()

        if rc != dlg.OKBUTTON:
            show_debugging_data("User scaped, nothing to do")
            dlg.Terminate()
            dlg.Dispose()
            return None

        show_debugging_data("good")
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
                "api_key": dlg.Controls("txt_token").Value or ANONYMOUS,
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
        show_debugging_data(f"Inserting {img_path} in writer")
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
            # The property was not present
            # Removed None in the type definition due to old python3.8 on Ubuntu 20.04
            # https://github.com/ikks/libreoffice-stable-diffusion/issues/1
            return None
        return value

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


class HordeClientSettings:
    """
    Store and load settings
    """

    def __init__(self, base_directory: Path = None):
        if base_directory is None:
            base_directory = tempfile.gettempdir()
        self.base_directory = base_directory
        self.settings_filename = "stablehordesettings.json"
        self.settings_file = base_directory / self.settings_filename

    def load(self) -> json:
        if not os.path.exists(self.settings_file):
            return {"api_key": ANONYMOUS}
        with open(self.settings_file) as myfile:
            return json.loads(myfile.read())

    def save(self, settings: json):
        with open(self.settings_file, "w") as myfile:
            myfile.write(json.dumps(settings))
        os.chmod(self.settings_file, 0o600)


def get_locale_dir(extid):
    show_debugging_data("here", important=True)

    ctx = uno.getComponentContext()
    pip = ctx.getByName(
        "/singletons/com.sun.star.deployment.PackageInformationProvider"
    )
    extpath = pip.getPackageLocation(extid)
    locdir = os.path.join(uno.fileUrlToSystemPath(extpath), "locale")
    show_debugging_data(f"Locales folder: {locdir}")
    return locdir


def create_image(desktop=None, context=None):
    """Creates an image from a prompt provided by the user, making use
    of AI Horde"""

    gettext.bindtextdomain(GETTEXT_DOMAIN, get_locale_dir(LIBREOFFICE_EXTENSION_ID))

    lo_manager = LibreOfficeInteraction(desktop, context)
    st_manager = HordeClientSettings(lo_manager.path_store_directory())
    saved_options = st_manager.load()
    sh_client = AiHordeClient(saved_options, lo_manager.base_info, lo_manager)

    show_debugging_data(lo_manager.base_info)

    options = lo_manager.prepare_options(saved_options)

    if options is None:
        # User cancelled, nothing to be done
        lo_manager.free()
        return
    elif len(options["prompt"]) < MIN_PROMPT_LENGTH:
        lo_manager.show_error(_("Please provide a prompt with at least 5 words"))
        lo_manager.free()
        return

    show_debugging_data(options)

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

        show_debugging_data(images_paths)
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
            create_image(self.desktop, self.context)
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
# * [X] Use thread to make more responsive the interface
# * [X] Check metadata response to avoid nsfw images
#   and showing the dialog with prefilled with the same
#   telling the user that it was NSFW
# * [X] Add information to the image https://discord.com/channels/781145214752129095/1401005281332433057/1406114848810467469
# * [X] Add to Impress and also to Sheets and Drawing
# * [ ] Use singleton path for the config path
# * [X] Add to Calc
# * [X] Add to Draw
# * [ ] Repo for client and use it as a submodule
#    -  Check how to add another source file in gimp and lo
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
# * [ ] Add option on the Dialog to show debug
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
