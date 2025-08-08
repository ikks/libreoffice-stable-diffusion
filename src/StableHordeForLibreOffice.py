# StableHorde client for LibreOffice
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
from urllib.error import HTTPError, URLError
from urllib.request import urlopen, Request

DEBUG = True
VERSION = "0.4"

log_file = os.path.join(tempfile.gettempdir(), "libreoffice_shotd.log")
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

PROPERTY_CURRENT_SESSION = "stable_horde_checked_update"

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

######## Fix this
# https://help.libreoffice.org/latest/en-US/text/sbasic/shared/03/sf_l10n.html?DbPAR=BASIC
import gettext  # noqa: E402

gettext.bindtextdomain("StableHordeForLibreOffice", "/path/to/my/language/directory")
gettext.textdomain("StableHordeForLibreOffice")
_ = gettext.gettext
########

API_ROOT = "https://stablehorde.net/api/v2/"

REGISTER_STABLE_HORDE_URL = "https://aihorde.net/register"


class InformerFrontendInterface(metaclass=abc.ABCMeta):
    """
    Implementing this interface for an application frontend
    gives StableHordeClient a way to inform progress.  It's
    expected that StableHordeClient receives as parameter
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
    def get_frontend_property(self, property_name: str) -> str | bool | None:
        """
        Gets a property from the frontend application, used to retrieved stored
        information during this session.  Used when checking for update.
        """
        raise NotImplementedError

    @abc.abstractclassmethod
    def set_frontend_property(self, property_name: str, value: str | bool):
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

MIN_WIDTH = 384
MAX_WIDTH = 1024
MIN_HEIGHT = 384
MAX_HEIGHT = 1024
MIN_PROMPT_LENGTH = 10
"""
It's  needed that the user writes down something to create an image from
"""

MODELS = [
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
]
"""
Initial list of models, new ones are downloaded from StableHorde API
"""


class StableHordeClient:
    """
    Interaction with Stable Horde platform, currently supports:
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
        Creates a Stable Horde client with the settings, if None, the API_KEY is
        set to ANONYMOUS, the name to identify the client to Stable Horde and
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
        self.progress_text: str = _("Starting")
        self.warnings: list[json] = []

        # Sync informer and async request
        self.finished_task: bool = True
        dt = self.headers.copy()
        del dt["apikey"]
        # Beware, not logging the api_key
        show_debugging_data(dt)

    def __url_open__(
        self, url: str | Request, timeout: float = 10, refresh_each: float = 0.5
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
            with urlopen(url, timeout=timeout) as response:
                show_debugging_data("Data arrived")
                self.response_data = json.loads(response.read().decode("utf-8"))
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

    def __update_models_requirements__(self):
        """
        Download Model requirements and store them in the proper space
        thanks to informer.
        usually it's a value to be update, take the lowest possible value.
        Add range when min and max are present as prefix of an attribute
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
        # max_cfg_scale puede estar sola
        # [samplers]   -> can be single
        # [schedulers] -> can be single
        pass

    def __get_model_requirements__(self, model):
        """
        Given the name of a model, fetch the requirements if any
        to have the opportunity to mix the requirements for the
        model. Give for now the lowest possible value in case of
        range restrictions with min and max where appropiate.
        """
        pass

    def __get_model_restrictions__(self, model):
        """
        This assumes there is an already downloaded file, if not,
        there are no restrictions. The restrictions can be:
         * Fixed Value
         * Range

        This requires know how to work with listeners
        """

    def refresh_models(self):
        """
        Refreshes the model list with the 50 more used including always stable_diffusion
        we update self.settings to store the date when the models were stored.
        """
        previous_update = self.settings.get(
            "local_settings", {"date_refreshed_models": "2025-07-01"}
        ).get("date_refreshed_models", "2025-07-01")
        today = datetime.now().date()
        days_updated = (
            today - date(*[int(i) for i in previous_update.split("-")])
        ).days
        if days_updated < StableHordeClient.MAX_DAYS_MODEL_UPDATE:
            show_debugging_data(f"No need to update models {previous_update}")
            return

        show_debugging_data("time to update models")
        locals = self.settings.get("local_settings", {"models": MODELS})
        locals["date_refreshed_models"] = today.strftime("%Y-%m-%d")

        url = API_ROOT + "/stats/img/models?model_state=known"
        self.headers["X-Fields"] = "month"

        self.progress_text = _("Updating models")
        self.__inform_progress__()
        try:
            self.__url_open__(url)
            del self.headers["X-Fields"]
        except (HTTPError, URLError):
            message = _(
                "Tried to get the latest models, check your Internet connection"
            )
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
            ][: StableHordeClient.MAX_MODELS_LIST]
        else:
            popular_models = [
                (key, val)
                for key, val in popular_models
                if key.lower().count("inpaint") == 0
            ][: StableHordeClient.MAX_MODELS_LIST]

        fetched_models = [model[0] for model in popular_models]
        default_model = self.settings.get("default_model", DEFAULT_MODEL)
        if default_model not in fetched_models:
            fetched_models.append(default_model)
        if len(fetched_models) > 3:
            compare = set(fetched_models)
            new_models = compare.difference(locals["models"])
            if new_models:
                show_debugging_data(f"New models {len(new_models)}")
                locals["models"] = sorted(fetched_models, key=lambda c: c.upper())
                if len(new_models) == 1:
                    message = _("We have a new model:\n\n * ") + new_models[0]
                else:
                    message = _("We have new models:\n * ") + "\n * ".join(new_models)

                self.informer.show_message(message)

        self.settings["local_settings"] = locals

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
                "Tried to check for most recent version, check your Internet connection"
            )
        return message

    def get_balance(self) -> str:
        """
        Given an Stable Horde token, present in the attribute api_key,
        returns the balance for the account. If happens to be an
        anonymous account, invites to register
        """
        if self.api_key == ANONYMOUS:
            return _("Register at ") + REGISTER_STABLE_HORDE_URL
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
        Executes the flow to get an image from Stable Horde

        1. Invokes endpoint to launch a work for image generation
        2. Reviews the status of the work
        3. Waits until the max_wait_minutes for the generation of
        the image
        4. Retrieves the resulting images and returns the local path of
        the downloaded images

        When no success, returns [].  raises exceptions, but tries to
        offer helpful messages
        """
        self.settings.update(options)
        self.api_key = options["api_key"]
        self.headers["apikey"] = self.api_key
        self.check_counter = 1
        self.check_max = (
            options["max_wait_minutes"] * 60
        ) / StableHordeClient.CHECK_WAIT
        # Id assigned when requesting the generation of an image
        self.id = ""

        # Used for the progressbar.  We depend on the max time the user indicated
        self.max_time = datetime.now().timestamp() + options["max_wait_minutes"] * 60
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
                width = int(options["image_width"] / 64) * 64
            else:
                width = options["image_width"]

            if options["image_height"] % 64 != 0:
                height = int(options["image_height"] / 64) * 64
            else:
                height = options["image_height"]

            params.update({"width": int(width)})
            params.update({"height": int(height)})

            data_to_send.update({"models": [options["model"]]})

            mode = options.get("mode", "")
            if mode == "MODE_IMG2IMG":
                data_to_send.update({"source_image": options["source_image"]})
                data_to_send.update({"source_processing": "img2img"})
                params.update(
                    {"denoising_strength": (1 - float(options["init_strength"]))}
                )
                params.update({"n": options["nimages"]})
            elif mode == "MODE_INPAINTING":
                data_to_send.update({"source_image": options["source_image"]})
                data_to_send.update({"source_processing": "inpainting"})
                params.update({"n": options["nimages"]})

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
                                    f"Register at { REGISTER_STABLE_HORDE_URL } and use your key to improve your rate success. Detail:"
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
                if self.api_key == ANONYMOUS and REGISTER_STABLE_HORDE_URL in message:
                    self.informer.show_error(
                        f"{ message }", url=REGISTER_STABLE_HORDE_URL
                    )
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
            self.informer.show_error(_("Stablehorde said: ") + f"'{ message }'.")
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

        show_debugging_data(f"{progress} {self.progress_text}")

        if self.informer and progress != self.progress:
            self.informer.update_status(self.progress_text, progress)
            self.progress = progress

    def __check_if_ready__(self) -> bool:
        """
        Queries Stable horde API to check if the requested image has been generated,
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
            self.progress_text = _("Downloading generated image")
            self.__inform_progress__()
            return True

        if data["processing"] == 0:
            if data["queue_position"] == 0:
                text = _("You are the first in the queue")
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
                        _("Get an Api key for free at ")
                        + REGISTER_STABLE_HORDE_URL
                        + _(
                            ".\n This model takes more time than your current configuration."
                        )
                    )
                    raise IdentifiedError(message, url=REGISTER_STABLE_HORDE_URL)
                else:
                    message = (
                        _("Please try with other model,")
                        + f"{self.settings['model']} would take more time than you configured,"
                        + _(" or try again later.")
                    )
                    raise IdentifiedError(message)

            if data["is_possible"] is True:
                # We still have time to wait, given that the status is processing, we
                # wait between 5 secs and 15 secs to check again
                wait_time = min(
                    max(StableHordeClient.CHECK_WAIT, int(data["wait_time"] / 2)),
                    StableHordeClient.MAX_TIME_REFRESH,
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
                        "Currently no worker available to generate your image. Please try again later."
                    )
                )
        else:
            if self.api_key == ANONYMOUS:
                message = (
                    _("Get an Api key for free at ")
                    + REGISTER_STABLE_HORDE_URL
                    + _(
                        ".\n This model takes more time than your current configuration."
                    )
                )
                raise IdentifiedError(message, url=REGISTER_STABLE_HORDE_URL)
            else:
                minutes = (self.check_max * StableHordeClient.CHECK_WAIT) / 60
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
        At this stage Stable horde has generated the images and it's time
        to download them all.
        """
        url = f"{ API_ROOT }generate/status/{ self.id }"
        self.progress_text = _("fetching images")
        self.__inform_progress__()
        self.__url_open__(url)
        data = self.response_data
        show_debugging_data(data)

        return data["generations"]

    def __get_images_filenames__(self, images: list[json]) -> list[str]:
        """
        Downloads the generated images and returns the full path of the
        downloaded images.
        """
        show_debugging_data("Start to download generated images")
        generated_filenames = []
        cont = 1
        nimages = len(images)
        for image in images:
            with tempfile.NamedTemporaryFile(
                "wb+", delete=False, suffix=".webp"
            ) as generated_file:
                if image["img"].startswith("https"):
                    show_debugging_data(f"Downloading { image['img'] }")
                    if nimages == 1:
                        self.progress_text = _("Downloading result")
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
                "Maybe you need to change some parameters to generate succesfully an image. Horde said:\n * "
            ) + "\n * ".join([i["message"] for i in self.warnings])
            show_debugging_data(self.warnings)
            self.informer.show_error(message, title="warning")
            self.warnings = []
        self.refresh_models()
        return generated_filenames

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
            # "calc": "com.sun.star.sheet.SpreadsheetDocument",
            # "draw": "com.sun.star.drawing.DrawingDocument",
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
            "NewDialog", "StableHordeOptionsDialog", (47, 10, 265, 206)
        )
        dlg.Caption = _("Stable Horde for LibreOffice - ") + VERSION
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
        lbl.Caption = _("ApiKey (Optional)")

        # Buttons
        button_ok = dlg.CreateButton("btn_ok", (145, 182, 45, 13), push="OK")
        button_ok.Caption = _("Process")
        button_ok.TabIndex = 12
        button_cancel = dlg.CreateButton("btn_cancel", (78, 182, 49, 12), push="CANCEL")
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
        ) + _(" Write at least 5 words or 10 characters.")
        # ctrl.OnTextChanged = onaction
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
        ctrl.Caption = _("NSFW")
        ctrl.TabIndex = 7
        ctrl.TipText = _(
            "If not marked, it's faster, when marked you are on the edge..."
        )

        ctrl = dlg.CreateCheckBox("bool_censure", (29, 145, 30, 10))
        ctrl.Caption = _("Censor NSFW")
        ctrl.TipText = _("Allow if you want to avoid unexpected images...")
        ctrl.TabIndex = 8

        return dlg

    def prepare_options(self, options: json = None) -> json:
        dlg = self.dlg
        dlg.Controls("txt_prompt").Value = self.selected
        api_key = options.get("api_key", "")
        dlg.Controls("txt_token").Value = "" if api_key == ANONYMOUS else api_key
        dlg.Controls("lst_model").RowSource = options.get(
            "local_settings", {"models": MODELS}
        ).get("models")
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

        added_image.setPropertyValue("Title", _("Stable Horde Generated Image"))
        added_image.setPropertyValue(
            "Description", self.options["prompt"] + _(" by ") + self.options["model"]
        )
        added_image.Visible = True
        self.model.Modified = True
        os.unlink(img_path)

    def __insert_image_in_calc__(self, img_path: str, width: int, height: int):
        self.bas.MsgBox("TBD")

    def __insert_image_in_draw__(self, img_path: str, width: int, height: int):
        self.bas.MsgBox("TBD")

    def __insert_image_in_text__(self, img_path: str, width: int, height: int):
        show_debugging_data(f"Inserting {img_path} in writer")
        image = self.doc.createInstance("com.sun.star.text.GraphicObject")
        image.GraphicURL = uno.systemPathToFileUrl(img_path)
        image.AnchorType = AS_CHARACTER
        image.Width = width * 10
        image.Height = height * 10

        ctrllr = self.model.CurrentController
        curview = ctrllr.ViewCursor
        curview.String = self.options["prompt"] + _(" by ") + self.options["model"]
        self.doc.Text.insertTextContent(curview, image, False)
        os.unlink(img_path)

    def get_frontend_property(self, property_name: str) -> str | bool | None:
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
            return None
        return value

    def set_frontend_property(self, property_name: str, value: str | bool):
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
        self.settingsfile = "stablehordesettings.json"
        self.file = base_directory / self.settingsfile

    def load(self) -> json:
        if not os.path.exists(self.file):
            return {"api_key": ANONYMOUS}
        with open(self.file) as myfile:
            return json.loads(myfile.read())

    def save(self, settings: json):
        with open(self.file, "w") as myfile:
            myfile.write(json.dumps(settings))
        os.chmod(self.file, 0o600)


def show_debugging_data(information, additional=""):
    if not DEBUG:
        return

    dnow = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(information, Exception):
        ln = information.__traceback__.tb_lineno
        logging.error(f"[{ dnow }]{ln}: { information }")
        logging.error(
            "".join(
                traceback.format_exception(None, information, information.__traceback__)
            )
        )
    else:
        logging.debug(f"[{ dnow }] { information }")
    if additional:
        logging.debug(f"[{ dnow }]{additional}")


def create_image(desktop=None, context=None):
    """Creates an image from a prompt provided by the user, making use
    of Stable Horde"""

    lo_manager = LibreOfficeInteraction(desktop, context)
    st_manager = HordeClientSettings(lo_manager.path_store_directory())
    saved_options = st_manager.load()
    sh_client = StableHordeClient(saved_options, lo_manager.base_info, lo_manager)

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
    images_paths = sh_client.generate_image(options)

    show_debugging_data(images_paths)
    if images_paths:
        lo_manager.insert_image(
            images_paths[0], options["image_width"], options["image_height"]
        )
        st_manager.save(sh_client.get_settings())

    lo_manager.free()


class StableHordeForLibreOffice(unohelper.Base, XJobExecutor, XEventListener):
    """Service that creates images from text. The url to be invoked is:
    service:org.fectp.StableHordeForLibreOffice
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
    StableHordeForLibreOffice,
    "org.fectp.StableHordeForLibreOffice",
    ("com.sun.star.task.JobExecutor",),
)

# TODO:
# * [X] Integrate changes from gimp work
# * [ ] Add option on the Dialog to show debug
# * [X] Issue bugs for Impress with placeholdertext bug 167809
# * [ ] Repo for client and use it as a submodule
#    -  Check how to add another source file in gimp and lo
# * [ ] Internationalization
# * [ ] Wayland transparent png - Not being reproduced...
# * [ ] Wishlist to have right alignment for numeric control option
# * [ ] Recommend to use a shared key to users
# * [X] Make release in Github
# * [ ] Modify Makefile to upload to github with gh
# * [ ] Automate version propagation when publishing
# * [ ] Handle Warnings.  For each model the restrictions are present in
# https://raw.githubusercontent.com/Haidra-Org/AI-Horde-image-model-reference/refs/heads/main/stable_diffusion.json
# https://discord.com/channels/781145214752129095/1081743238194536458/1402045915510083724
# https://github.com/Haidra-Org/hordelib/blob/a0555b474696257a2374f4d1d4bc10b3d3fae5e3/hordelib/horde.py#L198
# * [ ] Add to Calc
# * [ ] Add to Draw
# * [X] Announce in stablehorde
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
# oxt/build && unopkg add -s -f loshd.oxt && libreoffice --writer >> /tmp/libreoffice.log
#
