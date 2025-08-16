# Changelog

All notable changes to this project will be documented in this file.

This project is [semver](https://semver.org/) based

## [0.4.2] - Allow to work while image is being generated

### Added

* You can move around and continue working on other tasks while the
plugin is generating the image. If you have a Python installation
previous to 3.11, it's not guaranteed and patience will be needed.

### Changed

* Bug when stored settings had bad models data
* Process and Cancel buttons were swapped
* Initial models are chosen from the most popular ones

## [0.4.1] - Hotfix

### Changed

* Bug when getting one new model when updating

## [0.4] - Add support for Spanish

### Added

* Translation to spanish
* Support to get more languages used
* Instructions for translators

### Changed

* Improved proressbar ticks
* Fallback on older python versions for progressbar
* Warnings are fixed

## [0.3.1] - Bugfix that prevented using Windows, Mac and Older Python Version


## [0.3] - Converted macro to extension

### Added

* Menu Entry
* Toolbar Entry
* Shortcut Ctrl+Shift+H

### Changed

* Use of logging

## [0.2] - Model List is updated from Stable Horde API

### Added
* Selection Model list is updated from the API and user is
   informed when new ones are installed.
* User is invited to create an api_key with option to visit with
   the browser
* Warnings about model usage are seen by user

### Changed
* The settings file is readable only by the user, no others,
   except admins
* Bugfixes

## [0.1] - First Release of LibreOffice Macro

### Added

* Support for LibreOffice Writer
