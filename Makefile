#!/usr/bin/python3
# -*- coding: utf-8 -*-

# To begin with translation process
# xgettext -o src/po/messages.pot --add-comments=TRANSLATORS: --keyword=_ --flag=_:1:pass-python-format --directory=. src/$(SCRIPTNAME).py

# To create a new language
# THELANG=da msginit --input=po/messages.pot --locale=$THELANG.UTF-8 --output=src/po/$THELANG.po

EXEC=loshd.oxt
UNOPKG=unopkg
CURRENT_LANG=es
SCRIPTNAME=StableHordeForLibreOffice
GETTEXTDOMAIN=stablehordeforlibreoffice

all: $(EXEC)

oxt/locale/%/LC_MESSAGES/$(GETTEXTDOMAIN).mo: src/po/%.po
	mkdir -p oxt/locale/$(CURRENT_LANG)/LC_MESSAGES && msgfmt --output-file=$@ $<

src/po/$(CURRENT_LANG).po: src/po/messages.pot
	msgmerge --update $@ $<

src/po/messages.pot: src/$(SCRIPTNAME).py
	xgettext -j -o $@ --add-comments=TRANSLATORS: --keyword=_ --flag=_:1:pass-python-format --directory=. $<

$(EXEC): src/$(SCRIPTNAME).py  oxt/locale/$(CURRENT_LANG)/LC_MESSAGES/$(GETTEXTDOMAIN).mo
	rm -f loshd.oxt
	oxt/build

clean:
	rm loshd.oxt

install:
	unopkg add -s -f loshd.oxt

run:
	make
	make install
	libreoffice --writer

publish:
	echo hola

.PHONY: install publish run

