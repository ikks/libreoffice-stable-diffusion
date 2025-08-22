#!/usr/bin/python3
# -*- coding: utf-8 -*-

# To begin with translation process
# xgettext -o src/po/messages.pot --add-comments=TRANSLATORS: --keyword=_ --flag=_:1:pass-python-format --directory=. src/$(SCRIPTNAME).py

# To create a new language
# THELANG=da msginit --input=po/messages.pot --locale=$THELANG.UTF-8 --output=src/po/$THELANG.po

# When there are new strings to be translated run
#    make src/po/messages.pot
# Review the file to delete multiple entries of the same string pointing
# to different places in the file.
# Then run
#    make langs


EXEC=loshd.oxt
UNOPKG=unopkg
CURRENT_LANG=es
SCRIPTNAME=StableHordeForLibreOffice
GETTEXTDOMAIN=stablehordeforlibreoffice
PO_FILES := $(wildcard src/po/*.po)
MO_FILES := $(subst .po,/LC_MESSAGES/$(GETTEXTDOMAIN).mo, $(subst src/po,oxt/locale,$(PO_FILES)))

all: $(EXEC)

oxt/locale/%/LC_MESSAGES/$(GETTEXTDOMAIN).mo: src/po/%.po
	mkdir -p $(dir $@)
	msgfmt --output-file=$@ $<

src/po/$(CURRENT_LANG).po: src/po/messages.pot
	msgmerge --update $@ $<

src/po/messages.pot: src/$(SCRIPTNAME).py
	xgettext -j -o $@ --add-comments=TRANSLATORS: --keyword=_ --flag=_:1:pass-python-format --directory=. $<

$(EXEC): src/$(SCRIPTNAME).py 
	oxt/build

langs: src/po/messages.pot $(MO_FILES)
	
clean:
	rm  -rf oxt/locale loshd.oxt src/po/*~ src/po/*bak

install:
	unopkg add -s -f loshd.oxt

publish:
	sh scripts/publish

run:
	make
	make install
	libreoffice --writer


.PHONY: clean install langs publish run 



