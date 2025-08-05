
EXEC=loshd.oxt
UNOPKG=unopkg

all: $(EXEC)

$(EXEC): src/StableHordeForLibreOffice.py
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

