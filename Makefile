
src = $(wildcard *.c)
obj = $(src:.c=.o)

LDFLAGS = -static -lfreexl -lm -lxlsxwriter -lz

all: xls2xlsx

xls2xlsx: $(obj)
	$(CC) -o $@ $^ $(LDFLAGS)

.PHONY: clean
clean:
	rm -f $(obj) xls2xlsx


PREFIX = /usr/local


.PHONY: install
install: xls2xlsx
	mkdir -p $(DESTDIR)$(PREFIX)/bin
	cp $< $(DESTDIR)$(PREFIX)/bin/xls2xlsx


.PHONY: uninstall
uninstall:
	rm -f $(DESTDIR)$(PREFIX)/bin/xls2xlsx