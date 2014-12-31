AUTOMAKE_OPTIONS = foreign
CFLAGS=-Wall -I/usr/local/include -DVERSION=\"$(VERSION)\" -DPROG="\"$(PACKAGE)\"" 
LDFLAGS=-L/usr/local/lib -L/usr/lib -ljson -lfreexl -lstdc++
CPPFLAGS=-Wall -I/usr/local/include -I/usr/local/include/libjson -DNDEBUG
bin_PROGRAMS = xlstojson
xlstojson_SOURCES = xlstojson.cc
