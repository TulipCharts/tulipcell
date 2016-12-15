CC = gcc

IND_PATH = ../tulipindicators/
EXTRAFLAGS ?=
CCFLAGS = -Wall -O2 -lm $(EXTRAFLAGS)

all: tulipcell.dll

$(IND_PATH)libindicators.a:
	make -C $(IND_PATH) CCFLAGS="$(CCFLAGS)" CC=gcc

tulipcell.dll: tulipcell.c tulipcell.def $(IND_PATH)libindicators.a
	$(CC) $(CCFLAGS) -o $@ $^ -shared -I$(IND_PATH)
	strip $@

clean:
	del *.dll
	del -f *.o
