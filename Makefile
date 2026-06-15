FBC=fbc
LINK_EXE="C:/Program Files/Microsoft Visual Studio/VB98/LINK.EXE"
VB_EXE="C:/Program Files/Microsoft Visual Studio/VB98/VB6.EXE"

BUILD_DIR=build

COMMON_SRC=

FB_SRC=
FB_PROJ_SRC= $(COMMON_SRC) $(FB_SRC)

BUILD_DIR=build

# check the build system to see if it's a legacy OS
LEGACY_OS:=$(shell powershell.exe -NoProfile -Command '[int]((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").CurrentVersion -lt 6.2)')

BASICpy_VB6_SRC = $(wildcard *.bas) $(wildcard *.cls) $(wildcard *.frm)

.PHONY: all
all: fbproj vbproj

output_dir:
	mkdir -p $(BUILD_DIR)

fbproj: $(FB_PROJ_SRC)

build:
	mkdir $(BUILD_DIR)

build/BASICpy.dll: vb/BASICpy.vbp $(BASICpy_VB6_SRC) build
	$(VB_EXE) -MAKE vb/BASICpy.vbp -D LEGACY_OS=$(LEGACY_OS) -OUTDIR $(BUILD_DIR)

vbproj: build/BASICpy.dll

.PHONY: install
install: vbproj
	-cp build/BASICpy.dll "C:/Program Files/Common Files/"
	-regsvr32.exe -u -s "C:/Program Files/Common Files/BASICpy.dll"
	-regsvr32.exe -s "C:/Program Files/Common Files/BASICpy.dll"

.PHONY: tags
tags:
	ctags *.bas *.cls *.frm

.PHONY: clean
clean:
	-powershell.exe -NoProfile -Command \
                            "Remove-Item -ErrorAction SilentlyContinue -Force -Recurse $(BUILD_DIR); \
                            exit 0"