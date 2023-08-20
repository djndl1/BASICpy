##
# BASICpy
#

FBC=fbc
VB6=VB6.EXE

BUILD_DIR=build

COMMON_SRC=

FB_SRC=
FB_PROJ_SRC= $(COMMON_SRC) $(FB_SRC)

VB_PROJ_FILE=vb/BASICpy.vbp
VB_SRC=$(wildcard vb/*.cls)
VB_PROJ_SRC=$(COMMON_SRC) $(VB_SRC)

# end

.PHONY: all
all: fbproj vbproj

output_dir:
	mkdir -p $(BUILD_DIR)

fbproj: $(FB_PROJ_SRC)

vbproj: $(VB_PROJ_FILE) $(VB_PROJ_SRC) output_dir
	$(VB6) -make vb/BASICpy.vbp -OUTDIR $(BUILD_DIR)
