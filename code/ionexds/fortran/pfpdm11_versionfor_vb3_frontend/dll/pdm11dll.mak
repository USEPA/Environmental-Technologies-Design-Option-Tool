FL=FL
FDLL= /c /FPi87 /AH /Gt /G2 /AH /Aw /Gw
FDL = /c /FPi87 /AH /Gt /G2 /AH /Aw

OBJS = pfpdm11.obj dgear.obj

all : $(OBJS)
    LINK @PDM11DLL.LNK

pfpdm11.obj : pfpdm11.for
    $(FL) $(FDLL) pfpdm11.for
dgear.obj : dgear.for
    $(FL) $(FDL) dgear.for

