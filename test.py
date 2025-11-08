import re
t="{\fArial|b0|i0|c0|p34;\C84;1097/@674\P\C256;IN:WNDFL,26\POUT:ZF316,41-44\C84;\P\C256;TERMINAL OUTPUT:xxxdBm\C84;\PSBOPTIMA(S)=1\PFUSBM1X4=1\PHO1(A)=1\PSBTEST=1\PCO48(TE)=2}"

t = re.sub("N:([A-Z0-9]+,\d+|[A-Z]+,[A-Z0-9]+)", f"N:asd,2", t)
t = re.sub("OUT:([A-Z0-9]+,[A-Z0-9\-]+)", f"OUT:zdcasdc".upper(), t)
print(t)