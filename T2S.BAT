@echo off
REM
REM BATCH FILE FOR TIFF TO SDI TRANSLATION
REM
REM USAGE:
REM	For single file:      T2S <inputTiffFile> [XOrigin YOrigin]
REM		     or:      T2S <inputTiffFile> <outputSDIFile> [XOrigin YOrigin]
REM
REM	For multiple files:   T2S <wildcardspec> [XOrigin YOrigin]
REM
REM EXAMPLES:
REM
REM	T2S test.tif
REM	will translate test.tif -> test.sis
REM
REM	T2S test.sis testout.sio
REM	will translate test.tif -> testout.sio
REM
REM	T2S c:\*.tif
REM	will translate c:\*.tif -> c:\*.sis
REM
REM	T2S test.sis testout.sio 1000000.0 2000000.0
REM	will translate test.tif -> testout.sio
REM	and place the .sio image at origin of (1000000,2000000)
REM
for %%f in (%1) do tiff2sdi %%f %2 %3 %4
