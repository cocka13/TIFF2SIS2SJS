@echo off
REM
REM BATCH FILE FOR SDI JPEG COMPRESSION
REM
REM USAGE:
REM	For single file:      S2J <inputSDIFile>
REM		     or:      S2J <inputSDIFile> <outputSDIFile> [<quality>]
REM
REM	For multiple files:   S2J <wildcardspec> [<outputpath> <quality>]
REM
REM EXAMPLES:
REM
REM	S2J test.sis
REM	will translate test.sis -> test.sjs
REM	the default quality is 95
REM
REM	S2J test.sis testjpg.sis 60
REM	will translate test.sis -> testjpg.sis using a quality factor of 60
REM
REM	S2J c:\*.sis
REM	will translate c:\*.sis -> c:\*.sjs
REM
REM	S2J c:\*.sis d:\comp\ 50
REM	will translate c:\*.sis -> d:\comp\*.sjs using a quality factor of 50
REM
REM
for %%f in (%1) do sdi2jpeg %%f %2 %3 %4
