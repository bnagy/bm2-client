REM Author: Ben Nagy
REM Copyright: Copyright (c) Ben Nagy, 2006-2010.
REM License: The MIT License
REM (See README.TXT or http://www.opensource.org/licenses/mit-license.php for details.)
=====================
rmdir /s /q c:\fuzzbot_code
REM This ping is because the network goes offline for a while
REM with kvm-qemu after power cycling a bunch of fuzzbots
ping -n 90 192.168.122.1
xcopy /d /y /s \\192.168.122.1\ramdisk\fuzzbot_code c:\fuzzbot_code\
copy /y c:\fuzzbot_code\bm2-client\startfuzz.bat c:\AUTOEXEC.BAT
c:\compname /c BUGMINER-?8
rmdir /s /q r:\fuzzclient
mkdir r:\fuzzclient
REM CHANGE THIS TO THE CORRECT PATH
"C:\Program Files\Debugging Tools for Windows (x86)\gflags" /p /enable WINWORD.EXE /full
cd c:\fuzzbot_code\bm2-client
start cmd /k ruby word_fuzzclient.rb
