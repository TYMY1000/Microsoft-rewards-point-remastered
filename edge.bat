@echo off
setlocal enabledelayedexpansion

rem List of 100 four-letter words (all unique)
set words[0]=code
set words[1]=data
set words[2]=byte
set words[3]=loop
set words[4]=hack
set words[5]=fire
set words[6]=wire
set words[7]=link
set words[8]=disk
set words[9]=task
set words[10]=mesh
set words[11]=port
set words[12]=host
set words[13]=node
set words[14]=echo
set words[15]=sync
set words[16]=ping
set words[17]=mode
set words[18]=send
set words[19]=back
set words[20]=read
set words[21]=boot
set words[22]=flow
set words[23]=load
set words[24]=test
set words[25]=open
set words[26]=save
set words[27]=text
set words[28]=exit
set words[29]=lock
set words[30]=edit
set words[31]=push
set words[32]=pull
set words[33]=drag
set words[34]=drop
set words[35]=zoom
set words[36]=view
set words[37]=clip
set words[38]=scan
set words[39]=bold
set words[40]=copy
set words[41]=dark
set words[42]=fast
set words[43]=deep
set words[44]=sort
set words[45]=trim
set words[46]=show
set words[47]=hide
set words[48]=hash
set words[49]=wrap
set words[50]=next
set words[51]=time
set words[52]=grid
set words[53]=form
set words[54]=dash
set words[55]=undo
set words[56]=redo
set words[57]=dual
set words[58]=path
set words[59]=rate
set words[60]=mute
set words[61]=note
set words[62]=jump
set words[63]=rule
set words[64]=file
set words[65]=true
set words[66]=call
set words[67]=root
set words[68]=swap
set words[69]=base
set words[70]=tags
set words[71]=home
set words[72]=safe
set words[73]=send
set words[74]=fire
set words[75]=text
set words[76]=zone
set words[77]=fast
set words[78]=port
set words[79]=disk
set words[80]=task
set words[81]=grid
set words[82]=view
set words[83]=file
set words[84]=bold
set words[85]=link
set words[86]=read
set words[87]=safe
set words[88]=mute
set words[89]=wrap
set words[90]=hash
set words[91]=note
set words[92]=clip
set words[93]=time
set words[94]=hide
set words[95]=swap
set words[96]=zoom
set words[97]=sort
set words[98]=redo
set words[99]=jump

rem Track searched words to avoid repetition
for /L %%i in (0,1,99) do set used[%%i]=0

rem Set counter to track number of searches
set count=0

:loop
rem Stop when 100 searches are completed
if %count% GEQ 100 (
    echo Completed 100 unique searches. Exiting...
    exit /b
)

:find_word
rem Generate a random index (0 to 99)
set /a index=%random% %% 100

rem Check if the word at this index has already been used
if !used[%index%]!==1 goto find_word

rem Mark word as used
set used[%index%]=1
set query=!words[%index%]!

rem Open Bing in a new tab using Edge (without switching focus)
start "" msedge.exe --profile-directory="Default" --new-tab "https://www.bing.com"

rem Wait for Bing to load
ping 127.0.0.1 -n 5 >nul

rem Focus the search box (without activating Edge)
nircmd sendkeypress ctrl+l
ping 127.0.0.1 -n 2 >nul

rem **Fix: Clear previous search by selecting all text and deleting it**
nircmd sendkeypress ctrl+a
ping 127.0.0.1 -n 2 >nul
nircmd sendkeypress backspace
ping 127.0.0.1 -n 2 >nul

rem Type the search query with a slight delay per character
for /l %%i in (0,1,31) do (
    set char=!query:~%%i,1!
    if not "!char!"=="" (
        nircmd sendkeypress !char!
        ping 127.0.0.1 -n 2 -w 300 >nul
    )
)

rem Ensure typing is complete before pressing Enter
ping 127.0.0.1 -n 2 >nul

rem Press Enter using PowerShell (without switching focus)
powershell -command "$wshell = New-Object -ComObject WScript.Shell; Start-Sleep -Milliseconds 300; $wshell.SendKeys('{ENTER}')"

rem Increment the search counter
set /a count+=1

rem Random delay before the next search (10-25 seconds)
set /a delay= 10
ping 127.0.0.1 -n %delay% >nul

rem Repeat loop
goto loop
