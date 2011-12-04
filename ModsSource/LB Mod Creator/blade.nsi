if you are adding a campaign for ThirdSide, the user will need a duplicate of sidenc02.mix from the CD/DVD as sidenc03.mix for campaign progression to work. Might be worth adding a tick box for additional side campaigns and just have launch base duplicate that. I'd also like to make a highly specific request on behalf of my mod, I need the soviet and allied movies to be extracted automatically from the original CD's and renamed before being put back in to a new mix file in the users RA2 directory. This is specifically to support the merged allied and soviet campaigns so that all the movies are available. One way of implementing it might be to make a plugin style installer to carry out this process for mods that flag the plugin as required (i.e. the project) rather than having it built into the mod installer as standard.

Basically, the program needs to extract all the movies from both the allied and soviet CD's and then rename them all to end in x rather than e. The exception to this is the start movie a00_f00e.bik which needs to be left alone so it overrides the YR one. Instead its the YR one that needs to be extracted and renamed. Ideally it would build a new ecache mix file containing the movies too, but that will need the DCoder mix library. If it can't find the files it should abort and clean up, no point only having some of the movie files and not others. Perhaps return the error to the user saying it failed to find the files it needed and to insert the correct CD and try again and tell them that the mod requiring the plugin will not feature some videos.


step 1 is to get sidenc02.mix copied to sidenc03.mix
this will need dcoders dll to extract it (from ra2.mix I assume)

then we get the movies

ReadRegStr $R0 HKLM "SOFTWARE\Westwood\Red Alert 2" "InstallPath"
StrCmp $R0 "" CheckCD1 0
IfFileExists "$R0" 0 CheckCD1
StrCpy $R0 $R0 -9
IfFileExists "$R0\movies01.mix" 0 CheckCD1

CheckCD1:
MessageBox MB_OKCANCEL|MB_ICONINFORMATION "Please insert the Red Alert 2 Allied CD or The First Decade DVD." IDOK FindCD1 IDCANCEL GiveUp
FindCD1:
StrCpy $R0 "D"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "E"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "F"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "G"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "H"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "I"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "J"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "K"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "L"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "M"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "N"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "O"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "P"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "Q"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "R"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "S"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "T"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "U"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "V"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "W"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "X"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "Y"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
StrCpy $R0 "Z"
IfFileExists "$R0:\movies01.mix" FoundCD1 0
IfFileExists "$R0:\path to dvd folder\movies01.mix" FoundCD1 0
MessageBox MB_RETRYCANCEL|MB_ICONQUESTION "Could not detect a Red Alert 2 Allied CD or The First Decade DVD." IDRETRY FindCD1 IDCANCEL GiveUp

FoundCD1:
use dcoders dll to extract movies to $INSTDIR\video
rename the files to end in x rather than e (except for a00_f00e.bik)
now repeat for Soviet CD and Yuri CD
put all videos into an ecache mix file (dcoder dll again)

don't forget to clean up if this fails.