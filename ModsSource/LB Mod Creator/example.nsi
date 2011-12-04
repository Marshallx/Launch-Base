Function funcCustomScript

  CopyFiles "$LBDIR\md5.dll" $INSTDIR
  IntOp $UNINSTALLCOUNT $UNINSTALLCOUNT + 1
  WriteINIStr "$INSTDIR\launcher\liblist.gam" Uninstall $UNINSTALLCOUNT "md5.dll"

  CopyFiles "$LBDIR\comdlg32.ocx" $INSTDIR
  IntOp $UNINSTALLCOUNT $UNINSTALLCOUNT + 1
  WriteINIStr "$INSTDIR\launcher\liblist.gam" Uninstall $UNINSTALLCOUNT "comdlg32.ocx"

  CopyFiles "$LBDIR\mscomctl.ocx" $INSTDIR
  IntOp $UNINSTALLCOUNT $UNINSTALLCOUNT + 1
  WriteINIStr "$INSTDIR\launcher\liblist.gam" Uninstall $UNINSTALLCOUNT "mscomctl.ocx"

  CopyFiles "$LBDIR\flac.exe" $INSTDIR
  IntOp $UNINSTALLCOUNT $UNINSTALLCOUNT + 1
  WriteINIStr "$INSTDIR\launcher\liblist.gam" Uninstall $UNINSTALLCOUNT "flac.exe"

FunctionEnd