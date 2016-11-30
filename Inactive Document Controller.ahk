Gui, Show , w200 h90, Doc Controller
Gui, Add, Text, x10 y10 w160 Center, Choose Program to control
Gui, Add, DropDownList, vProgramGroup, Chrome PDF|Adobe Reader|PowerPoint
Gui, Add, Button, default, OK


return

bindKeys:

  Msgbox, % "Press OK and then the key you want to bind to Page Down"
  Input, LastKey, L1, {LControl}{RControl}{LAlt}{RAlt}{LShift}{RShift}{LWin}{RWin}{AppsKey}{F1}{F2}{F3}{F4}{F5}{F6}{F7}{F8}{F9}{F10}{F11}{F12}{Left}{Right}{Up}{Down}{Home}{End}{PgUp}{PgDn}{Del}{Ins}{BS}{Capslock}{Numlock}{PrintScreen}{Pause}
  if ErrorLevel
   prefix = EndKey`:
   StringReplace, LastKey, ErrorLevel, %prefix%
  Hotkey, %LastKey%,PP1,On

  Msgbox, % "Press OK and then the key you want to bind to Page Up"
  Input, LastKey, L1, {LControl}{RControl}{LAlt}{RAlt}{LShift}{RShift}{LWin}{RWin}{AppsKey}{F1}{F2}{F3}{F4}{F5}{F6}{F7}{F8}{F9}{F10}{F11}{F12}{Left}{Right}{Up}{Down}{Home}{End}{PgUp}{PgDn}{Del}{Ins}{BS}{Capslock}{Numlock}{PrintScreen}{Pause}
  if ErrorLevel
   prefix = EndKey`:
   StringReplace, LastKey, ErrorLevel, %prefix%
  Hotkey, %LastKey%,PP2,On

  MsgBox, Keys mapped.

return


TargetAdobe:
  SetTitleMatchMode, 2
  WinGet, Window, List, Adobe

  Loop %Window%

  {
    Id:=Window%A_Index%
    WinGetTitle, TVar , % "ahk_id " Id
    Window%A_Index%:=TVar ;use this if you want an array
    tList.=A_Index " " TVar "`n"  ;use this if you just want the list
  }

  InputBox, values, Enter number of window you want to target:, %tList%
  Msgbox, % "You selected: " Window%values%
return

TargetChrome:

  SetTitleMatchMode, 2
  WinGet, Window, List, Google Chrome

  Loop %Window%
  {
    Id:=Window%A_Index%
    WinGetTitle, TVar , % "ahk_id " Id
    Window%A_Index%:=TVar ;use this if you want an array
    tList.=A_Index " " TVar "`n"  ;use this if you just want the list
  }

  InputBox, values, Enter number of window you want to target:, %tList%
  Msgbox, % "You selected: " Window%values%

return


TargetPowerPoint:

  ppt:=ComObjActive("PowerPoint.Application")
  ppcol:=ppt.Presentations

  done = 0
  for Presentation, in ppcol
    if done = 0
    {
    MsgBox,4,Use this presentation?,% Presentation.name
      IfMsgBox Yes
      {
        target:=Presentation
        done = 1
      }
    }
return

buttonok:
  Gui, Submit

  GoSub, BindKeys
  if (ProgramGroup == "Chrome PDF")
    GoSub, TargetChrome
  else if (ProgramGroup == "Adobe Reader")
    GoSub, TargetAdobe
  else if (ProgramGroup == "PowerPoint")
    GoSub, TargetPowerPoint

return


PP1:
  if (ProgramGroup == "Chrome PDF")
  {
    title := % Window%values%
    SetTitleMatchMode, 3
    ControlSend, Chrome_RenderWidgetHostHWND2, {Left}, %title%
  }
  else if (ProgramGroup == "Adobe Reader")
  {
    title := % Window%values%
    SetTitleMatchMode, 3
    ControlSend, AVL_AVView25, {Up}, %title%
  }
  else if (ProgramGroup == "PowerPoint")
    target.SlideShowWindow.View.Next
return

PP2:
  if (ProgramGroup == "Chrome PDF")
  {
    title := % Window%values%
    SetTitleMatchMode, 3
    ControlSend, Chrome_RenderWidgetHostHWND2, {Right}, %title%
  }
  else if (ProgramGroup == "Adobe Reader")
  {
    title := % Window%values%
    SetTitleMatchMode, 3
    ControlSend, AVL_AVView25, {Down}, %title%
  }
  else if (ProgramGroup == "PowerPoint")
    target.SlideShowWindow.View.Previous
return
