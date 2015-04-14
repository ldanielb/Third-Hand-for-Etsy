gui, font, s12, Verdana
Gui, Add, GroupBox, w320 h55 x955 y0,
Gui, Add, Text, w300 h40 x965 y13 vstatus, Status:Ready
gui, font, s8, Verdana
Gui, Add, Edit, w930 x5 y5 r1 vURL, https://etsy.com
Gui, Add, Button, x5 y30 w19 gBrB, <
Gui, Add, Button, x25 y30 w19 gBrF, >
Gui Add, Button, x45 y30 yp w44 Default, Go
Gui Add, Button, x90 y30 yp w100 , Favorite All
Gui Add, Button, x191 y30 yp w100 , Unfavorite All
Gui Add, Button, x292 y30 yp w100 , Stop
Gui Add, ActiveX, xm x5 y60 w1270 h680 vWB, Shell.Explorer
ComObjConnect(WB, WB_events)  ; Connect WB's events to the WB_events class object.
stop:=0
Gui, Show, ,Third Hand by Doldinsoft
; Continue on to load the initial page:

ButtonGo:
Gui Submit, NoHide
WB.Navigate(URL)
return

Brb:
try WB.GoBack()
return

BrF:
try WB.GoForward()
return

ButtonStop:
stop := 1
return

ButtonUnfavoriteAll:
Gui, Submit, NoHide
stop:=0
gui, font, s12, Verdana
GuiControl,, status, Status:Unfavoriteing every item on the page.
gui, font, s8, Verdana

try unfavedonex:=wb.document.getElementsByClassName("btn-fave").length
try unfavedonex1:=wb.document.getElementsByClassName("button-fave").length
loop
{
If stop = 1
{
	GuiControl,, status, Status:Ready
    Break
}
if (unfavedonex < 0)
{
    if (unfavedonex1 < 0)
		{
		   break
		}
}
unfavedonex:=unfavedonex-1
unfavedonex1:=unfavedonex1-1
if (unfavedonex >= 0)
{
	try unfavedonename:=wb.document.getElementsByClassName("btn-fave")[ unfavedonex ].className
	Sleep, 1500
	if (unfavedonename = "btn-fave done")
	{
      try favchange:=wb.document.getElementsByClassName("btn-fave")[ unfavedonex ].click()
	  Sleep, 1500
	}
}
if (unfavedonex1 >= 0)
{
	try unfavedonename1:=wb.document.getElementsByClassName("button-fave")[ unfavedonex1 ].className
	sleep, 1500
	if (unfavedonename1 = "button-fave favorited-button")
	{
      try favchange1:=wb.document.getElementsByClassName("button-fave")[ unfavedonex1 ].click()
	  Sleep, 1500
	}
}
}
If stop != 1
{
sleep, 3500
gui, font, s12, Verdana
GuiControl,, status, Status:Ready
gui, font, s8, Verdana
WB.Refresh()
Sleep, 6000
}
return

ButtonFavoriteAll:
stop:=0
loop
{
Gui, Submit, NoHide
gosub, ButtonUnfavoriteAll
gui, font, s12, Verdana
GuiControl,, status, Status:Favoriteing every item on the page.
gui, font, s8, Verdana

try favedonex:=wb.document.getElementsByClassName("btn-fave").length
try favedonex1:=wb.document.getElementsByClassName("button-fave").length

loop
{
If stop = 1
{
	GuiControl,, status, Status:Ready
    Break
}
if (favedonex < 0)
{
   if (favedonex1 < 0)
   {
	  break
   }
}
favedonex:=favedonex-1
favedonex1:=favedonex1-1
if (favedonex >= 0)
{
    try favedonename:=wb.document.getElementsByClassName("btn-fave")[ favedonex ].className
	Sleep, 1500
	if (favedonename = "btn-fave")
	{
      try favchange:=wb.document.getElementsByClassName("btn-fave")[ favedonex ].click()
	  Sleep, 1500
	}
}
if (favedonex1 >= 0)
{
    try favedonename1:=wb.document.getElementsByClassName("button-fave")[ favedonex1 ].className
	sleep, 1500
	if (favedonename1 = "button-fave unfavorited-button")
	{
      try favchange1:=wb.document.getElementsByClassName("button-fave")[ favedonex1 ].click()
	  Sleep, 1500
	}
}
}
If stop != 1
{
sleep, 3500
WB.Refresh()
sleep, 6000
while wb.ReadyState != 4
	sleep, 400
GuiControl,, status, Status:Checking to see if all hearts have been used.
try favedonex:=wb.document.getElementsByClassName("btn-fave").length
try favedonex1:=wb.document.getElementsByClassName("button-fave").length
}
loop
{
If stop = 1
{
	GuiControl,, status, Status:Ready
    Break
}
if (favedonex < 0)
{
   if (favedonex1 < 0)
   {
	  break
   }
}
favedonex:=favedonex-1
favedonex1:=favedonex1-1
if (favedonex >= 0)
{
    try favedonename:=wb.document.getElementsByClassName("btn-fave")[ favedonex ].className
	Sleep, 1500
	if (favedonename = "btn-fave")
	{
		Msgbox, You are out of hearts, try again in a little bit!
		gui, font, s12, Verdana
		GuiControl,, status, Status:Ready
		gui, font, s8, Verdana
		return
	}
}
if (favedonex1 >= 0)
{
    try favedonename1:=wb.document.getElementsByClassName("button-fave")[ favedonex1 ].className
	sleep, 1500
	if (favedonename1 = "button-fave unfavorited-button")
	{
		Msgbox, You are out of hearts, try again in a little bit!
		gui, font, s12, Verdana
		GuiControl,, status, Status:Ready
		gui, font, s8, Verdana
		return
	}
}
}
sleep, 4500	
If stop != 1
{
try itemnum:=wb.document.getElementsByClassName("listing-title").length
itemnumsav:=itemnum
wb1 := ComObjCreate("InternetExplorer.Application")
wb1.visible := false
wb1.navigate("http://www.etsy.com")
sleep, 2000
gui, font, s12, Verdana
GuiControl,, status, Status:Loading pages in a hidden window for 45 seconds.
gui, font, s8, Verdana
timer:=0
}
loop
{
if (itemnum <= 0)
{
 loop, 45
 {
 If stop = 1
  {
	GuiControl,, status, Status:Ready
	Process, Close, iexplore.exe
    Break
  }
 timer:=timer+1
 GuiControl,, status, Status:Loading pages in a hidden window for 45 seconds.(%timer%)
 sleep, 1500
 }
 Process, Close, iexplore.exe
 gui, font, s12, Verdana
 GuiControl,, status, Status:Ready
 gui, font, s8, Verdana
 break
}

if (itemnum > 0)
{
If stop = 1
  {
	GuiControl,, status, Status:Ready
	Process, Close, iexplore.exe
    Break
  }
try loadpage:=wb.document.getElementsByClassName("listing-title")[ itemnum-1 ].getElementsByTagName("a")[0].href
wb1.navigate(loadpage, 2048)
Sleep, 1500
itemnum:=(itemnum-1)
}
}


    If stop = 1
	{
	 GuiControl,, status, Status:Ready
     Break
	}
	If stop = 0
	{
	MsgBox, 4, , Do you want to continue? (Press YES or NO), 10
    try nextcheck:=wb.document.getElementsByClassName("next").length
	sleep, 2500
	IfMsgBox No
	{
       nextcheck := 0
	   sleep, 500
	}
	if nextcheck > 0
	{
	 try nextclick:=wb.document.getElementsByClassName("next")[0].click()
	 sleep, 5000
	}
	if nextcheck = 0
	{
	return
	}
	}

}
return

class WB_events
{
    NavigateComplete2(wb, NewURL)
    {
        GuiControl,, URL, %NewURL%  ; Update the URL edit control.
    }
}

GuiClose:
ExitApp