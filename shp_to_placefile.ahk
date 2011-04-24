;~     shp_to_placefile - Convert an Esri Shapefile into a GRLevelX Placefile.
;~     Copyright (C) 2011 Bryan Perry <ih57452[AT]gmail.com>

;~     This program is free software: you can redistribute it and/or modify
;~     it under the terms of the GNU General Public License as published by
;~     the Free Software Foundation, either version 3 of the License, or
;~     (at your option) any later version.

;~     This program is distributed in the hope that it will be useful,
;~     but WITHOUT ANY WARRANTY; without even the implied warranty of
;~     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
;~     GNU General Public License for more details.

;~     You should have received a copy of the GNU General Public License
;~     along with this program.  If not, see <http://www.gnu.org/licenses/>.

#NoEnv
#NoTrayIcon
SendMode Input
SetWorkingDir %A_ScriptDir%
SetBatchLines, -1
SetFormat, FloatFast, 0.15
DetectHiddenWindows, On
#Include .\lib\Dlg.ahk

load_options("shp_to_placefile.ini", "placefile_name:shapefile.txt|threshold:999|poly_mode:0|poly_color:0x00ff00|poly_alpha:255|line_color:0x0000ff|line_width:1|limit_n:0|limit_s:0|limit_e:0|limit_w:0|hide_shp2text:1")

Gui, Add, GroupBox, x12 y12 w160 h80, Convert with shp2text
Gui, Add, DropDownList, x22 y32 w140 h20 r10 vshp_file
Gui, Add, Button, x22 y62 w90 h20, Convert
Gui, Add, GroupBox, x12 y102 w160 h140, Settings
Gui, Add, DropDownList, x22 y122 w140 h21 r10 gload_xls vxls_file
Gui, Add, Text, x22 y152 w50 h20, Placefile:
Gui, Add, Edit, x72 y149 w90 h20 vplacefile_name, %placefile_name%
Gui, Add, Text, x22 y175 w80 h20, View Threshold:
Gui, Add, Edit, x102 y172 w40 h20 Number vthreshold
Gui, Add, UpDown, Range1-999, %threshold%
Gui, Add, Radio, x22 y212 w90 h20 Checked vpoly_mode, Draw Polygons
Gui, Add, Radio, % "x22 y192 w80 h20 vline_mode Checked" . !poly_mode, Draw Lines
Gui, Add, GroupBox, x182 y12 w200 h50, Polygon Options
Gui, Add, Button, x192 y32 w40 h20 gpoly_color, Color
Gui, Add, Progress, x233 y32 w15 h20 C%poly_color% vpoly_color_preview, 100
Gui, Add, Text, x252 y32 w30 h20, Alpha:
Gui, Add, Slider, x282 y32 w90 h20 Range0-255 TickInterval64 ToolTip vpoly_alpha, %poly_alpha%
Gui, Add, GroupBox, x182 y72 w200 h170, Line Options
Gui, Add, Button, x192 y92 w40 h20 gline_color, Color
Gui, Add, Progress, x233 y92 w15 h20 C%line_color% vline_color_preview, 100
Gui, Add, Text, x252 y92 w40 h20, Width:
Gui, Add, Edit, x292 y92 w40 h20 Number vline_width
Gui, Add, UpDown, Range1-50, %line_width%
Gui, Add, Text, x192 y122 w60 h20, Hover Text:
Gui, Font,, Courier New
Gui, Add, Edit, x192 y142 w180 h90 -Wrap vhover_text
Gui, Font
Gui, Add, GroupBox, x12 y252 w370 h150, Filter Options
Gui, Add, ListView, x22 y272 w350 h100 Grid -Multi glist_click, #|Field|Search Type|Search String
Gui, Add, Button, x72 y375 w100 h20, Edit Filter
Gui, Add, Button, x222 y375 w100 h20, Clear Filter
Gui, Add, GroupBox, x12 y412 w370 h50, Limits
Gui, Add, Text, x22 y435 w50 h20, North:
Gui, Add, Edit, x52 y432 w50 h20 vlimit_n, %limit_n%
Gui, Add, Text, x112 y435 w50 h20, South:
Gui, Add, Edit, x145 y432 w50 h20 vlimit_s, %limit_s%
Gui, Add, Text, x202 y435 w50 h20, East:
Gui, Add, Edit, x232 y432 w50 h20 vlimit_e, %limit_e%
Gui, Add, Text, x292 y435 w50 h20, West:
Gui, Add, Edit, x322 y432 w50 h20 vlimit_w, %limit_w%
Gui, Add, Button, x22 y472 w170 h30, Add To Placefile
Gui, Add, Button, x202 y472 w80 h30, Help
Gui, Add, Button, x292 y472 w80 h30 gGuiClose, Exit
; Generated using SmartGUI Creator 4.0

LV_ModifyCol(1, "Integer")
LV_ModifyCol(3, "Desc")
LV_ModifyCol(4, "Desc")
file_list("shp")
file_list("xls")
Gui, +LastFound
hwnd := WinExist()
WinSet, Trans, 0
Gui, Show,, Shapefile to Placefile
SetTimer, fade_in, 30
Return

ButtonHelp:
Run, https://github.com/ih57452/shp_to_placefile/wiki
Return

GuiClose:
GuiEscape:
SetTimer, fade_in, Off
SetTimer, loading_bar, Off
If shp2text_pid
	WinClose, ahk_pid %shp2text_pid%
SetTimer, fade_out, 30
Return

ButtonConvert:
Gui, +OwnDialogs
IfNotExist, shp2text.exe
{
	MsgBox, 20, Error, This requires shp2text.exe. Would you like to visit the site to download it?
	IfMsgBox, Yes
		Run, http://www.obviously.com/gis/shp2text/
	Return
}
Gosub, save
shp_file := SubStr(shp_file, 1, -4)
IfExist, %shp_file%.xls
{
	MsgBox, 36, File Exists, This file has already been converted. Do you want to convert it again?
	IfMsgBox, No
		Goto, converted
	IfMsgBox, Yes
		FileDelete, %shp_file%.xls
}
Gosub, loading
RunWait, % comspec . (hide_shp2text ? " /c " : " /k ") . "shp2text.exe --spreadsheet " . shp_file . ".shp > " . shp_file . ".xls",, % (hide_shp2text ? "Hide" : ""), shp2text_pid
shp2text_pid =
Gosub, loaded
converted:
file_list("xls", shp_file . ".xls")
Return

load_xls:
Gosub, save
FileReadLine, lv, %xls_file%, 1
StringSplit, lv, lv, %A_Tab%, %A_Space%
LV_Delete()
Loop, %lv0%
{
	If (A_Index > 1) And (A_Index < 6)
		Continue
	LV_Add("", A_Index, lv%A_Index%)
}
Loop, 4
	LV_ModifyCol(A_Index, "AutoHdr")
Return

list_click:
If (A_GuiEvent != "DoubleClick")
	Return
ButtonEditFilter:
lv_num := LV_GetNext(0, "Focused")
LV_GetText(lv_text, lv_num, 2)
LV_GetText(lv_type, lv_num, 3)
LV_GetText(lv_string, lv_num, 4)
Gui, +Disabled
Gui, 2:Default
Gui, +Owner1 -MinimizeBox
Gui, Add, Text,, %lv_text%
Gui, Add, DropDownList, ys w100 r11 vlv_type, Equals||NotEqual|Contains|NotContain|StartsWith|EndsWith|GreaterThan|LessThan|GreaterOrEqual|LessOrEqual|RegExp
Gui, Add, Edit, ys w150 vlv_string, %lv_string%
Gui, Add, Button, ys w50 Default, OK
Gui, Add, Button, ys w50, Cancel
If lv_type
	GuiControl, ChooseString, lv_type, %lv_type%
Gui, Show,, Edit Filter
Return

2ButtonOK:
Gui, 1:-Disabled
Gui, Submit
Gui, Destroy
Gui, 1:Default
If (lv_type = "RegExp")
{
	RegExMatch("", lv_string)
	If ErrorLevel
	{
		MsgBox, 16, RegExp Error, There was an error with your Regular Expression search string:`n%lv_string%`n%ErrorLevel%
		lv_string =
		lv_type =
	}
}
LV_Modify(lv_num, "Col3", lv_type, lv_string)
Loop, 2
	LV_ModifyCol(A_Index + 2, "AutoHdr")
Return

2ButtonCancel:
2GuiClose:
2GuiEscape:
loaded:
SetTimer, loading_bar, Off
Gui, 1:-Disabled
Gui, Destroy
Gui, 1:Default
Return

ButtonClearFilter:
LV_Modify(LV_GetNext(0, "Focused"), "Col3", "", "")
Loop, 2
	LV_ModifyCol(A_Index + 2, "AutoHdr")
Return

poly_color:
line_color:
Dlg_Color(%A_ThisLabel%)
GuiControl, % "+C" . %A_ThisLabel%, %A_ThisLabel%_preview
Return

ButtonAddToPlacefile:
Gosub, save
Gosub, loading
Loop, Read, %xls_file%, %placefile_name%
{
	If (A_Index = 1)
	{
		id = -1
		points =
		draw =
		ht =
		If (limited := limit_n Or limit_s Or limit_e Or limit_w)
		{
			limit_n ? : limit_n := 90
			limit_s ? : limit_s := -90
			limit_e ? : limit_e := 180
			limit_w ? : limit_w := -180
		}
		FileAppend, Threshold: %threshold%`n
		If (!poly_mode)
			FileAppend, % "Color: " . rgb(line_color) . "`n"
		Continue
	}
	StringSplit, field, A_LoopReadLine, %A_Tab%, %A_Space%
	If (id != field1)
	{
		If points
		{
			points := SubStr(points, 2)
			If limited
			{
				Loop, Parse, points, |
				{
					StringSplit, point, A_LoopField, `,, %A_Space%
					If (draw := (point2 <= limit_e And point2 >= limit_w And point1 >= limit_s And point1 <= limit_n))
						Break
				}
			}
			If (!limited Or draw)
			{
				Loop, Parse, points, |
				{
					If (A_Index = 1)
					{
						If poly_mode
							FileAppend, % "Polygon:`n" . A_LoopField . ", " . rgb(poly_color, 1) . ", " . poly_alpha . "`n"
						Else
							FileAppend, % "Line: " . line_width . ", 0, """ . ht . """`n" . A_LoopField . "`n"
						Continue
					}
					FileAppend, %A_LoopField%`n
				}
				FileAppend, End:`n
			}
			points =
			ht =
		}
		id := field1
		Loop, %field0%
		{
			If (!(match := search(search_type%A_Index%, search_string%A_Index%, field%A_Index%)))
				Break
		}
	}
	If match
	{
		points .= "|" . field3 . ", " . field2
		If (!poly_mode And !ht)
			ht := hover_text(hover_text)
	}
}
Gosub, loaded
Return

loading:
Gui, +Disabled
Gui, 3:Default
Gui, +Owner1 -MinimizeBox +LabelGui
Gui, Add, Progress, % "w250 h30 +0x8 vloading_bar C" . (poly_mode ? poly_color : line_color)
Gui, Add, Button, w50 h30 x112 gGuiClose, Exit
Gui, Show,, Loading
SetTimer, loading_bar, 100
Return

loading_bar:
GuiControl, 3:, loading_bar, 1
Return

save:
Gui, Submit, NoHide
Loop, % LV_GetCount()
{
	LV_GetText(lv_num, A_Index, 1)
	LV_GetText(lv_text, A_Index, 2)
	LV_GetText(lv_type, A_Index, 3)
	LV_GetText(lv_string, A_Index, 4)
	search_type%lv_num% := lv_type
	search_string%lv_num% := lv_string
	StringReplace, hover_text, hover_text, #%lv_text%#, #%lv_num%#, 1
}
save_options("shp_to_placefile.ini", "placefile_name|threshold|poly_mode|poly_color|poly_alpha|line_color|line_width|limit_n|limit_s|limit_e|limit_w|hide_shp2text")
Return

fade_in:
trans += 16
WinSet, Trans, %trans%, ahk_id %hwnd%
If (trans >= 255)
{
	WinSet, Trans, Off, ahk_id %hwnd%
	SetTimer, fade_in, Off
}
Return

fade_out:
trans -= 16
WinSet, Trans, %trans%, ahk_id %hwnd%
If (trans <= 0)
	ExitApp
Return

load_options(file, options) {
	local s
	Loop, Parse, options, |
	{
		StringSplit, s, A_LoopField, :
		IniRead, %s1%, %file%, options, %s1%, %s2%
	}
	Return
}

save_options(file, options) {
	global
	Loop, Parse, options, |
		IniWrite, % %A_LoopField%, %file%, options, %A_LoopField%
	Return
}

file_list(ext, select = 1) {
	Loop, *.%ext%
	{
		If (A_LoopFileExt != ext)
			Continue
		list .= "|" . A_LoopFileName
	}
	GuiControl,, %ext%_file, %list%
	GuiControl, Choose, %ext%_file, |%select%
	Return
}

hover_text(s) {
	global
	Loop, %field0%
		StringReplace, s, s, #%A_Index%#, % field%A_Index%, 1
	StringReplace, s, s, `n, \n, 1
	Return, s
}

rgb(hex, commas = 0) {
	dec := ((0x0 . SubStr(hex, 3, 2)) + 0) . ", " . ((0x0 . SubStr(hex, 5, 2)) + 0) . ", " . ((0x0 . SubStr(hex, 7, 2)) + 0)
	If (!commas)
		StringReplace, dec, dec, `,,, 1
	Return, dec
}

search(search_type, search_string, string) {
	If (!search_type)
		Return, 1
	Return, %search_type%(search_string, string)
}

Equals(search_string, string) {
	Return, (search_string = string)
}

GreaterThan(search_string, string) {
	Return, (search_string < string)
}

LessThan(search_string, string) {
	Return, (search_string > string)
}

GreaterOrEqual(search_string, string) {
	Return, (search_string <= string)
}

LessOrEqual(search_string, string) {
	Return, (search_string >= string)
}

NotEqual(search_string, string) {
	Return, (search_string != string)
}

Contains(search_string, string) {
	Return, InStr(string, search_string)
}

NotContain(search_string, string) {
	Return, !InStr(string, search_string)
}

StartsWith(search_string, string) {
	Return, (search_string = SubStr(string, 1, StrLen(search_string)))
}

EndsWith(search_string, string) {
	Return, (search_string = SubStr(string, -StrLen(search_string) + 1))
}

RegExp(search_string, string) {
	Return, RegExMatch(string, search_string)
}