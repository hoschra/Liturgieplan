;SET cursor on first day

WinWait, Umfragedaten (2 von 3) - nuudel - Google Chrome, 
IfWinNotActive, Umfragedaten (2 von 3) - nuudel - Google Chrome, , WinActivate, Umfragedaten (2 von 3) - nuudel - Google Chrome, 
WinWaitActive, Umfragedaten (2 von 3) - nuudel - Google Chrome, 
;MouseClick, left,  163,  211
;Sleep, 100
;MouseClick, left,  131,  262
;Sleep, 100
;Send, {TAB}{TAB}{TAB}{TAB}
Loop
{
   FileReadLine, line, input.txt, %A_Index%
   if ErrorLevel
		break
	StringSplit,line_array,line,";"
	;MsgBox, 4, , Line #%A_Index% is "%line_array1% - "%line_array2%-"%line_array3%".  Continue?
   ;IfMsgBox, No
   ;   return
	sleep,500
	send,%line_array1%{TAB}{TAB}
	sleep,500
	send,%line_array2%{TAB}%line_array3%{TAB}%line_array4%{TAB}{TAB}{TAB}
	line_array2:=""
	line_array3:=""
	line_array4:=""
	
}

