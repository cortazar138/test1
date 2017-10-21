$ScriptControl = New-Object -comobject MSScriptControl.ScriptControl
$ScriptControl.language = "vbscript"
$SAPGUI = $ScriptControl.Eval('(GetObject("SAPGUI")).GetScriptingEngine')
$d = $SAPGUI.FindById("ses[0]/wnd[1]/tbar[0]/btn[0]")
$d | Get-Member