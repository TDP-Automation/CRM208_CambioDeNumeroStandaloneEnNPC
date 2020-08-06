

If ucase(DataTable("Tipo_Plan", dtLocalSheet)) = "POSTPAGO" Then
	Call Potspago()
	ElseIf ucase(DataTable("Tipo_Plan", dtLocalSheet)) = "PREPAGO" Then
	Call Prepago()
	ElseIf ucase(DataTable("Tipo_Plan", dtLocalSheet)) = "PORTABILIDAD" Then
	Call Portabilidad()
End If



Sub Potspago()
While Window("Ejecutivo de interacción").InsightObject("InsightObject").Exist = false
	wait 1
Wend
wait 1
 Window("Ejecutivo de interacción").InsightObject("InsightObject").Click
 While Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Exist = false
 	wait 1
 Wend
 wait 1
 Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Click
While  Window("Ejecutivo de interacción").InsightObject("InsightObject_16").Exist = false
	wait 1
Wend
wait 1
Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ScoreCalculado.png", True
imagenToWord "Score Calculado", RutaEvidencias() & "ScoreCalculado.png"
Set shell = CreateObject("Wscript.Shell") 
shell.SendKeys "{PGDN}"
If Window("Ejecutivo de interacción").InsightObject("InsightObject_5").Exist = True Then
	Window("Ejecutivo de interacción").InsightObject("InsightObject_5").Click
End If

wait 3
Set shell = CreateObject("Wscript.Shell") 
shell.SendKeys "{PGDN}"

If Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Exist = true then
	wait 1
End If
wait 1
If Window("Ejecutivo de interacción").InsightObject("InsightObject_19").Exist = true Then
	Call Validacion()
End If
If Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Exist = true Then
	Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Click
End If

wait 5
Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ValidacionDatos.png", True
imagenToWord "Validación de Datos Exitosa", RutaEvidencias() & "ValidacionDatos.png"
Window("Ejecutivo de interacción").InsightObject("InsightObject_8").Click
End Sub
Sub Validacion()
	Window("Ejecutivo de interacción").InsightObject("InsightObject_19").Click
	wait 1
	Window("Ejecutivo de interacción").InsightObject("InsightObject_21").Click
	wait 1
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "prueba"
	wait 1
	Window("Ejecutivo de interacción").InsightObject("InsightObject_22").Type micCtrlDwn + micAltDwn + "q" + micCtrlUp + micAltUp
	wait 2
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "gmail.com"
	wait 1

End Sub

Sub Prepago()
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_3").Exist = false
		wait 1
	Wend
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Prepago.png", True
	imagenToWord "Click Prepago", RutaEvidencias() & "Prepago.png"
	Window("Ejecutivo de interacción").InsightObject("InsightObject_3").Click
	wait 2
	If Window("Ejecutivo de interacción").InsightObject("InsightObject_5").Exist = true Then
		Window("Ejecutivo de interacción").InsightObject("InsightObject_5").Click
		wait 3
	End If
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGDN}"
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Exist = false
		wait 1
	Wend
	If Window("Ejecutivo de interacción").InsightObject("InsightObject_23").Exist = true Then
		Window("Ejecutivo de interacción").InsightObject("InsightObject_23").Click
		wait 1
		Window("Ejecutivo de interacción").InsightObject("InsightObject_21").Click
		wait 1
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "prueba"
		wait 1
		Window("Ejecutivo de interacción").InsightObject("InsightObject_22").Type micCtrlDwn + micAltDwn + "q" + micCtrlUp + micAltUp
		wait 2
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "gmail.com"
		wait 1
	End If
	Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Click
	wait 5
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ValidacionDatos.png", True
	imagenToWord "Validación de Datos Exitosa", RutaEvidencias() & "ValidacionDatos.png"
	Window("Ejecutivo de interacción").InsightObject("InsightObject_8").Click
End Sub
Sub Portabilidad()
	Window("Ejecutivo de interacción").InsightObject("InsightObject_2").Click
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_9").Exist = False
		wait 1
	Wend
	If  ucase(DataTable("Tipo_Venta", dtLocalSheet)) = "CHIP" Then
		Window("Ejecutivo de interacción").InsightObject("InsightObject_9").Click
	ElseIf ucase(DataTable("Tipo_Venta", dtLocalSheet)) = "COMBO" Then
		Window("Ejecutivo de interacción").InsightObject("InsightObject_10").Click
	End If
	Window("Ejecutivo de interacción").InsightObject("InsightObject_11").Click
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys DataTable("Num_Porta", dtLocalSheet)
	Window("Ejecutivo de interacción").InsightObject("InsightObject_12").Click
	wait 1
	If ucase(DataTable("Producto_Origen", dtLocalSheet)) = "POSTPAGO" Then
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "{DOWN 2}"
		shell.SendKeys "{ENTER}"
	ElseIf ucase(DataTable("Producto_Origen", dtLocalSheet)) = "PREPAGO" Then
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "{DOWN 1}"
		shell.SendKeys "{ENTER}"
	End If
	Window("Ejecutivo de interacción").InsightObject("InsightObject_14").Click
	If ucase(DataTable("Operador_Origen", dtLocalSheet)) = "OPERADOR VIRTUAL" Then
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "{DOWN 8}"
		shell.SendKeys "{ENTER}"
	End If

	Window("Ejecutivo de interacción").InsightObject("InsightObject_15").Click


	
End Sub

'Window("Ejecutivo de interacción").InsightObject("InsightObject_11").Click
'wait 1
'Set shell = CreateObject("Wscript.Shell") 
'shell.SendKeys "prueba"
'Window("Ejecutivo de interacción").InsightObject("InsightObject_11").Type micCtrlDwn + micAltDwn + "Q" + micCtrlUp + micAltUp
'Set shell = CreateObject("Wscript.Shell") 
'shell.SendKeys "gmail.com"
'wait 1

	
