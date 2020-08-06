'1098821232	
'650003195
Dim CasoPrueba, Num_Iter
Dim p, v,  F12, d, F18, v18, F6, m, f
Dim validPrecio
Dim Proceso

Num_Iter = Environment.Value("ActionIteration") 
CasoPrueba = DataTable("e_CasoPrueba",dtLocalSheet)

Call ValidacionesCasosPrueba()

Sub ValidacionesCasosPrueba()

Call ValidacionContado()
Call Validacion12Cuotas()
Call AntesNPC()
Call ValidaNPC(18)
Call ValidaNPC(12)
Call finNpc()


'	Select Case CasoPrueba
'		Case 5	
'		
'			Call CambioPlan("Plan Mi Movistar S/ 65.9 :")
'			Call CambioEquipo("SAMSUNG GXY A20 NEGRO SM-A205G")	
'			
'			'-    VALIDA CONTADO
		'	Call ValidacionContado()
'			
'			'-    VALIDA 12 CUOTAS
			'Call Validacion12Cuotas()
'	
'			Call CambioPlan("Plan Mi Movistar S/ 89.9 :")
'			
'			Call  CambioEquipo("SAMSUNG GXY A20 NEGRO SM-A205G")	
'			
'			'-    VALIDA CONTADO
'			Call ValidacionContado()
'			
'			'-    VALIDA 12 CUOTAS
'			Call Validacion12Cuotas()
'			
'			'-    VALIDA 18 CUOTAS
'			Call Validacion18Cuotas()
'				
'			'- 	  VALIDA 12 CUOTAS	
'			Call Validacion12Cuotas()
''			if validPrecio = "1" then
''				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ValFlujoCont.png", True
''				imagenToWord "No se realizo financiamiento porque no existe precio se terminara el flujo con Contado.",RutaEvidencias() &Num_Iter&"_"&"ValFlujoCont.png"
''			End If
			'Call AntesNPC()
		'	Call ValidaNPC(18)
			
'			Call finNpc()
'		Case 7
'			Call CambioPlan("Plan Mi Movistar Total 30GB:")
'			Call  CambioEquipo("SAMSUNG GXY A20 NEGRO SM-A205G")	
'			'-    VALIDA CONTADO
'			Call ValidacionContado()
'			'- 	  VALIDA 12 CUOTAS	
'			Call Validacion12Cuotas()
'			'-    VALIDA 18 CUOTAS
'			Call Validacion18Cuotas()
'			Call AntesNPC()
'			Call ValidaNPC(18)
'			Call finNpc()
'		Case 3
'			Call CambioPlan("Plan Mi Movistar S/75.9 :")
'			Call  CambioEquipo("SAMSUNG GXY S10 PLUS NEG SM-G975FZ 128GB")
'						'-    VALIDA CONTADO
'			Call ValidacionContado()
'			'- 	  VALIDA 12 CUOTAS	
'			Call Validacion12Cuotas()
'			'-    VALIDA 18 CUOTAS
'			Call Validacion18Cuotas()
'			Call AntesNPC()
'			Call ValidaNPC(1)
'			Call ValidaNPC(12)
'			Call ValidaNPC(18)
'			Call finNpc()
'		Case 6
'			Call CambioPlan("Plan Mi Movistar S/89.9 :")
'			Call CambioEquipo("SAMSUNG GXY A20 NEGRO SM-A205G")
'			'-    VALIDA CONTADO
'			Call ValidacionContado()
'			'-    VALIDA 18 CUOTAS
'			Call Validacion18Cuotas()
'			'- 	  VALIDA 12 CUOTAS	
'			Call Validacion12Cuotas()
'			Call AntesNPC()
'			Call ValidaNPC(18)
'			Call ValidaNPC(1)
'			Call ValidaNPC(12)
'			Call finNpc()
'		Case 15
'			Call CambioPlan("Plan Mi Movistar S/65.9 :")
'			Call CambioEquipo("SAMSUNG GXY S10 PLUS NEG SM-G975FZ 128GB")
'			'- 	  VALIDA 12 CUOTAS	
'			Call Validacion12Cuotas()
'			'-    VALIDA 18 CUOTAS
'			Call Validacion18Cuotas()
'			'-    VALIDA CONTADO
'			Call ValidacionContado()
'		Case 32
'			
'			Call ValidaNPCCorp(1)
'			Call ValidaNPCCorp(D1)
'			Call ValidaNPCCorp(18)
'			Call ValidaNPCCorp(12)
'			Call ValidaNPCCorp(6)
'		
'		Case 29
'			
'			Call ValidaNPCCorp(1)
'			Call ValidaNPCCorp(D1)
'			Call ValidaNPCCorp(6)
'			Call ValidaNPCCorp(12)
'			Call ValidaNPCCorp(18)
'			
'			
'	End Select
'
End Sub
Sub Carga()

End Sub

Sub ValidacionContado()
	
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaList("Plan de Financiamiento:").Exist = False
		wait 1
	Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaList("Plan de Financiamiento:").Select "Contado"	
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Calcular").Click
    While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("568,00(st)").Exist = false
    	wait 1
    Wend
	m = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("568,00(st)").GetROProperty("text")
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"Contado.png", True
	imagenToWord "Validación Lista",RutaEvidencias() &"Contado.png"
	If m <> "0,00" Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("568,00(st)").CaptureBitmap RutaEvidencias() &"ValContado.png", True
	    imagenToWord "Monto Calculado Contado: ",RutaEvidencias() &"ValContado.png"
	End If	
	
End Sub
Sub Validacion12Cuotas()
    dim d
    d = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("568,00(st)").GetROProperty("text")
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaList("Plan de Financiamiento:").Select "MOVISTAR-12 cuotas"
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"CambioFin12.png", True
	imagenToWord "Cambiamos a financiamiento 12 cuotas",RutaEvidencias() &"CambioFin12.png"
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Cambiar forma de pago").GetROProperty("enabled") = "0" Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Cambiar forma de pago").CaptureBitmap RutaEvidencias() &"Check0.png", True
	    imagenToWord "Verificamos que el Check cambiar forma de pago esté marcado y deshabilitado",RutaEvidencias() &"Check0.png"
	End If
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Calcular").Click
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("568,00(st)").GetROProperty("text") = d
		wait 1
	Wend
	
	If ((JavaDialog("JDialog").JavaStaticText("No existe precio para").Exist) or(JavaWindow("Ejecutivo de interacción").JavaDialog("JDialog").JavaStaticText("No existe precio para").Exist)) = true Then
		if JavaWindow("Ejecutivo de interacción").JavaDialog("JDialog").JavaStaticText("No existe precio para").Exist = True Then
			v = JavaWindow("Ejecutivo de interacción").JavaDialog("JDialog").JavaStaticText("No existe precio para").GetROProperty("text")
			validPrecio = "1"
		Else 
			validPrecio = "0"
		End If
		If JavaDialog("JDialog").JavaStaticText("No existe precio para").Exist = True Then
			v = JavaDialog("JDialog").JavaStaticText("No existe precio para").GetROProperty("text")
			validPrecio = "1"
		Else 
			validPrecio = "0"
		End If
		
		wait 1
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ValMensajeFInan12.png", True
		imagenToWord "Se muestra mensaje: "&v,RutaEvidencias() &Num_Iter&"_"&"ValMensajeFInan12.png"	
		If JavaDialog("JDialog").JavaButton("Aceptar").Exist = True Then
			JavaDialog("JDialog").JavaButton("Aceptar").Click
		End If
		If JavaWindow("Ejecutivo de interacción").JavaDialog("JDialog").JavaButton("Aceptar").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("JDialog").JavaButton("Aceptar").Click
		End If
			
	else
        JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"Financiado12.png", True
	    imagenToWord "Validación Lista",RutaEvidencias() &"Financiado12.png"
	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("568,00(st)").CaptureBitmap RutaEvidencias() &"ValFin12.png", True
	    imagenToWord "Monto Calculado Financiado 12 cuotas: ",RutaEvidencias() &"ValFin12.png"
	End If
	wait 1	
End Sub
Sub Validacion18Cuotas()
	Dim c
	c = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("568,00(st)").GetROProperty("text")
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaList("Plan de Financiamiento:").Select "MOVISTAR-18 cuotas"
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"CambioFin18.png", True
	imagenToWord "Cambiamos a financiamiento 18 cuotas",RutaEvidencias() &"CambioFin18.png"	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Cambiar forma de pago").GetROProperty("enabled") = "0" Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Cambiar forma de pago").CaptureBitmap RutaEvidencias() &"Check0.png", True
	    imagenToWord "Verificamos que el Check cambiar forma de pago esté marcado y deshabilitado",RutaEvidencias() &"Check0.png"
	End If
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Calcular").Click
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("568,00(st)").GetROProperty("text") = c
		wait 1
	Wend
		
	If ((JavaWindow("Ejecutivo de interacción").JavaDialog("JDialog").JavaStaticText("No existe precio para").Exist) or (JavaDialog("JDialog").JavaStaticText("No existe precio para").Exist)) = true Then
		if JavaWindow("Ejecutivo de interacción").JavaDialog("JDialog").JavaStaticText("No existe precio para").Exist = True Then
			d = JavaWindow("Ejecutivo de interacción").JavaDialog("JDialog").JavaStaticText("No existe precio para").GetROProperty("text")
			validPrecio = "1"
		Else 
			validPrecio = "0"
		End If
		If JavaDialog("JDialog").JavaStaticText("No existe precio para").Exist = True Then
			d = JavaDialog("JDialog").JavaStaticText("No existe precio para").GetROProperty("text")
			validPrecio = "1"
		Else 
			validPrecio = "0"
		End If
		
		wait 1
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ValMensajeFInan18.png", True
		imagenToWord "Se muestra mensaje: "&d,RutaEvidencias() &Num_Iter&"_"&"ValMensajeFInan18.png"	
		If JavaDialog("JDialog").JavaButton("Aceptar").Exist = True Then
			JavaDialog("JDialog").JavaButton("Aceptar").Click
		End If
		If JavaWindow("Ejecutivo de interacción").JavaDialog("JDialog").JavaButton("Aceptar").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("JDialog").JavaButton("Aceptar").Click
		End If
		
	else
		 JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"Financiado18.png", True
	    imagenToWord "Validación Lista",RutaEvidencias() &"Financiado18.png"
	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("568,00(st)").CaptureBitmap RutaEvidencias() &"ValFin18.png", True
	    imagenToWord "Monto Calculado Financiado 18 cuotas: ",RutaEvidencias() &"ValFin18.png"
	End If
		
End Sub
Sub CambioPlan(plan)
	While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaEdit("Dispositivo seleccionado:").Exist or JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaEdit("TextFieldNative$1").Exist) = False
			wait 1
		Wend
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaEdit("TextFieldNative$1").Exist = true Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaEdit("TextFieldNative$1").Set plan	
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaEdit("Dispositivo seleccionado:").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaEdit("Dispositivo seleccionado:").Set plan	
		End If
			
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Buscar").Click
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Seleccionar").Exist = False
			wait 1
		Wend
		
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Seleccionar").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Seleccionar").Set "ON"	
		End If 
'		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Seleccionar_2").Exist = True Then
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Seleccionar_2").Set "ON"	
'		End If
'
		
		wait 1
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SelPlan.png", True
		imagenToWord "Selección de Plan",RutaEvidencias() &Num_Iter&"_"&"SelPlan.png"
		
End Sub
Sub CambioEquipo(Equipo)
	
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Reemplazar dispositivo").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Reemplazar dispositivo").Click		
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Agregar Equipo").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Agregar Equipo").Click
		End If
		
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaEdit("TextFieldNative$1").Exist = False
			wait 1
		Wend
		
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaEdit("TextFieldNative$1").Set Equipo
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaButton("Buscar").Click
		
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaCheckBox("Seleccionar").Exist = False
			wait 1
		Wend 
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaCheckBox("Seleccionar").Set "ON"
		wait 2
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SelEquipo.png", True
		imagenToWord "Selección de Equipo",RutaEvidencias() &Num_Iter&"_"&"SelEquipo.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaButton("Agregar").Click
End Sub
Sub AntesNPC()

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Negociar configuración").Set "ON"
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"NegConf.png", True
	imagenToWord "Negociar configuración",RutaEvidencias() &Num_Iter&"_"&"NegConf.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Siguiente >").Click
	Proceso = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").GetROProperty("text")
	wait 2
	Proceso = left(proceso, 6)
	If Proceso <> "Alta E" Then
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaTable("(Nuevo)").Exist = false
			wait 1
		Wend
		wait 2
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ValidaCambPlan.png", True
		imagenToWord "Se valida el cambio de plan",RutaEvidencias() &Num_Iter&"_"&"ValidaCambPlan.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Siguiente >").Click	
	End If
	
	
End Sub
Sub ValidaNPC(cuota)
	

	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Exist = False
		wait 1
	Wend
	wait 2
	
	Select Case cuota
		Case 12
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "MOVISTAR-12 cuotas"
			wait 1
		Case 18
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "MOVISTAR-18 cuotas"
			wait 1
		Case 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "Contado"
		
	End Select
 	wait 1
 	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ClicValidar.png", True
	imagenToWord "Cambiamos financiamiento a: "&cuota&" cuotas",RutaEvidencias() &Num_Iter&"_"&"ClicValidar.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
	Call Carga()
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist = True Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CambPlanVal.png", True
		imagenToWord "Validación para el cambio de plan",RutaEvidencias() &Num_Iter&"_"&"CambPlanVal.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Aceptar").Click
		Call Carga()
	End If
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Calcular").Click
	Call Carga()
	
		Select Case cuota
		Case 12
			wait 1
			While  JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaStaticText("788,00(st)").Exist = false
				wait 1
			Wend
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"Calculo18.png", True
	        imagenToWord "Monto calculado financiado 12 cuotas ",RutaEvidencias() &"Calculo18.png"
	        JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaStaticText("788,00(st)").CaptureBitmap RutaEvidencias() &"Monto18.png", True
	        imagenToWord "El monto calculado es: ",RutaEvidencias() &"Monto18.png"
		Case 18
			wait 2
			If JavaWindow("Ejecutivo de interacción").JavaDialog("JDialog").JavaStaticText("No existe precio para").Exist Then
			    JavaWindow("Ejecutivo de interacción").JavaDialog("JDialog").CaptureBitmap RutaEvidencias() &"Dialogo.png", True
	            imagenToWord "Validación de Diálogo al calcular financiamiento: ",RutaEvidencias() &"Dialogo.png"
				JavaWindow("Ejecutivo de interacción").JavaDialog("JDialog").JavaButton("Aceptar").Click
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").GetROProperty("enabled") = "0" Then
					 JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"ValBot.png", True
	                 imagenToWord "Validación del botón siguiente ",RutaEvidencias() &"ValBot.png"
	                 JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").CaptureBitmap RutaEvidencias() &"ValBot2.png", True
	                 imagenToWord "Botón <<Siguiente>> deshabilitado ",RutaEvidencias() &"ValBot2.png"
				End If
			End If
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Validaprecio.png", True
			imagenToWord "Se Valida Precio para el financiamiento de 18 cuotas: "& valor,RutaEvidencias() &Num_Iter&"_"&"Validaprecio.png"
			
		Case 1	
			wait 1
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Validaprecio.png", True
			imagenToWord "Se Valida Precio para el pago Contado: "& valor,RutaEvidencias() &Num_Iter&"_"&"Validaprecio.png"
	End Select
wait 1
End Sub
Sub ValidaNPCCorp(cuota)
	

	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Exist = False
		wait 1
	Wend
	wait 2
	
	Select Case cuota
		Case 12
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "Corporativo-12 cuotas"
			wait 1
		Case 6
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "Corporativo-6 cuotas"
			wait 1
		Case 18
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "Corporativo-18 cuotas"
			wait 1
		Case D1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "Diferido-1 cuota"
			wait 1
		Case 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "Contado"
		
	End Select
 	wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ClicValidar.png", True
	imagenToWord "Se da clic en el boton Validar",RutaEvidencias() &Num_Iter&"_"&"ClicValidar.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
		
	While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = false
		wait 1
	Wend
'	While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").GetROProperty("visible") = 1
'		wait 1
'	Wend
	wait 10
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist = True Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CambPlanVal.png", True
		imagenToWord "Validación para el cambio de plan",RutaEvidencias() &Num_Iter&"_"&"CambPlanVal.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Aceptar").Click
		wait 3
	End If
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ClicCalcular.png", True
	imagenToWord "Se da clic en el boton Calcular",RutaEvidencias() &Num_Iter&"_"&"ClicCalcular.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Calcular").Click
	While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = false
		wait 1
	Wend
	wait 2
	'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaStaticText("667,80(st)").Click
	
	Dim valor
	valor = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaStaticText("667,80(st)").GetROProperty("text")
	'msgbox valor
		Select Case cuota
		Case 12
			wait 1
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Validaprecio12.png", True
			imagenToWord "Se Valida Precio para el financiamiento de 12 cuotas: "& valor,RutaEvidencias() &Num_Iter&"_"&"Validaprecio12.png"
		Case 18
			wait 1
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Validaprecio18.png", True
			imagenToWord "Se Valida Precio para el financiamiento de 18 cuotas: "& valor,RutaEvidencias() &Num_Iter&"_"&"Validaprecio18.png"
		Case 1	
			wait 1
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Validaprecio1.png", True
			imagenToWord "Se Valida Precio para el pago Contado: "& valor,RutaEvidencias() &Num_Iter&"_"&"Validaprecio1.png"
		Case D1
			wait 1
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ValidaprecioD1.png", True
			imagenToWord "Se Valida Precio para el financiamiento : "& valor,RutaEvidencias() &Num_Iter&"_"&"ValidaprecioD1.png"
		Case 6
			wait 1
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Validaprecio6.png", True
			imagenToWord "Se Valida Precio para el financiamiento de 6 cuotas: "& valor,RutaEvidencias() &Num_Iter&"_"&"Validaprecio6.png"
	End Select

End Sub
Sub finNpc()
	wait 2
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ValidaNPCDF.png", True
	imagenToWord "Se culmina la validación de la NPC",RutaEvidencias() &Num_Iter&"_"&"ValidaNPCDF.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Click
End Sub


