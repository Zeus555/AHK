;~ No lea variables de ambiente.
#NoEnv

;~ Una sola instancia del programa, en caso que este corriendo previamente, la fuerza a terminar, y se ejecuta con normalidad.
#SingleInstance Force

;Directiva que da mas tiempo al esperar la disponibilidad del porta papeles.
#ClipboardTimeout 2000

;Directiva que indica que el script estará permanentemente en ejecución hasta que encuentre un fin de ejecución.
#Persistent

;Directiva que indica que si encuentra otra instancia ejecutandose la cierre y cargue nuevamente.
#SingleInstance force

;~ Tiempo que se dilata en presionar una tecla en milisegundos.
SetKeyDelay, 50

;~ Elevar los privilegios como administrador, para evitar bloqueos en comandos.
if not A_IsAdmin
{
    Run *RunAs "%A_ScriptFullPath%"
    ExitApp
}

; Calcular posición centrada
winX := (A_ScreenWidth - 300) // 2
;~ winY := (A_ScreenHeight - 100) // 2
winY := 50

; Mostrar ventana GUI superpuesta sin robar foco
StatusLog := "Validar pre-requisitos."
Gui, CopilotMsg:New
Gui, CopilotMsg:+ToolWindow +AlwaysOnTop -Caption +E0x08000000 ; WS_EX_NOACTIVATE
Gui, CopilotMsg:Color, FFFFE0
Gui, CopilotMsg:Font, s10, Segoe UI
Gui, CopilotMsg:Add, Edit, vStatusText ReadOnly Multi Left w300 h60, %StatusLog%
Gui, CopilotMsg:Show, x%winX% y%winY% w310 h70 NoActivate

; Verifica si Google Chrome está instalado
chromePaths := ["C:\Program Files\Google\Chrome\Application\chrome.exe","C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
chromeInstalled := false

;~ Valida la existencia de path de Google Chrome.
for index, path in chromePaths {
    if FileExist(path) {
        chromeInstalled := true
        break
    }
}

;~ Validar si tenemos Google Chrome instalado o no.
if (chromeInstalled = false)
{
	MsgBox, 16, Requisito indispensable, Debes tener instalado Google Chrome.
	ExitApp 1
}

;Verificamos si existe carpeta de aplicación, sino la creamos.
IfNotExist , %A_AppData%\MassiveCopilot\
{
    ;Creamos directorio de aplicación
	FileCreateDir , %A_AppData%\MassiveCopilot\

	;Verificamos si directorio de aplicación fue creado con exito
	if ErrorLevel
	{
		;~ Reportamos el error.
		Msgbox , 16, Requisito indispensable, No se pudo crear directorio de la aplicación en %A_AppData%.
		;~ Salimos de la aplicación.
		ExitApp 2
	}
}

;~ Quitamos menu de opciones de barra de estado estandar y dejamos solamente botones ayuda, licencia y salir.
menu, tray, NoStandard
menu, tray, add, Ayuda
menu, tray, add, Licencia
menu, tray, add, Salir

;~ Especificar directorio donde se encuentran las imagenes a buscar y agregarlas a tu programa compilado.
FileInstall , D:\PRC Monitor Price\ImgSearch\Copilot_Input_Question.bmp, %A_AppData%\MassiveCopilot\Copilot_Input_Question.bmp, 1
FileInstall , D:\PRC Monitor Price\ImgSearch\Copilot_btnSend.bmp, %A_AppData%\MassiveCopilot\Copilot_btnSend.bmp, 1
FileInstall , D:\PRC Monitor Price\ImgSearch\Copilot_btnCopiar.bmp, %A_AppData%\MassiveCopilot\Copilot_btnCopiar.bmp, 1
FileInstall , D:\PRC Monitor Price\ImgSearch\Copilot_btnAbajo.bmp, %A_AppData%\MassiveCopilot\Copilot_btnAbajo.bmp, 1

;~ Especificar imagenes a buscar.
ImgInputQuestion := A_AppData . "\MassiveCopilot\Copilot_Input_Question.bmp"
ImgbtnSend := A_AppData . "\MassiveCopilot\Copilot_btnSend.bmp"
ImgbtnCopiar := A_AppData . "\MassiveCopilot\Copilot_btnCopiar.bmp"
ImgbtnAbajo := A_AppData . "\MassiveCopilot\Copilot_btnAbajo.bmp"

LogStatus("Pedir prompt al usuario.")

;~ Obtener del usuario la pregunta que se le hara a Copilot.
Gui, CopilotInput:+ToolWindow -SysMenu  ; Oculta los botones de cerrar, minimizar y maximizar
Gui, CopilotInput:Font, s14, Segoe UI 
Gui, CopilotInput:Add, Text,, Massive Copilot - Escribe una pregunta.
Gui, CopilotInput:Add, Edit, vUserInput w500 h150, ¿Cómo puedo automatizar tareas repetitivas en mi trabajo para ahorrar tiempo y ser más eficiente? soy vendedor para Sony

; Añadir los botones en la misma línea
Gui, CopilotInput:Add, Button, x100 y+10 w80 gSubmit, OK
Gui, CopilotInput:Add, Button, x+10 w80 gCancel, Cancelar

;~ Mostrar el formulario.
Gui, CopilotInput:Show,, Massive Copilot
return

Submit:
Gui, CopilotInput:Submit
ErrorLevel := 0

if !ErrorLevel
{
	;~ Buscar ventana de Google Chrome si existe, o no. 
	IfWinNotExist, ahk_exe chrome.exe
	{
		LogStatus("Abrir nueva ventana de Google Chrome.")
		
		;~ Abrir Google Chrome en modo depuracion con un perfil de usuario especifico y en la pagina de Copilot.
		Run, chrome.exe --remote-debugging-port=9222 --remote-allow-origins=* --user-data-dir="%A_AppData%\MassiveCopilot\ChromeProfile" --do-not-de-elevate https://copilot.microsoft.com/chats/
		WinActivate, ahk_exe chrome.exe
		WinWaitActive, ahk_exe chrome.exe
		sleep, 1000
		;~ Obtener el ID de la ventana actual.
		WinGetTitle, TitleWinCopilot, A
	}
	else
	{
		LogStatus("Buscar pestaña Copilot.")
		
		;~ Si la ventana de Google Chrome existe, entonces la activamos.
		WinActivate, ahk_exe chrome.exe
		WinWaitActive, ahk_exe chrome.exe
		Sleep, 1000
		;~ Esperar hasta que se encuentre la pestaña de Copilot.
		Loop, 20
		{
			;~ Obtener el ID de la ventana actual.
			WinGetTitle, TitleWinCurr, A
			
			LogStatus("Ventana actual:" . TitleWinCurr)
			
			;~ En caso que coincida el ID de la primera ventana encontra con la ventana actual.
			if (TitleWinCurr = TitleFirstWin)
			{
				;~ Abrir una nueva pestaña de Google Chrome en modo depuracion con un perfil de usuario especifico y en la pagina de Copilot.
				Run, chrome.exe --remote-debugging-port=9222 --remote-allow-origins=* --user-data-dir="%A_AppData%\MassiveCopilot\ChromeProfile" --do-not-de-elevate https://copilot.microsoft.com/chats/
				sleep, 2000
				WinWaitActive, ahk_exe chrome.exe
				;~ Obtener el ID de la ventana actual.
				WinGetTitle, TitleWinCopilot, A
			}
			
			;~ Si es la primera iteracion obtener el ID de la primera ventana correspondiente al titulo de la ventana de la primera pestaña encontrada.
			if (A_Index = 1)
			{
				TitleFirstWin := TitleWinCurr
			}
			
			;~ Buscar ID de la ventana que indique esta en la pestaña de Copilot.
			SetTitleMatchMode RegEx
			WinGet, IDCopilot, ID,^Microsoft.Copilot+
			if IDCopilot
			{
				LogStatus("Pestaña de Copilot encontrada.")
				
				WinActivate, ahk_id %IDCopilot%
				TitleWinCopilot := TitleWinCurr
				break
			}
			else 
			{
				LogStatus("Cambiar de pestaña.")
				Send, ^{Tab}
				Sleep, 2000
			}
		}
	}
	
	LogStatus("Limpiar el portapapeles.")

	;~ Limpiar el portapapeles.
	Clipboard := ""

	;~ Activar ventana de Copilot.
	IfWinExist, %TitleWinCopilot%
	{
		WinActivate, %TitleWinCopilot%
		WinWaitActive, %TitleWinCopilot%,,5
		WinMaximize, %TitleWinCopilot%
		Sleep, 1000
	}
	
	LogStatus("Ventana actual:" . TitleWinCopilot)
	
	;~ Validar si la ventana de Copilot es la ventana activa.
	IfWinActive, %TitleWinCopilot%
	{
		LogStatus("Buscar inputbox en web de Copilot.")
		
		;Seteamos modo de coordenadas para comando de busqueda de imagenes
		CoordMode, Pixel, Screen
		
		loop, 5
		{
			LogStatus("Buscar inputbox ciclo#." . A_Index)
			;~ Buscar inputbox de Copilot para especificar las preguntas necesarias.
			v1X=
			v1Y=
			dif1 := 10 * A_Index
			Indice1 := A_Index
			ImageSearch, v1X, v1Y, 0, 0, A_ScreenWidth, A_ScreenHeight, *%dif1% %ImgInputQuestion%
			if (v1X >0 and v1Y >0)
			{
				LogStatus("Escribiendo prompt en inputbox.")
				CoordMode, Mouse
				MouseClick, left, v1X+25, v1Y+40,1,25
				Sleep, 200
				Send, ^a{Delete}
				SendMode Event
				Send, %UserInput%
				SendMode Input
				sleep, 500
				
				loop, 5
				{
					LogStatus("Buscar boton enviar ciclo#." . A_Index)
					
					dif2 := 10 * A_Index
					Indice2 := A_Index
					;~ Dar click en boton enviar.
					v2X=
					v2Y=
					ImageSearch, v2X, v2Y, 0, 50, A_ScreenWidth, A_ScreenHeight, *%dif2% %ImgbtnSend%
					if (v2X >0 and v2Y >0)
					{
						LogStatus("Dar click en boton enviar.")
						
						MouseClick, left, v2X+25, v2Y+25,1,25
						Sleep, 1000
				     
						loop, 20
						{
							LogStatus("Buscar boton copiar respuesta ciclo#" . A_Index)
							
							dif3 := 10 * A_Index
							Indice3 := A_Index
							;~ Dar click en boton copiar al portapapeles.
							v3X=
							v3Y=
							ImageSearch, v3X, v3Y, 0, 0, A_ScreenWidth, A_ScreenHeight, *60 %ImgbtnCopiar%
							if (v3X >0 and v3Y >0)
							{
								;~ Dar click en boton copiar al portapapeles.
								MouseClick, left, v3X+25, v3Y+25,1,35
								;~ Esperar a que el portapapeles contenga el texto de la respuesta.
								ClipWait, 0.25
								if !Errorlevel
								{
									LogStatus("Respuesta copiada en el portapapeles.")
									break
								}
							}
							
							LogStatus("Buscar boton bajar ciclo#" . A_Index)
							;~ Dar click en boton ir hacia abajo en la respuesta.
							v4X=
							v4Y=
							ImageSearch, v4X, v4Y, 0, 0, A_ScreenWidth, A_ScreenHeight, *30 %ImgbtnAbajo%
							if (v4X >0 and v4Y >0)
							{
								MouseClick, left, v4X+25, v4Y+25,1,35
							}
							
							LogStatus("Escrolear hacia abajo ciclo#" . A_Index)
							;~ Mover hacia abajo la rueda del raton para desplazar la pagina hacia abajo.
							MouseClick, WheelDown,% A_ScreenWidth//2, % A_ScreenHeight//2, 5
							
							Sleep, 1000
						}
						break
					}
					
					;~ Si se llega al ultimo intento por encontrar el boton send, reportar el error y salir.
					if (Indice2 = 5)
					{
						LogStatus("Error al buscar boton enviar luego de " . Indice2 . " intentos.")
						MsgBox, 16, Boton Enviar, Luego de %Indice2% intentos no se pudo encontrar.
						ExitApp, 3
					}
					sleep, 1000
				}
				break
			}
			
			;~ Si se llega al ultimo intento por encontrar el inputbox, reportar el error y salir.
			if (Indice1 = 5)
			{
				LogStatus("Error al buscar boton enviar luego de " . Indice2 . " intentos.")
				MsgBox, 16, Boton inputbox, Luego de %Indice1% intentos no se pudo encontrar el inputbox para poner el prompt.
				ExitApp, 4
			}
			Sleep, 1000
		}
	}
}

Salir:
	MsgBox, Fin
	Gui, CopilotMsg:Destroy
	Gui, CopilotInput:Destroy
	ExitApp, 0

; Acción al presionar Cancelar
Cancel:
	Gui, CopilotInput:Destroy
	LogStatus("Prompt cancelado.")
	MsgBox, Fin
	Gui, CopilotMsg:Destroy
	ExitApp, 1

Esc::  ; Se activa al presionar Esc
	Gosub, Cancel
return

Enter::  ; Se activa al presionar Enter
	Gosub, Submit
return

Ayuda:
MsgBox, 64, Ayuda — Massive Copilot,
( LTrim
---------------------------------------------------
        Massive Copilot — Ayuda
---------------------------------------------------

Esta aplicación de escritorio para Windows utiliza Google Chrome
para conectarse al sitio web oficial de Copilot de Microsoft.

* ¿Qué hace?
- Solicita al usuario una pregunta mediante un InputBox.
- Accede automáticamente a Copilot en línea (con conexión a Internet).
- Envía la pregunta usando la versión más avanzada: GPT-5.
- Copia la respuesta generada directamente al portapapeles.
- Finaliza el programa tras completar la operación.

* Requisitos:
- Google Chrome instalado
- Conexión activa a Internet
- Acceso al sitio web de Copilot
- Resolución minima de pantalla de 800 x 600
- Ejecutarse como administrador

Tuani Labs no almacena datos ni respuestas. Todo se gestiona
localmente y se copia al portapapeles para tu comodidad.

---------------------------------------------------
)
return

Licencia:
MsgBox, 64, Información de Licencia,
( LTrim
---------------------------------------------------
        © 2011–2025 Tuani Labs Technologies
---------------------------------------------------

Producto: Massive Copilot Suite
Versión: 1.0 — Edición Profesional 32-bit
Actualización: 1.0.0 (Estable)
ID del Producto: 00000-000-0000000-00000

Licencia registrada para: Publicidad para Redes Sociales
Tipo de licencia: Uso autorizado — No transferible

Gracias por confiar en Tuani Labs.
---------------------------------------------------
)
return

LogStatus(msg)
{
	global
	GuiControl, CopilotMsg:, StatusText, %msg%
}
