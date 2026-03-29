#Requires AutoHotkey v2.0
#SingleInstance Force

; ===== CONFIGURACIÓN =====
global libreriaDefecto := "LIBRERIA"
global versionNatural := "213"  ; Valor por defecto
global offsetLibreriaEncontrado := 0
global destinoBase := ObtenerRutaDestino()  ; ← Se calcula dinámicamente
global extensionesNatural := [".NSN", ".NSP", ".NSM", ".NSL", ".NSG", ".NSD", ".NSC", ".NSA", ".NSH"]
global rutasDefecto := ["C:\Users\" . A_UserName . "\workspace110", "C:\Users\" . A_UserName . "\git"]

; ===== VARIABLES GLOBALES =====
global archivosEncontrados := []
global archivosValidos     := []   ; Subconjunto de archivosEncontrados que pasaron ValidarNombreNatural
global txtLibreria := ""  ; Campo de texto para la librería
global rutaOrigen := ""
global nombreCarpetaSeleccionada := ""
global mainGui := ""
global listView := ""
global esCarpetaPersonalizada := false  ; Flag para saber si es carpeta personalizada
global proyectoSeleccionado := ""       ; Proyecto elegido en la pantalla intermedia
global origenSeleccion := ""            ; "workspace" | "personalizada" — para saber a dónde vuelve Volver

; ===== INICIO DEL SCRIPT =====
; Detectar versión de Natural primero
DetectarVersionNatural()

; Leer librería actual desde NATPARM.SAG al inicio
libreriaActual := LeerLibreriaDeNATPARM()

if libreriaActual != ""
    libreriaDefecto := libreriaActual
else
    libreriaDefecto := ""   ; ← Campo queda vacío si no hay librería asignada

; Leer máximo de workfiles desde NATPARM.SAG al inicio (evita que quede en el default hardcodeado)
workMaxEntries := LeerMaxWorkfilesDeNATPARM()

MostrarMenuPrincipal()

; ===== FUNCIÓN: DETECTAR VERSIÓN DE NATURAL =====
DetectarVersionNatural() {
    global versionNatural
    
    rutaBase := A_ScriptDir . "\dos\NATURAL"
    
    if DirExist(rutaBase) {
        Loop Files, rutaBase . "\*", "D"
        {
            if RegExMatch(A_LoopFileName, "^\d+$") {
                versionNatural := A_LoopFileName
                return true
            }
        }
    }
    
    ; Si no encuentra, mantener el valor por defecto (213)
    return false
}

; ===== FUNCIÓN: FORMATEAR VERSIÓN PARA MOSTRAR =====
FormatearVersion(version) {
    ; Convertir "213" a "2.1.3"
    if StrLen(version) = 3 {
        return SubStr(version, 1, 1) . "." . SubStr(version, 2, 1) . "." . SubStr(version, 3, 1)
    }
    ; Si tiene otro formato, devolverlo tal cual
    return version
}

; ===== FUNCIÓN: OBTENER OFFSET POR VERSIÓN (FALLBACK) =====
ObtenerOffsetLibreriaPorVersion() {
    global versionNatural
    
    offsets := Map(
        "213", 0x205A,
        "210", 0x2031,
        "211", 0x2031,
        "212", 0x2031,
        "214", 0x205A
    )
    
    if offsets.Has(versionNatural)
        return offsets[versionNatural]
    
    versionBase := SubStr(versionNatural, 1, 3)
    if offsets.Has(versionBase)
        return offsets[versionBase]
    
    return 0x205A  ; Default
}

; ===== FUNCIÓN: BUSCAR OFFSET DE LIBRERÍA POR PATRÓN =====
BuscarOffsetLibreriaPorPatron() {
    global versionNatural
    
    rutaNATPARM := A_ScriptDir . "\dos\NATURAL\" . versionNatural . "\PROF\NATPARM.SAG"
    
    if !FileExist(rutaNATPARM) {
        return 0
    }
    
    try {
        ; Leer el archivo completo como bytes
        archivo := FileOpen(rutaNATPARM, "r")
        tamaño := archivo.Length
        contenido := Buffer(tamaño)
        archivo.RawRead(contenido, tamaño)
        archivo.Close()
        
        ; Patrón a buscar: 69 00 09 00 (en hexadecimal)
        ; En bytes: [0x69, 0x00, 0x09, 0x00]
        
        ; Buscar el patrón en el archivo
        offsetEncontrado := 0
        
        ; Recorrer el buffer buscando el patrón
        ; Dejamos espacio para los 4 bytes del patrón + 8 bytes de nombre
        Loop tamaño - 12 {
            i := A_Index - 1  ; Índice base 0 para el buffer
            
            ; Verificar si encontramos el patrón
            if (NumGet(contenido, i, "UChar") = 0x69 &&
                NumGet(contenido, i + 1, "UChar") = 0x00 &&
                NumGet(contenido, i + 2, "UChar") = 0x09 &&
                NumGet(contenido, i + 3, "UChar") = 0x00) {
         
                ; El nombre de la librería está inmediatamente después del patrón
                offsetEncontrado := i + 4
                break
            }
        }
        
        if offsetEncontrado > 0 {
            ; Leer el nombre en ese offset para verificar
            nombreBytes := ""
            Loop 8 {
                byte := NumGet(contenido, offsetEncontrado + A_Index - 1, "UChar")
                if byte = 0
                    break
  
               nombreBytes .= Chr(byte)
            }
            
            nombreLimpio := Trim(nombreBytes)
            
            ; Verificar que el nombre parezca válido (letras mayúsculas, números, #, -)
            if RegExMatch(nombreLimpio, "^[A-Z0-9# -]+$") {
                return offsetEncontrado
            }
        }
        
        return 0  ; No encontrado
        
    } catch as err {
        ; Error interno silencioso (no afecta al usuario)
        return 0
    }
}

; ===== FUNCIÓN: LEER LIBRERÍA DE NATPARM.SAG (CORREGIDA PARA AHK v2) =====
LeerLibreriaDeNATPARM() {
    global versionNatural, offsetLibreriaEncontrado
    
    rutaNATPARM := A_ScriptDir . "\dos\NATURAL\" . versionNatural . "\PROF\NATPARM.SAG"
    
    if !FileExist(rutaNATPARM)
        return ""
    
    try {
        ; Buscar el offset por patrón
        offsetLibreria := BuscarOffsetLibreriaPorPatron()
        
        if offsetLibreria = 0 {
            ; Silencioso: usamos el offset por versión pero NO mostramos cartel
            offsetLibreria := ObtenerOffsetLibreriaPorVersion()
        } else {
            offsetLibreriaEncontrado := offsetLibreria
        }
        
        archivo := FileOpen(rutaNATPARM, "r")
        if !archivo
            return ""
    
        archivo.Seek(offsetLibreria)
        
        nombreLimpio := ""
        Loop 8 {
            byte := archivo.ReadUChar()
            if byte = 0
                break
            nombreLimpio .= Chr(byte)
 
        }
        archivo.Close()
        
        nombreLimpio := Trim(nombreLimpio)
        
        ; Si no hay nombre válido → devolvemos vacío (para que el campo quede en blanco)
        if nombreLimpio = "" || !RegExMatch(nombreLimpio, "^[A-Z0-9# -]+$")
            return ""
        
        return nombreLimpio
        
    } catch as err {
        ; Solo error real (no el aviso normal)
        return ""
    }
}

; ===== FUNCIÓN: ESCRIBIR LIBRERÍA EN NATPARM.SAG (CORREGIDA PARA AHK v2) =====
EscribirLibreriaEnNATPARM(nombreLibreria) {
    global versionNatural, offsetLibreriaEncontrado
    
    rutaNATPARM := A_ScriptDir . "\dos\NATURAL\" . versionNatural . "\PROF\NATPARM.SAG"
    
    if !FileExist(rutaNATPARM) {
        MsgBox("No se encontró NATPARM.SAG en:`n" . rutaNATPARM, "Error", "Icon!")
        return false
    }
    
    try {
        ; Determinar offset a usar
        offsetLibreria := offsetLibreriaEncontrado
        
        if offsetLibreria = "" || offsetLibreria = 0 {
            ; Si no tenemos offset guardado, buscar ahora
            offsetLibreria := BuscarOffsetLibreriaPorPatron()
            
            if offsetLibreria = 0 {
                ; Silencioso: usamos el offset por versión sin mostrar nada al usuario
                offsetLibreria := ObtenerOffsetLibreriaPorVersion()
            }
        }
        
        archivo := FileOpen(rutaNATPARM, "rw")
        if !archivo {
            throw Error("No se pudo abrir el archivo")
     
        }
        
        ; Limpiar y validar nombre
        nombreLimpio := StrUpper(Trim(nombreLibreria))
        if StrLen(nombreLimpio) > 8
            nombreLimpio := SubStr(nombreLimpio, 1, 8)
        
        ; Ir al offset encontrado
        archivo.Seek(offsetLibreria)
        
        ; Escribir el nombre byte a byte
        Loop StrLen(nombreLimpio) {
            char := SubStr(nombreLimpio, A_Index, 1)
            archivo.WriteUChar(Ord(char))
        }
        
        ; Rellenar con 0x00 hasta completar los 8 bytes del campo
        ; Esto evita dejar basura de una librería anterior más larga
        Loop (8 - StrLen(nombreLimpio)) {
            archivo.WriteUChar(0x00)
        }
        
        archivo.Close()
        
        ; Verificar
        nombreVerificado := LeerLibreriaDeNATPARM()
        if nombreVerificado != nombreLimpio {
            throw Error("No se pudo verificar la escritura")
        }
        
        ; Guardar el offset para futuras operaciones
        offsetLibreriaEncontrado := offsetLibreria
        
        return true
        
    } catch as err {
        MsgBox("Error al escribir NATPARM.SAG:`n" . err.Message, "Error", "IconX")
        return false
    }
}

MostrarMenuPrincipal() {
    global mainGui, txtRutaActual, btnEscanear, txtLibreria
	versionFormateada := FormatearVersion(versionNatural)

    mainGui := Gui("-MaximizeBox", "Migrador: NaturalONE → Natural " . versionFormateada)
    mainGui.BackColor := "F5F5F7"
    mainGui.SetFont("s12 c333333", "Segoe UI")

    ; Título
    mainGui.SetFont("s14 bold c0D47A1", "Segoe UI")
    mainGui.Add("Text", "x40 y25 w560 Center", "MIGRACIÓN DE CÓDIGO NATURAL")
    mainGui.SetFont("s12 c555555", "Segoe UI")
    mainGui.Add("Text", "x40 y60 w560 Center", "NaturalONE  →  Natural " . versionFormateada . " (Windows 3.1)")

    mainGui.Add("Progress", "x60 y100 w520 h2 BackgroundE0E0E0 -Smooth Disabled")

    ; Sección origen
    mainGui.SetFont("s11 c444444", "Segoe UI")
    mainGui.Add("Text", "x60 y125", "Seleccione el proyecto de origen:")

    ; === Botón unificado Seleccionar Workspace ===
    margenIzq    := 60
    anchoBoton   := 240
    espacioEntre := 20
    altoBoton    := 50
    yBotones     := 160

    btnWorkspace := mainGui.Add("Button", "x" margenIzq " y" yBotones " w" (anchoBoton*2 + espacioEntre) " h" altoBoton, "📁  Seleccionar Proyecto")
    btnWorkspace.Opt("+Background4CAF50")
    btnWorkspace.SetFont("s11 cFFFFFF bold", "Segoe UI")

    ; Botón carpeta personalizada (debajo, ocupa todo el ancho disponible)
    yCustom := yBotones + altoBoton + 15
    btnCustom := mainGui.Add("Button", "x" margenIzq " y" yCustom " w" (anchoBoton*2 + espacioEntre) " h" altoBoton, "📂  Seleccionar carpeta personalizada...")
    btnCustom.Opt("+Background607D8B")
    btnCustom.SetFont("s11 cFFFFFF", "Segoe UI")

    ; Ruta actual
    yRuta := yCustom + altoBoton + 25
    mainGui.SetFont("s11 c666666", "Segoe UI")
    mainGui.Add("Text", "x" margenIzq " y" yRuta, "Ruta seleccionada:")

    txtRutaActual := mainGui.Add("Edit", "x" margenIzq " y" (yRuta+25) " w520 h32 ReadOnly -TabStop BackgroundFFFFFF c333333", "No seleccionada")
    txtRutaActual.SetFont("s11", "Consolas")

    mainGui.Add("Progress", "x60 y" (yRuta+70) " w520 h2 BackgroundE0E0E0 -Smooth Disabled")

    ; Sección de librería de destino
    yLibreria := yRuta + 90
    mainGui.SetFont("s11 c666666", "Segoe UI")
    mainGui.Add("Text", "x" margenIzq " y" yLibreria, "Librería de destino:")

    txtLibreria := mainGui.Add("Edit", "x" margenIzq " y" (yLibreria+25) " w520 h32 BackgroundFFFFFF c333333 Uppercase", libreriaDefecto)
    txtLibreria.SetFont("s11", "Segoe UI")
    txtLibreria.OnEvent("Change", FiltrarLibreria)

    mainGui.Add("Progress", "x60 y" (yLibreria+70) " w520 h2 BackgroundE0E0E0 -Smooth Disabled")

    ; Botón ESCANEAR
    yEscanear := yLibreria + 90
    btnEscanear := mainGui.Add("Button", "x" margenIzq " y" yEscanear " w520 h60 Disabled", "🔍   ESCANEAR OBJETOS NATURAL")
    btnEscanear.Opt("+BackgroundFF5722")
    btnEscanear.SetFont("s13 bold cFFFFFF", "Segoe UI")

    mainGui.Add("Progress", "x60 y" (yEscanear+75) " w520 h2 BackgroundE0E0E0 -Smooth Disabled")

    ; Botón Gestionar Workfiles
    yWorkfiles := yEscanear + 90
    btnWorkfiles := mainGui.Add("Button", "x" margenIzq " y" yWorkfiles " w520 h42", "🗂  Administrar Workfiles (WORK)")
    btnWorkfiles.Opt("+Background37474F")
    btnWorkfiles.SetFont("s10 cFFFFFF", "Segoe UI")

    mainGui.Add("Progress", "x60 y" (yWorkfiles+55) " w520 h2 BackgroundE0E0E0 -Smooth Disabled")

    ; Botón Salir
    ySalir := yWorkfiles + 75
    btnSalir := mainGui.Add("Button", "x" margenIzq " y" ySalir " w520 h45", "❌  Salir")
    btnSalir.Opt("+Background757575")
    btnSalir.SetFont("s11 cFFFFFF", "Segoe UI")

    ; Eventos
    btnWorkspace.OnEvent("Click", (*) => MostrarSeleccionWorkspace())
    btnCustom.OnEvent("Click", (*) => SeleccionarRutaPersonalizada())
    btnEscanear.OnEvent("Click", (*) => EscanearArchivos())
    btnWorkfiles.OnEvent("Click", (*) => MostrarGestionWorkfiles())
    btnSalir.OnEvent("Click", (*) => ExitApp())
    txtLibreria.OnEvent("LoseFocus", (*) => ActualizarDestino())

    mainGui.OnEvent("Close", (*) => ExitApp())
    mainGui.Show("Center w640 h740")
    
    ; Actualizar destinoBase con la librería actual
    destinoBase := ObtenerRutaDestino()
}

; =============================================================================
;   GESTIÓN DE WORKFILES (WORK) — NATPARM.SAG
; =============================================================================

; ===== CONSTANTES DE WORKFILES =====
; Estructura real en NATPARM.SAG:
;   Patrón de localización: 00 54 00 04 00 20 00 35 00
; - Byte ANTERIOR al patrón: número máximo de workfiles
;   - Byte POSTERIOR al patrón: inicio del slot del Workfile 1
; Cada slot ocupa 53 bytes fijos:
;   - 51 bytes de ruta (terminada en 0x00, resto relleno con 0x00)
; - 2 bytes fijos 0x00 0x00 al final del slot

global WORK_MAX_DEFAULT  := 30
global WORK_SLOT_SIZE    := 53   ; bytes por slot en NATPARM.SAG (51 ruta + 2 fijos)
global WORK_PATH_SIZE    := 51   ; bytes disponibles para la ruta (incluye terminador 0x00)
global workfilesData     := []   ; Array global con los workfiles cargados
global workfileOffsets   := []   ; Offset donde empieza cada slot
global workfileLengths   := []   ; Siempre WORK_SLOT_SIZE
global workOffsetBase    := 0    ; Offset del primer slot (inicio bloque)
global workMaxOffset     := 0    ; Offset del byte de máximo workfiles en NATPARM.SAG
global workMaxEntries    := 30   ; Máximo de workfiles leído de NATPARM.SAG

; ===== FUNCIÓN: LEER WORKFILES DE NATPARM.SAG =====
LeerWorkfilesDeNATPARM(maxEntries := 30) {
    global versionNatural, workfileOffsets, workfileLengths, workMaxOffset, workOffsetBase
    global WORK_SLOT_SIZE, WORK_PATH_SIZE

    rutaNATPARM := A_ScriptDir . "\dos\NATURAL\" . versionNatural . "\PROF\NATPARM.SAG"
    resultado        := []
    workfileOffsets  := []
    workfileLengths  := []

    if !FileExist(rutaNATPARM)
        return resultado

    archivo := FileOpen(rutaNATPARM, "r")
    if !archivo
        return resultado
    tamaño := archivo.Length
    buf    := Buffer(tamaño)
    archivo.RawRead(buf, tamaño)
    archivo.Close()

    ; ── Localizar patrón completo: 00 54 00 04 00 20 00 35 00 ───────────────
    patron := [0x00, 0x54, 0x00, 0x04, 0x00, 0x20, 0x00, 0x35, 0x00]
    patronLen := patron.Length
    inicioBloque := 0

    Loop tamaño - patronLen - 1 {
        i := A_Index - 1
        coincide := true
        for k, b in patron {
            if NumGet(buf, i + k - 1, "UChar") != b {
                coincide := false
                break
            }
        }
        if coincide {
            if workMaxOffset = 0
                workMaxOffset := i - 1
            inicioBloque := i + patronLen
            break
        }
    }

    if inicioBloque = 0
        return resultado

    workOffsetBase := inicioBloque

    ; ── Recorrer slots de 53 bytes ───────────────────────────────────────────
    offsetActual := inicioBloque
    Loop maxEntries {
        if offsetActual + WORK_PATH_SIZE > tamaño
            break

        ruta := ""
        Loop WORK_PATH_SIZE {
            b := NumGet(buf, offsetActual + A_Index - 1, "UChar")
            if b = 0
                break
            ruta .= Chr(b)
        }

        resultado.Push(ruta)
        workfileOffsets.Push(offsetActual)
        workfileLengths.Push(WORK_SLOT_SIZE)

        offsetActual += WORK_SLOT_SIZE
    }

    while resultado.Length < maxEntries {
        resultado.Push("")
        workfileOffsets.Push(0)
        workfileLengths.Push(WORK_SLOT_SIZE)
    }

    return resultado
}

; ===== FUNCIÓN: ESCRIBIR UN WORKFILE EN NATPARM.SAG =====
EscribirWorkfileEnNATPARM(numero, rutaCompleta) {
    global versionNatural, workfileOffsets, workfileLengths
    global WORK_SLOT_SIZE, WORK_PATH_SIZE

    rutaNATPARM := A_ScriptDir . "\dos\NATURAL\" . versionNatural . "\PROF\NATPARM.SAG"

    if !FileExist(rutaNATPARM) {
        MsgBox("No se encontró NATPARM.SAG en:`n" . rutaNATPARM, "Error", "IconX")
        return false
    }

    if numero < 1 || numero > workfileOffsets.Length {
        MsgBox("Número de workfile inválido: " . numero, "Error", "IconX")
        return false
    }

    if workfileOffsets[numero] = 0 {
        MsgBox("No se conoce el offset del Workfile #" . numero . ".`n"
             . "Abra la ventana de Workfiles para que se relean los offsets.", "Error", "IconX")
        return false
    }

    rutaLimpia := Trim(rutaCompleta)

    if StrLen(rutaLimpia) >= WORK_PATH_SIZE {
        MsgBox("La ruta excede el máximo permitido.`n`n"
             . "Máximo: " . (WORK_PATH_SIZE - 1) . " caracteres`n"
             . "Su ruta: " . StrLen(rutaLimpia) . " caracteres",
             "Error - Ruta demasiado larga", "IconX")
        return false
    }

    try {
        archivo := FileOpen(rutaNATPARM, "rw")
        if !archivo
            throw Error("No se pudo abrir NATPARM.SAG para escritura")

        archivo.Seek(workfileOffsets[numero])

        Loop WORK_PATH_SIZE {
            archivo.WriteUChar(0x00)
        }

        archivo.Seek(workfileOffsets[numero])

        Loop StrLen(rutaLimpia) {
            archivo.WriteUChar(Ord(SubStr(rutaLimpia, A_Index, 1)))
        }

        archivo.Close()
        return true

    } catch as err {
        MsgBox("Error al escribir workfile en NATPARM.SAG:`n" . err.Message, "Error", "IconX")
        return false
    }
}

; ===== FUNCIÓN: LEER MÁXIMO DE WORKFILES =====
LeerMaxWorkfilesDeNATPARM() {
    global versionNatural, workMaxOffset
    rutaNATPARM := A_ScriptDir . "\dos\NATURAL\" . versionNatural . "\PROF\NATPARM.SAG"
    workMaxOffset := 0
    if !FileExist(rutaNATPARM)
        return WORK_MAX_DEFAULT
    try {
        archivo := FileOpen(rutaNATPARM, "r")
        if !archivo
            return WORK_MAX_DEFAULT
        tamaño := archivo.Length
        buf    := Buffer(tamaño)
        archivo.RawRead(buf, tamaño)
        archivo.Close()
  
        patron := [0x00, 0x54, 0x00, 0x04, 0x00, 0x20, 0x00, 0x35, 0x00]
        patronLen := patron.Length
        Loop tamaño - patronLen {
            i := A_Index - 1
            coincide := true
            for k, b in patron {
                if NumGet(buf, i + k - 1, "UChar") != b {
                    coincide := false
                    break
                }
            }
            if coincide {
                workMaxOffset := i - 1
                valor := NumGet(buf, workMaxOffset, "UChar")
                if (valor >= 0 && valor <= 32)
                    return valor
                return WORK_MAX_DEFAULT
            }
        }
        return WORK_MAX_DEFAULT
    } catch {
        return WORK_MAX_DEFAULT
    }
}

; ===== FUNCIÓN: ESCRIBIR MÁXIMO DE WORKFILES =====
EscribirMaxWorkfilesEnNATPARM(valor) {
    global versionNatural, workMaxOffset
    rutaNATPARM := A_ScriptDir . "\dos\NATURAL\" . versionNatural . "\PROF\NATPARM.SAG"
    if !FileExist(rutaNATPARM) {
        MsgBox("No se encontró NATPARM.SAG en:`n" . rutaNATPARM, "Error", "IconX")
        return false
    }
    if workMaxOffset = 0 {
        return false
    }
    try {
        archivo := FileOpen(rutaNATPARM, "rw")
        if !archivo
            throw Error("No se pudo abrir NATPARM.SAG para escritura")
        archivo.Seek(workMaxOffset)
        archivo.WriteUChar(valor)
        archivo.Close()
        return true
    } catch as err {
        MsgBox("Error al escribir el máximo de workfiles en NATPARM.SAG:`n" . err.Message, "Error", "IconX")
        return false
    }
}

; =============================================================================
;   VALIDACIÓN DE RUTAS WINDOWS 3.1 (formato 8.3)
; =============================================================================

global WF_NOMBRES_RESERVADOS := ["CON","PRN","AUX","NUL",
    "COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9",
    "LPT1","LPT2","LPT3","LPT4","LPT5","LPT6","LPT7","LPT8","LPT9"]

global WF_CHARS_PERMITIDOS := "_-!@#$%&'(){}^~"

; ===== FUNCIÓN: Validar un segmento de nombre 8.3 (archivo o carpeta) =====
; Ahora requiere obligatoriamente un punto, extensión .SAG prohibida, y extensión vacía o de 3 caracteres
WF_ValidarSegmento83(segmento) {
    global WF_NOMBRES_RESERVADOS, WF_CHARS_PERMITIDOS

    if segmento = ""
        return "El nombre no puede estar vacío."
        
    puntoPos := InStr(segmento, ".")
    if puntoPos = 0
        return "El nombre debe contener obligatoriamente un punto y extensión de 3 caracteres (ej: WORKFILE.TXT)."

    base := SubStr(segmento, 1, puntoPos - 1)
    ext  := SubStr(segmento, puntoPos + 1)

    ; Restringir específicamente la extensión SAG (case-insensitive)
    if StrUpper(ext) = "SAG"
        return "La extensión .SAG no está permitida para los Workfiles."

    ; El nombre no puede empezar con punto
    if SubStr(segmento, 1, 1) = "."
        return "El nombre no puede empezar con punto: '" . segmento . "'"

    ; Longitud del nombre base: 1–8 caracteres
    if StrLen(base) = 0
        return "El nombre base no puede estar vacío en '" . segmento . "'"
    if StrLen(base) > 8
        return "El nombre base excede 8 caracteres: '" . base . "' (" . StrLen(base) . " caracteres)"

    ; Longitud de la extensión: exactamente 3 caracteres obligatorios
    if StrLen(ext) != 3
        return "La extensión debe tener exactamente 3 caracteres (ej: WORKFILE.TXT)."

    ; Caracteres permitidos en nombre base y extensión
    for parte in [base, ext] {
        if parte = ""
            continue
        Loop StrLen(parte) {
            c := SubStr(parte, A_Index, 1)
            if !RegExMatch(c, "i)[A-Z0-9]") && !InStr(WF_CHARS_PERMITIDOS, c) {
                return "Carácter no permitido '" . c . "' en '" . parte . "'"
            }
        }
    }

    ; Nombre reservado (comparar en mayúsculas)
    baseMayus := StrUpper(base)
    for reservado in WF_NOMBRES_RESERVADOS {
        if baseMayus = reservado
            return "Nombre reservado del sistema: '" . segmento . "'"
    }

    return ""   ; Válido
}

MostrarGestionWorkfiles() {
    global versionNatural, workfilesData, workMaxEntries, mainGui, txtFilename

    ; Leer el máximo real desde la dirección 0x3D9
    workMaxEntries := LeerMaxWorkfilesDeNATPARM()

    ; Siempre leer los 32 slots posibles para que el array nunca quede vacío
    workfilesData := LeerWorkfilesDeNATPARM(32)
    while workfilesData.Length < 32
        workfilesData.Push("")

    ; ── Ocultar ventana principal (igual que al escanear) ─────────────────────
    mainGui.Hide()

    ; ── Crear ventana ─────────────────────────────────────────────────────────
    wfGui := Gui("-MaximizeBox +OwnDialogs", "Administración de Workfiles (WORK)")
    wfGui.BackColor := "F0F0F0"
    wfGui.SetFont("s11 c222222", "Segoe UI")

    ; ── Barra de título interna (estilo imagen) ────────────────────────────────
    wfGui.SetFont("s11 bold cFFFFFF", "Segoe UI")
    encTitle := wfGui.Add("Text", "x0 y0 w600 h28 Background1A237E", "")
    titleLbl := wfGui.Add("Text", "x10 y6 w370 Background1A237E cFFFFFF", "Número máximo de Workfiles [WORK]")

    ; Campo editable de cantidad máxima (en la barra azul, a la derecha)
    wfGui.SetFont("s10 bold c000080", "Segoe UI")
    txtMaxWF := wfGui.Add("Edit", "x440 y4 w40 h20 Border Center +Number", workMaxEntries)

    ; Etiqueta aclaratoria del rango DESPUÉS del campo (a su derecha)
    wfGui.SetFont("s9 cCCCCCC", "Segoe UI")
    wfGui.Add("Text", "x485 y7 w100 Background1A237E cCCCCCC Left", "(0-32)")

    ; ── Fila Number + Filename ─────────────────────────────────────────────────
    wfGui.SetFont("s11 bold c222222", "Segoe UI")
    wfGui.Add("Text", "x15 y40", "Número:")

    wfGui.SetFont("s11 c222222", "Segoe UI")
    txtNumber := wfGui.Add("Edit", "x85 y38 w50 h22 Border ReadOnly", "1")

    wfGui.SetFont("s11 bold c222222", "Segoe UI")
    wfGui.Add("Text", "x150 y40", "Archivo:")

    wfGui.SetFont("s11 c222222", "Segoe UI")
    txtFilename := wfGui.Add("Edit", "x220 y38 w360 h22 Border ReadOnly -TabStop BackgroundFFFFFF c222222", "")

    ; ── Separador ─────────────────────────────────────────────────────────────
    wfGui.Add("Progress", "x15 y68 w565 h2 BackgroundCCCCCC -Smooth Disabled")

    ; ── ListView ──────────────────────────────────────────────────────────────
    wfLV := wfGui.Add("ListView",
        "x15 y75 w565 h330 -Multi Grid NoSortHdr",
        ["Nr", "Nombre de Archivo"])

    ; Poblar lista (mostrando SOLO el nombre del archivo)
    Loop workMaxEntries {
        rutaCom := (workfilesData.Length >= A_Index) ? workfilesData[A_Index] : ""
        nombreCorto := RegExReplace(rutaCom, ".*[\\/]", "")
        wfLV.Add("", A_Index, nombreCorto)
    }

    wfLV.ModifyCol(1, "45 Left")    ; Nr
    wfLV.ModifyCol(2, "500 Left")   ; Nombre de archivo

    ; Seleccionar fila 1 al inicio y deshabilitar caja si está vacío
    if workMaxEntries > 0 {
        wfLV.Modify(1, "Select Focus Vis")
        nombreInit := wfLV.GetText(1, 2)
        RefrescarEstadoTxtFilename(nombreInit)
    } else {
        RefrescarEstadoTxtFilename("")
    }

    ; ── Separador ─────────────────────────────────────────────────────────────
    wfGui.Add("Progress", "x15 y413 w565 h2 BackgroundCCCCCC -Smooth Disabled")

    ; ── Botones centrados ─────────────────────────────────────────────────────
    wfGui.SetFont("s10 c222222", "Segoe UI")
    btnCrear   := wfGui.Add("Button", "x20  y423 w100 h32", "&Crear")
    btnUpdate  := wfGui.Add("Button", "x135 y423 w100 h32", "&Actualizar")
    btnImport  := wfGui.Add("Button", "x250 y423 w100 h32", "&Importar")
    btnDelete  := wfGui.Add("Button", "x365 y423 w100 h32", "&Eliminar")
    btnClose   := wfGui.Add("Button", "x480 y423 w100 h32", "&Volver")

    ; ── Etiqueta de estado (debajo de los botones, centrada) ─────────────────
    wfGui.SetFont("s10 bold c444444", "Segoe UI")
    infoLbl := wfGui.Add("Text", "x15 y465 w565 Center", "")

    ; ── Eventos ───────────────────────────────────────────────────────────────
    wfLV.OnEvent("Click",      ActualizarCamposWF)
    wfLV.OnEvent("ItemSelect", ActualizarCamposWF)
    btnCrear.OnEvent("Click",  CrearWorkfile)
    btnUpdate.OnEvent("Click", ActualizarWorkfile)
    btnImport.OnEvent("Click", ImportarWorkfile)
    btnDelete.OnEvent("Click", EliminarWorkfile)
    btnClose.OnEvent("Click",  CerrarWfGui)
    wfGui.OnEvent("Close",     CerrarWfGui)

    wfGui.Show("Center w600 h495")
    SendMessage(0xB1, -1, 0, txtMaxWF)

    ; Guardado automático del máximo de Workfiles al perder el foco
    txtMaxWF.OnEvent("LoseFocus", ActualizarMaxWF)

    ; ── Función interna: mostrar nombre en txtFilename (siempre ReadOnly) ────────
    RefrescarEstadoTxtFilename(nombre) {
        if nombre = "" {
            txtFilename.Opt("BackgroundF0F0F0 c888888")
            txtFilename.Value := ""
        } else {
            txtFilename.Opt("BackgroundFFFFFF c222222")
            txtFilename.Value := nombre
        }
    }

    ; ── Función interna: guardar máximo de Workfiles al cambiar txtMaxWF ─────────
    ActualizarMaxWF(*) {
        maxStr := Trim(txtMaxWF.Value)
        if !RegExMatch(maxStr, "^\d+$") || Integer(maxStr) < 0 || Integer(maxStr) > 32
            return
        nuevoMax := Integer(maxStr)
        if nuevoMax = workMaxEntries
            return
        EscribirMaxWorkfilesEnNATPARM(nuevoMax)
        workMaxEntries := nuevoMax
        while workfilesData.Length < workMaxEntries
            workfilesData.Push("")
        wfLV.Delete()
        Loop workMaxEntries {
            rutaCom := workfilesData[A_Index]
            nombreList := RegExReplace(rutaCom, ".*[\\/]", "")
            wfLV.Add("", A_Index, nombreList)
        }
        wfLV.ModifyCol(1, "45 Left")
        wfLV.ModifyCol(2, "500 Left")
        if workMaxEntries > 0 {
            wfLV.Modify(1, "Select Focus Vis")
            txtNumber.Value := 1
            nombreInit := RegExReplace(workfilesData[1], ".*[\\/]", "")
            RefrescarEstadoTxtFilename(nombreInit)
        } else {
            txtNumber.Value := ""
            RefrescarEstadoTxtFilename("")
        }
        infoLbl.Value := "✓ Máximo de Workfiles actualizado a " . nuevoMax
        infoLbl.Opt("c006600")
    }

    ; ── Función interna: cerrar subventana y restaurar ventana principal ───────
    CerrarWfGui(*) {
        wfGui.Destroy()
        mainGui.Show()
    }

    ; ── Función interna: actualizar campos al hacer clic en una fila ──────────
    ActualizarCamposWF(*) {
        fila := wfLV.GetNext(0, "Focused")
        if fila = 0
            fila := wfLV.GetNext(0)
        if fila = 0
            return
            
        txtNumber.Value   := fila
        nombreArchivo := wfLV.GetText(fila, 2)
        RefrescarEstadoTxtFilename(nombreArchivo)
    }

    ; ── Función interna: crear Workfile (MODAL SIN PARPADEO) ────────────────────
    CrearWorkfile(*) {
        fila := wfLV.GetNext(0, "Focused")
        if fila = 0
            fila := wfLV.GetNext(0)
        if fila = 0 {
            MsgBox("Seleccione primero un Workfile de la lista.", "Crear Workfile", "IconX")
            return
        }

        ; Bloquear si el slot ya tiene un Workfile asignado
        nombreEnSlot := wfLV.GetText(fila, 2)
        if nombreEnSlot != "" {
            MsgBox("El Workfile #" . fila . " ya tiene asignado el archivo:`n`n"
                 . "    " . nombreEnSlot . "`n`n"
                 . "Use 'Actualizar' para cambiar el nombre o 'Eliminar' para vaciarlo primero.",
                 "Slot ocupado", "Icon!")
            return
        }

        crearDlg := Gui("-MaximizeBox +MinimizeBox", "Crear nuevo Workfile")
        crearDlg.BackColor := "F0F0F0"
        crearDlg.SetFont("s11 cFFFFFF bold", "Segoe UI")
        crearDlg.Add("Text", "x0 y0 w380 h28 Background1A237E", "")
        crearDlg.Add("Text", "x10 y6 w360 Background1A237E cFFFFFF", "Workfile #" . fila . " — Crear archivo nuevo")

        crearDlg.SetFont("s10 c444444", "Segoe UI")
        crearDlg.Add("Text", "x15 y42 w350", "Nombre del archivo (Ej: WORK1.TXT):")

        crearDlg.SetFont("s11 c222222", "Segoe UI")
        txtNombreNuevo := crearDlg.Add("Edit", "x15 y62 w350 h24 Border Uppercase")

        crearDlg.SetFont("s8 c777777", "Segoe UI")
        crearDlg.Add("Text", "x15 y90 w350", "Formato: máx. 8 caracteres + extensión obligatoria de 3 (no .SAG)")

        crearDlg.SetFont("s10 c222222", "Segoe UI")
        btnDlgOK     := crearDlg.Add("Button", "x15  y118 w160 h28 Default", "✔  Crear")
        btnDlgCancel := crearDlg.Add("Button", "x205 y118 w160 h28",         "✖  Cancelar")
        btnDlgOK.Opt("+Background1A237E")
        btnDlgOK.SetFont("s10 bold cFFFFFF", "Segoe UI")

        resultado := ""

        ; Filtrado en tiempo real: formato 8.3
        txtNombreNuevo.OnEvent("Change", FiltrarNombre83)

        ; Función de cierre para evitar parpadeos: reactivar ventana principal ANTES de ocultar la subventana
        CerrarCrear(*) {
            wfGui.Opt("-Disabled")
            WinActivate("ahk_id " wfGui.Hwnd)
            crearDlg.Hide()
        }
    
        btnDlgOK.OnEvent("Click",    (*) => (resultado := "OK", CerrarCrear()))
        btnDlgCancel.OnEvent("Click", CerrarCrear)
        crearDlg.OnEvent("Close",    CerrarCrear)

        wfGui.Opt("+Disabled") 
        crearDlg.Show("Center w380 h160")
        
        WinWaitClose(crearDlg)
        valorNombreNuevo := txtNombreNuevo.Value
        crearDlg.Destroy()

        if resultado != "OK" || Trim(valorNombreNuevo) = ""
            return

        nombreNuevo := StrUpper(Trim(valorNombreNuevo))

        errorValidacion := WF_ValidarSegmento83(nombreNuevo)
        if errorValidacion != "" {
            MsgBox("El nombre del archivo no cumple las reglas de validación de Windows 3.1:`n`n" . errorValidacion, "Error - Nombre inválido", "IconX")
            return
        }

        ; ── Validar que no exista ya un Workfile con el mismo nombre en otro slot ──
        Loop workMaxEntries {
            if A_Index = fila
                continue
            nombreExistente := wfLV.GetText(A_Index, 2)
            if (nombreExistente != "" && StrUpper(nombreExistente) = StrUpper(nombreNuevo)) {
                MsgBox(
                    "No se puede crear el Workfile porque ya existe uno con el mismo nombre:`n`n"
                    . "    Slot #" . A_Index . " → " . nombreExistente . "`n`n"
                    . "Cada Workfile debe tener un nombre único en la lista.",
                    "Nombre duplicado — Creación cancelada", "IconX")
                return
            }
        }

        carpetaWF := A_ScriptDir . "\dos\NATURAL\" . versionNatural . "\WF"
        if !DirExist(carpetaWF) {
            try {
                DirCreate(carpetaWF)
            } catch as errDir {
                MsgBox("No se pudo crear la carpeta destino:`n" . carpetaWF . "`n`n" . errDir.Message, "Error", "IconX")
                return
            }
        }

        rutaDestino := carpetaWF . "\" . nombreNuevo
        if FileExist(rutaDestino) {
            respuesta := MsgBox("El archivo físico ya existe en la carpeta WF.`n`n¿Desea enlazarlo de todos modos sin sobrescribir su contenido?", "Archivo existente", "YesNo Icon?")
            if respuesta = "No"
                return
        } else {
            try {
                FileAppend("", rutaDestino)
            } catch as errFile {
                MsgBox("Error al crear el archivo físico:`n" . errFile.Message, "Error", "IconX")
                return
            }
        }

        rutaCompleta := "C:\NATURAL\" . versionNatural . "\WF\" . nombreNuevo

        if workfileLengths.Length >= fila && workfileLengths[fila] > 0 {
            if StrLen(rutaCompleta) + 1 > workfileLengths[fila] {
                MsgBox("La ruta completa excede la longitud máxima permitida para este slot.`n`n"
                     . "Máximo: " . (workfileLengths[fila] - 1) . " caracteres`n"
                     . "Ruta: " . rutaCompleta . " (" . StrLen(rutaCompleta) . " caracteres)",
                     "Error - Ruta demasiado larga", "IconX")
                return
            }
        }

        wfGuardado := false
        if workfileOffsets.Length >= fila && workfileOffsets[fila] > 0 {
            wfGuardado := EscribirWorkfileEnNATPARM(fila, rutaCompleta)
        } else {
            MsgBox("No se encontró el offset para el Workfile #" . fila . " en NATPARM.SAG.`n"
                 . "No se actualizará la entrada.",
                 "Advertencia", "Icon!")
        }

        while workfilesData.Length < fila
            workfilesData.Push("")
        workfilesData[fila] := rutaCompleta

        wfLV.Modify(fila, "Col2", nombreNuevo)
        txtNumber.Value := fila
        RefrescarEstadoTxtFilename(nombreNuevo)

        if wfGuardado {
            infoLbl.Value := "✓ Workfile #" . fila . " creado → " . nombreNuevo
            infoLbl.Opt("c006600")
        } else {
            infoLbl.Value := "⚠ Archivo creado en WF\ pero no se actualizó NATPARM.SAG"
            infoLbl.Opt("c996600")
        }
    }

    ; ── Función interna: importar Workfile (MODAL SIN PARPADEO) ─────────────────
    ImportarWorkfile(*) {
        fila := wfLV.GetNext(0, "Focused")
        if fila = 0
            fila := wfLV.GetNext(0)
        if fila = 0 {
            MsgBox("Seleccione primero un Workfile de la lista.", "Importar Workfile", "IconX")
            return
        }
        if fila > workMaxEntries {
            MsgBox("El número de fila (" . fila . ") supera el máximo (" . workMaxEntries . ").", "Importar Workfile", "IconX")
            return
        }

        ; Bloquear si el slot ya tiene un Workfile asignado
        nombreEnSlot := wfLV.GetText(fila, 2)
        if nombreEnSlot != "" {
            MsgBox("El Workfile #" . fila . " ya tiene asignado el archivo:`n`n"
                 . "    " . nombreEnSlot . "`n`n"
                 . "Use 'Actualizar' para cambiar el nombre o 'Eliminar' para vaciarlo primero.",
                 "Slot ocupado", "Icon!")
            return
        }

        carpetaWF := A_ScriptDir . "\dos\NATURAL\" . versionNatural . "\WF"

        ; ── Recopilar archivos válidos en .\WF ───────────────────────────────
        archivosWF := []
        if DirExist(carpetaWF) {
            Loop Files, carpetaWF . "\*.*" {
                if WF_ValidarSegmento83(A_LoopFileName) = ""
                    archivosWF.Push(A_LoopFileName)
            }
        }

        impDlg := Gui("-MaximizeBox +MinimizeBox", "Importar Workfile")
        impDlg.BackColor := "F0F0F0"
        impDlg.SetFont("s11 cFFFFFF bold", "Segoe UI")
        impDlg.Add("Text", "x0 y0 w340 h28 Background1A237E", "")
        impDlg.Add("Text", "x10 y6 w320 Background1A237E cFFFFFF", "Workfile #" . fila . " — Seleccionar origen")

        impDlg.SetFont("s10 bold c222222", "Segoe UI")
        impDlg.Add("Text", "x15 y44 w310", "Archivos en carpeta de Workfiles:")

        hayArchivos := archivosWF.Length > 0
        if hayArchivos {
            lbArchivos := impDlg.Add("ListView", "x15 y64 w310 h150 -Multi -Hdr Grid NoSortHdr", ["Nombre"])
            for archivo in archivosWF
                lbArchivos.Add("", archivo)
            lbArchivos.ModifyCol(1, "290 Left")
            lbArchivos.Modify(1, "Select Focus Vis")
            impDlg.SetFont("s10 c222222", "Segoe UI")
            btnEnlazar := impDlg.Add("Button", "x15 y226 w310 h28 Default", "✔  Importar archivo seleccionado")
            btnEnlazar.Opt("+Background1A237E")
            btnEnlazar.SetFont("s10 bold cFFFFFF", "Segoe UI")
        } else {
            impDlg.SetFont("s10 c888888 italic", "Segoe UI")
            impDlg.Add("Text", "x15 y64 w310 h150", "(No hay archivos en la carpeta .\WF)")
            lbArchivos := ""
            impDlg.SetFont("s10 c222222", "Segoe UI")
            btnEnlazar := impDlg.Add("Button", "x15 y226 w310 h28", "✔  Importar archivo seleccionado")
            btnEnlazar.Opt("Disabled")
            btnEnlazar.SetFont("s10 cFFFFFF", "Segoe UI")
        }

        impDlg.Add("Progress", "x15 y268 w310 h2 BackgroundCCCCCC -Smooth Disabled")

        impDlg.SetFont("s10 bold c222222", "Segoe UI")
        impDlg.Add("Text", "x15 y282 w310", "Importar desde otra ubicación:")
        impDlg.SetFont("s10 c222222", "Segoe UI")
        btnExplorar := impDlg.Add("Button", "x15 y304 w310 h28", "📂  Explorar sistema de archivos...")
        btnExplorar.Opt("+Background607D8B")
        btnExplorar.SetFont("s10 cFFFFFF", "Segoe UI")

        impDlg.Add("Progress", "x15 y346 w310 h2 BackgroundCCCCCC -Smooth Disabled")

        impDlg.SetFont("s10 c222222", "Segoe UI")
        btnImpCancel := impDlg.Add("Button", "x15 y360 w310 h28", "✖  Cancelar")

        modoResultado := ""

        ; Función de cierre para evitar parpadeos
        CerrarImp(*) {
            wfGui.Opt("-Disabled")
            WinActivate("ahk_id " wfGui.Hwnd)
            impDlg.Hide()
        }

        btnEnlazar.OnEvent("Click",   (*) => (modoResultado := "local",   CerrarImp()))
        btnExplorar.OnEvent("Click",  (*) => (modoResultado := "externo", CerrarImp()))
        btnImpCancel.OnEvent("Click", CerrarImp)
        impDlg.OnEvent("Close",       CerrarImp)

        wfGui.Opt("+Disabled")
        impDlg.Show("Center w340 h402")
        
        WinWaitClose(impDlg)
        nombreLocal := (lbArchivos != "") ? lbArchivos.GetText(lbArchivos.GetNext(0, "Focused"), 1) : ""
        impDlg.Destroy()

        if modoResultado = ""
            return

        ; ── Modo local: enlace directo desde .\WF ────────────────────────────
        if modoResultado = "local" {
            if nombreLocal = ""
                return
            nombreArchivo := StrUpper(nombreLocal)
            errorValidacion := WF_ValidarSegmento83(nombreArchivo)
            if errorValidacion != "" {
                MsgBox("Nombre inválido:`n`n" . errorValidacion, "Error", "IconX")
                return
            }
            ProcesarImport(fila, nombreArchivo, "local")
            return
        }

        ; ── Modo externo: explorador de archivos ──────────────────────────────
        archivoOrigen := FileSelect(1, "", "Seleccionar Workfile para importar", "Todos los archivos (*.*)")
        if archivoOrigen = ""
            return

        nombreArchivo := StrUpper(RegExReplace(archivoOrigen, ".*[\\/]", ""))

        errorValidacion := WF_ValidarSegmento83(nombreArchivo)
        if errorValidacion != "" {
            MsgBox("El nombre del archivo no cumple las reglas de validación de Windows 3.1:`n`n"
                 . errorValidacion . "`n`nNombre: " . nombreArchivo,
                 "Error - Nombre de archivo inválido", "IconX")
            return
        }

        ; Si el archivo ya está en .\WF → enlace directo, sin copiar
        rutaEnWF := carpetaWF . "\" . nombreArchivo
        if archivoOrigen = rutaEnWF {
            ProcesarImport(fila, nombreArchivo, "local")
            return
        }

        if !DirExist(carpetaWF) {
            try {
                DirCreate(carpetaWF)
            } catch as errDir {
                MsgBox("No se pudo crear la carpeta destino:`n" . carpetaWF . "`n`n" . errDir.Message, "Error", "IconX")
                return
            }
        }

        if FileExist(rutaEnWF) {
            respuesta := MsgBox("El archivo ya existe en la carpeta WF.`n`n¿Desea sobrescribirlo?",
                                "Archivo existente", "YesNo Icon?")
            if respuesta = "No"
                return
        }

        try {
            FileCopy(archivoOrigen, rutaEnWF, 1)
        } catch as errCopy {
            MsgBox("Error al copiar el archivo:`n" . errCopy.Message, "Error", "IconX")
            return
        }

        ProcesarImport(fila, nombreArchivo, "externo")

        ProcesarImport(fila, nombreArchivo, modo) {
            ; ── Validar que no exista ya un Workfile con el mismo nombre en otro slot ──
            Loop workMaxEntries {
                if A_Index = fila
                    continue
                nombreExistente := wfLV.GetText(A_Index, 2)
                if (nombreExistente != "" && StrUpper(nombreExistente) = StrUpper(nombreArchivo)) {
                    MsgBox(
                        "No se puede importar el Workfile porque ya existe un Workfile con el mismo nombre:`n`n"
                        . "    Slot #" . A_Index . " → " . nombreExistente . "`n`n"
                        . "Cada Workfile debe tener un nombre único en la lista.",
                        "Nombre duplicado — Importación cancelada", "IconX")
                    return
                }
            }

            nombreActualEnSlot := wfLV.GetText(fila, 2)
            if nombreActualEnSlot != "" {
                respSobrescribir := MsgBox(
                    "El Workfile #" . fila . " ya tiene asignado el archivo:`n`n"
                    . "    " . nombreActualEnSlot . "`n`n"
                    . "La entrada en NATPARM.SAG será sobrescrita con:`n`n"
                    . "    " . nombreArchivo . "`n`n"
                    . "¿Desea continuar?",
                    "Workfile existente — Confirmar sobrescritura", "YesNo Icon!")
                if respSobrescribir = "No"
                    return
            }

            rutaCompleta := "C:\NATURAL\" . versionNatural . "\WF\" . nombreArchivo

            if workfileLengths.Length >= fila && workfileLengths[fila] > 0 {
                if StrLen(rutaCompleta) + 1 > workfileLengths[fila] {
                    MsgBox("La ruta completa excede la longitud máxima permitida para este slot.`n`n"
                         . "Máximo: " . (workfileLengths[fila] - 1) . " caracteres`n"
                         . "Ruta: " . rutaCompleta . " (" . StrLen(rutaCompleta) . " caracteres)",
                         "Error - Ruta demasiado larga", "IconX")
                    return
                }
            }

            wfGuardado := false
            if workfileOffsets.Length >= fila && workfileOffsets[fila] > 0 {
                wfGuardado := EscribirWorkfileEnNATPARM(fila, rutaCompleta)
            } else {
                MsgBox("No se encontró el offset para el Workfile #" . fila . " en NATPARM.SAG.`n"
                     . "No se actualizará la entrada.",
                     "Advertencia", "Icon!")
            }

            while workfilesData.Length < fila
                workfilesData.Push("")
            workfilesData[fila] := rutaCompleta

            wfLV.Modify(fila, "Col2", nombreArchivo)
            txtNumber.Value := fila
            RefrescarEstadoTxtFilename(nombreArchivo)

            if wfGuardado {
                etiqueta := "importado"
                infoLbl.Value := "✓ Workfile #" . fila . " " . etiqueta . " → " . nombreArchivo
                infoLbl.Opt("c006600")
            } else {
                infoLbl.Value := "⚠ No se actualizó NATPARM.SAG"
                infoLbl.Opt("c996600")
            }
        }
    }

    ; ── Función interna: actualizar Workfile (subventana modal) ───────────────
    ActualizarWorkfile(*) {
        fila := wfLV.GetNext(0, "Focused")
        if fila = 0
            fila := wfLV.GetNext(0)
        if fila = 0 {
            MsgBox("Seleccione primero un Workfile de la lista.", "Actualizar Workfile", "IconX")
            return
        }

        nombreActualEnLista := wfLV.GetText(fila, 2)
        if nombreActualEnLista = "" {
            MsgBox("El Workfile #" . fila . " está vacío. Use 'Crear' o 'Importar' para asignarlo.", "Workfile vacío", "Icon!")
            return
        }

        actuDlg := Gui("-MaximizeBox +MinimizeBox", "Actualizar Workfile")
        actuDlg.BackColor := "F0F0F0"
        actuDlg.SetFont("s11 cFFFFFF bold", "Segoe UI")
        actuDlg.Add("Text", "x0 y0 w380 h28 Background1A237E", "")
        actuDlg.Add("Text", "x10 y6 w360 Background1A237E cFFFFFF", "Workfile #" . fila . " — Actualizar nombre")

        actuDlg.SetFont("s10 c444444", "Segoe UI")
        actuDlg.Add("Text", "x15 y38 w350", "Nombre actual:")
        actuDlg.SetFont("s10 bold c1A237E", "Segoe UI")
        actuDlg.Add("Text", "x15 y56 w350", nombreActualEnLista)

        actuDlg.SetFont("s10 c444444", "Segoe UI")
        actuDlg.Add("Text", "x15 y82 w350", "Nuevo nombre (Ej: WORK1.TXT):")

        actuDlg.SetFont("s11 c222222", "Segoe UI")
        txtNombreActu := actuDlg.Add("Edit", "x15 y102 w350 h24 Border Uppercase")
        txtNombreActu.Value := nombreActualEnLista

        actuDlg.SetFont("s8 c777777", "Segoe UI")
        actuDlg.Add("Text", "x15 y130 w350", "Formato: máx. 8 caracteres + extensión obligatoria de 3 (no .SAG)")

        actuDlg.SetFont("s10 c222222", "Segoe UI")
        btnActuOK     := actuDlg.Add("Button", "x15  y155 w160 h28 Default", "✔  Actualizar")
        btnActuCancel := actuDlg.Add("Button", "x205 y155 w160 h28",         "✖  Cancelar")
        btnActuOK.Opt("+Background1A237E")
        btnActuOK.SetFont("s10 bold cFFFFFF", "Segoe UI")

        ; Filtrado en tiempo real: formato 8.3
        txtNombreActu.OnEvent("Change", FiltrarNombre83)

        resultadoActu := ""

        CerrarActu(*) {
            wfGui.Opt("-Disabled")
            WinActivate("ahk_id " wfGui.Hwnd)
            actuDlg.Hide()
        }

        btnActuOK.OnEvent("Click",    (*) => (resultadoActu := "OK", CerrarActu()))
        btnActuCancel.OnEvent("Click", CerrarActu)
        actuDlg.OnEvent("Close",      CerrarActu)

        wfGui.Opt("+Disabled")
        actuDlg.Show("Center w380 h196")

        WinWaitClose(actuDlg)
        valorNombreActu := StrUpper(Trim(txtNombreActu.Value))
        actuDlg.Destroy()

        if resultadoActu != "OK" || valorNombreActu = ""
            return

        ; ── Validación formato 8.3 ────────────────────────────────────────────
        errorValidacion := WF_ValidarSegmento83(valorNombreActu)
        if errorValidacion != "" {
            MsgBox("El nombre no cumple el formato requerido:`n`n" . errorValidacion, "Error — Nombre inválido", "IconX")
            return
        }

        ; ── Sin cambio ────────────────────────────────────────────────────────
        if StrUpper(valorNombreActu) = StrUpper(nombreActualEnLista)
            return

        ; ── Validar duplicado en lista ────────────────────────────────────────
        Loop workMaxEntries {
            if A_Index = fila
                continue
            nombreExistente := wfLV.GetText(A_Index, 2)
            if (nombreExistente != "" && StrUpper(nombreExistente) = StrUpper(valorNombreActu)) {
                MsgBox(
                    "Ya existe un Workfile con ese nombre:`n`n"
                    . "    Slot #" . A_Index . " → " . nombreExistente . "`n`n"
                    . "Cada Workfile debe tener un nombre único.",
                    "Nombre duplicado — Actualización cancelada", "IconX")
                return
            }
        }

        ; ── Validar que no exista ya en la carpeta WF\ ───────────────────────
        carpetaWF  := A_ScriptDir . "\dos\NATURAL\" . versionNatural . "\WF"
        rutaNueva  := carpetaWF . "\" . valorNombreActu
        if FileExist(rutaNueva) {
            MsgBox(
                "Ya existe un archivo físico con ese nombre en la carpeta WF:`n`n"
                . "    " . valorNombreActu . "`n`n"
                . "Por favor, elija un nombre distinto.",
                "Archivo existente — Actualización cancelada", "IconX")
            return
        }

        ; ── Renombrar archivo físico si existe ────────────────────────────────
        rutaAnterior := carpetaWF . "\" . nombreActualEnLista
        if FileExist(rutaAnterior) {
            try {
                FileMove(rutaAnterior, rutaNueva, 0)
            } catch as errMv {
                MsgBox("No se pudo renombrar el archivo físico:`n" . errMv.Message
                     . "`n`nNo se actualizará NATPARM.SAG para mantener la integridad.",
                     "Error crítico", "IconX")
                return
            }
        }

        ; ── Actualizar NATPARM.SAG ────────────────────────────────────────────
        rutaCompleta := "C:\NATURAL\" . versionNatural . "\WF\" . valorNombreActu

        longitudOrig := (workfileLengths.Length >= fila && workfileLengths[fila] > 0) ? workfileLengths[fila] : 51
        if StrLen(rutaCompleta) > longitudOrig {
            MsgBox("La ruta resultante excede la longitud máxima del slot (" . longitudOrig . " caracteres).",
                   "Error — Nombre demasiado largo", "IconX")
            return
        }

        wfGuardado := false
        if workfileOffsets.Length >= fila && workfileOffsets[fila] > 0
            wfGuardado := EscribirWorkfileEnNATPARM(fila, rutaCompleta)

        while workfilesData.Length < fila
            workfilesData.Push("")
        workfilesData[fila] := rutaCompleta

        wfLV.Modify(fila, "", fila, valorNombreActu)
        RefrescarEstadoTxtFilename(valorNombreActu)

        if wfGuardado {
            infoLbl.Value := "✓ Workfile #" . fila . " actualizado → " . valorNombreActu
            infoLbl.Opt("c006600")
        } else {
            infoLbl.Value := "⚠ Archivo renombrado pero no se actualizó NATPARM.SAG"
            infoLbl.Opt("c996600")
        }
    }

    ; ── Función interna: eliminar Workfile (subventana modal) ────────────────
    EliminarWorkfile(*) {
        fila := wfLV.GetNext(0, "Focused")
        if fila = 0
            fila := wfLV.GetNext(0)
        if fila = 0 {
            MsgBox("Seleccione primero un Workfile de la lista.", "Eliminar Workfile", "IconX")
            return
        }

        nombreActual := wfLV.GetText(fila, 2)
        if nombreActual = "" {
            MsgBox("El Workfile #" . fila . " ya está vacío.", "Eliminar Workfile", "Icon!")
            return
        }

        carpetaWF := A_ScriptDir . "\dos\NATURAL\" . versionNatural . "\WF"
        rutaFisica := carpetaWF . "\" . nombreActual
        existeFisico := FileExist(rutaFisica) ? true : false

        elimDlg := Gui("-MaximizeBox +MinimizeBox", "Eliminar Workfile")
        elimDlg.BackColor := "F0F0F0"

        ; Barra de título interna
        elimDlg.SetFont("s11 cFFFFFF bold", "Segoe UI")
        elimDlg.Add("Text", "x0 y0 w390 h28 Background1A237E", "")
        elimDlg.Add("Text", "x10 y6 w370 Background1A237E cFFFFFF", "Workfile #" . fila . " — Eliminar")

        ; Info del workfile seleccionado
        elimDlg.SetFont("s10 bold c222222", "Segoe UI")
        elimDlg.Add("Text", "x15 y38 w360", "Workfile seleccionado:")
        elimDlg.SetFont("s10 c222222", "Segoe UI")
        elimDlg.Add("Text", "x15 y56 w30", "#" . fila)
        elimDlg.SetFont("s10 bold c1A237E", "Segoe UI")
        elimDlg.Add("Text", "x50 y56 w325", nombreActual)

        ; Separador
        elimDlg.Add("Progress", "x15 y74 w360 h2 BackgroundCCCCCC -Smooth Disabled")

        ; ── Opción 1: Eliminación LÓGICA ─────────────────────────────────────
        elimDlg.SetFont("s9 bold cB8860B", "Segoe UI")
        elimDlg.Add("Text", "x15 y84 w360",
            "⚠  El archivo físico NO se elimina del sistema.")

        elimDlg.SetFont("s10 c222222", "Segoe UI")
        btnLogico := elimDlg.Add("Button", "x15 y104 w360 h28", "🗒  Eliminación lógica")
        btnLogico.Opt("+Background37474F")
        btnLogico.SetFont("s10 bold cFFFFFF", "Segoe UI")

        ; Separador
        elimDlg.Add("Progress", "x15 y142 w360 h2 BackgroundCCCCCC -Smooth Disabled")

        ; ── Opción 2: Eliminación FÍSICA ─────────────────────────────────────
        elimDlg.SetFont("s9 bold cB00000", "Segoe UI")
        elimDlg.Add("Text", "x15 y152 w360",
            "⛔  Esta acción es irreversible. El archivo físico se perderá.")

        elimDlg.SetFont("s10 c222222", "Segoe UI")
        btnFisico := elimDlg.Add("Button", "x15 y172 w360 h28", "🗑  Eliminación física")
        btnFisico.Opt("+BackgroundB00000")
        btnFisico.SetFont("s10 bold cFFFFFF", "Segoe UI")
        ; Deshabilitar eliminación física si no existe el archivo
        if !existeFisico {
            btnFisico.Opt("Disabled")
        }

        ; Separador
        elimDlg.Add("Progress", "x15 y210 w360 h2 BackgroundCCCCCC -Smooth Disabled")

        ; Botón Cancelar
        elimDlg.SetFont("s10 c222222", "Segoe UI")
        btnElimCancel := elimDlg.Add("Button", "x15 y220 w360 h28", "✖  Cancelar")

        tipoElim := ""

        ; Función de cierre sin parpadeo
        CerrarElim(*) {
            wfGui.Opt("-Disabled")
            WinActivate("ahk_id " wfGui.Hwnd)
            elimDlg.Hide()
        }

        btnLogico.OnEvent("Click",     (*) => (tipoElim := "logica",   CerrarElim()))
        btnFisico.OnEvent("Click",     (*) => (tipoElim := "fisica",   CerrarElim()))
        btnElimCancel.OnEvent("Click", CerrarElim)
        elimDlg.OnEvent("Close",       CerrarElim)

        wfGui.Opt("+Disabled")
        elimDlg.Show("Center w390 h262")

        WinWaitClose(elimDlg)
        elimDlg.Destroy()

        if tipoElim = ""
            return

        ; ── Eliminación LÓGICA ────────────────────────────────────────────────
        if tipoElim = "logica" {
            respConf := MsgBox(
                "¿Confirmar eliminación lógica del Workfile #" . fila . "?`n`n"
                . "Nombre: " . nombreActual . "`n`n"
                . "El archivo físico NO será eliminado.",
                "Confirmar eliminación lógica", "YesNo Icon?")
            if respConf = "No"
                return

            limpiado := false
            if workfileOffsets.Length >= fila && workfileOffsets[fila] > 0
                limpiado := EscribirWorkfileEnNATPARM(fila, "")

            while workfilesData.Length < fila
                workfilesData.Push("")
            workfilesData[fila] := ""

            wfLV.Modify(fila, "Col2", "")
            RefrescarEstadoTxtFilename("")

            if limpiado {
                infoLbl.Value := "✓ Workfile #" . fila . " eliminado lógicamente"
                infoLbl.Opt("c006600")
            } else {
                infoLbl.Value := "⚠ No se pudo limpiar la entrada en NATPARM.SAG"
                infoLbl.Opt("c996600")
            }
            return
        }

        ; ── Eliminación FÍSICA ────────────────────────────────────────────────
        if tipoElim = "fisica" {
            respConf := MsgBox(
                "¿Confirmar eliminación física del Workfile #" . fila . "?`n`n"
                . "Nombre: " . nombreActual . "`n`n"
                . "Se eliminará permanentemente el archivo.",
                "Confirmar eliminación física", "YesNo Icon!")
            if respConf = "No"
                return

            ; Eliminar archivo físico
            if FileExist(rutaFisica) {
                try {
                    FileDelete(rutaFisica)
                } catch as errDel {
                    MsgBox("No se pudo eliminar el archivo físico:`n" . rutaFisica
                         . "`n`n" . errDel.Message
                         . "`n`nNo se modificará NATPARM.SAG para mantener la integridad.",
                         "Error al eliminar archivo", "IconX")
                    return
                }
            }

            ; Limpiar entrada en NATPARM.SAG
            limpiado := false
            if workfileOffsets.Length >= fila && workfileOffsets[fila] > 0
                limpiado := EscribirWorkfileEnNATPARM(fila, "")

            while workfilesData.Length < fila
                workfilesData.Push("")
            workfilesData[fila] := ""

            wfLV.Modify(fila, "Col2", "")
            RefrescarEstadoTxtFilename("")

            if limpiado {
                infoLbl.Value := "✓ Workfile #" . fila . " eliminado físicamente"
                infoLbl.Opt("c800000")
            } else {
                infoLbl.Value := "⚠ Archivo físico eliminado pero no se limpió NATPARM.SAG"
                infoLbl.Opt("c996600")
            }
        }
    }
}

; ===== FUNCIÓN: OBTENER RUTA DE DESTINO SEGÚN VERSIÓN =====
ObtenerRutaDestino() {
    global versionNatural, libreriaDefecto
    
    rutaBase := ".\dos\NATAPPS"
    
    ; Determinar estructura según versión
    if (versionNatural = "210" || versionNatural = "211" || versionNatural = "212") {
        ; Versiones 2.1.0, 2.1.1, 2.1.2: guardar directamente en NATAPPS
        return rutaBase . "\" . libreriaDefecto
    } else {
        ; Versiones 2.1.3+: usar estructura FUSER\SRC
        return rutaBase . "\FUSER\" . libreriaDefecto . "\SRC"
    }
}

; ===== FUNCIÓN: ACTUALIZAR DESTINO =====
ActualizarDestino() {
    global txtLibreria, destinoBase, libreriaDefecto, mainGui, versionNatural
    
    ; Verificar si la ventana principal está minimizada o inactiva
    try {
        if !WinActive("ahk_id " . mainGui.Hwnd)
            return
    }
    
    libreria := Trim(txtLibreria.Value)
    
    ; Si está vacío, restaurar valor por defecto sin mensaje
    if libreria = "" {
        txtLibreria.Value := libreriaDefecto
        destinoBase := ObtenerRutaDestino()
        return
    }
    
    ; Convertir a mayúsculas
    libreria := StrUpper(libreria)
    
    ; Validar reglas de Natural (ya filtrado, pero por si acaso)
    errores := []
    
    ; Verificar que solo contenga caracteres válidos (A-Z, 0-9, -, #)
    if !RegExMatch(libreria, "^[A-Z0-9#-]+$") {
        errores.Push("- Solo se permiten letras, números, guión (-) y numeral (#)")
    }
    
    ; Verificar que empiece con letra o numeral
    if !RegExMatch(libreria, "^[A-Z#]") {
        errores.Push("- Debe empezar con una letra (A-Z) o numeral (#)")
    }
    
    ; Verificar longitud (máximo 8)
    if StrLen(libreria) > 8 {
        errores.Push("- No puede exceder 8 caracteres")
    }
    
    ; Si hay errores, mostrar mensaje y restaurar valor anterior
    if errores.Length > 0 {
        mensaje := "El nombre de librería '" . libreria . "' no es válido:`n`n"
        for error in errores {
            mensaje .= error . "`n"
        }
        mensaje .= "`nPor favor, corrija el nombre según las reglas de Natural."
        MsgBox(mensaje, "Nombre de librería inválido", "IconX 48")
        
        ; Restaurar valor por defecto
        txtLibreria.Value := libreriaDefecto
        destinoBase := ObtenerRutaDestino()
        
        ; Dar foco de nuevo al campo para que el usuario lo corrija
        txtLibreria.Focus()
        return
    }
    
    ; Si pasó todas las validaciones, actualizar
    txtLibreria.Value := libreria
    libreriaDefecto := libreria
    destinoBase := ObtenerRutaDestino()
}

; ===== FUNCIÓN AUXILIAR: Habilitar botón escanear solo si hay ruta Y librería =====
ActualizarEstadoBotonEscanear() {
    global rutaOrigen, txtLibreria, btnEscanear
    tieneRuta     := (rutaOrigen != "")
    tieneLibreria := (Trim(txtLibreria.Value) != "")
    btnEscanear.Enabled := (tieneRuta && tieneLibreria)
}

; ===== FUNCIÓN: SELECCIONAR RUTA PREDEFINIDA =====
SeleccionarRuta(ruta, proyectoDirecto := "") {
    global rutaOrigen, txtRutaActual, esCarpetaPersonalizada
    global nombreCarpetaSeleccionada, proyectoSeleccionado, origenSeleccion

    if DirExist(ruta) {
        esCarpetaPersonalizada := false
        partes := StrSplit(ruta, "\")
        nombreCarpetaSeleccionada := partes.Length > 0 ? partes[partes.Length] : ruta

        if proyectoDirecto != "" {
            ; Viene desde MostrarSeleccionWorkspace — proyecto ya elegido
            proyectoSeleccionado := proyectoDirecto
            rutaOrigen := ruta . "\" . proyectoDirecto
            txtRutaActual.Value := rutaOrigen
            ActualizarEstadoBotonEscanear()
            origenSeleccion := "workspace"
            EscanearArchivos()
        } else {
            ; Flujo normal — mostrar ventana de selección de proyecto
            rutaOrigen := ruta
            MostrarSeleccionProyecto(ruta)
        }
    } else {
        MsgBox("La ruta no existe:`n" . ruta . "`n`n¿Desea seleccionar una carpeta personalizada?", "Error", "IconX 4")
        if MsgBox("", "", "YesNo Icon?") = "Yes"
            SeleccionarRutaPersonalizada()
    }
}

; ===== FUNCIÓN: MOSTRAR VENTANA DE SELECCIÓN DE PROYECTO =====
; ===== FUNCIÓN: MOSTRAR VENTANA DE SELECCIÓN DE WORKSPACE =====
MostrarSeleccionWorkspace() {
    global mainGui, extensionesNatural

    rutaUsuario := "C:\Users\" . A_UserName
    proyectos   := []

    ; Recorrer carpetas de primer nivel: solo "git" (exacto) o que empiecen con "workspace"
    Loop Files, rutaUsuario . "\*", "D" {
        nombreBase := A_LoopFileName
        esGit       := (StrLower(nombreBase) = "git")
        esWorkspace := (SubStr(StrLower(nombreBase), 1, 9) = "workspace")
        if !esGit && !esWorkspace
            continue

        rutaBase := rutaUsuario . "\" . nombreBase

        ; Buscar subcarpetas que contengan al menos un objeto Natural
        Loop Files, rutaBase . "\*", "D" {
            subcarpeta     := A_LoopFileName
            rutaSubcarpeta := rutaBase . "\" . subcarpeta
            tieneObjetos   := false
            Loop Files, rutaSubcarpeta . "\*.*", "R" {
                for ext in extensionesNatural {
                    if InStr(A_LoopFileExt, SubStr(ext, 2)) {
                        tieneObjetos := true
                        break
                    }
                }
                if tieneObjetos
                    break
            }
            if tieneObjetos
                proyectos.Push({nombre: subcarpeta, rutaBase: rutaBase})
        }
    }

    if proyectos.Length = 0 {
        MsgBox("No se encontraron proyectos con objetos Natural en carpetas 'git' o 'workspace*' de:`n" . rutaUsuario, "Sin proyectos", "Icon!")
        return
    }

    mainGui.Hide()

    ; Ordenar array de objetos alfabéticamente por nombre (todos los workspaces combinados)
    n := proyectos.Length
    i := 1
    while i <= n - 1 {
        minIdx := i
        k := i + 1
        while k <= n {
            if StrCompare(proyectos[k].nombre, proyectos[minIdx].nombre, "Locale") < 0
                minIdx := k
            k++
        }
        if minIdx != i {
            temp := proyectos[i]
            proyectos[i] := proyectos[minIdx]
            proyectos[minIdx] := temp
        }
        i++
    }

    wsGui := Gui("-MaximizeBox", "Seleccionar Proyecto")
    wsGui.BackColor := "F8F9FA"

    ; Barra de título interna
    wsGui.SetFont("s13 bold cFFFFFF", "Segoe UI")
    wsGui.Add("Text", "x0 y0 w380 h36 Background0D47A1", "")
    wsGui.Add("Text", "x15 y8 w350 Background0D47A1 cFFFFFF", "Seleccionar Proyecto")

    wsGui.SetFont("s10 c555555", "Segoe UI")
    wsGui.Add("Text", "x15 y50 w350", "Seleccione un proyecto para ver sus objetos Natural:")

    ; ListView — una sola columna con el nombre del proyecto
    wsLV := wsGui.Add("ListView", "x15 y72 w350 h360 -Multi Grid NoSortHdr", ["Proyecto"])
    wsLV.SetFont("s10", "Segoe UI")
    for p in proyectos
        wsLV.Add("", p.nombre)
    wsLV.ModifyCol(1, "330 Left")
    if proyectos.Length > 0
        wsLV.Modify(1, "Select Focus Vis")

    wsGui.SetFont("s10 c222222", "Segoe UI")
    btnWSCancel := wsGui.Add("Button", "x15  y444 w160 h32",        "←  Volver")
    btnWSOK     := wsGui.Add("Button", "x205 y444 w160 h32 Default", "✔  Seleccionar")
    btnWSOK.Opt("+Background0D47A1")
    btnWSOK.SetFont("s10 bold cFFFFFF", "Segoe UI")

    resultadoWS := ""

    CerrarWsGui(*) {
        wsGui.Hide()
        mainGui.Show()
    }

    ConfirmarWorkspace(*) {
        filaWS := wsLV.GetNext(0, "Focused")
        if filaWS = 0
            filaWS := wsLV.GetNext(0)
        if filaWS = 0 {
            MsgBox("Seleccione un proyecto de la lista.", "Sin selección", "IconX")
            return
        }
        resultadoWS := filaWS
        wsGui.Hide()
    }

    btnWSOK.OnEvent("Click", ConfirmarWorkspace)
    btnWSCancel.OnEvent("Click", CerrarWsGui)
    wsGui.OnEvent("Close", CerrarWsGui)

    wsGui.Show("Center w380 h490")
    WinWaitClose(wsGui)
    wsGui.Destroy()

    if resultadoWS = "" {
        mainGui.Show()
        return
    }

    ; Pasar al flujo con la ruta elegida
    proyectoElegido := proyectos[resultadoWS]
    SeleccionarRuta(proyectoElegido.rutaBase, proyectoElegido.nombre)
}

; ===== FUNCIÓN: MOSTRAR VENTANA DE SELECCIÓN DE PROYECTO =====
MostrarSeleccionProyecto(rutaBase) {
    global mainGui, proyectoSeleccionado, rutaOrigen, txtRutaActual

    ; Recopilar solo subcarpetas que contengan al menos un objeto Natural
    proyectos := []
    Loop Files, rutaBase . "\*", "D" {
        carpeta := A_LoopFileName
        tieneObjetos := false
        Loop Files, rutaBase . "\" . carpeta . "\*.*", "R" {
            for ext in extensionesNatural {
                if InStr(A_LoopFileExt, SubStr(ext, 2)) {
                    tieneObjetos := true
                    break
                }
            }
            if tieneObjetos
                break
        }
        if tieneObjetos
            proyectos.Push(carpeta)
    }

    if proyectos.Length = 0 {
        MsgBox("No se encontraron proyectos en:`n" . rutaBase, "Sin proyectos", "Icon!")
        return
    }

    mainGui.Hide()

    projGui := Gui("-MaximizeBox", "Seleccionar Proyecto")
    projGui.BackColor := "F8F9FA"

    ; Barra de título interna
    projGui.SetFont("s13 bold cFFFFFF", "Segoe UI")
    projGui.Add("Text", "x0 y0 w500 h36 Background0D47A1", "")
    projGui.Add("Text", "x15 y8 w470 Background0D47A1 cFFFFFF", "Seleccionar Proyecto")

    projGui.SetFont("s10 c555555", "Segoe UI")
    projGui.Add("Text", "x15 y50 w470", "Seleccione un proyecto para ver sus objetos Natural:")

    ; ListView de proyectos (sin selección múltiple con SHIFT — se gestiona con hook)
    projLV := projGui.Add("ListView", "x15 y72 w470 h350 -Multi Grid NoSortHdr", ["Proyecto"])
    projLV.SetFont("s10", "Segoe UI")
    for proyecto in proyectos
        projLV.Add("", proyecto)
    projLV.ModifyCol(1, "450 Left")
    if proyectos.Length > 0
        projLV.Modify(1, "Select Focus Vis")

    projGui.SetFont("s10 c222222", "Segoe UI")
    btnProjCancel := projGui.Add("Button", "x15  y434 w225 h32",        "✖  Cancelar")
    btnProjOK     := projGui.Add("Button", "x260 y434 w225 h32 Default", "✔  Ver objetos del proyecto")
    btnProjOK.Opt("+Background0D47A1")
    btnProjOK.SetFont("s10 bold cFFFFFF", "Segoe UI")

    resultadoProj := ""

    CerrarProjGui(*) {
        projGui.Hide()
        mainGui.Show()
    }

    ConfirmarProyecto(*) {
        filaProj := projLV.GetNext(0, "Focused")
        if filaProj = 0
            filaProj := projLV.GetNext(0)
        if filaProj = 0 {
            MsgBox("Seleccione un proyecto de la lista.", "Sin selección", "IconX")
            return
        }
        resultadoProj := projLV.GetText(filaProj, 1)
        projGui.Hide()
    }

    btnProjOK.OnEvent("Click", ConfirmarProyecto)
    btnProjCancel.OnEvent("Click", CerrarProjGui)
    projGui.OnEvent("Close", CerrarProjGui)

    projGui.Show("Center w500 h480")
    WinWaitClose(projGui)
    projGui.Destroy()

    if resultadoProj = "" {
        mainGui.Show()
        return
    }

    ; Guardar proyecto y actualizar ruta de origen a la subcarpeta del proyecto
    proyectoSeleccionado := resultadoProj
    rutaOrigen := rutaBase . "\" . resultadoProj
    txtRutaActual.Value := rutaOrigen
    ActualizarEstadoBotonEscanear()

    ; Escanear directamente
    EscanearArchivos()
}

; ===== FUNCIÓN: SELECCIONAR RUTA PERSONALIZADA =====
SeleccionarRutaPersonalizada() {
    global rutaOrigen, txtRutaActual, btnEscanear, esCarpetaPersonalizada
    global nombreCarpetaSeleccionada, proyectoSeleccionado, origenSeleccion

    carpeta := DirSelect("*", 3, "Seleccione la carpeta de origen (workspace110 o git)")

    if carpeta != "" {
        rutaOrigen := carpeta
        txtRutaActual.Value := carpeta
        ActualizarEstadoBotonEscanear()
        esCarpetaPersonalizada := true

        partes := StrSplit(carpeta, "\")
        nombreCarpetaSeleccionada := partes.Length > 0 ? partes[partes.Length] : carpeta
        proyectoSeleccionado := nombreCarpetaSeleccionada
        origenSeleccion := "personalizada"

        EscanearArchivos()
    }
}

; ===== FUNCIÓN: ESCANEAR ARCHIVOS =====
EscanearArchivos() {
    global archivosEncontrados, rutaOrigen, mainGui, txtLibreria, proyectoSeleccionado
    
    if rutaOrigen = "" {
        MsgBox("Por favor, seleccione primero una carpeta de origen.", "Error", "IconX")
        return
    }
    
    if Trim(txtLibreria.Value) = "" {
        MsgBox("Por favor, ingrese una librería de destino antes de escanear.", "Librería requerida", "IconX")
        txtLibreria.Focus()
        return
    }
    
    ; Limpiar array
    archivosEncontrados := []
    
    ; Mostrar progreso
    mainGui.Hide()
    progressGui := Gui("+AlwaysOnTop +Disabled", "Escaneando...")
    progressGui.Add("Text", "x20 y20 w360", "Buscando objetos Natural...")
    txtProgreso := progressGui.Add("Text", "x20 y50 w360", "Iniciando escaneo...")
    progressGui.Show("w400 h100")
    
    ; Buscar archivos recursivamente
    contadorArchivos := 0
    Loop Files, rutaOrigen . "\*.*", "R"
    {
        ; Verificar extensión
        for ext in extensionesNatural {
            if InStr(A_LoopFileExt, SubStr(ext, 2)) {
                contadorArchivos++
                archivosEncontrados.Push({
                    ruta: A_LoopFileFullPath,
                    nombre: A_LoopFileName,
                    dir: A_LoopFileDir,
                    ext: A_LoopFileExt,
                    tamano: A_LoopFileSize,
                    fecha: A_LoopFileTimeModified
                })
                
                if Mod(contadorArchivos, 10) = 0
                    txtProgreso.Value := "Encontrados: " . contadorArchivos . " archivos..."
                
                break
            }
        }
    }
    
    progressGui.Destroy()
    
    if archivosEncontrados.Length = 0 {
        MsgBox("No se encontraron objetos Natural en la ruta seleccionada.`n`nExtensiones buscadas: " . ArrayToString(extensionesNatural), 
        "Sin resultados", "Icon!")
        mainGui.Show()
        return
    }
    
    ; Mostrar ventana de selección
    MostrarVentanaSeleccion()
}

MostrarVentanaSeleccion() {
    global archivosEncontrados, archivosValidos, listView, txtContador
    global esCarpetaPersonalizada, proyectoSeleccionado, origenSeleccion
    archivosValidos := []   ; Resetear en cada apertura de la ventana

    ; Contar solo archivos válidos
    contadorValidos := 0
    for archivo in archivosEncontrados {
        nombreSinExt := StrReplace(archivo.nombre, "." . archivo.ext, "")
        if ValidarNombreNatural(nombreSinExt, "." . archivo.ext)
            contadorValidos++
    }

    ; =============================================
    ;     VENTANA DE SELECCIÓN - ESTILO MODERNO
    ; =============================================
    selGui := Gui("-MaximizeBox +MinSize800x580", "Seleccionar Objetos Natural para Migrar")
    selGui.BackColor := "F8F9FA"
    selGui.SetFont("s12 c333333", "Segoe UI")

    ; Header: "Objetos Natural encontrados (N)" a la izquierda, "PROYECTO: XXXX" a la derecha
    selGui.SetFont("s13 bold c0D47A1", "Segoe UI")
    selGui.Add("Text", "x30 y20 w480", "Objetos Natural encontrados (" . contadorValidos . ")")

    ; Etiqueta de proyecto a la derecha (Workspace/Git y carpeta personalizada)
    if proyectoSeleccionado != "" {
        selGui.SetFont("s12 bold c0D47A1", "Segoe UI")
        selGui.Add("Text", "x490 y22 w280 Right", "PROYECTO: " . proyectoSeleccionado)
    }

    ; Instrucción clara
    selGui.SetFont("s11 c555555", "Segoe UI")
    selGui.Add("Text", "x30 y55 w740", "Marque los objetos que desea copiar. Puede usar SHIFT+Click para selección múltiple.")

    ; ────────────────────────────────────────────
    ;               LISTVIEW PRINCIPAL
    ; ────────────────────────────────────────────
    listView := selGui.Add("ListView", "x30 y85 w740 h415 Checked Grid", ["    NOMBRE", "TIPO", "TAMAÑO", "FECHA", "__RUTA__"])
    listView.SetFont("s10", "Segoe UI")

    ; Poblar ListView con validación de nombres
    for archivo in archivosEncontrados {
        nombreSinExt := StrReplace(archivo.nombre, "." . archivo.ext, "")
        if !ValidarNombreNatural(nombreSinExt, "." . archivo.ext)
            continue

        tamanoKB := Round(archivo.tamano / 1024, 1) . " KB"
        fechaModif := FileGetTime(archivo.ruta, "M")
        fechaFormateada := FormatTime(fechaModif, "dd/MM/yyyy")
        tipoObjeto := ConvertirExtensionATipo(archivo.ext)

        listView.Add("", archivo.nombre, tipoObjeto, tamanoKB, fechaFormateada, archivo.ruta)
        archivosValidos.Push(archivo)
    }

    listView.ModifyCol(1, "310 Left")   ; Nombre
    listView.ModifyCol(2, "130 Left")   ; Tipo
    listView.ModifyCol(3, "100 Left")   ; Tamaño
    listView.ModifyCol(4, "160 Left")   ; Fecha
    listView.ModifyCol(5, "0")          ; Ruta — columna oculta

    ; ────────────────────────────────────────────
    ; BARRA DE ACCIONES INFERIOR
    ; ────────────────────────────────────────────
    selGui.SetFont("s10", "Segoe UI")

    ; Seleccionar / Deseleccionar
    btnSelTodos := selGui.Add("Button", "x30 y520 w160 h45", "✓  Seleccionar todos")
    btnSelTodos.Opt("+Background4CAF50")   ; Verde
    btnSelTodos.SetFont("s10 cFFFFFF", "Segoe UI")

    btnDeselTodos := selGui.Add("Button", "x200 y520 w200 h45", "✗  Deseleccionar todos")
    btnDeselTodos.Opt("+BackgroundE53935")   ; Rojo suave
    btnDeselTodos.SetFont("s10 cFFFFFF", "Segoe UI")

    ; Contador de seleccionados — se agrega DESPUÉS de los botones para que quede
    ; por encima en el orden Z y no provoque borrado visual al actualizarse
    btnVolver := selGui.Add("Button", "x560 y520 w100 h45", "  ←  Volver")
    btnVolver.Opt("+Background607D8B")   ; Gris azulado
    btnVolver.SetFont("s10 cFFFFFF", "Segoe UI")

    btnCopiar := selGui.Add("Button", "x670 y520 w100 h45", "📋  COPIAR")
    btnCopiar.Opt("+BackgroundFF5722")   ; Naranja acción
    btnCopiar.SetFont("s10 bold cFFFFFF", "Segoe UI")

    ; txtContador se crea AL FINAL para que su redibujado no tape los botones
    txtContador := selGui.Add("Text", "x420 y533 w130 BackgroundF8F9FA c0D47A1", "Seleccionados: 0")
    txtContador.SetFont("s11 bold", "Segoe UI")

    ; ────────────────────────────────────────────
    ;               EVENTOS
    ; ────────────────────────────────────────────
    btnSelTodos.OnEvent("Click", (*) => SeleccionarTodos(true))
    btnDeselTodos.OnEvent("Click", (*) => SeleccionarTodos(false))
    btnVolver.OnEvent("Click", (*) => VolverAlMenu(selGui, origenSeleccion))
    btnCopiar.OnEvent("Click", (*) => CopiarArchivosSeleccionados(selGui))

    listView.OnEvent("ItemCheck", (*) => ActualizarContador())

    ; Hook para Shift+Click: subclassing del control ListView
    global lvUltimoClick := 0
    global lvProcOrig := 0
    lvHwnd := listView.Hwnd
    lvProcOrig := DllCall("SetWindowLongPtr", "Ptr", lvHwnd, "Int", -4,
                          "Ptr", CallbackCreate(LV_ShiftClickProc, , 4), "Ptr")

    selGui.OnEvent("Close", (*) => VolverAlMenu(selGui, origenSeleccion))

    selGui.Show("Center w800 h600")
}

; ===== FUNCIÓN: ACTUALIZAR CONTADOR =====
ActualizarContador() {
    global listView, txtContador
    
    totalEnLista := listView.GetCount()   ; Solo los ítems válidos visibles en el ListView
    contador := 0
    Loop totalEnLista {
        if listView.GetNext(A_Index - 1, "Checked") = A_Index
            contador++
    }
    
    txtContador.Value := "Seleccionados: " . contador
    if (contador = 0)
        txtContador.Opt("c555555")
    else if (contador = totalEnLista)
        txtContador.Opt("c2E7D32")   ; verde cuando todos están seleccionados
    else
        txtContador.Opt("c0D47A1")   ; azul cuando hay selección parcial
}

; ===== FUNCIÓN: SELECCIONAR TODOS =====
SeleccionarTodos(estado) {
    global listView
    Loop listView.GetCount() {
        listView.Modify(A_Index, estado ? "Check" : "-Check")
    }
    ActualizarContador()
}

; ===== FUNCIÓN: SHIFT+CLICK EN LISTVIEW (via subclassing) =====
; Intercepta WM_LBUTTONDOWN sobre el ListView para detectar Shift+Click
; y marcar/desmarcar el rango entre el último ítem clickeado y el actual.
LV_ShiftClickProc(hwnd, msg, wParam, lParam) {
    global listView, lvUltimoClick, lvProcOrig

    ; Solo interceptar WM_LBUTTONDOWN (0x0201)
    if msg != 0x0201
        return DllCall("CallWindowProc", "Ptr", lvProcOrig, "Ptr", hwnd,
                       "UInt", msg, "Ptr", wParam, "Ptr", lParam, "Ptr")

    ; Determinar qué fila fue clickeada mediante LVM_HITTEST (0x1012)
    ; Construir estructura LVHITTESTINFO: x (Int32) + y (Int32) + flags (UInt32) + iItem (Int32)
    htInfo := Buffer(24, 0)
    NumPut("Int", lParam & 0xFFFF,        htInfo, 0)   ; x
    NumPut("Int", (lParam >> 16) & 0xFFFF, htInfo, 4)  ; y
    filaHit := DllCall("SendMessage", "Ptr", hwnd, "UInt", 0x1012, "Ptr", 0, "Ptr", htInfo.Ptr, "Ptr") + 1

    ; Si no se clickeó ninguna fila válida, comportamiento normal
    if filaHit <= 0 {
        lvUltimoClick := 0
        return DllCall("CallWindowProc", "Ptr", lvProcOrig, "Ptr", hwnd,
                       "UInt", msg, "Ptr", wParam, "Ptr", lParam, "Ptr")
    }

    ; Si Shift NO está presionado: guardar fila y comportamiento normal
    if !GetKeyState("Shift", "P") {
        lvUltimoClick := filaHit
        return DllCall("CallWindowProc", "Ptr", lvProcOrig, "Ptr", hwnd,
                       "UInt", msg, "Ptr", wParam, "Ptr", lParam, "Ptr")
    }

    ; Shift+Click: aplicar rango desde lvUltimoClick hasta filaHit
    desde := lvUltimoClick > 0 ? lvUltimoClick : filaHit
    hasta  := filaHit

    ; Ordenar rango
    if desde > hasta {
        tmp  := desde
        desde := hasta
        hasta := tmp
    }

    ; El estado de referencia es el OPUESTO al estado actual de filaHit:
    ; si el ítem clickeado está marcado → el Shift+Click lo desmarca (y al rango también)
    ; si está desmarcado → lo marca (y al rango también)
    ; Esto replica exactamente el comportamiento nativo de Windows
    estaCheckeado := (listView.GetNext(filaHit - 1, "Checked") = filaHit)
    estadoRef     := !estaCheckeado

    ; Aplicar ese estado a todo el rango
    Loop (hasta - desde + 1) {
        listView.Modify(desde + A_Index - 1, estadoRef ? "Check" : "-Check")
    }

    lvUltimoClick := filaHit
    ActualizarContador()

    ; Retornar 0 para suprimir el comportamiento por defecto del click
    return 0
}

; ===== FUNCIÓN: COPIAR ARCHIVOS SELECCIONADOS =====
CopiarArchivosSeleccionados(ventana) {
    global listView, archivosEncontrados, archivosValidos, destinoBase, versionNatural
    
    ; Contar seleccionados — leer ruta desde columna oculta para respetar el orden de sort
    archivosCopiar := []
    Loop listView.GetCount() {
        if listView.GetNext(A_Index - 1, "Checked") = A_Index {
            rutaFila := listView.GetText(A_Index, 5)
            ; Buscar el objeto correspondiente en archivosValidos por ruta
            for av in archivosValidos {
                if av.ruta = rutaFila {
                    archivosCopiar.Push(av)
                    break
                }
            }
        }
    }
    
    if archivosCopiar.Length = 0 {
        MsgBox("No hay archivos seleccionados.", "Aviso", "Icon!")
        return
    }
    
    ; Obtener ruta de destino actualizada
    destinoBase := ObtenerRutaDestino()
    
    ; Mostrar información según versión
    versionFormateada := FormatearVersion(versionNatural)
    libreriaActual := Trim(txtLibreria.Value)
    
    mensajeConfirmacion := "¿Desea copiar " . archivosCopiar.Length . " archivo(s) a:`n`n"
    mensajeConfirmacion .= destinoBase . "`n`n"
    mensajeConfirmacion .= "Librería: " . libreriaActual . "`n"
    mensajeConfirmacion .= "Versión Natural: " . versionFormateada . "`n"
    
    mensajeConfirmacion .= "`nSe eliminarán automáticamente los encabezados de NaturalONE."
    respuesta := MsgBox(mensajeConfirmacion, "Confirmar", "YesNo Icon?")
    if respuesta = "No"
        return
    
    ; Crear directorios necesarios según la versión
    if !DirExist(destinoBase) {
        try {
            ; Para versiones antiguas, crear solo la carpeta de librería
            if (versionNatural = "210" || versionNatural = "211" || versionNatural = "212") {
                DirCreate(destinoBase)
            } else {
                ; Para versiones nuevas, crear toda la estructura FUSER\LIBRERIA\SRC
                rutaNATAPPS := ".\dos\NATAPPS"
                rutaFUSER := rutaNATAPPS . "\FUSER"
                rutaLibreria := rutaFUSER . "\" . libreriaActual
                
                if !DirExist(rutaNATAPPS)
                    DirCreate(rutaNATAPPS)
                if !DirExist(rutaFUSER)
                    DirCreate(rutaFUSER)
        
                if !DirExist(rutaLibreria)
                    DirCreate(rutaLibreria)
                if !DirExist(destinoBase)
                    DirCreate(destinoBase)
            }
        } catch as err {
          
            MsgBox("Error al crear el directorio destino:`n" . err.Message, "Error", "IconX")
            return
        }
    } else {
        ; Si existe, borrar todo el contenido anterior
        try {
            Loop Files, destinoBase . "\*" {
                FileDelete(A_LoopFileFullPath)
            }
        } catch as err {
            MsgBox("Error al eliminar archivos anteriores:`n" . err.Message, "Error", "IconX")
            return
        }
    }
    
    ; Copiar archivos
    ventana.Hide()
    progressGui := Gui("+AlwaysOnTop", "Copiando archivos...")
    progressGui.Add("Text", "x20 y20 w360", "Copiando archivos...")
    txtArchivo := progressGui.Add("Text", "x20 y50 w360", "")
    txtLineas := progressGui.Add("Text", "x20 y75 w360", "")
    progressGui.Show("w400 h120")
    
    copiados := 0
    errores := []
    totalLineasEliminadas := 0
    
    for archivo in archivosCopiar {
        txtArchivo.Value := "Procesando: " . archivo.nombre
        
        rutaDestino := destinoBase . "\" . archivo.nombre
        
        try {
            ; Leer archivo original con codificación UTF-8
            contenidoOriginal := FileRead(archivo.ruta, "UTF-8")
            
            ; Eliminar BOM si existe (caracteres invisibles al inicio)
            if (SubStr(contenidoOriginal, 1, 1) = Chr(0xFEFF))
                contenidoOriginal := SubStr(contenidoOriginal, 2)
            
            ; Limpiar encabezados de NaturalONE y quitar tildes
            resultado := LimpiarEncabezadosNatural(contenidoOriginal)
            contenidoLimpio := resultado.contenido
            lineasEliminadas := resultado.lineasEliminadas

            ; Traducir asignaciones := a COMPUTE = en objetos de código
            extUpper := StrUpper(archivo.ext)
            if (extUpper = "NSP" || extUpper = "NSN" || extUpper = "NSS"
             || extUpper = "NSC" || extUpper = "NSH" || extUpper = "NST") {
                contenidoLimpio := TraducirAsignaciones(contenidoLimpio)
                contenidoLimpio := EliminarDefineWorkFile(contenidoLimpio)
                contenidoLimpio := TraducirDefineWindow(contenidoLimpio)
                contenidoLimpio := LimpiarTiposEnView(contenidoLimpio)
            }
            
            ; Si es un área de datos (.NSL, .NSG, .NSA), procesar estructura especial
            esAreaDatos := InStr(archivo.ext, "NSL") || InStr(archivo.ext, "NSG") || InStr(archivo.ext, "NSA")
            if esAreaDatos {
                contenidoLimpio := ProcesarAreaDatos(contenidoLimpio, rutaOrigen)
            }
            
            ; Si es una DDM (.NSD), procesar formato DDM
            esDDM := InStr(archivo.ext, "NSD")
            if esDDM {
                contenidoLimpio := ProcesarDDM(contenidoLimpio)
            }
            
            totalLineasEliminadas += lineasEliminadas
        
            txtLineas.Value := "Líneas eliminadas en total: " . totalLineasEliminadas
            
            ; Eliminar archivo si existe
            if FileExist(rutaDestino)
                FileDelete(rutaDestino)
            
            ; DDMs: escribir en binario puro para preservar 0C, 0D, 0A tal como están
            if esDDM {
                bufDDM := Buffer(StrPut(contenidoLimpio, "CP0") - 1)
                StrPut(contenidoLimpio, bufDDM, "CP0")
                f := FileOpen(rutaDestino, "w")
                f.RawWrite(bufDDM, bufDDM.Size)
                f.Close()
            } else {
                ; Resto de archivos: codificación del sistema (CP1252 en Windows español)
                FileAppend(contenidoLimpio, rutaDestino, "CP0")
            }
            
            copiados++
        } catch as err {
            errores.Push(archivo.nombre . ": " . err.Message)
        }
    }
    
    progressGui.Destroy()
    
    ; ===== GENERAR FILEDIR.SAG =====
    if copiados > 0 {
        resultado := GenerarFILEDIRSAG(archivosCopiar)
        
        ; Escribir librería en NATPARM.SAG
        libreriaActual := Trim(txtLibreria.Value)
        libreriaEscrita := false  ; Inicializar antes del bloque condicional
        if libreriaActual != "" {
            if EscribirLibreriaEnNATPARM(libreriaActual) {
                libreriaEscrita := true
            } else {
                libreriaEscrita := false
            }
        }
        
        if resultado.exito {
            mensaje := "✓ Copiados exitosamente: " . copiados . " de " . archivosCopiar.Length
            ; Extraer solo el directorio sin FILEDIR.SAG
            rutaSinArchivo := StrReplace(resultado.ruta, "\FILEDIR.SAG", "")
            mensaje .= "`n✂ Líneas de encabezado eliminadas: " . totalLineasEliminadas
            mensaje .= "`n📋 FILEDIR.SAG generado con " . resultado.totalObjetos . " objeto(s)"
            mensaje .= "`n📁 Ubicación: " . rutaSinArchivo
            if libreriaEscrita
                mensaje .= "`n📚 Librería establecida: " . libreriaActual
            else
                mensaje .= "`n⚠ No se pudo establecer la librería en NATPARM.SAG"
        } else {
            mensaje := "✓ Copiados exitosamente: " . copiados . " de " . archivosCopiar.Length
            mensaje .= "`n✂ Líneas de encabezado eliminadas: " . totalLineasEliminadas
            mensaje .= "`n❌ Error al generar FILEDIR.SAG: " . resultado.error
            if libreriaEscrita
                mensaje .= "`n📚 Librería establecida: " . libreriaActual
            else
                mensaje .= "`n⚠ No se pudo establecer la librería en NATPARM.SAG"
        }
    } else {
        mensaje := "❌ No se pudo copiar ningún archivo"
    }
    
    if errores.Length > 0
        mensaje .= "`n`n❌ Errores: " . errores.Length . "`n" . ArrayToString(errores, "`n")
    
    MsgBox(mensaje, "Resultado", copiados = archivosCopiar.Length ? "Icon!" : "Icon!")
    
    ; Guardar librería actual para cargarla al volver
    global libreriaDefecto
    libreriaActualMigracion := Trim(txtLibreria.Value)
    if libreriaActualMigracion != ""
        libreriaDefecto := libreriaActualMigracion
    
    ventana.Destroy()
    MostrarMenuPrincipal()
}

; ===== FUNCIÓN: GENERAR FILEDIR.SAG =====
GenerarFILEDIRSAG(archivos) {
    global destinoBase, versionNatural
    
    try {
		; Determinar dónde guardar FILEDIR.SAG según versión
        if (versionNatural = "210" || versionNatural = "211" || versionNatural = "212") {
            ; Versiones antiguas: FILEDIR.SAG va en la carpeta de librería
            rutaNATUSER := destinoBase  ; Ya es .\dos\NATAPPS\LIBRERIA
        } else {
            ; Versiones nuevas: FILEDIR.SAG va en NATUSER (sin \SRC)
            rutaNATUSER := StrReplace(destinoBase, "\SRC", "")
        }
        
        archivoSAG := rutaNATUSER . "\FILEDIR.SAG"
	
        ; Preparar lista de objetos ordenados alfabéticamente
        objetosParaFILEDIR := []
        
        for archivo in archivos {
            ; Extraer extensión y determinar tipo
            ext := "." . archivo.ext
            tipoObjeto := DeterminarTipoObjeto(ext)
            
            if tipoObjeto = 0 {
                continue  ; Extensión no reconocida
            }
            
            ; Extraer nombre sin extensión y normalizar a 8 caracteres
            nombreSinExt := StrReplace(archivo.nombre, ext, "")
            nombreSinExt := StrUpper(nombreSinExt)  ; Mayúsculas
            nombreSinExt := SubStr(nombreSinExt, 1, 8)  ; Máximo 8 caracteres
            
            ; Validar que el nombre cumpla las reglas de Natural
            if !ValidarNombreNatural(nombreSinExt, "." . archivo.ext) {
                continue  ; Nombre inválido
            }
            
            ; Si es DDM, extraer nombre lógico del contenido
            nombreLogico := ""  ; Inicializar variable
            if tipoObjeto = 0x44 {  ; DDM
                nombreLogico := ExtraerNombreLogicoDDM(archivo.ruta)
            }
            
            objetosParaFILEDIR.Push({
                nombre: nombreSinExt,
                nombreLogico: nombreLogico,
                tipo: tipoObjeto,
                esDDM: (tipoObjeto = 0x44)
            })
        }
        
        ; Ordenar alfabéticamente
        objetosOrdenados := OrdenarObjetosAlfabeticamente(objetosParaFILEDIR)
        
        ; Separar DDMs de otros objetos
        ddms := []
        otrosObjetos := []
        
        for obj in objetosOrdenados {
            if obj.esDDM {
                ddms.Push(obj)
            } else {
                otrosObjetos.Push(obj)
            }
        }
        
        ; Combinar: primero DDMs, luego otros objetos
        objetosOrdenados := []
        for ddm in ddms {
            objetosOrdenados.Push(ddm)
        }
        for otro in otrosObjetos {
            objetosOrdenados.Push(otro)
        }
        
        if objetosOrdenados.Length = 0 {
             return {exito: false, error: "No hay objetos válidos para generar FILEDIR.SAG"}
        }
        
        ; Generar archivo binario EN LA CARPETA NATUSER (sin \SRC)
        rutaNATUSER := StrReplace(destinoBase, "\SRC", "")
        archivoSAG := rutaNATUSER . "\FILEDIR.SAG"
        
        if FileExist(archivoSAG)
            FileDelete(archivoSAG)
        
        ; Crear archivo
        file := FileOpen(archivoSAG, "w")
        
        ; ===== HEADER (20 bytes) =====
        headerBytes := [0x01, 0x00, 0x2A, 0x00, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x00, 0x53, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]
        
        ; Actualizar contador de objetos en byte 0
        headerBytes[1] := objetosOrdenados.Length
        
        for byte in headerBytes {
            file.WriteUChar(byte)
        }
        
        ; ===== OBJETOS (64 bytes cada uno = 0x40) =====
        for obj in objetosOrdenados {
            
            ; Para DDMs, usar nombre lógico (hasta 28 chars); para otros, nombre físico (hasta 8 chars)
            if obj.esDDM && obj.nombreLogico != "" {
                ; DDM: nombre lógico (hasta 28 caracteres) desde byte 0x00
                nombreLogicoBytess := StrSplit(obj.nombreLogico)
                Loop 28 {
                    if A_Index <= nombreLogicoBytess.Length
                        file.WriteUChar(Ord(nombreLogicoBytess[A_Index]))
                    else
                        file.WriteUChar(0x00)
                }
            } else {
                ; Otros objetos: nombre normal (8 caracteres) + padding
                nombreBytes := StrSplit(obj.nombre)
                Loop 8 {
                    if A_Index <= nombreBytes.Length
                        file.WriteUChar(Ord(nombreBytes[A_Index]))
                    else
                        file.WriteUChar(0x00)
                }
                
                ; Padding hasta 28 bytes (20 bytes adicionales)
                Loop 20 {
                    file.WriteUChar(0x00)
                }
            }
            
            ; Último byte del bloque de nombre: 00
              file.WriteUChar(0x00)
            
            ; --- Offset después del nombre (ahora ajustado) ---
            ; Rellenar hasta llegar al offset 0x20 (donde va el nombre físico)
            ; Ya escribimos 32 bytes (28 nombre + 4 padding), necesitamos 0 más para llegar a 0x20
            
            ; --- Offset 0x20-0x2F: Nombre físico ---
            file.WriteUChar(0x00)
            file.WriteUChar(0x00)
            file.WriteUChar(0x00)
            file.WriteUChar(0x00)
            
            ; Nombre físico (siempre 8 caracteres)
            nombreFisicoBytes := StrSplit(obj.nombre)
            Loop 8 {
                if A_Index <= nombreFisicoBytes.Length
                    file.WriteUChar(Ord(nombreFisicoBytes[A_Index]))
                else
                    file.WriteUChar(0x00)
            }
            
            file.WriteUChar(0x00)
            file.WriteUChar(0x00)
            file.WriteUChar(0x00)
            
            ; --- Offset 0x30-0x3F: Tipo + UserID ---
            file.WriteUChar(0x00)
            file.WriteUChar(0x00)
            
            ; Byte 0x32: TIPO DE OBJETO
            file.WriteUChar(obj.tipo)
            
            file.WriteUChar(0x01)
            file.WriteUChar(0x01)
            
            ; UserID "SAGPC"
            file.WriteUChar(0x53)  ; S
            file.WriteUChar(0x41)  ; A
            file.WriteUChar(0x47)  ; G
            file.WriteUChar(0x50)  ; P
            file.WriteUChar(0x43)  ; C
            file.WriteUChar(0x20)  ; espacio
            file.WriteUChar(0x20)  ; espacio
            file.WriteUChar(0x20)  ; espacio
            
            file.WriteUChar(0x00)
            file.WriteUChar(0x00)
            file.WriteUChar(0x01)
            file.WriteUChar(0x7E)
            file.WriteUChar(0x00)
            file.WriteUChar(0x00)
            file.WriteUChar(0x00)
        }
        
        file.Close()
        
        return {exito: true, totalObjetos: objetosOrdenados.Length, ruta: archivoSAG}
        
    } catch as err {
        return {exito: false, error: err.Message}
    }
}

; ===== FUNCIÓN: VALIDAR NOMBRE NATURAL =====
ValidarNombreNatural(nombre, extension := "") {
    ; Determinar límite de longitud según extensión
    maxLongitud := 8  ; Por defecto para la mayoría de objetos
    
    ; Para DDMs (.NSD) permitir hasta 28 caracteres
    if extension != "" && StrUpper(extension) = ".NSD" {
        maxLongitud := 28
    }
    
    ; Validar longitud mínima y máxima
    if StrLen(nombre) < 1 || StrLen(nombre) > maxLongitud
        return false
    
    ; Primer carácter debe ser A-Z o #
    primerChar := SubStr(nombre, 1, 1)
    if !RegExMatch(primerChar, "^[A-Z#]$")
        return false
    
    ; Caracteres siguientes: A-Z, 0-9, _, -, #
    ; Para DDMs también permitir espacios internos
    if StrLen(nombre) > 1 {
        resto := SubStr(nombre, 2)
        
        if extension != "" && StrUpper(extension) = ".NSD" {
            ; Para DDMs: permitir letras, números, numeral, guión y espacios
            if !RegExMatch(resto, "^[A-Z0-9# -]+$")
                return false
        } else {
            ; Para otros objetos: reglas normales
            if !RegExMatch(resto, "^[A-Z0-9_#-]+$")
                return false
        }
    }
    
    return true
}

; ===== FUNCIÓN: DETERMINAR TIPO DE OBJETO =====
DeterminarTipoObjeto(extension) {
    ; Mapeo de extensiones a bytes de tipo
    tipos := Map(
        ".NSP", 0x50,  ; Program
        ".NSN", 0x4E,  ; Subprogram
        ".NSM", 0x4D,  ; Map
        ".NSC", 0x43,  ; Copycode
        ".NSH", 0x48,  ; Helproutine
        ".NSS", 0x53,  ; Subroutine
        ".NST", 0x54,  ; Text
        ".NSG", 0x47,  ; Global Data Area
        ".NSL", 0x4C,  ; Local Data Area
        ".NSA", 0x41,  ; Parameter Data Area
        ".NSD", 0x44   ; DDM
    )
    
    extensionUpper := StrUpper(extension)
    
    if tipos.Has(extensionUpper)
        return tipos[extensionUpper]
    else
        return 0  ; No reconocido
}

; ===== FUNCIÓN: EXTRAER NOMBRE LÓGICO DDM =====
ExtraerNombreLogicoDDM(rutaArchivo) {
    try {
        contenido := FileRead(rutaArchivo, "UTF-8")
        lineas := StrSplit(contenido, "`n")
        
        ; El nombre lógico está en la primera línea, columna 22 en adelante
        if lineas.Length >= 1 {
            primeraLinea := lineas[1]
            
            ; Extraer desde columna 22 (índice 22)
            if StrLen(primeraLinea) >= 22 {
                nombreLogico := SubStr(primeraLinea, 22)
                
                ; Buscar hasta el primer espacio múltiple o final
                posEspacio := InStr(nombreLogico, "  ")
                if posEspacio > 0 {
                    nombreLogico := SubStr(nombreLogico, 1, posEspacio - 1)
                }
                
                nombreLogico := Trim(nombreLogico)
                nombreLogico := StrUpper(nombreLogico)
                nombreLogico := SubStr(nombreLogico, 1, 28)  ; Máximo 28 caracteres
                
                return nombreLogico
            }
        }
    }
    
    return ""
}

; ===== FUNCIÓN: ORDENAR OBJETOS ALFABÉTICAMENTE =====
OrdenarObjetosAlfabeticamente(objetos) {
    ; Bubble sort simple para ordenar por nombre
    n := objetos.Length
    
    Loop n - 1 {
        i := A_Index
        Loop n - i {
            j := A_Index
            if StrCompare(objetos[j].nombre, objetos[j + 1].nombre) > 0 {
                ; Intercambiar
                temp := objetos[j]
                objetos[j] := objetos[j + 1]
                objetos[j + 1] := temp
            }
        }
    }
    
    return objetos
}

; ===== FUNCIÓN: TRADUCIR DEFINE WINDOW A SET CONTROL =====
TraducirDefineWindow(contenido) {
    lineas := StrSplit(contenido, "`n")

    ; ── Paso 1: recolectar todos los bloques DEFINE WINDOW ───────────────────
    ; ventanas es un Map: nombre → {control: "WB...", lineasBloque: [índices]}
    ventanas := Map()
    i := 1
    while i <= lineas.Length {
        linea := lineas[i]

        ; Ignorar líneas comentadas
        if SubStr(linea, 1, 1) = "*" {
            i++
            continue
        }

        ; Detectar inicio de bloque DEFINE WINDOW
        if RegExMatch(Trim(linea), "^DEFINE\s+WINDOW\s+(\S+)", &m) {
            nombreVentana := StrUpper(m[1])
            columnas := ""
            filas    := ""
            baseVal  := ""
            framed   := false
            lineasBloque := [i]   ; índice de la línea DEFINE WINDOW

            ; Leer las siguientes líneas del bloque (SIZE, BASE, FRAMED)
            j := i + 1
            while j <= lineas.Length {
                lb := Trim(lineas[j])
                if SubStr(lineas[j], 1, 1) = "*" {
                    j++
                    continue
                }
                if RegExMatch(lb, "^SIZE\s+(\d+)\s*\*\s*(\d+)", &ms) {
                    columnas := ms[1]
                    filas    := ms[2]
                    lineasBloque.Push(j)
                } else if RegExMatch(lb, "^BASE\s+(\d+)\s*/\s*(\d+)", &mb) {
                    baseVal := mb[1] . "/" . mb[2]
                    lineasBloque.Push(j)
                } else if RegExMatch(lb, "^FRAMED") {
                    framed := true
                    lineasBloque.Push(j)
                } else {
                    ; Primera línea que no pertenece al bloque — fin del bloque
                    break
                }
                j++
            }

            ; Construir string de control
            control := "WL" . filas . "WC" . columnas . "WB" . baseVal
            if framed
                control .= "F"

            ventanas[nombreVentana] := {control: control, lineasBloque: lineasBloque}
            i := j
            continue
        }
        i++
    }

    ; ── Paso 2: reemplazar SET WINDOW y eliminar bloques DEFINE WINDOW ───────
    ; Marcar índices a eliminar
    eliminar := Map()
    for nombre, datos in ventanas {
        for idx in datos.lineasBloque
            eliminar[idx] := true
    }

    resultado := []
    i := 1
    while i <= lineas.Length {
        if eliminar.Has(i) {
            i++
            continue
        }

        linea := lineas[i]

        if SubStr(linea, 1, 1) != "*" {
            ; Reemplazar SET WINDOW OFF → SET CONTROL 'WC'
            if RegExMatch(linea, "i)SET\s+WINDOW\s+OFF") {
                indentacion := ""
                k := 1
                while k <= StrLen(linea) && SubStr(linea, k, 1) = " " {
                    indentacion .= " "
                    k++
                }
                linea := indentacion . "SET CONTROL 'WB'"
            } else {
                ; Reemplazar SET WINDOW '<nombre>' → SET CONTROL '<control>'
                for nombre, datos in ventanas {
                    patron := "i)SET\s+WINDOW\s+'" . nombre . "'"
                    if RegExMatch(linea, patron) {
                        indentacion := ""
                        k := 1
                        while k <= StrLen(linea) && SubStr(linea, k, 1) = " " {
                            indentacion .= " "
                            k++
                        }
                        linea := indentacion . "SET CONTROL '" . datos.control . "'"
                        break
                    }
                }
            }
        }

        resultado.Push(linea)
        i++
    }

    contenidoFinal := ""
    for i, linea in resultado {
        contenidoFinal .= linea
        if i < resultado.Length
            contenidoFinal .= "`n"
    }
    return contenidoFinal
}

; ===== FUNCIÓN: LIMPIAR TIPO Y LONGITUD EN CAMPOS DE VIEW =====
LimpiarTiposEnView(contenido) {
    lineas     := StrSplit(contenido, "`n")
    resultado  := []
    dentroView := false

    ; Códigos de formato Natural válidos (sin longitud: D y T)
    tiposSinLong  := "D|T"
    tiposConLong  := "A|B|F|I|N|P"
    ; Patrón de tipo con longitud:  (A10)  (N5)  (I2)  (B4)  (F8)  (P3)
    ; Patrón de tipo sin longitud:  (D)    (T)
    ; NO tocar: (1:20)  (1:5)  (0:10)  — contienen ':'
    patronConLong  := "\((" . tiposConLong . ")\d+(\.\d+)?\)"
    patronSinLong  := "\((" . tiposSinLong . ")\)"

    for linea in lineas {
        lineaTrim := Trim(linea)

        ; Detectar inicio de bloque VIEW (no comentada)
        if SubStr(linea, 1, 1) != "*" {
            if RegExMatch(lineaTrim, "i)^\d+\s+\S+\s+VIEW(\s+OF)?\s+\S+")
                dentroView := true
        }

        ; Detectar fin de bloque VIEW: END-DEFINE u otro nivel 1 que no sea VIEW
        if dentroView && SubStr(linea, 1, 1) != "*" {
            if InStr(lineaTrim, "END-DEFINE")
                dentroView := false
            else if RegExMatch(lineaTrim, "^1\s+") && !RegExMatch(lineaTrim, "i)^1\s+\S+\s+VIEW") {
                dentroView := false
            }
        }

        ; Dentro de un bloque VIEW: eliminar tipo/longitud entre paréntesis
        ; Solo si el paréntesis NO contiene ':'
        if dentroView && SubStr(linea, 1, 1) != "*" {
            ; Eliminar (A10), (N5), (I2), (B4), (F8), (P3), (D), (T), etc.
            ; Pero respetar (1:20), (0:10), etc.
            linea := RegExReplace(linea, "\((?![^)]*:)(?:" . tiposConLong . ")\d+(?:\.\d+)?\)", "")
            linea := RegExReplace(linea, "\((?![^)]*:)(?:" . tiposSinLong . ")\)", "")
        }

        resultado.Push(linea)
    }

    contenidoFinal := ""
    for i, linea in resultado {
        contenidoFinal .= linea
        if i < resultado.Length
            contenidoFinal .= "`n"
    }
    return contenidoFinal
}

; ===== FUNCIÓN: ELIMINAR LÍNEAS 'DEFINE WORK FILE' NO COMENTADAS =====
EliminarDefineWorkFile(contenido) {
    lineas := StrSplit(contenido, "`n")
    resultado := []

    for linea in lineas {
        ; Si la línea está comentada (columna 1 = '*'), se conserva siempre
        if SubStr(linea, 1, 1) = "*" {
            resultado.Push(linea)
            continue
        }
        ; Si contiene 'DEFINE WORK FILE' (no comentada), se elimina
        if InStr(linea, "DEFINE WORK FILE") {
            continue
        }
        resultado.Push(linea)
    }

    contenidoFinal := ""
    for i, linea in resultado {
        contenidoFinal .= linea
        if i < resultado.Length
            contenidoFinal .= "`n"
    }
    return contenidoFinal
}

; ===== FUNCIÓN: TRADUCIR ASIGNACIONES := A COMPUTE = =====
TraducirAsignaciones(contenido) {
    lineas := StrSplit(contenido, "`n")
    resultado := []

    for linea in lineas {
        ; Solo procesar líneas que no estén comentadas (columna 1 != '*')
        ; y que contengan el patrón ' := '
        if SubStr(linea, 1, 1) != "*" && InStr(linea, " := ") {
            ; Extraer la parte izquierda y derecha del ' := '
            pos := InStr(linea, " := ")
            izq := Trim(SubStr(linea, 1, pos - 1))
            der := Trim(SubStr(linea, pos + 4))
            ; Reconstruir como: COMPUTE <izq> = <der>
            ; Preservar la indentación original
            indentacion := ""
            i := 1
            while i <= StrLen(linea) && SubStr(linea, i, 1) = " " {
                indentacion .= " "
                i++
            }
            linea := indentacion . "COMPUTE " . izq . " = " . der
        }
        resultado.Push(linea)
    }

    contenidoFinal := ""
    for i, linea in resultado {
        contenidoFinal .= linea
        if i < resultado.Length
            contenidoFinal .= "`n"
    }
    return contenidoFinal
}

; ===== FUNCIÓN: LIMPIAR ENCABEZADOS DE NATURALONE =====
LimpiarEncabezadosNatural(contenido) {
    lineasEliminadas := 0
    lineas := StrSplit(contenido, "`n")
    lineasLimpias := []
    
    ; Patrones a eliminar
    patronesEliminar := [
        ">Natural Source Header",
        ":Mode",
        ":CP",
        "<Natural Source Header"
    ]
    
    for linea in lineas {
        eliminar := false
        
        ; Verificar si la línea contiene algún patrón
        for patron in patronesEliminar {
            if InStr(linea, patron) {
                eliminar := true
                lineasEliminadas++
                break
            }
         }
        
        ; Si no se debe eliminar, agregar a las líneas limpias
        if !eliminar {
            lineasLimpias.Push(linea)
        }
    }
    
    ; Reconstruir contenido sin tildes
    contenidoLimpio := ""
    for i, linea in lineasLimpias {
        contenidoLimpio .= linea
        if i < lineasLimpias.Length
            contenidoLimpio .= "`n"
    }
    
    ; Quitar tildes de TODO el contenido limpio
    contenidoLimpio := QuitarTildes(contenidoLimpio)
    
    return {contenido: contenidoLimpio, lineasEliminadas: lineasEliminadas}
}

; ===== FUNCIÓN: QUITAR TILDES =====
QuitarTildes(texto) {
    ; Reemplazar todas las vocales con tilde por vocales sin tilde
    ; Minúsculas
    texto := StrReplace(texto, "á", "a")
    texto := StrReplace(texto, "é", "e")
    texto := StrReplace(texto, "í", "i")
    texto := StrReplace(texto, "ó", "o")
    texto := StrReplace(texto, "ú", "u")
    ; Mayúsculas
    texto := StrReplace(texto, "Á", "A")
    texto := StrReplace(texto, "É", "E")
    texto := StrReplace(texto, "Í", "I")
    texto := StrReplace(texto, "Ó", "O")
    texto := StrReplace(texto, "Ú", "U")
    ; Ñ y diéresis
    texto := StrReplace(texto, "ñ", "n")
    texto := StrReplace(texto, "Ñ", "N")
    texto := StrReplace(texto, "ü", "u")
    texto := StrReplace(texto, "Ü", "U")
    
    return texto
}

; ===== FUNCIÓN: PROCESAR ÁREA DE DATOS =====
ProcesarAreaDatos(contenido, rutaOrigen) {
    lineas := StrSplit(contenido, "`n")
    lineasDF := []
    dentroDefineData := false
    nivelEstructuraActual := 0
    i := 1
    
    while i <= lineas.Length {
        linea := lineas[i]
        lineaTrim := Trim(linea)
        
        ; Detectar inicio de DEFINE DATA
        if InStr(lineaTrim, "DEFINE DATA LOCAL") || InStr(lineaTrim, "DEFINE DATA GLOBAL") || InStr(lineaTrim, "DEFINE DATA PARAMETER") {
            dentroDefineData := true
            i++
            continue
        }
        
        ; Detectar fin de DEFINE DATA
        if InStr(lineaTrim, "END-DEFINE") {
            dentroDefineData := false
            break
        }
        
        ; Procesar campos dentro de DEFINE DATA
        if dentroDefineData && StrLen(lineaTrim) > 0 {
            ; Verificar si es un comentario
            if SubStr(lineaTrim, 1, 1) = "*" {
                textoComentario := SubStr(lineaTrim, 2)
                textoComentario := Trim(textoComentario)
                lineaComentario := "**C           0   *" . textoComentario
                lineasDF.Push(lineaComentario)
                i++
                continue
            }
            
            ; Verificar si es CONST o INIT (línea independiente)
            if lineaTrim = "CONST" || lineaTrim = "INIT" {
                i++
                continue
            }
            
            ; Verificar si es una VISTA (VIEW OF)
            if InStr(lineaTrim, "VIEW OF") {
                resultadoVista := ProcesarVista(lineas, i, rutaOrigen)
                
                if resultadoVista.exito {
                    ; Agregar líneas generadas de la vista
                    for lineaVista in resultadoVista.lineas {
                        lineasDF.Push(lineaVista)
                    }
                    
                    ; Saltar las líneas procesadas
                    i := resultadoVista.siguienteLinea
                    continue
                }
            }
            
            ; Parsear campo
            resultado := ParsearCampoNatural(lineaTrim, nivelEstructuraActual)
            
            if resultado.linea != "" {
                inicializacion := ""
                esConstante := false
                saltarLineas := 0
                
                ; CASO 1: Verificar si CONST/INIT está en la misma línea
                ; Ejemplo: 1 VARIABLE(N2) INIT <12>
                if InStr(lineaTrim, " CONST ") || InStr(lineaTrim, " INIT ") {
                    esConstante := InStr(lineaTrim, " CONST ") > 0
                    
                    ; Buscar el valor en la misma línea
                    posInicio := InStr(lineaTrim, "<")
                    if posInicio > 0 {
                        posFin := InStr(lineaTrim, ">", , posInicio)
                        if posFin > 0 {
                            inicializacion := SubStr(lineaTrim, posInicio + 1, posFin - posInicio - 1)
                            inicializacion := StrReplace(inicializacion, "'", "")
                            inicializacion := Trim(inicializacion)
                        }
                    }
                }
                ; CASO 2 y 3: CONST/INIT en línea separada
                else if i + 1 <= lineas.Length {
                    lineaSig := Trim(lineas[i + 1])
                    
                    ; Verificar si la siguiente línea es CONST o INIT
                    if lineaSig = "CONST" || lineaSig = "INIT" {
                        esConstante := (lineaSig = "CONST")
                        saltarLineas := 1
                        
                        ; Buscar el valor en la línea siguiente al CONST/INIT
                        if i + 2 <= lineas.Length {
                            lineaValor := lineas[i + 2]
                            
                            ; Si tiene <, leer el valor
                            if InStr(lineaValor, "<") {
                                valorCompleto := ""
                                linActual := i + 2
                                
                                while linActual <= lineas.Length {
                                    lineaTemp := lineas[linActual]
                                    valorCompleto .= lineaTemp
                                    saltarLineas++
                                    
                                    if InStr(lineaTemp, ">") {
                                        break
                                    }
                                    linActual++
                                }
                                
                                inicializacion := ProcesarInicializacion(valorCompleto, esConstante)
                            }
                        }
                    }
                    ; CASO 2b: CONST/INIT en la misma línea pero valor en línea separada
                    ; Ejemplo: 1 VARIABLE(N2) INIT
                    ; <12>
                    else if InStr(lineaTrim, " CONST") = StrLen(lineaTrim) - 4 || InStr(lineaTrim, " INIT") = StrLen(lineaTrim) - 3 {
                        esConstante := InStr(lineaTrim, " CONST") > 0
                        
                        ; Buscar valor en siguiente línea
                        if i + 1 <= lineas.Length {
                            lineaValor := lineas[i + 1]
                            
                            if InStr(lineaValor, "<") {
                                valorCompleto := ""
                                linActual := i + 1
                                
                                while linActual <= lineas.Length {
                                    lineaTemp := lineas[linActual]
                                    valorCompleto .= lineaTemp
                                    saltarLineas++
                                    
                                    if InStr(lineaTemp, ">") {
                                        break
                                    }
                                    linActual++
                                }
                                
                                inicializacion := ProcesarInicializacion(valorCompleto, esConstante)
                            }
                        }
                    }
                }
                
                ; Generar línea
                if inicializacion != "" {
                    lineaGenerada := GenerarLineaConInicializacion(resultado, esConstante, inicializacion)
                    lineasDF.Push(lineaGenerada)
                    i += saltarLineas
                } else {
                    lineasDF.Push(resultado.linea)
                }
                
                ; Actualizar nivel de estructura
                if resultado.esEstructura {
                    nivelEstructuraActual := resultado.nivel
                } else if resultado.nivel <= nivelEstructuraActual && resultado.nivel > 0 {
                    nivelEstructuraActual := 0
                }
            }
        }
        
        i++
    }
    
    ; Reconstruir contenido
    contenidoFinal := ""
    for lineaDF in lineasDF {
        contenidoFinal .= lineaDF . "`n"
    }
    contenidoFinal .= contenido
    
    return contenidoFinal
}

; ===== FUNCIÓN: PROCESAR VISTA =====
ProcesarVista(lineas, indiceInicio, rutaOrigen) {
    ; Procesa una vista VIEW OF y genera las líneas **DV y **DD correspondientes
    
    lineaActual := Trim(lineas[indiceInicio])
    lineasGeneradas := []
    
    ; Parsear línea de vista: "1 NOMBRE-VISTA VIEW OF NOMBRE-DDM"
    RegExMatch(lineaActual, "^(\d+)\s+([A-Z0-9#_-]+)\s+VIEW\s+OF\s+([A-Z0-9#_-]+)", &matchVista)
    
    if !matchVista
        return {exito: false, lineas: [], siguienteLinea: indiceInicio + 1}
    
    nivelVista := matchVista[1]
    nombreVista := matchVista[2]
    nombreDDM := matchVista[3]
    
    ; Leer la DDM
    resultadoDDM := LeerDDM(nombreDDM, rutaOrigen)
    
    if !resultadoDDM.encontrado {
        ; Si no se encuentra la DDM, agregar comentario de error
        lineasGeneradas.Push("**C           0   * ERROR: DDM " . nombreDDM . " NO ENCONTRADA")
        return {exito: true, lineas: lineasGeneradas, siguienteLinea: indiceInicio + 1}
    }
    
    camposDDM := resultadoDDM.campos
    
    ; Generar línea **DV
    lineaDV := "**DV          0        V" . nivelVista . nombreVista
    
    ; Rellenar con espacios hasta columna 58
    while StrLen(lineaDV) < 57 {
        lineaDV .= " "
    }
    
    lineaDV .= nombreDDM
    lineasGeneradas.Push(lineaDV)
    
    ; Procesar campos de la vista
    i := indiceInicio + 1
    
    while i <= lineas.Length {
        lineaCampo := Trim(lineas[i])
        
        ; Si línea vacía, ignorar
        if lineaCampo = "" {
            i++
            continue
        }
        
        ; Si encuentra END-DEFINE, terminar
        if InStr(lineaCampo, "END-DEFINE") {
            break
        }
        
        ; Verificar si es un campo de la vista (debe empezar con un número)
        RegExMatch(lineaCampo, "^(\d+)\s+([A-Z0-9#_-]+)", &matchCampo)
        
        if !matchCampo {
            i++
            continue
        }
        
        nivelCampo := matchCampo[1]
        nombreCampo := matchCampo[2]
        
        ; Si el nivel es menor o igual al de la vista, terminó la vista
        if Integer(nivelCampo) <= Integer(nivelVista) {
            break
        }
        
        ; Extraer dimensiones si las tiene (ej: CAMPO(1:5) o CAMPO(1:3,1:12))
        dimensionesArray := ""
        if InStr(lineaCampo, "(") && InStr(lineaCampo, ":") {
            RegExMatch(lineaCampo, "(\(\d+:\d+[^)]*\))", &matchDim)
            if matchDim {
                dimensionesArray := matchDim[1]
            }
        }
        
        ; Buscar información del campo en la DDM
        if !camposDDM.Has(nombreCampo) {
            ; Campo no encontrado en DDM — construir **DD con nivel y nombre, sin tipo ni longitud
            lineaDD := "**DD"
            lineaDD .= "          "   ; Col 5-14: 10 espacios
            lineaDD .= "0"            ; Col 15
            lineaDD .= "   "          ; Col 16-18
            lineaDD .= " "            ; Col 19: sin tipo
            lineaDD .= " "            ; Col 20
            lineaDD .= "   "          ; Col 21-23: sin longitud
            lineaDD .= " "            ; Col 24: sin tipo estructura
            lineaDD .= nivelCampo     ; Col 25: nivel
            lineaDD .= nombreCampo    ; Col 26+: nombre
            if dimensionesArray != "" {
                while StrLen(lineaDD) < 57
                    lineaDD .= " "
                lineaDD .= dimensionesArray
            }
            lineasGeneradas.Push(lineaDD)
            i++
            continue
        }
        
        infoCampo := camposDDM[nombreCampo]
        
        ; Generar línea **DD
        lineaDD := GenerarLineaDD(infoCampo, nivelCampo, nombreCampo, dimensionesArray)
        lineasGeneradas.Push(lineaDD)
        
        i++
    }
    
    return {exito: true, lineas: lineasGeneradas, siguienteLinea: i}
}

; ===== FUNCIÓN: GENERAR LÍNEA DD =====
GenerarLineaDD(infoCampo, nivel, nombreCampo, dimensionesArray) {
    ; Genera una línea **DD para un campo de vista
    
    tipoEstructura := infoCampo.tipoEstructura
    tipoDato := infoCampo.tipoDato
    longitud := infoCampo.longitud
    
    lineaDD := "**DD"
    
    ; Determinar si es array (tiene dimensiones con formato 1:5)
    esArray := dimensionesArray != ""
    cantidadDimensiones := 0
    
    if esArray {
        ; Contar dimensiones (cantidad de comas + 1)
        cantidadDimensiones := 1
        Loop Parse, dimensionesArray {
            if A_LoopField = ","
                cantidadDimensiones++
        }
    }
    
    ; Columnas 5-14: Array info o espacios
    if esArray {
	    ; Columna 5: espacio
	    lineaDD .= " "
		
        ; Columna 6: I
        lineaDD .= "I"
        
        ; Columna 7: cantidad dimensiones
        lineaDD .= cantidadDimensiones
        
        ; Columna 8: espacio
        lineaDD .= " "
        
        ; Columna 9: C si es compuesto (M/P con array)
        if tipoEstructura = "M" || tipoEstructura = "P"
            lineaDD .= "C"
        else
            lineaDD .= " "
        
        ; Columnas 10-14: espacios
        lineaDD .= "     "
    } else if tipoEstructura = "G" {
        ; Para GROUP: columnas 5-14 = "    C     " (4 espacios + C + 5 espacios)
        lineaDD .= "    C     "
    } else {
        ; Para campos normales: columnas 5-14 = "          " (10 espacios)
        lineaDD .= "          "
    }
    
    ; Columna 15: 0
    lineaDD .= "0"
    
    ; Columnas 16-23: espacios y tipo de dato con longitud
    lineaDD .= "   "  ; Columnas 16-18
    
    ; Columna 19: tipo de dato (o espacio si es GROUP sin tipo)
    if tipoEstructura = "G" {
        lineaDD .= " "
    } else {
        lineaDD .= SubStr(tipoDato . " ", 1, 1)
    }
    
    ; Columna 20: espacio
    lineaDD .= " "
    
    ; Columnas 21-23: longitud (alineada a derecha, o espacios si no tiene)
    if longitud != "" {
        longitudStr := longitud
        while StrLen(longitudStr) < 3 {
            longitudStr := " " . longitudStr
        }
        lineaDD .= SubStr(longitudStr, 1, 3)
    } else {
        lineaDD .= "   "
    }
    
    ; Columna 24: G/M/P o espacio
    lineaDD .= SubStr(tipoEstructura . " ", 1, 1)
    
    ; Columna 25: nivel
    lineaDD .= nivel
    
    ; Columna 26+: nombre del campo
    lineaDD .= nombreCampo
    
    ; Si tiene dimensiones de array, agregar en columna 58
    if esArray {
        ; Rellenar con espacios hasta columna 58
        while StrLen(lineaDD) < 57 {
            lineaDD .= " "
        }
        lineaDD .= dimensionesArray
    }
    
    return lineaDD
}

; ===== FUNCIÓN: LEER DDM =====
LeerDDM(nombreDDM, rutaOrigen) {
    ; Buscar archivo .NSD que coincida con el nombre de la DDM
    ; El nombre puede estar en el nombre del archivo o dentro del contenido
    
    archivoEncontrado := ""
    encontrado := false
    
    ; Buscar recursivamente archivos .NSD
    Loop Files, rutaOrigen . "\*.NSD", "R"
    {
        if encontrado
            break
            
        ; Leer el contenido del archivo
        contenido := FileRead(A_LoopFileFullPath, "UTF-8")
        
        ; Buscar el nombre de la DDM en las primeras líneas
        lineas := StrSplit(contenido, "`n", "`r", 10)  ; Leer primeras 10 líneas
        
        for linea in lineas {
            ; Buscar patrón: "FILE: XXX  - NOMBREDDM"
            if InStr(linea, "FILE:") && InStr(linea, nombreDDM) {
                archivoEncontrado := A_LoopFileFullPath
                encontrado := true
                break
            }
        }
    }
    
    if archivoEncontrado = ""
        return {encontrado: false, campos: Map()}
    
    ; Leer el contenido completo
    contenido := FileRead(archivoEncontrado, "UTF-8")
    
    ; Extraer campos de la DDM
    campos := ExtraerCamposDDM(contenido)
    
    return {encontrado: true, campos: campos}
}

; ===== FUNCIÓN: EXTRAER CAMPOS DE DDM =====
ExtraerCamposDDM(contenidoDDM) {
    ; Parsear la DDM y extraer información de cada campo
    ; Retorna un Map donde la clave es el nombre del campo y el valor es un objeto con sus propiedades
    
    campos := Map()
    lineas := StrSplit(contenidoDDM, "`n")
    
    for linea in lineas {
        lineaTrim := Trim(linea)
        
        ; Eliminar números de línea al inicio (formato: 0001, 0002, etc.)
        lineaTrim := RegExReplace(lineaTrim, "^\d{4}\s*", "")
        lineaTrim := Trim(lineaTrim)
        
        ; Ignorar líneas vacías, comentarios y encabezados
        if lineaTrim = "" || SubStr(lineaTrim, 1, 1) = "*" || InStr(lineaTrim, "---")
            continue
        
        ; Ignorar encabezados
        if InStr(lineaTrim, "DB:") || InStr(lineaTrim, "TYPE:") || InStr(lineaTrim, "T L DB")
            continue
        
        ; Separar por espacios
        camposLinea := []
        palabraActual := ""
        
        Loop Parse, lineaTrim {
            char := A_LoopField
            if char = " " || char = "`t" {
                if palabraActual != "" {
                    camposLinea.Push(palabraActual)
                    palabraActual := ""
                }
            } else {
                 palabraActual .= char
            }
        }
        
        if palabraActual != ""
            camposLinea.Push(palabraActual)
        
        ; Debe tener al menos 4 campos (nivel, nombreCorto, nombreLargo, tipo o G/M/P)
        if camposLinea.Length < 3
            continue
        
        ; Parsear estructura del campo
        indice := 1
        tipoEstructura := ""  ; G, M o P
        nivel := ""
        nombreCorto := ""
        nombreLargo := ""
        tipoDato := ""
        longitud := ""
        opcion := ""
        descriptor := ""
        dimensiones := ""
        
        ; Verificar si empieza con G, M o P
        if camposLinea[1] = "G" || camposLinea[1] = "M" || camposLinea[1] = "P" {
            tipoEstructura := camposLinea[1]
            indice := 2
        }
        
        ; Nivel
        if indice <= camposLinea.Length && RegExMatch(camposLinea[indice], "^\d+$") {
            nivel := camposLinea[indice]
            indice++
        }
        
        ; Nombre corto
        if indice <= camposLinea.Length {
            nombreCorto := camposLinea[indice]
            indice++
        }
        
        ; Nombre largo
        if indice <= camposLinea.Length {
            nombreLargo := camposLinea[indice]
            indice++
        }
        
        ; Tipo de dato
        if indice <= camposLinea.Length && StrLen(camposLinea[indice]) = 1 {
            tipoDato := camposLinea[indice]
            indice++
        }
        
        ; Longitud (solo si tiene tipo de dato)
        if tipoDato != "" && indice <= camposLinea.Length && (RegExMatch(camposLinea[indice], "^\d+\.?\d*$") || RegExMatch(camposLinea[indice], "^\d+$")) {
            longitud := camposLinea[indice]
            indice++
        }
        
        ; Opción y descriptor (solo si no es estructura contenedora)
        esEstructuraContenedora := (tipoEstructura = "G" || tipoEstructura = "P") && tipoDato = ""
        
        if !esEstructuraContenedora {
            while indice <= camposLinea.Length {
                campo := camposLinea[indice]
                
                if campo = "N" || campo = "F" {
                    if opcion = ""
                        opcion := campo
                } else if campo = "D" || campo = "S" {
                    if descriptor = ""
                        descriptor := campo
                }
                
                indice++
            }
        }
        
        ; Detectar dimensiones de arrays (buscar patrón como (1:5) o (1:3,1:12))
        if InStr(nombreLargo, "(") {
            RegExMatch(nombreLargo, "^([^(]+)(\(.+\))$", &matchArray)
            if matchArray {
                nombreLargo := matchArray[1]
                dimensiones := matchArray[2]
            }
        }
        
        ; Guardar campo en el Map (usar nombreLargo como clave)
        campos[nombreLargo] := {
            tipoEstructura: tipoEstructura,
            nivel: nivel,
            nombreCorto: nombreCorto,
            tipoDato: tipoDato,
            longitud: longitud,
            opcion: opcion,
            descriptor: descriptor,
            dimensiones: dimensiones
        }
    }
    
    return campos
}

; ===== FUNCIÓN: PARSEAR CAMPO NATURAL =====
ParsearCampoNatural(linea, nivelEstructuraActual := 0) {
    ; Extraer nivel (primer número)
    RegExMatch(linea, "^(\d+)\s+", &matchNivel)
    if !matchNivel
        return {linea: "", esEstructura: false, nivel: 0}
    
    nivel := matchNivel[1]
    restoLinea := SubStr(linea, StrLen(matchNivel[0]) + 1)
    
    ; Verificar si tiene paréntesis (campo con tipo) o no (estructura)
    tieneParentesis := InStr(restoLinea, "(")
    
    if !tieneParentesis {
        ; Es una ESTRUCTURA (no tiene tipo de dato)
        ; Extraer solo el nombre
        nombreEstructura := Trim(restoLinea)
        
        ; Construir línea **DS
        lineaDF := "**DS          0         "
        lineaDF .= nivel
        lineaDF .= nombreEstructura
        
        return {linea: lineaDF, esEstructura: true, nivel: Integer(nivel)}
    }
    
    ; Es un campo normal (con tipo de dato)
    ; Extraer nombre y tipo (nombre(tipo))
    RegExMatch(restoLinea, "^([A-Z0-9#_-]+)\s*\(([^)]+)\)", &matchCampo)
    if !matchCampo
        return {linea: "", esEstructura: false, nivel: 0}
    
    nombreCampo := matchCampo[1]
    tipoDato := matchCampo[2]
    
    ; Verificar si es un array (contiene '/')
    esArray := InStr(tipoDato, "/")
    dimensionesArray := ""
    numeroDimensiones := 0
    
    if esArray {
        ; Separar tipo/longitud de las dimensiones
        partes := StrSplit(tipoDato, "/")
        tipoDato := partes[1]
        dimensionesArray := partes[2]
        
        ; Contar dimensiones (cantidad de comas + 1)
        numeroDimensiones := 1
        Loop Parse, dimensionesArray {
            if A_LoopField = ","
                numeroDimensiones++
        }
    }
    
    ; Parsear tipo de dato (formato: LETRA + LONGITUD)
    tipoLetra := SubStr(tipoDato, 1, 1)
    longitud := SubStr(tipoDato, 2)
    
    ; Si no hay longitud, dejar vacío
    if longitud = "" {
        longitud := "    "
    } else {
        ; Alinear longitud a la derecha en 4 espacios
        espaciosNecesarios := 4 - StrLen(longitud)
        Loop espaciosNecesarios {
            longitud := " " . longitud
        }
    }
    
    ; Determinar si es miembro de estructura (DK) o campo normal (DF)
    esMiembroEstructura := (nivelEstructuraActual > 0 && Integer(nivel) > nivelEstructuraActual)
    
    ; Construir línea según el tipo
    if esMiembroEstructura {
        ; Miembro de estructura: **DK
        lineaDF := "**DK "
    } else {
        ; Campo normal: **DF
        lineaDF := "**DF "
    }
    
    ; Columnas 6-7: Indicador de array o espacios
    if esArray {
        lineaDF .= "I" . numeroDimensiones
    } else {
        lineaDF .= "  "
    }
    
    ; Columnas 8-18: espacios + 0 + espacios
    lineaDF .= "       0   "
    
    ; Columna 19: Tipo de dato
    lineaDF .= tipoLetra
    
    ; Columnas 20-23: Longitud
    lineaDF .= longitud
    
    ; Columna 24: Espacio
    lineaDF .= " "
    
    ; Columna 25: Nivel
    lineaDF .= nivel
    
    ; Columna 26+: Nombre del campo
    lineaDF .= nombreCampo
    
    ; Si es array, agregar dimensiones a partir de columna 58
    if esArray {
        longitudActual := StrLen(lineaDF)
        espaciosNecesarios := 57 - longitudActual
        
        if espaciosNecesarios > 0 {
            Loop espaciosNecesarios {
                lineaDF .= " "
            }
        }
        
        lineaDF .= "(" . dimensionesArray . ")"
    }
    
    return {linea: lineaDF, esEstructura: false, nivel: Integer(nivel)}
}

; ===== FUNCIÓN: VOLVER AL MENÚ =====
VolverAlMenu(ventana, origen := "") {
    global rutaOrigen, txtLibreria, destinoBase, lvProcOrig, lvUltimoClick

    ; Guardar valores actuales
    rutaOrigenPreservada := rutaOrigen
    libreriaPreservada   := txtLibreria.Value

    ; Limpiar el subclassing del ListView antes de destruir la ventana
    lvProcOrig    := 0
    lvUltimoClick := 0

    ventana.Destroy()

    if origen = "workspace" {
        ; Volver a la subventana Seleccionar Workspace
        MostrarSeleccionWorkspace()
    } else {
        ; Carpeta personalizada o cualquier otro caso → menú principal
        MostrarMenuPrincipal()

        ; Restaurar valores
        if rutaOrigenPreservada != "" {
            global txtRutaActual, btnEscanear
            rutaOrigen := rutaOrigenPreservada
            txtRutaActual.Value := rutaOrigenPreservada
            ActualizarEstadoBotonEscanear()
        }

        if libreriaPreservada != "" {
            txtLibreria.Value := libreriaPreservada
            destinoBase := ObtenerRutaDestino()
        }
    }
}

; ===== FUNCIÓN: CONVERTIR EXTENSIÓN A TIPO =====
ConvertirExtensionATipo(ext) {
    ext := StrUpper(ext)
    switch ext {
        case "NSP": return "Programa"
        case "NSN": return "Subprograma"
        case "NSM": return "Mapa"
        case "NSC": return "Copycode"
        case "NSH": return "Helproutine"
        case "NSS": return "Subrutina"
        case "NST": return "Texto"
        case "NSG": return "Global"
        case "NSL": return "Local"
        case "NSA": return "Parameter"
        case "NSD": return "DDM"
        default: return ext
    }
}

; ===== FUNCIÓN: EXTRAER NOMBRE DE PROYECTO =====
ExtraerNombreProyecto(ruta) {
    ; Si la ruta está vacía o no existe, devolver mensaje apropiado
    if ruta = "" || !DirExist(ruta)
        return "Sin seleccionar"
		
    partes := StrSplit(ruta, "\")
	
    ; Buscar si es workspace o git
    for i, parte in partes {
        if InStr(parte, "workspace") or InStr(parte, "git") {
            if i < partes.Length
                return partes[i + 1]
        }
    }
	
	; Si no es workspace ni git, devolver la última carpeta del path
    if partes.Length > 0
        return partes[partes.Length]  ; Última carpeta del path
	
    return "Sin seleccionar"
}

; ===== FUNCIÓN: ARRAY TO STRING =====
ArrayToString(arr, separador := ", ") {
    resultado := ""
    for item in arr {
        resultado .= item . separador
    }
    return SubStr(resultado, 1, -StrLen(separador))
}

; ===== FUNCIÓN: PROCESAR DDM =====
ProcesarDDM(contenido) {
    ; Limpiar caracteres de retorno de carro
    contenido := StrReplace(contenido, "`r", "")
    
    lineas := StrSplit(contenido, "`n")
    lineasProcesadas := []
    
    for indice, linea in lineas {
        ; Convertir todo a mayúsculas
        linea := StrUpper(linea)
        
        ; Primera línea: contiene 'DEFAULT SEQUENCE:' -> agregarle 3 espacios (20 20 20)
        ; Va precedida por 0x0C (inicio del archivo) y NO lleva 0D 0A
        if indice = 1 {
            linea := StrReplace(linea, "DEFAULT SEQUENCE:", "DEFAULT SEQUENCE:   ")
            lineasProcesadas.Push(Chr(0x0C) . linea)
            continue
        }
        
        ; Segunda línea: descartar contenido original, escribir solo 20 (sin 0D 0A propio)
        ; La línea 3 aportará el 0D 0A que separa ambas
        if indice = 2 {
            lineasProcesadas.Push(Chr(0x20))
            continue
        }
        
        ; Tercera línea: mantener igual, lleva 0D 0A al principio
        if indice = 3 {
            lineasProcesadas.Push(Chr(0x0D) . Chr(0x0A) . linea)
            continue
        }
        
        ; Cuarta línea: reemplazar con encabezado específico, lleva 0D 0A al principio
        if indice = 4 {
            lineasProcesadas.Push(Chr(0x0D) . Chr(0x0A) . "TYL  DB  NAME                              F LENG  S D REMARKS")
            continue
        }
        
        ; Quinta línea: reemplazar con separador, lleva 0D 0A al principio
        if indice = 5 {
            lineasProcesadas.Push(Chr(0x0D) . Chr(0x0A) . "---  --  --------------------------------  - ----  - - --------------------")
            continue
        }
        
        ; Resto de líneas: procesar campos o eliminar comentarios
        lineaTrim := Trim(linea)
        
        ; Si es comentario (empieza con *), ELIMINAR
        if SubStr(lineaTrim, 1, 1) = "*" {
            continue
        }
        
        ; Si está vacía, mantener con 0D 0A al principio
        if lineaTrim = "" {
            lineasProcesadas.Push(Chr(0x0D) . Chr(0x0A) . linea)
            continue
        }
        
        ; Es un campo: reformatear y agregar 0D 0A al principio
        lineaReformateada := ReformatearCampoDDM(linea)
        lineasProcesadas.Push(Chr(0x0D) . Chr(0x0A) . lineaReformateada)
    }
    
    ; Agregar línea de cierre al final (sin 0D 0A al principio para no duplicar el salto)
    lineasProcesadas.Push("******DDM OUTPUT TERMINATED******" . Chr(0x0D) . Chr(0x0A))

    ; ── Pase de corrección: G/P sin subcampo de nivel 2 → reemplazar por espacio ──
    ; Las líneas de campos llevan Chr(0x0D).Chr(0x0A) al principio (2 bytes de prefijo),
    ; por lo que col 1 del contenido real está en posición 3 del string,
    ; y col 3 (nivel) está en posición 5.
    Loop lineasProcesadas.Length {
        idx  := A_Index
        lin  := lineasProcesadas[idx]

        ; Solo nos interesan líneas que empiecen con 0D 0A (líneas de campo)
        if StrLen(lin) < 5 || (SubStr(lin, 1, 1) != Chr(0x0D))
            continue

        col1 := SubStr(lin, 3, 1)   ; carácter en columna 1 (tras 0D 0A)
        if col1 != "G" && col1 != "P"
            continue

        ; Buscar el siguiente campo válido (que también empiece con 0D 0A)
        siguienteNivel := ""
        Loop (lineasProcesadas.Length - idx) {
            sig := lineasProcesadas[idx + A_Index]
            if StrLen(sig) >= 5 && SubStr(sig, 1, 1) = Chr(0x0D) {
                siguienteNivel := SubStr(sig, 5, 1)   ; col 3 del siguiente campo
                break
            }
        }

        ; Si el siguiente nivel no es "2", reemplazar G/P por espacio en col 1
        ; Estructura del string: pos1=0D, pos2=0A, pos3=col1, pos4=col2...
        if siguienteNivel != "2"
            lineasProcesadas[idx] := SubStr(lin, 1, 2) . " " . SubStr(lin, 4)
    }

    ; Reconstruir contenido: las líneas ya llevan su 0D 0A al principio,
    ; se concatenan directamente sin separador adicional
    contenidoFinal := ""
    for linea in lineasProcesadas {
        contenidoFinal .= linea
    }
    
    return contenidoFinal
}

; ===== FUNCIÓN: REFORMATEAR CAMPO DDM =====
ReformatearCampoDDM(lineaOriginal) {
    lineaTrim := Trim(lineaOriginal)
    
    ; Si la línea está vacía o es muy corta, devolverla tal cual
    if StrLen(lineaTrim) < 10 {
        return lineaOriginal
    }
    
    ; Extraer componentes usando posiciones aproximadas de NaturalONE
    ; Separar por espacios múltiples para identificar campos
    campos := []
    palabraActual := ""
    
    Loop Parse, lineaTrim {
        char := A_LoopField
        if char = " " || char = "`t" {
            if palabraActual != "" {
                campos.Push(palabraActual)
                palabraActual := ""
            }
        } else {
            palabraActual .= char
        }
    }
    
    ; Agregar última palabra si existe
    if palabraActual != "" {
        campos.Push(palabraActual)
    }
    
    ; Identificar componentes
    if campos.Length < 4 {
        return lineaOriginal
    }
    
    indice := 1
    tipoEstructura := ""
    nivel := ""
    nombreCorto := ""
    nombreLargo := ""
    tipoDato := ""
    longitud := ""
    opcion := ""
    descriptor := ""
    remarks := []
    
    ; Primer campo: puede ser G, M, o P
    if campos[1] = "G" || campos[1] = "M" || campos[1] = "P" {
        tipoEstructura := campos[1]
        indice := 2
    }
    
    ; Nivel (un dígito)
    if indice <= campos.Length && StrLen(campos[indice]) <= 2 && RegExMatch(campos[indice], "^\d+$") {
        nivel := campos[indice]
        indice++
    }
    
    ; Nombre corto (2 caracteres típicamente, puede tener @ o letras/números)
    if indice <= campos.Length {
        nombreCorto := campos[indice]
        indice++
    }
    
    ; Nombre largo (puede tener guiones)
    if indice <= campos.Length {
        nombreLargo := campos[indice]
        indice++
    }
    
    ; Tipo de dato (una letra: A, N, P, B, I, etc)
    if indice <= campos.Length && StrLen(campos[indice]) = 1 {
        tipoDato := campos[indice]
        indice++
    }
    
    ; Longitud (puede tener punto decimal como 4.0 o 7.0)
    if indice <= campos.Length && (RegExMatch(campos[indice], "^\d+\.?\d*$") || RegExMatch(campos[indice], "^\d+$")) {
        longitud := campos[indice]
        indice++
    }
    
    opcion := ""
    descriptor := ""
    remarks := []
    
    ; Determinar si este campo debe procesar opción y descriptor
    ; Solo si NO es una estructura contenedora (G o P sin tipo de dato)
    esEstructuraContenedora := (tipoEstructura = "G" || tipoEstructura = "P") && tipoDato = ""
    
    while indice <= campos.Length {
        campo := campos[indice]
        
        ; Si es una estructura contenedora (G o P sin tipo de dato), todos los flags van a remarks
        if esEstructuraContenedora {
            remarks.Push(campo)
        }
        ; Para campos con tipo de dato (M, P con tipo, o regulares), procesar opción y descriptor
        else {
            ; Opción: N (Null Value Suppression) o F (Fixed Storage)
            if campo = "N" || campo = "F" {
                if opcion = "" {
                    opcion := campo
                } else {
                    remarks.Push(campo)
                }
            }
            ; Descriptor: D (Descriptor) o S (Super/Sub descriptor)
            else if campo = "D" || campo = "S" {
                if descriptor = "" {
                    descriptor := campo
                } else {
                    remarks.Push(campo)
                }
            }
            ; Todo lo demás son remarks
            else {
                remarks.Push(campo)
            }
        }
        
        indice++
    }
    
    ; Construir remarks
    remarksTexto := ""
    for rem in remarks {
        if remarksTexto != ""
            remarksTexto .= " "
        remarksTexto .= rem
    }
    
    ; Construir línea reformateada con columnas exactas
    nuevaLinea := ""
    
    ; Columna 1: G/M/P o espacio
    nuevaLinea .= (tipoEstructura != "" ? tipoEstructura : " ")
    
    ; Columna 2: espacio
    nuevaLinea .= " "
    
    ; Columna 3: nivel
    nuevaLinea .= (nivel != "" ? nivel : " ")
    
    ; Columnas 4-5: espacios
    nuevaLinea .= "  "
    
    ; Columnas 6-7: nombre corto (2 caracteres)
    nuevaLinea .= SubStr(nombreCorto . "  ", 1, 2)
    
    ; Columnas 8-9: espacios
    nuevaLinea .= "  "
    
    ; Columnas 10-41: nombre largo (32 caracteres, pad derecha)
    nuevaLinea .= SubStr(nombreLargo . "                                ", 1, 32)
    
    ; Columnas 42-43: espacios
    nuevaLinea .= "  "
    
    ; Columna 44: tipo de dato
    nuevaLinea .= SubStr(tipoDato . " ", 1, 1)
    
    ; Columna 45: espacio
    nuevaLinea .= " "
    
    ; Columnas 46-49: longitud (4 caracteres, alineado derecha)
    if longitud != "" {
        ; Alinear a la derecha con espacios
        while StrLen(longitud) < 4 {
            longitud := " " . longitud
        }
        nuevaLinea .= SubStr(longitud, 1, 4)
    } else {
        nuevaLinea .= "    "
    }
    
    ; Columnas 50-51: espacios
    nuevaLinea .= "  "
    
    ; Columna 52: opción (debajo de S)
    nuevaLinea .= SubStr(opcion . " ", 1, 1)
    
    ; Columna 53: espacio
    nuevaLinea .= " "
    
    ; Columna 54: descriptor (debajo de D)
    nuevaLinea .= SubStr(descriptor . " ", 1, 1)
    
    ; Columna 55: espacio y columnas 56-75: remarks
    if remarksTexto != "" {
        nuevaLinea .= " " . remarksTexto
    }
    
    return nuevaLinea
}

; ===== FUNCIÓN: PROCESAR INICIALIZACIÓN =====
ProcesarInicializacion(textoCompleto, esConstante) {
    ; Extraer el contenido entre < y >
    inicio := InStr(textoCompleto, "<")
    fin := InStr(textoCompleto, ">", , inicio)
    
    if inicio = 0 || fin = 0
        return ""
    
    contenido := SubStr(textoCompleto, inicio + 1, fin - inicio - 1)
    
    ; Limpiar comillas simples y guiones de continuación
    contenido := StrReplace(contenido, "'", "")
    contenido := StrReplace(contenido, "-", "")
    contenido := StrReplace(contenido, "`n", "")
    contenido := StrReplace(contenido, "`r", "")
    contenido := Trim(contenido)
    
    return contenido
}

; ===== FUNCIÓN: GENERAR LÍNEA CON INICIALIZACIÓN =====
GenerarLineaConInicializacion(resultado, esConstante, valorInicializacion) {
    lineaBase := resultado.linea
    
    ; Extraer componentes
    prefijo := SubStr(lineaBase, 1, 4)           ; **DF o **DK
    arrayInfo := SubStr(lineaBase, 6, 2)         ; "  " o "I1", "I2", "I3"
    tipoDato := SubStr(lineaBase, 19, 1)         ; A, N, F, etc
    longitud := SubStr(lineaBase, 21, 3)         ; "  20" o " 2.7" etc
    nivel := SubStr(lineaBase, 25, 1)            ; 1, 2, 3, etc
    nombreYResto := SubStr(lineaBase, 26)        ; CHANGO o ARRAY(1:5,1:5,1:5)
    
    ; Dividir valor en fragmentos de 50 caracteres
    lineasInit := []
    posicion := 1
    longitudValor := StrLen(valorInicializacion)
    
    while posicion <= longitudValor {
        fragmento := SubStr(valorInicializacion, posicion, 50)
        lineasInit.Push(fragmento)
        posicion += 50
    }
    
    cantidadLineasI := lineasInit.Length
    totalLineas := 1 + cantidadLineasI  ; 1 **HS + N **I
    
    ; Construir línea **DF/DK modificada
    nuevaLinea := ""
    
    ; Col 1-4: prefijo
    nuevaLinea .= prefijo
    
    ; Col 5: C si CONST, espacio si INIT
    if esConstante {
        nuevaLinea .= "C"
    } else {
        nuevaLinea .= " "
    }
    
    ; Col 6-7: array info
    nuevaLinea .= arrayInfo
    
    ; Col 8: S (inicializado)
    nuevaLinea .= "S"
    
    ; Col 9-14: espacios (6 espacios)
    nuevaLinea .= "      "
    
    ; Col 15: número total de líneas
    nuevaLinea .= Format("{:1}", totalLineas)
    
    ; Col 16-18: espacios (3 espacios)
    nuevaLinea .= "   "
    
    ; Col 19: tipo de dato
    nuevaLinea .= tipoDato
    
    ; Col 20: espacio
    nuevaLinea .= " "
    
    ; Col 21-23: longitud
    nuevaLinea .= longitud
    
    ; Col 24: C si CONST, espacio si INIT
    if esConstante {
        nuevaLinea .= "C"
    } else {
        nuevaLinea .= " "
    }
    
    ; Col 25: nivel
    nuevaLinea .= nivel
    
    ; Col 26+: nombre y resto
    nuevaLinea .= nombreYResto
    
    ; Construir resultado completo
    resultadoFinal := nuevaLinea . "`n"
    
    ; Línea **HS
    resultadoFinal .= "**HS" . cantidadLineasI . "   `n"
    
    ; Líneas **I
    for fragmento in lineasInit {
        ; **I (col 1-3) + 22 espacios (col 4-25) + contenido (col 26+)
        lineaI := "**I                      " . fragmento
        
        ; Rellenar con espacios hasta columna 75
        while StrLen(lineaI) < 75 {
            lineaI .= " "
        }
        
        resultadoFinal .= lineaI . "`n"
    }
    
    ; Quitar último salto de línea
    resultadoFinal := SubStr(resultadoFinal, 1, -1)
    
    return resultadoFinal
}

; ===== FUNCIÓN: FILTRAR LIBRERÍA EN TIEMPO REAL (permite borrar todo) =====
FiltrarLibreria(ctrl, *) {
    global libreriaDefecto
    
    texto := ctrl.Value
    
    ; Guardar posición actual del cursor ANTES de cualquier modificación
    ; EM_GETSEL (0xB0): wParam y lParam reciben inicio y fin de selección
    cursorPos := SendMessage(0xB0, 0, 0, ctrl) & 0xFFFF   ; byte bajo = inicio de selección
    
    ; Convertir a mayúsculas
    textoUpper := StrUpper(texto)
    
    ; Filtrar caracteres válidos, rastreando cuántos se eliminaron ANTES del cursor
    nuevo := ""
    eliminadosAntesCursor := 0
    Loop Parse, textoUpper {
        posChar := A_Index   ; posición 1-based en el texto original
        esValido := RegExMatch(A_LoopField, "[A-Z0-9#-]")
        if esValido {
            nuevo .= A_LoopField
        } else {
            ; Si el carácter inválido estaba antes o en el cursor, ajustar
            if posChar <= cursorPos
                eliminadosAntesCursor++
        }
    }
    
    ; Máximo 8 caracteres
    if StrLen(nuevo) > 8
        nuevo := SubStr(nuevo, 1, 8)
    
    ; Primer carácter debe ser A-Z o #
    primerEliminado := 0
    if (nuevo != "" && !RegExMatch(SubStr(nuevo, 1, 1), "^[A-Z#]$")) {
        nuevo := SubStr(nuevo, 2)   ; quitar primer carácter inválido
        primerEliminado := 1
    }
    
    ; Aplicar el texto filtrado solo si cambió
    if (ctrl.Value != nuevo) {
        ctrl.Value := nuevo
    }
    
    ; Restaurar cursor en la posición correcta, descontando caracteres eliminados
    nuevoCursor := cursorPos - eliminadosAntesCursor - primerEliminado
    if nuevoCursor < 0
        nuevoCursor := 0
    if nuevoCursor > StrLen(nuevo)
        nuevoCursor := StrLen(nuevo)
    SendMessage(0xB1, nuevoCursor, nuevoCursor, ctrl)
    
    ; Actualizar estado del botón escanear según librería
    ActualizarEstadoBotonEscanear()
}

; ===== FUNCIÓN: FILTRAR NOMBRE EN FORMATO 8.3 EN TIEMPO REAL =====
FiltrarNombre83(ctrl, *) {
    texto := ctrl.Value

    ; Guardar posición del cursor ANTES de modificar
    cursorPos := SendMessage(0xB0, 0, 0, ctrl) & 0xFFFF

    ; Convertir a mayúsculas
    textoUpper := StrUpper(texto)

    ; Separar en nombre y extensión según el primer punto
    puntoPosOrig := InStr(textoUpper, ".")
    if puntoPosOrig > 0 {
        parteNombreOrig := SubStr(textoUpper, 1, puntoPosOrig - 1)
        parteExtOrig    := SubStr(textoUpper, puntoPosOrig + 1)
    } else {
        parteNombreOrig := textoUpper
        parteExtOrig    := ""
    }

    ; Caracteres válidos para nombre FAT 8.3 (excluye punto, se trata aparte)
    caracteresValidos := "[A-Z0-9!#$%&'()\-@^_`{}~]"

    ; Filtrar parte nombre (máx. 8 chars válidos), rastreando eliminados antes del cursor
    nuevoNombre := ""
    eliminadosAntesCursor := 0
    Loop Parse, parteNombreOrig {
        posChar := A_Index
        if RegExMatch(A_LoopField, caracteresValidos) {
            if StrLen(nuevoNombre) < 8
                nuevoNombre .= A_LoopField
            else if posChar <= cursorPos  ; char válido pero truncado por el límite
                eliminadosAntesCursor++
        } else {
            if posChar <= cursorPos
                eliminadosAntesCursor++
        }
    }

    ; Construir resultado con o sin extensión
    if puntoPosOrig > 0 {
        ; Hay punto: filtrar extensión (máx. 3 chars válidos)
        nuevoExt := ""
        offsetExt := puntoPosOrig  ; los chars de la extensión en el texto original empiezan aquí
        Loop Parse, parteExtOrig {
            posChar := offsetExt + A_Index
            if RegExMatch(A_LoopField, caracteresValidos) {
                if StrLen(nuevoExt) < 3
                    nuevoExt .= A_LoopField
                else if posChar <= cursorPos
                    eliminadosAntesCursor++
            } else {
                if posChar <= cursorPos
                    eliminadosAntesCursor++
            }
        }
        ; Eliminar segundos puntos (solo se permite uno)
        ; ya están excluidos porque el punto no está en caracteresValidos y se trató aparte
        nuevo := nuevoNombre . "." . nuevoExt
    } else {
        nuevo := nuevoNombre
    }

    ; Aplicar solo si cambió
    if ctrl.Value != nuevo
        ctrl.Value := nuevo

    ; Restaurar cursor
    nuevoCursor := cursorPos - eliminadosAntesCursor
    if nuevoCursor < 0
        nuevoCursor := 0
    if nuevoCursor > StrLen(nuevo)
        nuevoCursor := StrLen(nuevo)
    SendMessage(0xB1, nuevoCursor, nuevoCursor, ctrl)
}