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
global txtLibreria := ""  ; Campo de texto para la librería
global rutaOrigen := ""
global nombreCarpetaSeleccionada := ""
global mainGui := ""
global listView := ""
global esCarpetaPersonalizada := false  ; Flag para saber si es carpeta personalizada
global WORK_SLOT_SIZE := 68

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
; Si soloVerificar=true, devuelve el bool sin modificar versionNatural (para el indicador visual)
DetectarVersionNatural(soloVerificar := false) {
    global versionNatural
    
    rutaBase := A_ScriptDir . "\dos\NATURAL"
    
    if DirExist(rutaBase) {
        Loop Files, rutaBase . "\*", "D"
        {
            if RegExMatch(A_LoopFileName, "^\d+$") {
                if !soloVerificar
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
        ; MsgBox("Error al leer NATPARM.SAG:`n" . err.Message, "Error", "IconX")
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
    mainGui.Add("Text", "x60 y125", "Seleccione la carpeta de origen:")

    ; === Botones Workspace y Git uno al lado del otro ===
    margenIzq   := 60
    anchoBoton  := 240
    espacioEntre := 20
    altoBoton   := 50
    yBotones    := 160

    ; Botón Workspace (izquierda)
    btnWS := mainGui.Add("Button", "x" margenIzq " y" yBotones " w" anchoBoton " h" altoBoton, "📁  Workspace110")
    btnWS.Opt("+Background4CAF50")
    btnWS.SetFont("s11 cFFFFFF bold", "Segoe UI")

    ; Botón Git (derecha)
    xGit := margenIzq + anchoBoton + espacioEntre
    btnGit := mainGui.Add("Button", "x" xGit " y" yBotones " w" anchoBoton " h" altoBoton, "📁  Git / Repositorio")
    btnGit.Opt("+Background2196F3")
    btnGit.SetFont("s11 cFFFFFF bold", "Segoe UI")

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
    
    ; Agregar evento Change para filtrar en tiempo real
    txtLibreria.OnEvent("Change", FiltrarLibreria)
    
    mainGui.Add("Progress", "x60 y" (yLibreria+70) " w520 h2 BackgroundE0E0E0 -Smooth Disabled")

	; Botón ESCANEAR (ahora arriba)
    yEscanear := yLibreria + 90
    btnEscanear := mainGui.Add("Button", "x" margenIzq " y" yEscanear " w520 h60 Disabled", "🔍   ESCANEAR OBJETOS NATURAL")
    btnEscanear.Opt("+BackgroundFF5722")
    btnEscanear.SetFont("s13 bold cFFFFFF", "Segoe UI")

    mainGui.Add("Progress", "x60 y" (yEscanear+75) " w520 h2 BackgroundE0E0E0 -Smooth Disabled")

    ; Botón Gestionar Workfiles (ahora abajo)
    yWorkfiles := yEscanear + 90
    btnWorkfiles := mainGui.Add("Button", "x" margenIzq " y" yWorkfiles " w520 h42", "🗂  Administrar Archivos de Trabajo (WORK)")
    btnWorkfiles.Opt("+Background37474F")
    btnWorkfiles.SetFont("s10 cFFFFFF", "Segoe UI")

    mainGui.Add("Progress", "x60 y" (yWorkfiles+55) " w520 h2 BackgroundE0E0E0 -Smooth Disabled")

    ; Botón Salir
    ySalir := yWorkfiles + 75
    btnSalir := mainGui.Add("Button", "x" margenIzq " y" ySalir " w520 h45", "❌  Salir")
    btnSalir.Opt("+Background757575")
    btnSalir.SetFont("s11 cFFFFFF", "Segoe UI")

    ; Eventos
    btnWS.OnEvent("Click", (*) => SeleccionarRuta(rutasDefecto[1]))
    btnGit.OnEvent("Click", (*) => SeleccionarRuta(rutasDefecto[2]))
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
; Cada entrada de workfile en NATPARM.SAG ocupa 68 bytes:
;   - 8 bytes de nombre (padded con 0x00)
;   - 60 bytes restantes de la estructura interna del slot
; El bloque de workfiles comienza en un offset que se busca por patrón.
; Patrón de inicio del bloque WORK: 57 4F 52 4B (= "WORK" en ASCII)
; seguido del byte de cantidad máxima (0x1E = 30 por defecto).

global WORK_MAX_DEFAULT  := 30
global WORK_SLOT_SIZE    := 68   ; bytes por entrada en NATPARM.SAG
global WORK_NAME_SIZE    := 8    ; bytes de nombre por entrada
global workfilesData     := []   ; Array global con los workfiles cargados
global workfileOffsets   := []   ; Offset donde empieza el nombre de cada workfile
global workfileLengths   := []   ; Longitud original (bytes) de cada nombre
global workOffsetBase    := 0    ; (reservado, ya no se usa para leer)
global workMaxOffset     := 0    ; Offset del byte de máximo workfiles en NATPARM.SAG
global workMaxEntries    := 30   ; Máximo de workfiles leído de NATPARM.SAG

; ===== FUNCIÓN AUXILIAR: leer bytes hasta 0x00 desde un offset en un buffer =====
WF_LeerHasta00(buf, tamaño, offset) {
    path := ""
    pos  := offset
    while pos < tamaño {
        b := NumGet(buf, pos, "UChar")
        if b = 0
            break
        path .= Chr(b)
        pos++
    }
    return {path: path, len: pos - offset}
}

; ===== FUNCIÓN AUXILIAR: buscar rutas por patrón 3A 5C (:\) en el buffer =====
; Reglas:
;  - El byte en (pos-1) debe ser una letra A-Z/a-z (letra de unidad, sin Ñ)
;  - Se lee desde (pos-1) hasta encontrar 0x00, máximo 51 caracteres
;  - El nombre de archivo sigue las reglas 8.3 (SFN)
WF_BuscarRutasPorPatron(buf, tamaño) {
    rutas := []
    offsets := []
    lens := []

    i := 1   ; empezar en 1 para poder leer buf[i-1]
    while i < tamaño - 1 {
        b0 := NumGet(buf, i,     "UChar")   ; posible 3A (:)
        b1 := NumGet(buf, i + 1, "UChar")   ; posible 5C (\)
        if (b0 = 0x3A && b1 = 0x5C) {
            ; byte anterior debe ser letra A-Z o a-z (letra de unidad)
            letraUnidad := NumGet(buf, i - 1, "UChar")
            if ((letraUnidad >= 0x41 && letraUnidad <= 0x5A)   ; A-Z
             || (letraUnidad >= 0x61 && letraUnidad <= 0x7A)) { ; a-z
                ; La ruta comienza en (i-1): letra + :\ + resto
                offsetInicio := i - 1
                path := ""
                pos  := offsetInicio
                while (pos < tamaño && StrLen(path) < 51) {
                    b := NumGet(buf, pos, "UChar")
                    if b = 0
                        break
                    path .= Chr(b)
                    pos++
                }
                ; Truncar a 51 si es necesario
                if StrLen(path) > 51
                    path := SubStr(path, 1, 51)
                rutas.Push(path)
                offsets.Push(offsetInicio)
                lens.Push(pos - offsetInicio)
                ; Saltar al final de esta ruta para no reutilizar bytes
                i := pos + 1
                continue
            }
        }
        i++
    }
    return {rutas: rutas, offsets: offsets, lens: lens}
}

; ===== FUNCIÓN AUXILIAR: buscar rutas desde un offset inicial =====
; Igual que WF_BuscarRutasPorPatron pero arranca desde offsetInicial.
; lens[] almacena el ESPACIO REAL disponible en cada slot:
;   = longitud del texto + bytes nulos de relleno que siguen (hasta 51 total)
WF_BuscarRutasPorPatron_Desde(buf, tamaño, offsetInicial) {
    rutas   := []
    offsets := []
    lens    := []

    i := offsetInicial + 1   ; +1 para poder leer buf[i-1]
    while i < tamaño - 1 {
        b0 := NumGet(buf, i,     "UChar")   ; posible 3A (:)
        b1 := NumGet(buf, i + 1, "UChar")   ; posible 5C (\)
        if (b0 = 0x3A && b1 = 0x5C) {
            letraUnidad := NumGet(buf, i - 1, "UChar")
            if ((letraUnidad >= 0x41 && letraUnidad <= 0x5A)   ; A-Z mayúsculas
             || (letraUnidad >= 0x61 && letraUnidad <= 0x7A)) { ; a-z minúsculas
                offsetRuta := i - 1
                path := ""
                pos  := offsetRuta
                ; Leer el texto hasta 0x00 (máx 51 chars)
                while (pos < tamaño && StrLen(path) < 51) {
                    b := NumGet(buf, pos, "UChar")
                    if b = 0
                        break
                    path .= Chr(b)
                    pos++
                }
                ; Truncar a 51 si es necesario
                if StrLen(path) > 51
                    path := SubStr(path, 1, 51)
                    
                textoLen := pos - offsetRuta   ; longitud del texto
                
                ; Contar bytes nulos de relleno que siguen al terminador
                ; para conocer el espacio total disponible en el slot
                posNull := pos   ; pos apunta al primer 0x00 terminador
                espacioTotal := textoLen
                while (posNull < tamaño && espacioTotal < 51) {
                    if NumGet(buf, posNull, "UChar") != 0
                        break
                    espacioTotal++
                    posNull++
                }
                
                rutas.Push(path)
                offsets.Push(offsetRuta)
                lens.Push(espacioTotal)   ; espacio real = texto + relleno nulo
                i := pos + 1
                continue
            }
        }
        i++
    }
    return {rutas: rutas, offsets: offsets, lens: lens}
}

; ===== FUNCIÓN: LEER WORKFILES DE NATPARM.SAG (por patrón 3A 5C = :\) =====
; Algoritmo:
;   1. Localizar el patrón 54 00 04 00 → byte siguiente = número máximo de WF
;   2. Localizar el primer workfile con [maxByte] 00 35 00 desde ese offset
;      → a partir de ahí buscar rutas por patrón 3A 5C (:\)
;   3. Cada ruta comienza en la letra de unidad anterior a :\ y termina en 0x00
;      (máximo 51 caracteres, nombres en formato 8.3)
LeerWorkfilesDeNATPARM(maxEntries := 30) {
    global versionNatural, workfileOffsets, workfileLengths, workMaxOffset

    rutaNATPARM := A_ScriptDir . "\dos\NATURAL\" . versionNatural . "\PROF\NATPARM.SAG"
    resultado        := []
    workfileOffsets  := []
    workfileLengths  := []

    if !FileExist(rutaNATPARM)
        return resultado

    ; Leer todo el archivo en un buffer
    archivo := FileOpen(rutaNATPARM, "r")
    if !archivo
        return resultado
    tamaño := archivo.Length
    buf    := Buffer(tamaño)
    archivo.RawRead(buf, tamaño)
    archivo.Close()

    ; ── 1. Localizar el byte de máximo (patrón 00 54 00 04, valor en byte anterior) ──
    if workMaxOffset = 0 {
        Loop tamaño - 4 {
            j := A_Index - 1
            if (NumGet(buf, j,     "UChar") = 0x00
             && NumGet(buf, j + 1, "UChar") = 0x54
             && NumGet(buf, j + 2, "UChar") = 0x00
             && NumGet(buf, j + 3, "UChar") = 0x04) {
                workMaxOffset := j - 1   ; byte justo ANTES del patrón
                break
            }
        }
    }
    if workMaxOffset = 0
        return resultado

    maxByte := NumGet(buf, workMaxOffset, "UChar")

    ; ── 2. Localizar inicio del bloque: [maxByte] 00 35 00 ───────────────────
    inicioBloque := 0
    Loop tamaño - workMaxOffset - 4 {
        i := workMaxOffset + A_Index - 1
        if (NumGet(buf, i,     "UChar") = maxByte
         && NumGet(buf, i + 1, "UChar") = 0x00
         && NumGet(buf, i + 2, "UChar") = 0x35
         && NumGet(buf, i + 3, "UChar") = 0x00) {
            inicioBloque := i + 4   ; primer byte de la primera entrada
            break
        }
    }
    if inicioBloque = 0
        return resultado

    ; ── 3. Tamaño fijo de cada slot (según observaciones) ───────────────────
    global WORK_SLOT_SIZE := 68   ; Definir esta constante si no existe

    ; ── 4. Recorrer todos los slots secuencialmente ─────────────────────────
    offsetActual := inicioBloque
    Loop maxEntries {
        offsetPath := offsetActual

        ; Leer la cadena hasta 51 caracteres o hasta encontrar 0x00
        ruta := ""
        pos := offsetActual
        while pos < offsetActual + 51 && pos < tamaño {
            byte := NumGet(buf, pos, "UChar")
            if byte = 0
                break
            ruta .= Chr(byte)
            pos++
        }

        resultado.Push(ruta)
        workfileOffsets.Push(offsetPath)
        workfileLengths.Push(WORK_SLOT_SIZE)   ; Espacio total del slot

        offsetActual += WORK_SLOT_SIZE
        if offsetActual >= tamaño
            break
    }

    ; Rellenar con vacíos si no se llegó al máximo
    while resultado.Length < maxEntries {
        resultado.Push("")
        workfileOffsets.Push(0)
        workfileLengths.Push(WORK_SLOT_SIZE)
    }

    return resultado
}

; ===== FUNCIÓN: ESCRIBIR UN WORKFILE EN NATPARM.SAG =====
; Escribe el nuevo path en el offset previamente localizado por LeerWorkfilesDeNATPARM.
; Regla: solo sobreescribe hasta la longitud original (rellena con 0x00 si es más corto,
; trunca silenciosamente si es más largo).
EscribirWorkfileEnNATPARM(numero, rutaCompleta) {
    global versionNatural, workfileOffsets, workfileLengths

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

    offsetPath  := workfileOffsets[numero]
    longitudSlot := workfileLengths[numero]   ; Debe ser WORK_SLOT_SIZE (68)

    rutaLimpia := Trim(rutaCompleta)
    if StrLen(rutaLimpia) + 1 > longitudSlot {   ; +1 por el terminador nulo
        MsgBox("La ruta completa excede la longitud máxima permitida para este slot.`n`n"
             . "Máximo: " . (longitudSlot - 1) . " caracteres`n"
             . "Su ruta: " . StrLen(rutaLimpia) . " caracteres", 
             "Error - Ruta demasiado larga", "IconX")
        return false
    }

    try {
        archivo := FileOpen(rutaNATPARM, "rw")
        if !archivo
            throw Error("No se pudo abrir NATPARM.SAG para escritura")

        archivo.Seek(offsetPath)

        ; Escribir la ruta byte a byte
        Loop StrLen(rutaLimpia) {
            char := SubStr(rutaLimpia, A_Index, 1)
            archivo.WriteUChar(Ord(char))
        }

        ; Terminador nulo
        archivo.WriteUChar(0x00)

        ; Rellenar el resto del slot con ceros
        bytesEscritos := StrLen(rutaLimpia) + 1
        if bytesEscritos < longitudSlot {
            Loop longitudSlot - bytesEscritos {
                archivo.WriteUChar(0x00)
            }
        }

        archivo.Close()
        return true

    } catch as err {
        MsgBox("Error al escribir workfile en NATPARM.SAG:`n" . err.Message, "Error", "IconX")
        return false
    }
}

; ===== FUNCIÓN: LEER MÁXIMO DE WORKFILES (patrón 00 54 00 04) =====
; Busca el patrón 00 54 00 04 en NATPARM.SAG.
; El byte inmediatamente ANTERIOR a ese patrón es el número máximo de Workfiles,
; almacenado en formato hexadecimal (ej: 0x1E = 30, 0x20 = 32).
; Guarda el offset de ese byte en workMaxOffset para poder escribirlo después.
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
        ; Buscar patrón 00 54 00 04 (no se repite en el archivo)
        ; El valor buscado está en el byte ANTERIOR al patrón (offset - 1)
        Loop tamaño - 4 {
            i := A_Index - 1
            if (NumGet(buf, i,     "UChar") = 0x00
             && NumGet(buf, i + 1, "UChar") = 0x54
             && NumGet(buf, i + 2, "UChar") = 0x00
             && NumGet(buf, i + 3, "UChar") = 0x04) {
                workMaxOffset := i - 1   ; byte justo ANTES del patrón
                ; NumGet con "UChar" devuelve el byte como entero decimal sin signo.
                ; Ejemplo: byte 0x1E en el archivo → valor = 30, byte 0x20 → valor = 32.
                valor := NumGet(buf, workMaxOffset, "UChar")
                ; Rango válido: 0 a 32 (0x00 a 0x20)
                if (valor >= 0 && valor <= 32)
                    return valor
                ; Byte fuera de rango: usar default
                return WORK_MAX_DEFAULT
            }
        }
        ; Patrón 00 54 00 04 no encontrado en el archivo
        return WORK_MAX_DEFAULT
    } catch {
        return WORK_MAX_DEFAULT
    }
}

; ===== FUNCIÓN: ESCRIBIR MÁXIMO DE WORKFILES EN EL OFFSET DETECTADO =====
; Escribe en el byte anterior al patrón 00 54 00 04, previamente localizado por LeerMaxWorkfilesDeNATPARM.
EscribirMaxWorkfilesEnNATPARM(valor) {
    global versionNatural, workMaxOffset
    rutaNATPARM := A_ScriptDir . "\dos\NATURAL\" . versionNatural . "\PROF\NATPARM.SAG"
    if !FileExist(rutaNATPARM) {
        MsgBox("No se encontró NATPARM.SAG en:`n" . rutaNATPARM, "Error", "IconX")
        return false
    }
    if workMaxOffset = 0 {
        ; No se localizó el patrón 54 00 04 00 en NATPARM.SAG. No se puede guardar el máximo.
        return false
    }
    try {
        archivo := FileOpen(rutaNATPARM, "rw")
        if !archivo
            throw Error("No se pudo abrir NATPARM.SAG para escritura")
        archivo.Seek(workMaxOffset)
        archivo.WriteUChar(valor)   ; Guardado como byte hexadecimal
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

; Nombres reservados del sistema que no pueden usarse
global WF_NOMBRES_RESERVADOS := ["CON","PRN","AUX","NUL",
    "COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9",
    "LPT1","LPT2","LPT3","LPT4","LPT5","LPT6","LPT7","LPT8","LPT9"]

; Caracteres especiales permitidos en nombres 8.3
; (además de letras A-Z y dígitos 0-9)
global WF_CHARS_PERMITIDOS := "_-!@#$%&'(){}^~"

; ===== FUNCIÓN: Validar un segmento de nombre 8.3 (archivo o carpeta) =====
; Devuelve "" si es válido, o un mensaje de error si no lo es.
WF_ValidarSegmento83(segmento) {
    global WF_NOMBRES_RESERVADOS, WF_CHARS_PERMITIDOS

    if segmento = ""
        return "El nombre no puede estar vacío."

    ; Separar nombre base y extensión
    puntoPos := InStr(segmento, ".")
    if puntoPos > 0 {
        base := SubStr(segmento, 1, puntoPos - 1)
        ext  := SubStr(segmento, puntoPos + 1)
    } else {
        base := segmento
        ext  := ""
    }

    ; El nombre no puede empezar con punto
    if SubStr(segmento, 1, 1) = "."
        return "El nombre no puede empezar con punto: '" . segmento . "'"

    ; Longitud del nombre base: 1–8 caracteres
    if StrLen(base) = 0
        return "El nombre base no puede estar vacío en '" . segmento . "'"
    if StrLen(base) > 8
        return "El nombre base excede 8 caracteres: '" . base . "' (" . StrLen(base) . " caracteres)"

    ; Longitud de la extensión: 0–3 caracteres
    if StrLen(ext) > 3
        return "La extensión excede 3 caracteres: '" . ext . "' (" . StrLen(ext) . " caracteres)"

    ; Caracteres permitidos en nombre base y extensión
    for parte in [base, ext] {
        if parte = ""
            continue
        Loop StrLen(parte) {
            c := SubStr(parte, A_Index, 1)
            ; Letra A-Z/a-z, dígito, o carácter especial permitido
            if !RegExMatch(c, "i)[A-Z0-9]") && !InStr(WF_CHARS_PERMITIDOS, c) {
                return "Carácter no permitido '" . c . "' en '" . parte . "'"
            }
        }
    }

    ; Nombre reservado (comparar en mayúsculas, con y sin extensión)
    baseMayus := StrUpper(base)
    for reservado in WF_NOMBRES_RESERVADOS {
        if baseMayus = reservado
            return "Nombre reservado del sistema: '" . segmento . "'"
    }

    return ""   ; Válido
}

; ===== FUNCIÓN: Validar una ruta completa Windows 3.1 =====
; Formato esperado: [Letra]:\[carpeta\...]rchivo.ext
; Devuelve "" si es válida, o un mensaje de error descriptivo.
WF_ValidarRuta31(ruta) {
    ; Vacío = borrar entrada = permitido
    if ruta = ""
        return ""

    ; Longitud máxima de la ruta completa: 51 caracteres
    longitudRuta := StrLen(ruta)
    if longitudRuta > 51
        return "La ruta excede 51 caracteres (" . longitudRuta . "/51)."

    ; No se permiten espacios
    if InStr(ruta, " ")
        return "Los espacios no están permitidos en rutas de Windows 3.1."

    ; No se permiten barras hacia adelante
    if InStr(ruta, "/")
        return "Use barras invertidas (\) en lugar de barras (/)."

    ; Debe comenzar con letra de unidad (A-Z/a-z) seguida de :\
    primerChar  := SubStr(ruta, 1, 1)
    segundoChar := SubStr(ruta, 2, 1)
    tercerChar  := SubStr(ruta, 3, 1)
    esLetra     := RegExMatch(primerChar, "i)[A-Z]")
    if !esLetra || segundoChar != ":" || tercerChar != Chr(92)
        return "La ruta debe comenzar con una letra de unidad seguida de :\ (ej. C:\CARPETA\ARCH.EXT)."


    ; Separar en segmentos por barra invertida (quitar la letra de unidad + :\)
    sinUnidad := SubStr(ruta, 4)
    segmentos := StrSplit(sinUnidad, Chr(92))

    ; Validar cada segmento (carpetas y archivo final) con reglas 8.3
    for segmento in segmentos {
        if segmento = ""
            continue   ; ignorar dobles barras
        errorSeg := WF_ValidarSegmento83(segmento)
        if errorSeg != ""
            return "'" . segmento . "': " . errorSeg
    }

    return ""   ; Ruta válida
}

MostrarGestionWorkfiles() {
    global versionNatural, workfilesData, workMaxEntries, mainGui

    ; Leer el máximo real desde la dirección 0x3D9
    workMaxEntries := LeerMaxWorkfilesDeNATPARM()

    ; Siempre leer los 32 slots posibles para que el array nunca quede vacío
    workfilesData := LeerWorkfilesDeNATPARM(32)
    while workfilesData.Length < 32
        workfilesData.Push("")



    ; ── Ocultar ventana principal (igual que al escanear) ─────────────────────
    mainGui.Hide()

    ; ── Crear ventana ─────────────────────────────────────────────────────────
    wfGui := Gui("-MaximizeBox", "Administración de Archivos de Trabajo (WORK)")
    wfGui.BackColor := "F0F0F0"
    wfGui.SetFont("s11 c222222", "Segoe UI")

    ; ── Barra de título interna (estilo imagen) ────────────────────────────────
    ; Ancho de ventana: 950px
    wfGui.SetFont("s11 bold cFFFFFF", "Segoe UI")
    encTitle := wfGui.Add("Text", "x0 y0 w600 h28 Background1A237E", "")
    titleLbl := wfGui.Add("Text", "x10 y6 w370 Background1A237E cFFFFFF", "Número Máximo de Archivos de Trabajo [WORK]")

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
    wfGui.Add("Text", "x150 y40", "Ruta:")

    wfGui.SetFont("s11 c222222", "Segoe UI")
    txtFilename := wfGui.Add("Edit", "x230 y38 w350 h22 Border", "")

    ; ── Separador ─────────────────────────────────────────────────────────────
    wfGui.Add("Progress", "x15 y68 w565 h2 BackgroundCCCCCC -Smooth Disabled")

    ; ── Encabezado de tabla ────────────────────────────────────────────────────
    wfGui.SetFont("s10 bold c222222", "Segoe UI")
    wfGui.Add("Text", "x15 y75", "Nr")
    wfGui.Add("Text", "x60 y75", "Ruta en NATPARM.SAG")
    ; (sin columna Archivo Importado)

    ; ── ListView ──────────────────────────────────────────────────────────────
    wfLV := wfGui.Add("ListView",
        "x15 y95 w565 h280 -Multi -Hdr Grid NoSortHdr",
        ["Nr", "Ruta en NATPARM.SAG"])

    ; Poblar lista
    Loop workMaxEntries {
        nombre := (workfilesData.Length >= A_Index) ? workfilesData[A_Index] : ""
        wfLV.Add("", A_Index, nombre)
    }

    wfLV.ModifyCol(1, "45 Left")    ; Nr
    wfLV.ModifyCol(2, "500 Left")   ; Ruta NATPARM
    ; (sin col3)

    ; Seleccionar fila 1 al inicio (solo si hay filas)
    if workMaxEntries > 0
        wfLV.Modify(1, "Select Focus Vis")

    ; ── Separador ─────────────────────────────────────────────────────────────
    wfGui.Add("Progress", "x15 y383 w565 h2 BackgroundCCCCCC -Smooth Disabled")

    ; ── Botones centrados ─────────────────────────────────────────────────────
    ; Ventana 600px — 3 botones 110px + 2 gaps 20px = 370px → margen = 115
    wfGui.SetFont("s10 c222222", "Segoe UI")
    btnUpdate  := wfGui.Add("Button", "x115 y393 w110 h32", "&Actualizar")
    btnImport  := wfGui.Add("Button", "x245 y393 w110 h32", Chr(0x1F4E5) . "  &Importar")
    btnClose   := wfGui.Add("Button", "x375 y393 w110 h32", "&Cerrar")
    ; ── Etiqueta de estado (debajo de los botones, centrada) ─────────────────
    wfGui.SetFont("s10 bold c444444", "Segoe UI")
    infoLbl := wfGui.Add("Text", "x15 y435 w565 Center", "")

    ; ── Eventos ───────────────────────────────────────────────────────────────
    wfLV.OnEvent("Click",      ActualizarCamposWF)
    wfLV.OnEvent("ItemSelect", ActualizarCamposWF)
    btnUpdate.OnEvent("Click", GuardarWorkfile)
    btnImport.OnEvent("Click", ImportarWorkfile)
    btnClose.OnEvent("Click",  CerrarWfGui)
    wfGui.OnEvent("Close",     CerrarWfGui)

    wfGui.Show("Center w600 h465")
    ; Evitar que txtMaxWF aparezca seleccionado en azul al abrir
    SendMessage(0xB1, -1, 0, txtMaxWF)   ; EM_SETSEL: deseleccionar

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
        txtFilename.Value := wfLV.GetText(fila, 2)
        ; Mover caret al final sin resaltar (evita azul de selección)
        SendMessage(0xB1, 0, -1, txtFilename)   ; EM_SETSEL: seleccionar todo
        SendMessage(0xB1, -1, 0, txtFilename)   ; EM_SETSEL: deseleccionar (caret al inicio)
    }

    ; ── Función interna: importar Workfile desde el sistema anfitrión ────────
    ImportarWorkfile(*) {
        ; Obtener fila seleccionada
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

        ; ── Seleccionar archivo desde el sistema anfitrión ───────────────────
        archivoOrigen := FileSelect(1, "", "Seleccionar Workfile para importar", "Todos los archivos (*.*)")
        if archivoOrigen = ""
            return   ; cancelado

        ; ── Extraer SOLO el nombre del archivo (sin la ruta) ─────────────────
        nombreArchivo := RegExReplace(archivoOrigen, ".*[\\/]", "")
        
        ; ── Validar que el nombre del archivo cumpla reglas 8.3 ──────────────
        errorValidacion := WF_ValidarSegmento83(nombreArchivo)
        if errorValidacion != "" {
            MsgBox("El nombre del archivo no cumple las reglas 8.3 de Windows 3.1:`n`n"
                 . errorValidacion . "`n`n"
                 . "Nombre: " . nombreArchivo, 
                 "Error - Nombre de archivo inválido", "IconX")
            return
        }

        ; ── Carpeta destino: .\dos\NATURAL\[version]\WF ──────────────────────
        carpetaWF := A_ScriptDir . "\dos\NATURAL\" . versionNatural . "\WF"
        if !DirExist(carpetaWF) {
            try {
                DirCreate(carpetaWF)
            } catch as errDir {
                MsgBox("No se pudo crear la carpeta destino:`n" . carpetaWF . "`n`n" . errDir.Message, "Error", "IconX")
                return
            }
        }

        ; ── Verificar si el archivo ya existe en destino ─────────────────────
        rutaDestino := carpetaWF . "\" . nombreArchivo
        if FileExist(rutaDestino) {
            respuesta := MsgBox("El archivo ya existe en la carpeta WF.`n`n"
                              . "¿Desea sobrescribirlo?", 
                              "Archivo existente", "YesNo Icon?")
            if respuesta = "No"
                return
        }

        ; ── Copiar el archivo a la carpeta WF ────────────────────────────────
        try {
            FileCopy(archivoOrigen, rutaDestino, 1)  ; 1 = sobrescribir
        } catch as errCopy {
            MsgBox("Error al copiar el archivo:`n" . errCopy.Message, "Error", "IconX")
            return
        }

        ; ── Construir la ruta COMPLETA que debe guardarse en NATPARM ─────────
        ;    (con unidad y carpeta, exactamente como la espera Natural)
        rutaCompleta := "C:\NATURAL\" . versionNatural . "\WF\" . nombreArchivo

        ; ── Validar que la ruta completa quepa en el slot ────────────────────
        if workfileLengths.Length >= fila && workfileLengths[fila] > 0 {
            if StrLen(rutaCompleta) + 1 > workfileLengths[fila] {
                MsgBox("La ruta completa excede la longitud máxima permitida para este slot.`n`n"
                     . "Máximo: " . (workfileLengths[fila] - 1) . " caracteres`n"
                     . "Ruta: " . rutaCompleta . " (" . StrLen(rutaCompleta) . " caracteres)", 
                     "Error - Ruta demasiado larga", "IconX")
                return
            }
        }

        ; ── Escribir la RUTA COMPLETA en NATPARM.SAG ─────────────────────────
        wfGuardado := false
        if workfileOffsets.Length >= fila && workfileOffsets[fila] > 0 {
            wfGuardado := EscribirWorkfileEnNATPARM(fila, rutaCompleta)
        } else {
            MsgBox("No se encontró el offset para el Workfile #" . fila . " en NATPARM.SAG.`n"
                 . "No se actualizará la entrada.", 
                 "Advertencia", "Icon!")
        }

        ; ── Actualizar ListView y memoria ────────────────────────────────────
        while workfilesData.Length < fila
            workfilesData.Push("")
        workfilesData[fila] := rutaCompleta
        
        wfLV.Modify(fila, "Col2", rutaCompleta)
        
        txtNumber.Value   := fila
        txtFilename.Value := rutaCompleta
        SendMessage(0xB1, -1, 0, txtFilename)

        ; ── Mensaje de estado ────────────────────────────────────────────────
        if wfGuardado {
            infoLbl.Value := "✓ Workfile #" . fila . " importado → " . nombreArchivo . " guardado con ruta completa"
            infoLbl.Opt("c006600")
        } else {
            infoLbl.Value := "⚠ Archivo copiado a WF\ pero no se actualizó NATPARM.SAG"
            infoLbl.Opt("c996600")
        }
    }

    ; ── Función interna: guardar (Update) ─────────────────────────────────────
    GuardarWorkfile(*) {

        ; ── 1. Validar y guardar el MÁXIMO (campo txtMaxWF) ───────────────────
        maxStr := Trim(txtMaxWF.Value)
        if !RegExMatch(maxStr, "^\d+$") || Integer(maxStr) < 0 || Integer(maxStr) > 32 {
            MsgBox("El valor máximo '" . maxStr . "' no es válido.`n"
                 . "Debe ser un número entero entre 0 y 32.", "Error — Máximo de Workfiles", "IconX")
            txtMaxWF.Focus()
            return
        }
        nuevoMax := Integer(maxStr)

        ; Guardar en 0x3D9 como byte hexadecimal
        maxGuardado := EscribirMaxWorkfilesEnNATPARM(nuevoMax)

        ; Si el máximo cambió, actualizar la lista y el array en memoria
        if nuevoMax != workMaxEntries {
            workMaxEntries := nuevoMax
            ; Asegurar que workfilesData tiene suficientes slots
            while workfilesData.Length < workMaxEntries
                workfilesData.Push("")
            wfLV.Delete()
            Loop workMaxEntries {
                wfLV.Add("", A_Index, workfilesData[A_Index])
            }
            wfLV.ModifyCol(1, "45 Left")
            wfLV.ModifyCol(2, "500 Left")
            if workMaxEntries > 0 {
                wfLV.Modify(1, "Select Focus Vis")
                txtNumber.Value   := 1
                txtFilename.Value := workfilesData[1]
            }
        }

        ; ── 2. Validar y guardar el FILENAME de la fila seleccionada ──────────
        filaStr := Trim(txtNumber.Value)
        if !RegExMatch(filaStr, "^\d+$") || Integer(filaStr) < 1 || Integer(filaStr) > workMaxEntries {
            ; Si el número de fila no aplica (p.ej. max=0) solo reportar el max
            if maxGuardado {
                infoLbl.Value := "✓ Actualizado correctamente"
                infoLbl.Opt("c006600")
            } else {
                infoLbl.Value := "✗ Error al actualizar"
                infoLbl.Opt("c800000")
            }
            return
        }

        fila  := Integer(filaStr)
        path  := Trim(txtFilename.Value)

        ; Validar ruta con reglas completas de Windows 3.1 (formato 8.3)
        errorRuta := WF_ValidarRuta31(path)
        if errorRuta != "" {
            MsgBox("Ruta inválida para Windows 3.1:`n`n" . errorRuta
                 . "`n`nEjemplo válido: C:\NATAPPS\WORK1.WRK", "Error — Ruta de Workfile", "IconX")
            txtFilename.Focus()
            return
        }

        ; Verificar que no exceda la longitud del slot original (para no corromper NATPARM.SAG)
        longitudOrig := (workfileLengths.Length >= fila && workfileLengths[fila] > 0)
                       ? workfileLengths[fila] : 51
        if StrLen(path) > longitudOrig {
            MsgBox("La ruta excede la longitud máxima permitida para este slot ("
                 . longitudOrig . " caracteres).", "Error — Ruta demasiado larga", "IconX")
            txtFilename.Focus()
            return
        }

        nombre := path   ; Alias para compatibilidad con el resto del código

        ; Actualizar ListView y memoria
        wfLV.Modify(fila, "", fila, nombre)
        ; Extender el array si es necesario (evita "Invalid index")
        while workfilesData.Length < fila
            workfilesData.Push("")
        workfilesData[fila] := nombre

        ; Escribir workfile en NATPARM.SAG
        wfGuardado := false
        if workfileOffsets.Length >= fila && workfileOffsets[fila] > 0
            wfGuardado := EscribirWorkfileEnNATPARM(fila, nombre)

        ; ── Mensaje de estado combinado ────────────────────────────────────────
        if maxGuardado && (workfileOffsets.Length >= fila && workfileOffsets[fila] > 0 ? wfGuardado : true) {
            infoLbl.Value := "✓ Actualizado correctamente"
            infoLbl.Opt("c006600")
        } else {
            infoLbl.Value := "✗ Error al actualizar"
            infoLbl.Opt("c800000")
        }

        ; Avanzar selección a la siguiente fila
        if workMaxEntries > 0 {
            siguiente := Min(fila + 1, workMaxEntries)
            wfLV.Modify(siguiente, "Select Focus Vis")
            txtNumber.Value   := siguiente
            ; workfilesData siempre tiene al menos 32 slots, acceso seguro
            txtFilename.Value := workfilesData[siguiente]
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
SeleccionarRuta(ruta) {
    global rutaOrigen, txtRutaActual, btnEscanear, esCarpetaPersonalizada, nombreCarpetaSeleccionada
    
    if DirExist(ruta) {
        rutaOrigen := ruta
        txtRutaActual.Value := ruta
        ActualizarEstadoBotonEscanear()
        esCarpetaPersonalizada := false  ; Es Workspace o Git, NO personalizada
        
        ; Extraer nombre de la carpeta para mostrar después
        partes := StrSplit(ruta, "\")
        if partes.Length > 0
            nombreCarpetaSeleccionada := partes[partes.Length]
        else
            nombreCarpetaSeleccionada := ruta
            
        MsgBox("Ruta seleccionada correctamente:`n" . ruta, "Éxito", "Icon! 4096")
    } else {
        MsgBox("La ruta no existe:`n" . ruta . "`n`n¿Desea seleccionar una carpeta personalizada?", "Error", "IconX 4")
        if MsgBox("", "", "YesNo Icon?") = "Yes"
            SeleccionarRutaPersonalizada()
    }
}

; ===== FUNCIÓN: SELECCIONAR RUTA PERSONALIZADA =====
SeleccionarRutaPersonalizada() {
    global rutaOrigen, txtRutaActual, btnEscanear, esCarpetaPersonalizada, nombreCarpetaSeleccionada
    
    carpeta := DirSelect("*", 3, "Seleccione la carpeta de origen (workspace110 o git)")
    
    if carpeta != "" {
        rutaOrigen := carpeta
        txtRutaActual.Value := carpeta
        ActualizarEstadoBotonEscanear()
        esCarpetaPersonalizada := true  ; Es carpeta personalizada
        
        ; Extraer nombre de la carpeta seleccionada
        partes := StrSplit(carpeta, "\")
        if partes.Length > 0
            nombreCarpetaSeleccionada := partes[partes.Length]
        else
            nombreCarpetaSeleccionada := carpeta
            
        MsgBox("Ruta seleccionada correctamente:`n" . carpeta, "Éxito", "Icon! 4096")
    }
}

; ===== FUNCIÓN: ESCANEAR ARCHIVOS =====
EscanearArchivos() {
    global archivosEncontrados, rutaOrigen, mainGui, txtLibreria
    
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
        MsgBox("No se encontraron objetos Natural en la ruta seleccionada.`n`nExtensiones buscadas: " . ArrayToString(extensionesNatural), "Sin resultados", "Icon!")
        mainGui.Show()
        return
    }
    
    ; Mostrar ventana de selección
    MostrarVentanaSeleccion()
}

MostrarVentanaSeleccion() {
    global archivosEncontrados, listView, txtContador, esCarpetaPersonalizada

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
    selGui.BackColor := "F8F9FA"          ; Fondo muy claro
    selGui.SetFont("s12 c333333", "Segoe UI")

    ; Título con contador de válidos
    selGui.SetFont("s14 bold c0D47A1", "Segoe UI")
    selGui.Add("Text", "x30 y20 w740 Center", "Objetos Natural encontrados (" . contadorValidos . ")")

    ; Instrucción clara
    selGui.SetFont("s11 c555555", "Segoe UI")
    selGui.Add("Text", "x30 y60 w740", "Marque los objetos que desea copiar.")

    ; ────────────────────────────────────────────
    ;               LISTVIEW PRINCIPAL
    ; ────────────────────────────────────────────
    ; ListView con columnas fijas (PROYECTO aplica tanto a Workspace/Git como a carpeta personalizada)
    listView := selGui.Add("ListView", "x30 y100 w740 h400 Checked Grid", ["    NOMBRE", "TIPO", "TAMAÑO", "FECHA", "PROYECTO"])
    listView.SetFont("s10", "Segoe UI")

    ; Poblar ListView con validación de nombres
    for archivo in archivosEncontrados {
        ; Validar que el nombre sea válido para Natural (sin extensión)
        nombreSinExt := StrReplace(archivo.nombre, "." . archivo.ext, "")
        if !ValidarNombreNatural(nombreSinExt, "." . archivo.ext) {
            continue  ; Saltar archivos con nombres inválidos
        }
        
        tamanoKB := Round(archivo.tamano / 1024, 1) . " KB"
        
        ; Obtener fecha de última modificación
        fechaModif := FileGetTime(archivo.ruta, "M")
        fechaFormateada := FormatTime(fechaModif, "dd/MM/yyyy")
        
        ; Convertir extensión a tipo
        tipoObjeto := ConvertirExtensionATipo(archivo.ext)
        
        ; Obtener nombre del proyecto/carpeta
        ; Para carpetas personalizadas, mostrar la carpeta seleccionada
        if esCarpetaPersonalizada {
            ; Extraer el nombre de la carpeta de la ruta origen
            partesRuta := StrSplit(rutaOrigen, "\")
            if partesRuta.Length > 0
                nombreCarpeta := nombreCarpetaSeleccionada  ; Última carpeta
            else
                nombreCarpeta := nombreCarpetaSeleccionada
        } else {
            ; Para Workspace/Git, usar la función ExtraerNombreProyecto
            nombreCarpeta := ExtraerNombreProyecto(archivo.dir)
        }
		
		; Si no se pudo determinar, mostrar "Personalizada"
        if nombreCarpeta = "" || nombreCarpeta = "Sin seleccionar"
            nombreCarpeta := "Personalizada"
        
        listView.Add("", archivo.nombre, tipoObjeto, tamanoKB, fechaFormateada, nombreCarpeta)
    }

    ; Ajuste de anchos de columnas según el tipo de carpeta
    if esCarpetaPersonalizada {
        ; Sin PROYECTO - más espacio para otras columnas
        listView.ModifyCol(1, "200 Left")   ; Nombre
        listView.ModifyCol(2, "100 Left")   ; Tipo
        listView.ModifyCol(3, "80 Left")    ; Tamaño
        listView.ModifyCol(4, "100 Left")   ; Fecha
    } else {
        ; Con PROYECTO
        listView.ModifyCol(1, "200 Left")   ; Nombre (máx 28 chars)
        listView.ModifyCol(2, "100 Left")   ; Tipo (10 chars)
        listView.ModifyCol(3, "80 Left")    ; Tamaño
        listView.ModifyCol(4, "100 Left")   ; Fecha (10 chars)
        listView.ModifyCol(5, "220 Left")   ; Proyecto
    }

    ; ────────────────────────────────────────────
    ;          BARRA DE ACCIONES INFERIOR
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
    btnVolver.OnEvent("Click", (*) => VolverAlMenu(selGui))
    btnCopiar.OnEvent("Click", (*) => CopiarArchivosSeleccionados(selGui))

    listView.OnEvent("ItemCheck", (*) => ActualizarContador())

    ; Hook para Shift+Click: subclassing del control ListView
    global lvUltimoClick := 0
    global lvProcOrig := 0
    lvHwnd := listView.Hwnd
    lvProcOrig := DllCall("SetWindowLongPtr", "Ptr", lvHwnd, "Int", -4,
                          "Ptr", CallbackCreate(LV_ShiftClickProc, , 4), "Ptr")

    selGui.OnEvent("Close", (*) => VolverAlMenu(selGui))

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
    global listView, archivosEncontrados, destinoBase, versionNatural
    
    ; Contar seleccionados
    archivosCopiar := []
    Loop listView.GetCount() {
        ; Verificar si el ítem actual está marcado
        if listView.GetNext(A_Index - 1, "Checked") = A_Index
            archivosCopiar.Push(archivosEncontrados[A_Index])
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
            
            ; Escribir archivo limpio con codificación del sistema (CP1252 en Windows español)
            FileAppend(contenidoLimpio, rutaDestino, "CP0")
            
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
                    ;          <12>
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
            ; Campo no encontrado en DDM, agregar comentario
            lineasGeneradas.Push("**C           0   * CAMPO " . nombreCampo . " NO ENCONTRADO EN DDM")
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
    ; Ejemplos de entrada: 
    ; "1 FLOAT(F4)" → "**DF          0   F   4 1FLOAT"
    ; "1 ARRAY(A10/1:5,1:5,1:5)" → "**DF I3       0   A  10 1ARRAY                           (1:5,1:5,1:5)"
    ; "1 ESTRUCTURA" → "**DS          0         1ESTRUCTURA"
    ; "2 CAMPO(A20)" (dentro de estructura) → "**DK          0   A  20 2CAMPO"
    
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
VolverAlMenu(ventana) {
    global rutaOrigen, txtLibreria, destinoBase, lvProcOrig, lvUltimoClick
    
    ; Guardar valores actuales
    rutaOrigenPreservada := rutaOrigen
    libreriaPreservada := txtLibreria.Value
    
    ; Limpiar el subclassing del ListView antes de destruir la ventana
    lvProcOrig    := 0
    lvUltimoClick := 0

    ventana.Destroy()
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
        ; Actualizar destinoBase manualmente
        destinoBase := ObtenerRutaDestino()
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
        
        ; Primera línea: agregar byte 0x0C al inicio
        if indice = 1 {
            lineasProcesadas.Push(Chr(0x0C) . linea)
            continue
        }
        
        ; Segunda línea: eliminar (saltar)
        if indice = 2 {
            continue
        }
        
        ; Tercera línea: mantener igual
        if indice = 3 {
            lineasProcesadas.Push(linea)
            continue
        }
        
        ; Cuarta línea: reemplazar con encabezado específico
        if indice = 4 {
            lineasProcesadas.Push("TYL  DB  NAME                              F LENG  S D REMARKS")
            continue
        }
        
        ; Quinta línea: reemplazar con separador
        if indice = 5 {
            lineasProcesadas.Push("---  --  --------------------------------  - ----  - - --------------------")
            continue
        }
        
        ; Resto de líneas: procesar campos o eliminar comentarios
        lineaTrim := Trim(linea)
        
        ; Si es comentario (empieza con *), ELIMINAR (no agregar a lineasProcesadas)
        if SubStr(lineaTrim, 1, 1) = "*" {
            continue
        }
        
        ; Si está vacía, mantener
        if lineaTrim = "" {
            lineasProcesadas.Push(linea)
            continue
        }
        
        ; Es un campo, necesita reformateo
        lineaReformateada := ReformatearCampoDDM(linea)
        lineasProcesadas.Push(lineaReformateada)
    }
    
    ; Reconstruir contenido
    contenidoFinal := ""
    for linea in lineasProcesadas {
        contenidoFinal .= linea . "`n"
    }
    
    ; Eliminar último salto de línea
    contenidoFinal := SubStr(contenidoFinal, 1, -1)
    
    return contenidoFinal
}

; ===== FUNCIÓN: REFORMATEAR CAMPO DDM =====
ReformatearCampoDDM(lineaOriginal) {
    ; Formato original de NaturalONE (aproximado):
    ; "T L DB Name                              F Leng  S D Remark"
    ; Col 1: T (G/M/P o espacio)
    ; Col 3: L (nivel)
    ; Col 6-7: DB (nombre corto)
    ; Col 10+: Name (nombre largo)
    ; Col 44+: F (tipo dato)
    ; etc.
    
    lineaTrim := Trim(lineaOriginal)
    
    ; Si la línea está vacía o es muy corta, devolverla tal cual
    if StrLen(lineaTrim) < 10 {
        return lineaOriginal
    }
    
    ; Extraer componentes usando posiciones aproximadas de NaturalONE
    ; (Estas posiciones pueden variar, usamos un enfoque robusto)
    
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
    
    ; Resto son opción, descriptor y remarks
    ; IMPORTANTE: Solo campos con tipo de dato pueden tener opción y descriptor
    ; - G (GROUP) sin tipo de dato: NO procesa opción/descriptor
    ; - P (PERIODIC) sin tipo de dato: NO procesa opción/descriptor
    ; - M (MULTIPLE) normalmente tiene tipo de dato, procesa opción/descriptor
    ; - Campos regulares con tipo de dato: SÍ procesan opción/descriptor
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
    
    ; Estructura de columnas:
    ; Col 1-4: **DF o **DK
    ; Col 5: C si es CONST, espacio si no
    ; Col 6-7: I1/I2/I3 si es array, espacios si no
    ; Col 8: S si está inicializado
    ; Col 9-14: espacios
    ; Col 15: número total de líneas (0 normalmente)
    ; Col 16-18: espacios
    ; Col 19: tipo de dato (A, N, etc)
    ; Col 20: espacio
    ; Col 21-23: longitud (alineada a derecha)
    ; Col 24: C si es CONST, espacio si no
    ; Col 25: nivel
    ; Col 26+: nombre y dimensiones
    
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