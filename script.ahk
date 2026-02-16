#Requires AutoHotkey v2.0
#SingleInstance Force

; ===== CONFIGURACIÓN =====
global libreriaDefecto := "LIBRERIA"
global destinoBase := ".\dos\NATAPPS\FUSER\" . libreriaDefecto . "\SRC"
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

; ===== INICIO DEL SCRIPT =====
; Leer librería actual desde NATPARM.SAG al inicio
libreriaActual := LeerLibreriaDeNATPARM()
if libreriaActual != ""
    libreriaDefecto := libreriaActual

MostrarMenuPrincipal()

; ===== FUNCIÓN: LEER LIBRERÍA DE NATPARM.SAG (CORREGIDA PARA AHK v2) =====
LeerLibreriaDeNATPARM() {
    rutaNATPARM := A_ScriptDir . "\dos\NATURAL\213\PROF\NATPARM.SAG"
    
    if !FileExist(rutaNATPARM) {
        return ""
    }
    
    try {
        ; Usar FileRead para leer bytes específicos
        archivo := FileOpen(rutaNATPARM, "r")
        if !archivo
            return ""
        
        ; Mover al offset 0x205A
        archivo.Seek(0x205A)
        
        ; Leer 8 bytes
        bytesLeidos := archivo.Read(8)
        archivo.Close()
        
        ; Limpiar y retornar
        nombreLimpio := StrReplace(bytesLeidos, Chr(0), "")
        nombreLimpio := Trim(nombreLimpio)
        
        return nombreLimpio
        
    } catch {
        return ""
    }
}

; ===== FUNCIÓN: ESCRIBIR LIBRERÍA EN NATPARM.SAG (CORREGIDA PARA AHK v2) =====
EscribirLibreriaEnNATPARM(nombreLibreria) {
    ; Usar ruta absoluta
    rutaNATPARM := A_ScriptDir . "\dos\NATURAL\213\PROF\NATPARM.SAG"
    
    if !FileExist(rutaNATPARM) {
        MsgBox("No se encontró NATPARM.SAG en:`n" . rutaNATPARM, "Error", "Icon!")
        return false
    }
    
    try {
        ; Usar FileOpen con modo "rw" (lectura/escritura)
        archivo := FileOpen(rutaNATPARM, "rw")
        if !archivo {
            throw Error("No se pudo abrir el archivo")
        }
        
        ; Limpiar y validar nombre
        nombreLimpio := StrUpper(Trim(nombreLibreria))
        if StrLen(nombreLimpio) > 8
            nombreLimpio := SubStr(nombreLimpio, 1, 8)
        
        ; Rellenar con espacios si es necesario (8 caracteres fijos)
        while StrLen(nombreLimpio) < 8 {
            nombreLimpio .= " "
        }
        
        ; Ir a la posición 0x205A
        archivo.Seek(0x205A)
        
        ; Escribir los 8 caracteres
        archivo.Write(nombreLimpio)
        
        ; Escribir byte 0x00 en posición 0x2062 (siguiente byte después del nombre de librería)
        archivo.Seek(0x2062)
        archivo.WriteUChar(0x00)
        
        archivo.Close()
        
        ; Verificar que se escribió correctamente
        nombreVerificado := LeerLibreriaDeNATPARM()
        if nombreVerificado != Trim(nombreLimpio) {
            throw Error("No se pudo verificar la escritura")
        }
        
        return true
        
    } catch as err {
        ; Método alternativo si el anterior falla
        try {
            MsgBox("Intentando método alternativo...", "Información", "Icon!")
            
            ; Leer todo el archivo, modificar y reescribir
            contenido := FileRead(rutaNATPARM, "RAW")
            if StrLen(contenido) < 0x205A + 8 {
                ; Extender el archivo si es necesario
                while StrLen(contenido) < 0x205A + 8 {
                    contenido .= Chr(0)
                }
            }
            
            ; Preparar nombre (8 bytes fijos)
            nombreLimpio := StrUpper(Trim(nombreLibreria))
            if StrLen(nombreLimpio) > 8
                nombreLimpio := SubStr(nombreLimpio, 1, 8)
            
            ; Rellenar con espacios
            while StrLen(nombreLimpio) < 8 {
                nombreLimpio .= " "
            }
            
            ; Convertir string a bytes
            nombreBytes := Buffer(8)
            Loop Parse, nombreLimpio {
                NumPut("UChar", Ord(A_LoopField), nombreBytes, A_Index - 1)
            }
            
            ; Insertar en la posición correcta
            contenidoBytes := Buffer(StrLen(contenido))
            Loop StrLen(contenido) {
                NumPut("UChar", Ord(SubStr(contenido, A_Index, 1)), contenidoBytes, A_Index - 1)
            }
            
            ; Copiar los bytes del nombre
            Loop 8 {
                NumPut("UChar", NumGet(nombreBytes, A_Index - 1, "UChar"), 
                      contenidoBytes, 0x205A + A_Index - 1)
            }
            
            ; Escribir archivo completo
            FileDelete(rutaNATPARM)
            nuevoArchivo := FileOpen(rutaNATPARM, "w")
            nuevoArchivo.RawWrite(contenidoBytes, contenidoBytes.Size)
            nuevoArchivo.Close()
            
            return true
            
        } catch as err2 {
            MsgBox("Error crítico al escribir NATPARM.SAG:`n`n" 
                . "Método 1: " . err.Message . "`n"
                . "Método 2: " . err2.Message . "`n`n"
                . "Solución manual:`n"
                . "1. Cierra Natural completamente`n"
                . "2. Ve a: " . rutaNATPARM . "`n"
                . "3. Abre NATPARM.SAG con editor hexadecimal`n"
                . "4. Cambia los bytes en offset 205A a 2061`n"
                . "5. Escribe: " . nombreLibreria, 
                "Error Crítico", "IconX")
            return false
        }
    }
}

MostrarMenuPrincipal() {
    global mainGui, txtRutaActual, btnEscanear, txtLibreria

    mainGui := Gui("+Resize +MinSize640x500 -MaximizeBox", "Migrador: NaturalONE → Natural 2.1.3")
    mainGui.BackColor := "F5F5F7"
    mainGui.SetFont("s12 c333333", "Segoe UI")

    ; Título
    mainGui.SetFont("s14 bold c0D47A1", "Segoe UI")
    mainGui.Add("Text", "x40 y25 w560 Center", "MIGRACIÓN DE CÓDIGO NATURAL")
    mainGui.SetFont("s12 c555555", "Segoe UI")
    mainGui.Add("Text", "x40 y60 w560 Center", "NaturalONE  →  Natural 2.1.3 (Windows 3.1)")

    mainGui.Add("Progress", "x60 y100 w520 h2 BackgroundE0E0E0 -Smooth Disabled")

    ; Sección origen
    mainGui.SetFont("s11 c444444", "Segoe UI")
    mainGui.Add("Text", "x40 y125", "1. Seleccione la carpeta de origen:")

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
    
    mainGui.Add("Progress", "x60 y" (yLibreria+70) " w520 h2 BackgroundE0E0E0 -Smooth Disabled")

    ; Botón Escanear (principal)
    yEscanear := yLibreria + 90
    btnEscanear := mainGui.Add("Button", "x" margenIzq " y" yEscanear " w520 h60 Disabled", "🔍   ESCANEAR OBJETOS NATURAL")
    btnEscanear.Opt("+BackgroundFF5722")
    btnEscanear.SetFont("s13 bold cFFFFFF", "Segoe UI")

    mainGui.Add("Progress", "x60 y" (yEscanear+75) " w520 h2 BackgroundE0E0E0 -Smooth Disabled")

    ; Botón Salir
    ySalir := yEscanear + 95
    btnSalir := mainGui.Add("Button", "x" margenIzq " y" ySalir " w520 h45", "❌  Salir")
    btnSalir.Opt("+Background757575")
    btnSalir.SetFont("s11 cFFFFFF", "Segoe UI")

    ; Eventos
    btnWS.OnEvent("Click", (*) => SeleccionarRuta(rutasDefecto[1]))
    btnGit.OnEvent("Click", (*) => SeleccionarRuta(rutasDefecto[2]))
    btnCustom.OnEvent("Click", (*) => SeleccionarRutaPersonalizada())
    btnEscanear.OnEvent("Click", (*) => EscanearArchivos())
    btnSalir.OnEvent("Click", (*) => ExitApp())
    txtLibreria.OnEvent("LoseFocus", (*) => ActualizarDestino())

    mainGui.OnEvent("Close", (*) => ExitApp())
    mainGui.Show("Center w640 h660")
    
    ; Actualizar destinoBase con la librería actual
    destinoBase := ".\dos\NATAPPS\FUSER\" . libreriaDefecto . "\SRC"
}

; ===== FUNCIÓN: ACTUALIZAR DESTINO =====
ActualizarDestino() {
    global txtLibreria, destinoBase, libreriaDefecto, mainGui
    
    ; Verificar si la ventana principal está minimizada o inactiva
    ; Si es así, no validar (para evitar mensaje al minimizar)
    try {
        if !WinActive("ahk_id " . mainGui.Hwnd)
            return
    }
    
    libreria := Trim(txtLibreria.Value)
    
    ; Si está vacío, restaurar valor por defecto sin mensaje
    if libreria = "" {
        txtLibreria.Value := libreriaDefecto
        destinoBase := ".\dos\NATAPPS\FUSER\" . libreriaDefecto . "\SRC"
        return
    }
    
    ; Convertir a mayúsculas
    libreria := StrUpper(libreria)
    
    ; Validar reglas de Natural
    errores := []
    
    ; Verificar que solo contenga caracteres válidos (A-Z, 0-9, -, #)
    if !RegExMatch(libreria, "^[A-Z0-9#-]+$") {
        errores.Push("- Solo se permiten letras, números, guión (-) y numeral (#)")
    }
    
    ; Verificar que empiece con letra o numeral
    if !RegExMatch(libreria, "^[A-Z#]") {
        errores.Push("- Debe empezar con una letra (A-Z) o numeral (#)")
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
        destinoBase := ".\dos\NATAPPS\FUSER\" . libreriaDefecto . "\SRC"
        
        ; Dar foco de nuevo al campo para que el usuario lo corrija
        txtLibreria.Focus()
        return
    }
    
    ; Si pasó todas las validaciones, actualizar
    txtLibreria.Value := libreria
    destinoBase := ".\dos\NATAPPS\FUSER\" . libreria . "\SRC"
}

; ===== FUNCIÓN: SELECCIONAR RUTA PREDEFINIDA =====
SeleccionarRuta(ruta) {
    global rutaOrigen, txtRutaActual, btnEscanear, esCarpetaPersonalizada, nombreCarpetaSeleccionada
    
    if DirExist(ruta) {
        rutaOrigen := ruta
        txtRutaActual.Value := ruta
        btnEscanear.Enabled := true
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
        btnEscanear.Enabled := true
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
    global archivosEncontrados, rutaOrigen, mainGui
    
    if rutaOrigen = "" {
        MsgBox("Por favor, seleccione primero una carpeta de origen.", "Error", "IconX")
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
    selGui := Gui("+Resize +MinSize800x580", "Seleccionar Objetos Natural para Migrar")
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
    ; Definir columnas según el tipo de carpeta
    if esCarpetaPersonalizada {
        ; Sin columna PROYECTO para carpetas personalizadas
        ; listView := selGui.Add("ListView", "x30 y100 w740 h400 Checked -Multi Grid", ["    NOMBRE", "TIPO", "TAMAÑO", "FECHA"])
		listView := selGui.Add("ListView", "x30 y100 w740 h400 Checked -Multi Grid", ["    NOMBRE", "TIPO", "TAMAÑO", "FECHA", "PROYECTO"])
    } else {
        ; Con columna PROYECTO para Workspace/Git
        listView := selGui.Add("ListView", "x30 y100 w740 h400 Checked -Multi Grid", ["    NOMBRE", "TIPO", "TAMAÑO", "FECHA", "PROYECTO"])
    }
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

    ; Contador de seleccionados (con fondo destacado)
    txtContador := selGui.Add("Text", "x420 y528 w200 c0D47A1", "Seleccionados: 0")
    txtContador.SetFont("s11 bold", "Segoe UI")

    ; Botones principales (Volver y Copiar)
    btnVolver := selGui.Add("Button", "x590 y520 w90 h45", "  ←  Volver")
    btnVolver.Opt("+Background607D8B")   ; Gris azulado
    btnVolver.SetFont("s10 cFFFFFF", "Segoe UI")

    btnCopiar := selGui.Add("Button", "x700 y520 w90 h45", "📋  COPIAR")
    btnCopiar.Opt("+BackgroundFF5722")   ; Naranja acción
    btnCopiar.SetFont("s10 bold cFFFFFF", "Segoe UI")

    ; ────────────────────────────────────────────
    ;               EVENTOS
    ; ────────────────────────────────────────────
    btnSelTodos.OnEvent("Click", (*) => SeleccionarTodos(true))
    btnDeselTodos.OnEvent("Click", (*) => SeleccionarTodos(false))
    btnVolver.OnEvent("Click", (*) => VolverAlMenu(selGui))
    btnCopiar.OnEvent("Click", (*) => CopiarArchivosSeleccionados(selGui))

    listView.OnEvent("ItemCheck", (*) => ActualizarContador())

    selGui.OnEvent("Close", (*) => VolverAlMenu(selGui))

    selGui.Show("Center w800 h600")
}

; ===== FUNCIÓN: ACTUALIZAR CONTADOR =====
ActualizarContador() {
    global listView, txtContador
    
    contador := 0
    Loop listView.GetCount() {
        if listView.GetNext(A_Index - 1, "Checked") = A_Index
            contador++
    }
    
    txtContador.Value := "Seleccionados: " . contador
    if (contador = 0)
        txtContador.Opt("c555555")
    else if (contador = archivosEncontrados.Length)
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

; ===== FUNCIÓN: COPIAR ARCHIVOS SELECCIONADOS =====
CopiarArchivosSeleccionados(ventana) {
    global listView, archivosEncontrados, destinoBase
    
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
    
    ; Confirmar - mostrar ruta con librería actualizada
    libreriaActual := Trim(txtLibreria.Value)
    rutaCompletaConLibreria := ".\dos\NATAPPS\FUSER\" . libreriaActual . "\SRC"
    respuesta := MsgBox("¿Desea copiar " . archivosCopiar.Length . " archivo(s) a:`n`n" . rutaCompletaConLibreria . "`n`nLibrería: " . libreriaActual . "`n`nSe eliminarán automáticamente los encabezados de NaturalONE.", "Confirmar", "YesNo Icon?")
    if respuesta = "No"
        return
    
    ; Crear directorio destino si no existe
    if !DirExist(destinoBase) {
        try {
            DirCreate(destinoBase)
        } catch as err {
            MsgBox("Error al crear el directorio destino:`n" . err.Message, "Error", "IconX")
            return
        }
    } else {
        ; Si existe, borrar todo el contenido anterior
        try {
            ; Borrar TODOS los archivos en el directorio
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
    global destinoBase
    
    try {
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
            
            ; Últimos 4 bytes del bloque de nombre: 00 00 00 00
              file.WriteUChar(0x00)
            ; file.WriteUChar(0x00)
            ; file.WriteUChar(0x00)
            ; file.WriteUChar(0x00)
            
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
    global rutaOrigen, txtLibreria, destinoBase
    
    ; Guardar valores actuales
    rutaOrigenPreservada := rutaOrigen
    libreriaPreservada := txtLibreria.Value
    
    ventana.Destroy()
    MostrarMenuPrincipal()
    
    ; Restaurar valores
    if rutaOrigenPreservada != "" {
        global txtRutaActual, btnEscanear
        rutaOrigen := rutaOrigenPreservada
        txtRutaActual.Value := rutaOrigenPreservada
        btnEscanear.Enabled := true
    }
    
    if libreriaPreservada != "" {
        txtLibreria.Value := libreriaPreservada
        ; Actualizar destinoBase manualmente
        destinoBase := ".\dos\NATAPPS\FUSER\" . libreriaPreservada . "\SRC"
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
