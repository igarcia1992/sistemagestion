VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Clientes 
   Caption         =   "Alta de clientes"
   ClientHeight    =   10815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17160
   OleObjectBlob   =   "Clientes.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public abierto As Byte
Public procesar As Long

#If VBA7 And Win64 Then
    ' Declaración de función para buscar una ventana en sistemas de 64 bits
    Private Declare PtrSafe Function FindWindow Lib "USER32" _
    Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    ' Declaración de función para buscar una ventana en sistemas de 32 bits
    Private Declare Function FindWindow Lib "USER32" _
    Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

#If VBA7 And Win64 Then
    ' Declaración de función para dibujar la barra de menú de una ventana en sistemas de 64 bits
    Private Declare PtrSafe Function DrawMenuBar Lib "USER32" (ByVal hwnd As Long) As LongPtr
#Else
    ' Declaración de función para dibujar la barra de menú de una ventana en sistemas de 32 bits
    Private Declare Function DrawMenuBar Lib "USER32" (ByVal hwnd As Long) As Long
#End If

#If VBA7 And Win64 Then
    #If VBA7 Then
        #If Win64 Then
            ' Declaración de función para establecer un valor en una ventana en sistemas de 64 bits (Windows API)
            Private Declare PtrSafe Function SetWindowLongPtr Lib "USER32" Alias "SetWindowLongPtrA" _
            (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
        #Else
            ' Declaración de función para establecer un valor en una ventana en sistemas de 32 bits (Windows API)
            Private Declare Function SetWindowLongPtr Lib "USER32" Alias "SetWindowLongA" _
            (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
        #End If
    #Else
        ' Declaración de función para establecer un valor en una ventana en sistemas de 32 bits (Windows API)
        Private Declare Function SetWindowLongPtr Lib "USER32" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    #End If
#Else
    ' Declaración de función para establecer un valor en una ventana en sistemas de 32 bits (Windows API)
    Private Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

#If VBA7 And Win64 Then
    #If VBA7 Then
        #If Win64 Then
            ' Declaración de función para obtener un valor de una ventana en sistemas de 64 bits (Windows API)
            Private Declare PtrSafe Function GetWindowLongPtr Lib "USER32" _
            Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
        #Else
            ' Declaración de función para obtener un valor de una ventana en sistemas de 32 bits (Windows API)
            Private Declare PtrSafe Function GetWindowLongPtr Lib "USER32" _
            Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
        #End If
    #Else
        ' Declaración de función para obtener un valor de una ventana en sistemas de 32 bits (Windows API)
        Private Declare Function GetWindowLongPtr Lib "USER32" Alias "GetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    #End If
#Else
    ' Declaración de función para obtener un valor de una ventana en sistemas de 32 bits (Windows API)
    Private Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
#End If

' Constantes para configurar el estilo de la ventana
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const GWL_STYLE As Long = (-16)


Public edit As Byte
Public registro As Variant
Public ID As Variant
Private Sub Bt_limpiar_Click()
    ' Desactivar la actualización de la interfaz gráfica
    Application.ScreenUpdating = False

    ' Limpiar cuadros de texto
    Me.TextNombre = ""
    Me.TextTelefono = ""
    Me.TextRegistro = ""
    Me.TextDireccion = ""
    Me.TextCorreo = ""
    TextBox1.value = Format(Date, "dd/mm/yyyy")
    TextBox2.value = Format(Date, "dd/mm/yyyy")
    
    ' Ocultar y habilitar botones
    Me.CommandEditar.Visible = False
    Me.CommandAgregar.Visible = True

    ' Establecer el foco en el primer cuadro de texto
    Me.TextNombre.SetFocus

    ' Volver a activar la actualización de la interfaz gráfica
    Application.ScreenUpdating = True
End Sub
Sub nada()

   ' Declaración de variables
    Dim hoja As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim valorComboBox As String

    ' Obtener el valor seleccionado en el ComboBox
    valorComboBox = Me.ComboBox1.value

    ' Limpiar ListBox antes de agregar nuevos elementos
    Me.ListBox1.Clear

    ' Configuración de la hoja y obtención de la última fila con datos
    Set hoja = ThisWorkbook.Sheets("CLIENTES")
    ultimaFila = hoja.Cells(hoja.Rows.Count, "B").End(xlUp).Row

    ' Llenado del ListBox con datos de la hoja "CLIENTES" desde la fila 5
    For i = 5 To ultimaFila
        ' Verificar si el valor en la columna H coincide con el valor seleccionado en el ComboBox
        If hoja.Cells(i, "H").value = valorComboBox Then
            Me.ListBox1.AddItem
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 0) = hoja.Cells(i, "C").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = hoja.Cells(i, "D").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = hoja.Cells(i, "E").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = hoja.Cells(i, "F").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = hoja.Cells(i, "G").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = hoja.Cells(i, "A").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = hoja.Cells(i, "H").value
        End If
    Next i
    End Sub

Private Sub ComboCartera_Change()
    ' Este procedimiento se activa cada vez que se cambia la selección en ComboCartera

    ' Elimina los espacios en blanco al principio y al final del texto ingresado en ComboCartera
    Dim texto As String
    texto = Trim(ComboCartera.Text)

    ' Asigna el texto modificado de vuelta a ComboCartera
    ComboCartera.Text = texto

    ' Verifica si ComboCartera está vacío después de eliminar los espacios en blanco
    If texto = "" Then
        ' Si está vacío, muestra el mensaje en el Label80
        Label80.Visible = True
    Else
        ' Si no está vacío, oculta el mensaje en el Label80
        Label80.Visible = False
    End If
End Sub

Private Sub ComboBox1_Change()
 ' Declaración de variables
    Dim hoja As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim valorComboBox As String

    ' Obtener el valor seleccionado en el ComboBox
    valorComboBox = Me.ComboBox1.value

    ' Limpiar ListBox antes de agregar nuevos elementos
    Me.ListBox1.Clear

    ' Configuración de la hoja y obtención de la última fila con datos
    Set hoja = ThisWorkbook.Sheets("CLIENTES")
    ultimaFila = hoja.Cells(hoja.Rows.Count, "B").End(xlUp).Row

    ' Llenado del ListBox con datos de la hoja "CLIENTES" desde la fila 5
    For i = 5 To ultimaFila
        ' Verificar si el valor en la columna H coincide con el valor seleccionado en el ComboBox
        If hoja.Cells(i, "H").value = valorComboBox Then
            Me.ListBox1.AddItem
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 0) = hoja.Cells(i, "C").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = hoja.Cells(i, "D").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = hoja.Cells(i, "E").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = hoja.Cells(i, "F").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = hoja.Cells(i, "G").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = hoja.Cells(i, "A").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = hoja.Cells(i, "H").value
        End If
    Next i
End Sub

Private Sub CommandAgregar_Click()
    ' Comprueba si el nombre está vacío
    If Trim(Me.TextNombre.value) = "" Then
        MsgBox "El campo nombre no puede ir vacio!", vbInformation, "Antonio's"
        Exit Sub
    End If
    
    Dim hoja As Worksheet
    Dim tabla As ListObject
    Set hoja = ThisWorkbook.Sheets("CLIENTES")
    Set tabla = hoja.ListObjects("DESTINO")
    Me.ListBox1.Clear
    ComboBox1.value = ""
    
    ' Bandera de validación
    Dim phoneNumberExists As Boolean
    phoneNumberExists = False
    
    ' Comprueba si el número de teléfono ya existe en la columna D
    Dim ulfila As Long
    If hoja.Range("D5").value = "" Then
    ulfila = 5
    Else
    ulfila = hoja.Cells(hoja.Rows.Count, "D").End(xlUp).Row + 1
    End If
    
    Dim X As Long
    
    For X = 5 To ulfila
        If CStr(hoja.Range("D" & X).value) = CStr(Me.TextTelefono.value) Then
            ' El número de teléfono ya existe
            phoneNumberExists = True
            Exit For
        End If
    Next X
    
    ' Muestra el cuadro de mensaje y maneja la respuesta del usuario
    If phoneNumberExists Then
        Dim response As VbMsgBoxResult
        response = MsgBox("El numero de telefono ya existe en lña base de datos aun asi deseas agregarlo?", vbYesNo + vbQuestion, "Duplicate Phone Number")
        
        If response = vbNo Then
        UserForm_Initialize
            ' El usuario eligió no agregar el número de teléfono duplicado
            Exit Sub
        End If
    End If
    
    ' Agrega una nueva fila a la tabla
    Dim nuevaFila As listRow
    Set nuevaFila = tabla.ListRows.Add(1)
    
    ' Asigna los valores a la nueva fila
    nuevaFila.Range(1, 1).value = CStr(Me.Label66.Caption)
    nuevaFila.Range(1, 2).value = CDate(Me.LabelFecha.Caption)
    nuevaFila.Range(1, 3).value = CStr(Me.TextNombre.value)
    nuevaFila.Range(1, 4).value = CStr(Me.TextTelefono.value)
    nuevaFila.Range(1, 5).value = CStr(Me.TextRegistro.value)
    nuevaFila.Range(1, 6).value = VBA.CStr(Me.TextCorreo.value)
    nuevaFila.Range(1, 7).value = CStr(Me.TextDireccion.value)
      
    ' Limpia los cuadros de texto
    Me.TextCorreo.value = ""
    Me.TextNombre.value = ""
    Me.TextTelefono.value = ""
    Me.TextRegistro.value = ""
    Me.TextDireccion.value = ""
  
    ' Establece el enfoque en el cuadro de texto del nombre
    Me.TextNombre.SetFocus
    
   hoja.Range("I2").value = hoja.Range("I2").value + 1
    
    ' Llama al evento cargar datos al listbox
    cargarlistbox
    
    On Error Resume Next
    ThisWorkbook.Save
    
End Sub

Private Sub CommandAtras_Click()
abierto = 1
Unload Me
Panel_Principal.Show
End Sub

Private Sub CommandButton7_Click()
 'Declaración de una variable de tipo Date llamada Fecha.
    Dim Fecha As Date

    ' Se obtiene la fecha del formulario CalendarForm y se asigna a la variable Fecha.
    Fecha = CalendarForm.GetDate

    ' Se asigna el valor formateado de la variable Fecha al TextBox1.
    TextBox1.value = Format(Fecha, "dd/mm/yyyy")  ' "Short Date" usa el formato de fecha del sistema
End Sub

Private Sub CommandButton8_Click()
 'Declaración de una variable de tipo Date llamada Fecha.
    Dim Fecha As Date

    ' Se obtiene la fecha del formulario CalendarForm y se asigna a la variable Fecha.
    Fecha = CalendarForm.GetDate

    ' Se asigna el valor formateado de la variable Fecha al TextBox1.
    TextBox2.value = Format(Fecha, "dd/mm/yyyy") ' "Short Date" usa el formato de fecha del sistema
End Sub

 Private Sub CommandButton9_Click()
    
  ' Limpiar ListBox antes de agregar nuevos elementos
    Me.ListBox1.Clear

    ' Validar que las fechas estén ingresadas y el ListBox no esté vacío
    If TextBox1 = "" Or TextBox2 = "" Or ListBox1.ListCount > 0 Then
        MsgBox "Por favor inglese fechas validas..", vbExclamation
        Exit Sub
    End If
    
    ' Convertir las fechas de los TextBox a formato Date
    startDate = CDate(TextBox1.value)
    endDate = CDate(TextBox2.value)
    
    ' Referencia a la hoja de trabajo llamada "Ingresos"
    Set ws = ThisWorkbook.Sheets("Clientes")
    
    ' Última fila con datos en la columna C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Recorrer las filas en la columna C
    For i = 5 To lastRow ' Asumiendo que la primera fila contiene encabezados
        ' Validar la fecha en la columna B
        If ws.Cells(i, 2).value >= startDate And ws.Cells(i, 2).value <= endDate Then
           
                ' Agregar datos al ListBox
                Me.ListBox1.AddItem "" ' Agrega una fila vacía al ListBox
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 0) = ws.Cells(i, 3).value
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = ws.Cells(i, 4).value
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = ws.Cells(i, 5).value
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = ws.Cells(i, 6).value
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = ws.Cells(i, 7).value
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = ws.Cells(i, 1).value
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = ws.Cells(i, 8).value
                
           
            End If
    Next i
    MsgBox "Proceso Terminado", vbInformation, "Antonio's"
End Sub
Private Sub CommandEditar_Click()
    ' Search for the value of id in column A of the CLIENTES sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("CLIENTES")
    Dim searchId As Variant
    searchId = ID
    Dim rngSearch As Range
    Set rngSearch = ws.Columns("A").Find(What:=searchId, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Check if the match was found
    If Not rngSearch Is Nothing Then
        ' Get the row where the match was found
        Dim foundRow As Long
        foundRow = rngSearch.Row

        ' Update values in the found row with the values from the TextBoxes
        ws.Range("C" & foundRow).value = Me.TextNombre.value
        ws.Range("D" & foundRow).value = Me.TextTelefono.value
        ws.Range("E" & foundRow).value = Me.TextRegistro.value
        ws.Range("F" & foundRow).value = VBA.CStr(Me.TextCorreo.value)
        ws.Range("G" & foundRow).value = Me.TextDireccion.value

        
        ' Clear the TextBoxes
        Me.TextNombre = Empty
        Me.TextTelefono = Empty
        Me.TextRegistro = Empty
        Me.TextDireccion = Empty
        Me.TextCorreo = Empty


        ' Set focus on the name TextBox
        Me.TextNombre.SetFocus

        ' Hide the Edit button and show the Add button
        Me.CommandAgregar.Visible = True
        Me.CommandEditar.Visible = False

        ' Clear the ListBox and reload the data
        
        registro = ""
        edit = 0

        ' Show an informative message
        MsgBox "Cambios Realizados Correctamente", vbInformation, "MWCMD"
    Else
        ' Show a message if the id was not found
        MsgBox "The ID was not found in the CLIENTES sheet.", vbExclamation, "MWCMD"
    End If
    cargarlistbox
    On Error Resume Next
    ThisWorkbook.Save
    
End Sub


Private Sub CommandEliminar_Click()
    If ID = "" Then
        MsgBox "Selecciona al cliente que deseas eliminar haciendo doble clic sobre él", vbInformation, "MWCMD"
        Exit Sub
    End If
    
    ' Buscar el valor de id en la columna A de la hoja CLIENTES
    Dim ws As Worksheet
 
    Set ws = ThisWorkbook.Sheets("CLIENTES")

        
    Dim rngBuscar As Range
    Set rngBuscar = ws.Columns("A").Find(What:=ID, LookIn:=xlValues, LookAt:=xlWhole)
    
    'id - Cliente
    
    ' Verificar si se encontró la coincidencia
    If Not rngBuscar Is Nothing Then
        ' Mostrar un MsgBox de confirmación
        pregunta = MsgBox("Estás a punto de eliminar al cliente: " & Me.ListBox1.List(Me.ListBox1.ListIndex, 0) & vbNewLine & "¿Quieres continuar?", _
                          vbYesNo + vbQuestion, "Confirmación de Antonio:")
        If pregunta = vbYes Then
            Dim tbl As ListObject
            Dim tblRange As Range
            Set tbl = ws.ListObjects("DESTINO") ' Nombre de la tabla
            Set tblRange = tbl.Range
            Dim rowToDelete As Range
            Set rowToDelete = tblRange.Rows(rngBuscar.Row - tblRange.Row + 1) ' Ajuste para incluir el encabezado
            rowToDelete.Delete
            MsgBox "El cliente ha sido eliminado.", vbInformation, "Antonio's"
        End If
    End If
    
    ListBox1.Clear
    cargarlistbox
    
    Me.TextNombre = Empty
    Me.TextTelefono = Empty
    Me.TextRegistro = Empty
    Me.TextCorreo = Empty
    Me.TextDireccion = Empty
    Me.CommandAgregar.Visible = True
    Me.CommandEditar.Visible = False
    
    On Error Resume Next
    ThisWorkbook.Save
End Sub

Private Sub ExportarExcel_Click()
    ' Declarar variables
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Integer
    Dim j As Integer

    ' Verificar si el ListBox está vacío
    If ListBox1.ListCount = 0 Then
        MsgBox "No data to export.", vbExclamation, "MWCMD"
        Exit Sub
    End If

    ' Desactivar actualización de la aplicación para evitar que se muestre el libro
    Application.ScreenUpdating = False

    ' Crear un nuevo libro de Excel
    Set wb = Workbooks.Add
    ' Tomar la primera hoja del nuevo libro
    Set ws = wb.Sheets(1)

    ' Colocar encabezados en la primera fila
    ws.Cells(1, 1).value = "Cliente"
    ws.Cells(1, 2).value = "Telefono"
    ws.Cells(1, 3).value = "Tipo de cliente"
    ws.Cells(1, 4).value = "Correo"
    ws.Cells(1, 5).value = "Direccion"
    ws.Cells(1, 6).value = "Cartera"


   ' Copiar datos desde el ListBox a la hoja de Excel
    For i = 0 To ListBox1.ListCount - 1
        For j = 1 To 6
            ' Verificar si j es diferente de 6 para copiar la columna correspondiente
            If j <> 6 Then
                ws.Cells(i + 2, j).value = ListBox1.List(i, j - 1)
            Else
                ' Si j es igual a 6, saltamos la columna 6 y copiamos la columna 7
                ws.Cells(i + 2, j).value = ListBox1.List(i, j)
            End If
        Next j
    Next i


    ' Pedir al usuario la ubicación y el nombre del archivo
    Dim savePath As String
    savePath = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Guardar como")

    ' Verificar si el usuario cancela la operación de guardar
    If savePath = "Falso" Then
        ' Cerrar el libro sin guardar cambios
        wb.Close SaveChanges:=False
        MsgBox "Operacion cancelada por el usuario.", vbExclamation, "MWCMD"
        Exit Sub
    End If

    ' Guardar el nuevo libro de Excel en la ubicación proporcionada sin abrirlo
    wb.SaveAs savePath
    ' Cerrar el libro sin guardar cambios
    wb.Close SaveChanges:=False

    ' Reactivar la actualización de la aplicación
    Application.ScreenUpdating = True

    ' Mostrar mensaje de éxito
    MsgBox "Datos exportados correctamente.", vbInformation
End Sub

Private Sub Label84_Click()
Me.TextCorreo.SetFocus
End Sub

Private Sub Label85_Click()
Me.TextTelefono.SetFocus
End Sub

Private Sub Label86_Click()
Me.TextNombre.SetFocus
End Sub

Private Sub ListBox1_Click()
ID = ""

'listo
If edit = 1 Then
'nada

Else
Me.TextNombre = Empty
Me.TextTelefono = Empty
Me.TextRegistro = Empty
Me.TextCorreo = Empty
Me.TextDireccion = Empty
Me.CommandAgregar.Visible = True
Me.CommandEditar.Visible = False
End If


End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Me.ListBox1.ListIndex = -1 Then
'nada de nada padrino
Else

'listo
Dim fila As Long
fila = Me.ListBox1.ListIndex

Me.TextNombre.value = Me.ListBox1.List(fila, 0)
Me.TextTelefono.value = Me.ListBox1.List(fila, 1)
Me.TextRegistro.value = Me.ListBox1.List(fila, 2)
Me.TextCorreo.value = Me.ListBox1.List(fila, 3)
Me.TextDireccion.value = Me.ListBox1.List(fila, 4)


ID = Me.ListBox1.List(fila, 5)

Me.CommandEditar.Visible = True
Me.CommandAgregar.Visible = False
End If
End Sub

Private Sub TextBox1_Change()
If TextBox1.value = "30/12/1899" Then
TextBox1.value = Format(Date, "dd/mm/yyyy")
End If
End Sub
Private Sub TextBox2_Change()
If TextBox2.value = "30/12/1899" Then
TextBox2.value = Format(Date, "dd/mm/yyyy")
End If
End Sub

Private Sub TextBox8_Change()
    Dim searchTerm As String
    Dim hoja As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim encontrado As Boolean
    Dim j As Long
    
    ComboBox1.value = ""
    Me.ListBox1.Clear
    Me.ListBox1.ColumnCount = 7
    Me.ListBox1.ColumnWidths = "145;60;60;140;180;0;0"

    ' Obtener el término de búsqueda del TextBox8
    searchTerm = LCase(Me.TextBox8.Text)

    ' Limpiar el ListBox1 antes de aplicar el filtro
    Me.ListBox1.Clear

    ' Definir la hoja de trabajo donde están los datos
    Set hoja = ThisWorkbook.Sheets("Clientes")

    ' Encontrar la última fila con datos en la columna "C"
    ultimaFila = hoja.Cells(hoja.Rows.Count, "C").End(xlUp).Row

    ' Recorrer las filas para buscar coincidencias desde la fila 4
    For i = 5 To ultimaFila
        encontrado = False

        ' Recorre las columnas de A a F
        For j = 3 To 7
            If InStr(1, LCase(hoja.Cells(i, j).value), searchTerm) > 0 Then
                encontrado = True
                Exit For ' Salir del bucle interno si se encuentra una coincidencia en cualquier columna
            End If
        Next j

        ' Si se encontró una coincidencia en cualquier columna, agregar la fila al ListBox1
        If encontrado Then
            ' Agregar una nueva fila a la lista
            Me.ListBox1.AddItem
            
            ' Mover los valores hacia la izquierda, omitiendo la primera columna
            For j = 3 To 7
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, j - 3) = hoja.Cells(i, j).value
            Next j
            
            ' Agregar la columna A en la columna 5 del ListBox
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = hoja.Cells(i, 1).value
        End If
    Next i
End Sub

Private Sub TextTelefono_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Verifica si el caracter presionado es numérico
    If Not IsNumeric(Chr(KeyAscii)) Then
        ' Si no es numérico, cancela la entrada
        KeyAscii = 0
    End If
    
    ' Verifica si la longitud después de la entrada excede 13 caracteres
    If Len(TextTelefono.value & Chr(KeyAscii)) > 13 Then
        ' Si excede, cancela la entrada
        KeyAscii = 0
    End If
End Sub



Private Sub ToggleButton1_Click()
    ' Verifica si ToggleButton1 está activado
    If ToggleButton1.value = True Then
        ' Cambia el color del botón a un verde tenue
        ToggleButton1.BackColor = RGB(144, 238, 144)  ' Puedes ajustar los valores RGB según tu preferencia
        ' Llama a la subrutina cargarlistbox
        cargarlistbox
    Else
        ' Si ToggleButton1 no está activado, limpia el ListBox1
        ListBox1.Clear
        ' Restaura el color del botón al original (por ejemplo, gris claro)
        ToggleButton1.BackColor = RGB(240, 240, 240)  ' Puedes ajustar los valores RGB según tu preferencia
    End If
End Sub



Private Sub UserForm_Initialize()
      ThisWorkbook.Activate
      
      Dim valorActual As Long ' Variable global para almacenar el valor actual
    ' Declarar una variable para verificar si se ejecuta en Windows de 64 bits
    
    Dim Windows64 As Boolean

    ' Validamos la versión de Office
    #If VBA7 And Win64 Then
        Dim lngMyHandle As LongPtr, lngCurrentStyle As LongPtr, lngNewStyle As LongPtr
    #Else
        Dim lngMyHandle As Long, lngCurrentStyle As Long, lngNewStyle As Long
    #End If

    ' Verificar la versión de Office para determinar el nombre de la ventana
    If Application.Version < 9 Then
        lngMyHandle = FindWindow("THUNDERXFRAME", Me.Caption)
    Else
        lngMyHandle = FindWindow("THUNDERDFRAME", Me.Caption)
    End If

    ' Actualizar el estilo de la ventana para permitir minimizar y maximizar
    #If VBA7 And Win64 Then
        lngCurrentStyle = GetWindowLongPtr(lngMyHandle, GWL_STYLE)
        lngNewStyle = lngCurrentStyle Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
        SetWindowLongPtr lngMyHandle, GWL_STYLE, lngNewStyle
    #Else
        lngCurrentStyle = GetWindowLong(lngMyHandle, GWL_STYLE)
        lngNewStyle = lngCurrentStyle Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
        SetWindowLong lngMyHandle, GWL_STYLE, lngNewStyle
    #End If

    ' Determinar si se ejecuta en Windows de 64 bits
    #If Win64 Then
        Windows64 = True
    #Else
        Windows64 = False
    #End If

    ' Variable para controlar si el formulario está abierto o cerrado
    abierto = 0

    ' Declaración de variables
    Dim hoja As Worksheet
    Dim ultimaFila As Long
    Set hoja = ThisWorkbook.Sheets("CLIENTES")
   
    ' Configuración de otros elementos del formulario
    Me.LabelFecha.Caption = Format(Date, "dd/mm/yyyy")
    Me.CommandEditar.Visible = False
    
    ' Agregar 1 al valor de la última fila y asignar al Caption del Label66
    Me.Label66.Caption = hoja.Range("I2")
    
    TextBox1.value = Format(Date, "dd/mm/yyyy")
    TextBox2.value = Format(Date, "dd/mm/yyyy")
    Me.TextBox1.Locked = True
    Me.TextBox2.Locked = True
    
    ' Configuración del ListBox
    Me.ListBox1.ColumnCount = 7
    Me.ListBox1.ColumnWidths = "145;60;70;130;180;0;0"
    cargarlistbox
    
     ComboBox1.Clear
    
     ' Agregar elementos al ComboBox
    With Me.ComboBox1
        .AddItem "Activo"
        .AddItem "Inactivo"
    End With
    
     On Error Resume Next
  ' Carga la imagen desde la ruta guardada en la celda específica
    imgPath = ThisWorkbook.Sheets("Inicio").Range("A1").value
    
      If imgPath <> "" Then
        With Me.Image1
            .Picture = LoadPicture(imgPath)
            .PictureSizeMode = fmPictureSizeModeStretch
            .PictureAlignment = fmPictureAlignmentCenter
        End With
        End If
    On Error GoTo 0
    
End Sub
Private Sub UserForm_Terminate()
If abierto = 0 Then
Call MiModulo.Cerrartodo
Else
'nada
End If
End Sub

Sub cargarlistbox()
    ' Declaración de variables
    Dim hoja As Worksheet
    Dim ultimaFila As Long
    Dim i As Long


    ' Configuración del ListBox
    Me.ListBox1.Clear
    Me.ListBox1.ColumnCount = 7
    Me.ListBox1.ColumnWidths = "145;60;60;140;180;0;0"

    ' Configuración de la hoja y obtención de la última fila con datos
    Set hoja = ThisWorkbook.Sheets("CLIENTES")
    ultimaFila = hoja.Cells(hoja.Rows.Count, "B").End(xlUp).Row

    ' Llenado del ListBox con datos de la hoja "CLIENTES" desde la fila 5
    For i = 5 To ultimaFila
    
            Me.ListBox1.AddItem
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 0) = hoja.Cells(i, "C").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = hoja.Cells(i, "D").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = hoja.Cells(i, "E").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = hoja.Cells(i, "F").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = hoja.Cells(i, "G").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = hoja.Cells(i, "A").value
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = hoja.Cells(i, "H").value
    Next i

    ' Agregar 1 al valor de la última fila y asignar al Caption del Label66
    Me.Label66.Caption = hoja.Range("I2")
End Sub
Private Sub TextNombre_Change()
  If Len(TextNombre.value) > 0 Then
       Label13.Visible = False
   Else
       Label13.Visible = True
   End If
End Sub
Private Sub Label13_Click()
TextNombre.SetFocus
End Sub
Private Sub TextTelefono_Change()
    If Len(TextTelefono.value) > 0 Then
        Label74.Visible = False
    Else
        Label74.Visible = True
    End If
End Sub

Private Sub Label74_Click()
    TextTelefono.SetFocus
End Sub
Private Sub TextCorreo_Change()
    If Len(TextCorreo.value) > 0 Then
        Label75.Visible = False
    Else
        Label75.Visible = True
    End If
End Sub

Private Sub Label75_Click()
    TextCorreo.SetFocus
End Sub
Private Sub TextRegistro_Change()
    If Len(TextRegistro.value) > 0 Then
        Label76.Visible = False
    Else
        Label76.Visible = True
    End If
End Sub

Private Sub Label76_Click()
    TextRegistro.SetFocus
End Sub

Private Sub TextDireccion_Change()
    If Len(TextDireccion.value) > 0 Then
        Label77.Visible = False
    Else
        Label77.Visible = True
    End If
End Sub

Private Sub Label77_Click()
'Se pociona en el textbox de nombre TextDireccion
    TextDireccion.SetFocus
End Sub
Private Sub LlenarComboCartera()

    ' Declarar variables
    Dim ws As Worksheet
    Dim rngConfiguracion As Range
    Dim cell As Range

    ' Establecer la hoja de trabajo
    Set ws = ThisWorkbook.Sheets("Clientes")

    ' Limpiar el ComboBox antes de llenarlo
    ComboBox5.Clear

    ' Definir el rango de celdas en la columna G a partir de G2
    Set rngConfiguracion = ws.Range("C2", ws.Cells(ws.Rows.Count, "C").End(xlUp))

    ' Recorrer las celdas en el rango
    For Each cell In rngConfiguracion
        ' Verificar si la celda no está vacía
        If cell.value <> "" Then
            ' Agregar el valor al ComboBox
            ComboBox5.AddItem cell.value
        End
    Next cell
End Sub
Function AgregarValoresUnicosComboCartera()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim dict As Object
    
    ' Inicializar el diccionario para almacenar valores únicos
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Definir la hoja de trabajo
    Set ws = ThisWorkbook.Sheets("Clientes")
    
    ' Encontrar el último valor no vacío en la columna H
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    
    ' Verificar si lastRow es menor que 5, lo que indicaría que no hay datos en la columna H
    If lastRow < 5 Then
        Exit Function
    End If
    
    ' Definir el rango desde la fila 5 hasta el último valor en la columna H
    Set rng = ws.Range("J5:J" & lastRow)
    
    ' Iterar sobre cada celda en el rango y agregar valores únicos al diccionario
    For Each cell In rng
        If cell.value <> "" Then
            If Not dict.Exists(cell.value) Then
                dict.Add cell.value, Nothing
            End If
        End If
    Next cell
    
    ' Limpiar el ComboBox antes de agregar nuevos elementos
    ComboCartera.Clear
    ComboBox1.Clear
    
    ' Agregar los valores únicos del diccionario al ComboBox
    For Each key In dict.Keys
        ComboCartera.AddItem key
        ComboBox1.AddItem key
    Next key
End Function

