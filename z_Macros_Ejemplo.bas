Attribute VB_Name = "z_Macros_Ejemplo"
Sub Graba_Archivo()
' ------------------------------------------------------------------ '
' --- Macro para guardar con el control de Excel inglés (Ctrl+s) --- '
' ------------------------------------------------------------------ '
    ActiveWorkbook.Save
End Sub
Sub Guarda_Cierra_Archivo()
Attribute Guarda_Cierra_Archivo.VB_ProcData.VB_Invoke_Func = " \n14"
    a_abiertos = Workbooks.Count
    If a_abiertos > 2 Then
        MsgBox "Debes cerrar los demás archivos de Excel que tienes abiertos antes de cerrar la aplicación.", _
                    vbInformation, minombre
        Exit Sub
    End If
    G_Cambios = MsgBox("¿Deseas guardar los cambios?" & vbNewLine, vbYesNo + vbInformation, minombre)
    If G_Cambios = vbYes Then
        ActiveWorkbook.Close SaveChanges:=True
    Else
        ActiveWorkbook.Close SaveChanges:=False
    End If
End Sub
Sub Cambiar_Ctrl_V()
' Cambiar el comando Ctrl+V por pegar contenido (fórmulas o valores), para que no peque los formatos.
Dim Mensaje_CERO As Integer
    On Error GoTo MENSAJE
    Selection.PasteSpecial Paste:=xlPasteFormulas
    Mensaje_CERO = 1
MENSAJE:
   If Mensaje_CERO = 0 Then
       On Error Resume Next
       Application.CutCopyMode = False
       ActiveCell.Select
       MsgBox "No se pueden pegar estos valores.", vbExclamation, minombre
    Else
        Exit Sub
    End If
End Sub
Sub Activa_Menu()
Dim cancel As Integer
' Para activar el menu fuera del clic derecho, no se puede llamar desde ese apartado porque no funciona bien
    Cancela_Boton_Excel = 0                               ' Variable para cancelar el menu de EXCEL

    Call Crear_Menu

    On Error GoTo 0
    Muestra_Mensaje = "SI"
    Application.CommandBars("Menu_Desplegable").ShowPopup     ' Muesta el menu
    cancel = Cancela_Boton_Excel
    Exit Sub
ErrorHandler:
    If Err.Number = 5 Then
        Call Crear_Menu
    Else
        MsgBox "Ha ocurrido un error: " & Err.Description, vbExclamation, minombre
    End If

    On Error Resume Next
    Muestra_Mensaje = "NO"
End Sub
Sub CambiarCtrl_X_Ctrl_C()
' Cambiar el comando CORTAR por COPIAR para evitar que fallen las fórmulas
'    CAMBIAR_HOJA = "NO"         ' Se pone esta variable para que no haga nada y deje copiar entre hojas
    Selection.Copy
End Sub
Sub IR_INICIO()
    Worksheets("TABLAS").Activate
End Sub
Sub IR_PRIMERA()
    Worksheets("PRIMERA").Activate
End Sub
Sub IR_SEGUNDA()
    Worksheets("SEGUNDA").Activate
End Sub
Sub IR_TERCERA()
    Worksheets("TERCERA").Activate
End Sub
