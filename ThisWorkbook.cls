VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
' --------------------------------------------------- '
' --- Para cambiar los controles a lo configurado --- '
' --------------------------------------------------- '
    Call Activa_Funciones
End Sub
Private Sub Workbook_BeforeClose(cancel As Boolean)
' --------------------------------------------------- '
' --- Para devolver los controles a la normalidad --- '
' --------------------------------------------------- '
    Call Desactiva_Funciones
End Sub
Private Sub Workbook_Activate()
' ---------------------------------------------------------------------- '
' --- Variable que vale NO en las macros que hacen cambio de archivo --- '
' ---------------------------------------------------------------------- '
    If CAMBIAR_ARCHIVO <> "NO" Then
        ' --- Para cambiar los controles a lo configurado --- '
        Call Activa_Funciones
    End If
End Sub
Private Sub Workbook_Deactivate()
' ---------------------------------------------------------------------- '
' --- Variable que vale NO en las macros que hacen cambio de archivo --- '
' ---------------------------------------------------------------------- '
    If CAMBIAR_ARCHIVO <> "NO" Then
        ' --- Para devolver los controles a la normalidad --- '
        Call Desactiva_Funciones
    End If
End Sub
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
' ------------------------------------------------------------------- '
' --- Variable que vale NO en las macros que hacen cambio de hoja --- '
' ------------------------------------------------------------------- '
'   If CAMBIAR_HOJA <> "NO" Then
        Application.ScreenUpdating = False
        Call Pantalla_Mostar
'  End If
End Sub
Public Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, cancel As Boolean)
' --------------------------------------------------------------------------------------------------------- '
' --- Esta macro activa el menú desplegable personalizado, incluyendo el botón derecho y crea uno nuevo --- '
' --------------------------------------------------------------------------------------------------------- '
    Cancela_Boton_Excel = 1                                     ' TRUE - Variable para cancelar el menu de EXCEL
    Call Crear_Menu                                             ' El número "0" = "False", mientras "1" = "True"

    On Error GoTo 0
    Application.CommandBars("Menu_Desplegable").ShowPopup       ' Muesta el menu
    cancel = Cancela_Boton_Excel
    Exit Sub
ErrorHandler:
    If Err.Number = 5 Then
        Call Crear_Menu
    Else
        MsgBox "Ha ocurrido un error: " & Err.Description, vbExclamation, minombre
    End If
    Cancela_Boton_Excel = 0                                     ' FALSE - Variable para activar el menu de EXCEL
    On Error Resume Next
End Sub
