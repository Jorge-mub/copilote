Attribute VB_Name = "prote"
Sub proteccion()
    'ThisWorkbook.Protect Password:="3097", structure:=True
    'Conos.Unprotect Password:="3097"
    'Conos.Visible = xlSheetHidden
    'consumo.Unprotect Password:="3097"
    'consumo.Visible = xlSheetHidden
    
    ThisWorkbook.Unprotect Password:="3097"
    Exist.Unprotect Password:="3097"
    Exist.Visible = xlSheetVisible
    consumo.Unprotect Password:="3097"
    consumo.Visible = xlSheetVisible
    
    'gruposHerramientas.Unprotect Password:="3097"
    'gruposHerramientas.Visible = xlSheetHidden
    'Interfaz.Unprotect Password:="3097"
    'Interfaz.Visible = xlSheetVisible
    'OPsheet.Unprotect Password:="3097"
    'OPsheet.Visible = xlSheetHidden
    'sumaConsumos.Unprotect Password:="3097"
    'sumaConsumos.Visible = xlSheetHidden
End Sub

