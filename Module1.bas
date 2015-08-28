Attribute VB_Name = "DB_connect"
Option Explicit
Public DBconnect As New ADODB.Connection
Public direccion As String

Public Sub conectar()
    Dim sarchibobase As String  'variable para guardar
    sarchibobase = direccion    'url de la bd
    DBconnect.Provider = "microsoft.jet.oledb.4.0" 'driver
    DBconnect.CursorLocation = adUseServer
    DBconnect.ConnectionString = sarchibobase 'Nombre del archivo, string.
    DBconnect.Mode = adModeReadWrite     'modo lectura escritura
    DBconnect.Open
End Sub
