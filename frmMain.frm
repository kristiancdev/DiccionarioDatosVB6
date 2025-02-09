VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    TestDictionary
End Sub

' Agregar referencia a "Microsoft Scripting Runtime" (Tools > References)
Private Sub TestDictionary()
    Dim dict As New Scripting.Dictionary
    dict.Add "ID1", "Juan Pérez"
    dict.Add "ID2", "Ana Gómez"
    dict.Add "ID3", "Carlos López"
    
    ' Acceder a un valor por clave
    MsgBox dict("ID2") ' Muestra "Ana Gómez"
    
    ' Verificar si una clave existe
    If dict.Exists("ID1") Then
        MsgBox "La clave ID1 existe."
    End If
    
    ' Recorrer el diccionario
    Dim key As Variant
    For Each key In dict.Keys
        Debug.Print "Clave: " & key & ", Valor: " & dict(key)
    Next key
End Sub
