VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "New Record Data"
   ClientHeight    =   9444.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10488
   OleObjectBlob   =   "Data Entry Userform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Methods that execute runtime user form commands

Private Sub CommandButton1_Click()
    Worksheets("Base individuos").Activate
    'Find the new blank row where a new collection record will be added
    
    a = Range("C99999").End(xlUp).Row + 1
    If a < 3 Then
        a = 4
    End If
    'Save the date input data in the corresponding column of the attribute
    Cells(a, 2) = Date
    
    'Save user name
    Cells(a, 3) = TextBox2.Value & " " & TextBox3.Value
    
    'Save activities selected
    If CheckBox15.Value Then
        Cells(a, 4) = 1
    End If
    If CheckBox16.Value Then
        Cells(a, 5) = 1
    End If
    If CheckBox17.Value Then
        Cells(a, 6) = 1
    End If
    If CheckBox18.Value Then
        Cells(a, 7) = 1
    End If
    If CheckBox19.Value Then
        Cells(a, 8) = 1
    End If
    
    'ESTIMATE UPDATED INCOME WITH NEW USER DATA INPUT
    Cells(a, 9) = Cells(a, 9).Value
    '------------------------------------------------------
    
    'Save guide attribute boolean option
    Cells(a, 10) = "SI"
    
    'Save country attribute
    Cells(a, 11) = ListBox2.Value
    
    'Save survey info: ¿Primera vez que visita?
    If CheckBox1.Value Then
        Cells(a, 12) = "Si"
    Else
        Cells(a, 12) = "No"
    End If
    
    'Save survey info: ¿Alguna vez ha sembrado un árbol?
    If CheckBox6.Value Then
        Cells(a, 13) = "Si"
    Else
        Cells(a, 13) = "No"
    End If
    
    'save survey info: ¿Alguna vez ha realizado una caminata ecológica?
    If CheckBox7.Value Then
        Cells(a, 14) = "Si"
    Else
        Cells(a, 14) = "No"
    End If
    
    'Save survey info
    If CheckBox9.Value Then
        Cells(a, 15) = "Nada"
    ElseIf CheckBox10.Value Then
        Cells(a, 15) = "Poco"
    ElseIf CheckBox11.Value Then
        Cells(a, 15) = "Neutral"
    ElseIf CheckBox12.Value Then
        Cells(a, 15) = "Algo"
    Else
        Cells(a, 15) = "Mucho"
    End If
    
    
    'Save birth date attribute
    Cells(a, 16) = TextBox4.Value
    
    'Save Email attribute
    Cells(a, 17) = TextBox8.Value
    
    Worksheets("Inicio").Activate
    
    'Close interface (optional) evaluate if new form is opened in order to confirm synch with SQL
    'UserForm1.Hide
    Unload Me
End Sub

'Clear data command to execute when 'delete' button is clicked
Private Sub CommandButton2_Click()
    Worksheets("Base individuos").Activate
    
    b = Range("C99999").End(xlUp).Row
    If b <= 3 Then
        b = 4
    End If
    'Borrar último registro
    Range(Cells(b, 2), Cells(b, 8)).Clear
    Range(Cells(b, 10), Cells(b, 30)).Clear
    Cells(b + 1, 9).Select
    Selection.Copy
    Cells(b, 9).Select
    Selection.PasteSpecial
End Sub

'Load country data to user interface
Private Sub UserForm_Initialize()
    ListBox2.RowSource = "Hoja1!A1:A228"
End Sub
