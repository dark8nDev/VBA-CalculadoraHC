VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   12360
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numberOfEquipments As Integer
Dim emissionFactor As Double
Dim RunTimeComboBoxes As Collection

Private Sub Frame1_Click()

End Sub

Private Sub UserForm_Initialize()
    
    emissionFactor = 0.126
    
    TxEmiFactor = emissionFactor
    
    Call initializeEquipmentListing
    Call AddEquipment
    
End Sub

Private Sub BtAddEquipment_Click()
    
    Call AddEquipment
    
End Sub

Private Sub BtRemoveEquipment_Click()
    
    If numberOfEquipments > 0 Then
    
        Me.Controls.Remove Me.Controls(newCbName & numberOfEquipments).Name
        Me.Controls.Remove Me.Controls(newTxName1 & numberOfEquipments).Name
        Me.Controls.Remove Me.Controls(newTxName2 & numberOfEquipments).Name
        Me.Controls.Remove Me.Controls(newTxName3 & numberOfEquipments).Name
        Me.Controls.Remove Me.Controls(newTxName4 & numberOfEquipments).Name
        
        numberOfEquipments = numberOfEquipments - 1
        
        RunTimeComboBoxes.Remove RunTimeComboBoxes.Count
        
        Me.Frame1.ScrollHeight = Me.Frame1.ScrollHeight - (defHeight + defMargin)
        
        TxGrandTotal.Value = ""
    
    End If
    
End Sub

Private Sub BtReset_Click()
    
    If numberOfEquipments > 0 Then
    
        Dim i As Integer
        
        For i = 1 To numberOfEquipments
        
            Me.Controls.Remove Me.Controls(newCbName & i).Name
            Me.Controls.Remove Me.Controls(newTxName1 & i).Name
            Me.Controls.Remove Me.Controls(newTxName2 & i).Name
            Me.Controls.Remove Me.Controls(newTxName3 & i).Name
            Me.Controls.Remove Me.Controls(newTxName4 & i).Name
            
        Next i
        
        Call initializeEquipmentListing
        
    End If

End Sub

Private Sub BtCalculate_Click()
    
    If numberOfEquipments > 0 Then
    
        Dim total As Double
        Dim tbConsume As control
        Dim txQuantity As control
        Dim txUse As control
        Dim txTotal As control
        Dim i As Integer
        
        total = 0
        
        ' Iterar sobre los TextBox que se van a operar
        For i = 1 To numberOfEquipments
            Set tbConsume = Me.Controls(newTxName1 & i)
            Set txQuantity = Me.Controls(newTxName2 & i)
            Set txUse = Me.Controls(newTxName3 & i)
            Set txTotal = Me.Controls(newTxName4 & i)
            On Error Resume Next
            txTotal.Value = CDbl(tbConsume.Value) * CDbl(txQuantity.Value) * CDbl(txUse.Value) / 1000
            total = total + txTotal.Value
            On Error GoTo 0
        Next i
        
        total = total * emissionFactor
        
        TxGrandTotal.Value = total
        
    End If
    
End Sub

Sub initializeEquipmentListing()
    
    numberOfEquipments = 0
    
    Set RunTimeComboBoxes = New Collection
    
    With Me.Frame1
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = .InsideHeight
    End With
    
    TxGrandTotal.Value = ""
    
End Sub

Sub AddEquipment()

    numberOfEquipments = numberOfEquipments + 1

    ' Declarar variables para nuevos controles y clase ComboBoxHandler
    Dim container As Object
    Dim newComboBox As Object
    Dim newTextBox1 As Object
    Dim newTextBox2 As Object
    Dim newTextBox3 As Object
    Dim newTextBox4 As Object
    Dim cboHandler As ComboBoxHandler
    
    ' Definir conteneder de controles
    Set container = Me.Frame1
    
    ' Instanciar clase
    Set cboHandler = New ComboBoxHandler
    
    ' Crear nuevo ComboBox y establecer propiedades
    Set newComboBox = container.Controls.Add("Forms.ComboBox.1", , True)
    With newComboBox
        .Name = newCbName & numberOfEquipments ' Nombre dinamico
        .Top = defTop + ((defHeight + defMargin) * (numberOfEquipments - 1)) ' Posicion
        .Left = defLeftCb
        .Width = 200 ' Dimensiones
        .Height = defHeight
        .RowSource = "TablaEquipos[Equipo]" ' Lista de valores
    End With
    
    ' Crear nuevo TextBox de consumo y establecer propiedades
    Set newTextBox1 = container.Controls.Add("Forms.TextBox.1", , True)
    With newTextBox1
        .Name = newTxName1 & numberOfEquipments ' Nombre dinamico
        .Top = defTop + ((defHeight + defMargin) * (numberOfEquipments - 1)) ' Posición
        .Left = defLeftTx1
        .Width = 100 ' Dimensiones
        .Height = defHeight
        .Enabled = False ' Bloquear edicion
    End With
    
    ' Crear nuevo TextBox de cantidad y establecer propiedades
    Set newTextBox2 = container.Controls.Add("Forms.TextBox.1", , True)
    With newTextBox2
        .Name = newTxName2 & numberOfEquipments ' Nombre dinamico
        .Top = defTop + ((defHeight + defMargin) * (numberOfEquipments - 1)) ' Posición
        .Left = defLeftTx2
        .Width = 50 ' Dimensiones
        .Height = defHeight
    End With
    
    ' Crear nuevo TextBox de uso y establecer propiedades
    Set newTextBox3 = container.Controls.Add("Forms.TextBox.1", , True)
    With newTextBox3
        .Name = newTxName3 & numberOfEquipments 'Nombre dinamico
        .Top = defTop + ((defHeight + defMargin) * (numberOfEquipments - 1)) ' Posición
        .Left = defLeftTx3
        .Width = 50 ' Dimensiones
        .Height = defHeight
    End With
    
    ' Crear nuevo TextBox de total y establecer propiedades
    Set newTextBox4 = container.Controls.Add("Forms.TextBox.1", , True)
    With newTextBox4
        .Name = newTxName4 & numberOfEquipments ' Nombre dinamico
        .Top = defTop + ((defHeight + defMargin) * (numberOfEquipments - 1)) ' Posición
        .Left = defLeftTx4
        .Width = 100 ' Dimensiones
        .Height = defHeight
        .Enabled = False ' Bloquear edicion
    End With
    
    Set cboHandler.cbRef = newComboBox
    Set cboHandler.txRef = newTextBox1
    Set cboHandler.wsRef = ThisWorkbook.Sheets("Equipos")
    
    RunTimeComboBoxes.Add Item:=cboHandler
    
    Set cboHandler = Nothing
    
    container.ScrollHeight = container.ScrollHeight + (defHeight + defMargin)
    
End Sub
