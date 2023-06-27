VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Trade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pId As Integer       
Private pName As String       
Private pAmount As Integer    

Public Property Get Id() As Integer
    Id = pId
End Property
Public Property Let Id(Value As Integer)
    pId = Value
End Property

Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(Value As String)
    pName = Value
End Property

Public Property Get Amount() As Integer
    Amount = pAmount
End Property
Public Property Let Amount(Value As Integer)
    pAmount = Value
End Property
