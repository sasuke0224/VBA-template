VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCompanyId As Integer     
Private pCompanyName As String    
Private pTrades As Collection     


Public Property Get CompanyId() As Integer
    CompanyId = pCompanyId
End Property
Public Property Let CompanyId(Value As Integer)
    pCompanyId = Value
End Property

Public Property Get CompanyName() As String
    CompanyName = pCompanyName
End Property
Public Property Let CompanyName(Value As String)
    pCompanyName = Value
End Property

Public Property Get Trades() As Collection
    Set Trades = pTrades
End Property
Public Property Set Trades(Value As Collection)
    Set pTrades = Value
End Property

Private Sub Class_Initialize()
    Set pTrades = New Collection
End Sub

Public Sub AddTrade(objTrade As Trade)
    pTrades.Add objTrade
End Sub
