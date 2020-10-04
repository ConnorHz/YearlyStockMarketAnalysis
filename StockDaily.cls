VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StockDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private mTicker As String
Private mDay As Date
Private mOpenPrice As Currency
Private mHighPrice As Currency
Private mLowPrice As Currency
Private mClosePrice As Currency
Private mVolume As Long

Public Property Get ticker() As String
    ticker = mTicker
End Property

Property Let ticker(value As String)
     mTicker = value
End Property

Public Property Get Day() As Date
    Day = mDay
End Property

Property Let Day(value As Date)
     mDay = value
End Property

Public Property Get OpenPrice() As Currency
    OpenPrice = mOpenPrice
End Property

Property Let OpenPrice(value As Currency)
     mOpenPrice = value
End Property

Public Property Get HighPrice() As Currency
    HighPrice = mHighPrice
End Property

Property Let HighPrice(value As Currency)
     mHighPrice = value
End Property

Public Property Get LowPrice() As Currency
    LowPrice = mLowPrice
End Property

Property Let LowPrice(value As Currency)
     mLowPrice = value
End Property

Public Property Get ClosePrice() As Currency
    ClosePrice = mClosePrice
End Property

Property Let ClosePrice(value As Currency)
     mClosePrice = value
End Property

Public Property Get Volume() As Long
    Volume = mVolume
End Property

Property Let Volume(value As Long)
     mVolume = value
End Property

Public Function GetPriceChange(daily As StockDaily) As Currency
     
    GetPriceChange = 0
    
    ' If the passed in daily is the same, just return zero
    If (daily Is Me) Then
        MsgBox (Me.ticker & ": Identical Daily was passed as parameter")
        Exit Function
    End If
    
    If (Me.ticker <> daily.ticker) Then
        MsgBox ("Mismatched Ticker Symbol")
        Exit Function
    End If
    
    
    If (Me.Day > daily.Day) Then
        GetPriceChange = daily.OpenPrice - Me.ClosePrice
    Else
        GetPriceChange = daily.ClosePrice - Me.OpenPrice
    End If
        
End Function

' Percent Change can be calculate using the formula (new_value-old_value)/old_value
Public Function GetPercentChange(daily As StockDaily) As Double
     
    GetPercentChange = 0
    
    ' If the passed in daily is the same, just return zero
    If (daily Is Me) Then
        MsgBox (Me.ticker & ": Identical Daily was passed as parameter")
        Exit Function
    End If
    
    If (Me.ticker <> daily.ticker) Then
        MsgBox ("Mismatched Ticker Symbol")
        Exit Function
    End If
    
    
    If (Me.Day > daily.Day) Then
        If Me.ClosePrice <> 0 Then
            GetPercentChange = (daily.OpenPrice - Me.ClosePrice) / Me.ClosePrice
        Else
            GetPercentChange = 0
        End If
    Else
        If Me.OpenPrice <> 0 Then
            GetPercentChange = (daily.ClosePrice - Me.OpenPrice) / Me.OpenPrice
        Else
            GetPercentChange = 0
        End If
    End If
        
End Function




