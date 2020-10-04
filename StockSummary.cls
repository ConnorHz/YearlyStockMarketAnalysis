VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StockSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private mTicker As String
Private mYear As Integer
Private mPriceChange As Currency
Private mPercentChange As String
Private mTotalVolume As LongLong


Public Property Get ticker() As String
    ticker = mTicker
End Property

Property Let ticker(value As String)
     mTicker = value
End Property

Public Property Get Year() As Integer
    Year = mYear
End Property

Property Let Year(value As Integer)
     mYear = value
End Property


Public Property Get PriceChange() As Currency
    PriceChange = mPriceChange
End Property

Property Let PriceChange(value As Currency)
     mPriceChange = value
End Property

Public Property Get PercentChange() As Double
    PercentChange = mPercentChange
End Property

Property Let PercentChange(value As Double)
     mPercentChange = value
End Property


Public Property Get TotalVolume() As LongLong
    TotalVolume = mTotalVolume
End Property

Property Let TotalVolume(value As LongLong)
     mTotalVolume = value
End Property

