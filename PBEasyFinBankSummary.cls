VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBEasyFinBankSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public count            As Long
Public cntAccIn  As Long
Public cntAccOut         As Long
Public totalAccIn      As Long
Public totalAccOut     As Long

Public Sub fromDictionary(dic As Variant)

    count = dic.Item("count")
    cntAccIn = dic.Item("cntAccIn")
    cntAccOut = dic.Item("cntAccOut")
    totalAccIn = dic.Item("totalAccIn")
    totalAccOut = dic.Item("totalAccOut")
    
End Sub

