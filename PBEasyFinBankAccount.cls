VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBEasyFinBankAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public accountNumber    As String
Public bankCode     As String
Public accountName  As String
Public accountType As String
Public state       As Integer
Public regDT       As String
Public memo        As String


Public Sub fromDictionary(dic As Variant)
    accountNumber = dic.Item("accountNumber")
    bankCode = dic.Item("bankCode")
    accountName = dic.Item("accountName")
    accountType = dic.Item("accountType")
    state = dic.Item("state")
    regDT = dic.Item("regDT")
    memo = dic.Item("memo")
End Sub

