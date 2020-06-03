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

Public contractDT        As String
Public useEndDate         As String
Public baseDate         As Integer
Public contractState            As Integer
Public closeRequestYN   As Boolean
Public useRestrictYN    As Boolean
Public closeOnExpired   As Boolean
Public unPaidYN         As Boolean

Public Sub fromDictionary(dic As Variant)
    accountNumber = dic.Item("accountNumber")
    bankCode = dic.Item("bankCode")
    accountName = dic.Item("accountName")
    accountType = dic.Item("accountType")
    state = dic.Item("state")
    regDT = dic.Item("regDT")
    memo = dic.Item("memo")
    
    contractDT = dic.Item("contractDT")
    useEndDate = dic.Item("useEndDate")
    baseDate = dic.Item("baseDate")
    contractState = dic.Item("contractState")
    closeRequestYN = dic.Item("closeRequestYN")
    useRestrictYN = dic.Item("useRestrictYN")
    closeOnExpired = dic.Item("closeOnExpired")
    unPaidYN = dic.Item("unPaidYN")
End Sub

