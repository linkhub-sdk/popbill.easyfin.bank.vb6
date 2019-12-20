VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBEasyFinBankSearchDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public tid    As String
Public trdate        As String
Public trserial        As Long
Public trdt As String
Public accIn As String
Public accOut As String
Public balance As String
Public remark1 As String
Public remark2 As String
Public remark3 As String
Public remark4 As String
Public regDT As String
Public memo As String


Public Sub fromDictionary(dic As Variant)
    tid = dic.Item("tid")
    trdate = dic.Item("trdate")
    trserial = dic.Item("trserial")
    trdt = dic.Item("trdt")
    accIn = dic.Item("accIn")
    accOut = dic.Item("accOut")
    balance = dic.Item("balance")
    remark1 = dic.Item("remark1")
    remark2 = dic.Item("remark2")
    remark3 = dic.Item("remark3")
    remark4 = dic.Item("remark4")
    regDT = dic.Item("regDT")
    memo = dic.Item("memo")
End Sub


