VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBEasyFinBankJobState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public jobID            As String   '작업아이디
Public jobState         As Integer  '작업상태
Public startDate      As String   '시작일자
Public endDate      As String       '종료일자
Public errorCode        As Long     '오류코드
Public errorReason      As String   '오류메시지
Public jobStartDT       As String   '수집 시작일시
Public jobEndDT         As String   '수집 종료일시
Public regDT            As String   '수집 요청일시

Public Sub fromDictionary(dic As Variant)

    jobID = dic.Item("jobID")
    jobState = dic.Item("jobState")
    startDate = dic.Item("startDate")
    endDate = dic.Item("endDate")
    errorCode = dic.Item("errorCode")
    errorReason = dic.Item("errorReason")
    jobStartDT = dic.Item("jobStartDT")
    jobEndDT = dic.Item("jobEndDT")
    regDT = dic.Item("regDT")
End Sub



