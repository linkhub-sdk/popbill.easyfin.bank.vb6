VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public jobID            As String   '�۾����̵�
Public jobState         As Integer  '�۾�����
Public startDate      As String   '��������
Public endDate      As String       '��������
Public errorCode        As Long     '�����ڵ�
Public errorReason      As String   '�����޽���
Public jobStartDT       As String   '���� �����Ͻ�
Public jobEndDT         As String   '���� �����Ͻ�
Public regDT            As String   '���� ��û�Ͻ�

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


