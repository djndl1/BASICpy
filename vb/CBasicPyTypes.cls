VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBasicPyTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' StopIteration error codes — exhaustion sentinel for IIterator protocol
' Return CVErr(StopIteration) from IIterator_NextItem() to signal end of iteration
Public Enum StopIterationCode
    StopIteration = vbObjectError + 512
End Enum
