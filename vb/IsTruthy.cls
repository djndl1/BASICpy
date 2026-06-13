VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IsTruthy"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Implements IPredicate

Private Function IPredicate_Test(ByVal value As Variant) As Boolean
    IPredicate_Test = CBool(value)
End Function
