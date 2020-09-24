VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Arr() As Any) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Long)

Dim CV(0 To 31) As Long
Dim IA() As Integer, IH(0 To 5) As Long, IP As Long
Dim LA() As Long, LH(0 To 5) As Long, LP As Long

Dim CA(0 To 2047) As Long

' returns character position of match or 0 if not found
Public Function HasDuplicatesM(Text As String) As Long
    Dim C As Long, CP As Long, CP2 As Long, CPV As Long, L As Long, P As Long
    ' cache the length
    L = Len(Text)
    ' more than two!
    If L > 2 Then
        ' change Integer safe array header information
        IH(3) = StrPtr(Text): IH(4) = L
        ' then start our work
        For L = 0 To L - 1
            ' cache current character code
            C = IA(L) And &HFFFF&
            ' find out index in array
            CP = (C And Not 31&) \ 32&
            ' find out bit in array item
            CP2 = CV(C And 31&)
            ' get the value of the array item
            CPV = CA(CP)
            ' is the bit turned on?
            If (CPV And CP2) = 0 Then
                ' turn it on
                CA(CP) = CPV Or CP2
            Else
                ' it was turned on, a duplicate is found!
                HasDuplicatesM = L + 1
                ' jump off the vagon!
                Exit For
            End If
        Next L
        ' cleanup: basically the fastest way to nullify a memory block in VB6, I think?
        Erase CA
    ' only two!
    ElseIf L = 2 Then
        ' just check the short way
        HasDuplicatesM = -(AscW(Text) = AscW(Right$(Text, 1)))
    End If
End Function

Private Sub Class_Initialize()
    ' cache bit values
    CV(0) = &H1&
    CV(1) = &H2&
    CV(2) = &H4&
    CV(3) = &H8&
    CV(4) = &H10&
    CV(5) = &H20&
    CV(6) = &H40&
    CV(7) = &H80&
    CV(8) = &H100&
    CV(9) = &H200&
    CV(10) = &H400&
    CV(11) = &H800&
    CV(12) = &H1000&
    CV(13) = &H2000&
    CV(14) = &H4000&
    CV(15) = &H8000&
    CV(16) = &H10000
    CV(17) = &H20000
    CV(18) = &H40000
    CV(19) = &H80000
    CV(20) = &H100000
    CV(21) = &H200000
    CV(22) = &H400000
    CV(23) = &H800000
    CV(24) = &H1000000
    CV(25) = &H2000000
    CV(26) = &H4000000
    CV(27) = &H8000000
    CV(28) = &H10000000
    CV(29) = &H20000000
    CV(30) = &H40000000
    CV(31) = &H80000000
    ' prepare Integer array for Text data
    IH(0) = 1: IH(1) = 2
    IP = ArrPtr(IA)
    ' prepare Long array for hacking Integer array on/off
    LH(0) = 1: LH(1) = 4: LH(4) = &H7FFFFFFF
    LP = ArrPtr(LA)
    ' enable safe array hack for Long array
    PutMem4 LP, VarPtr(LH(0))
    ' now do the same operation for Integer array... without PutMem4!
    LH(3) = IP: LA(0) = VarPtr(IH(0))
End Sub

Private Sub Class_Terminate()
    ' disable Integer safe array hack
    LH(3) = IP
    LA(0) = 0
    ' disable Long safe array hack
    LH(3) = LP
    LA(0) = 0
End Sub
