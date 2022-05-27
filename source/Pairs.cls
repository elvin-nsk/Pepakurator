VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pairs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type typeThis
  Collection As Collection
End Type
Private This As typeThis

'===============================================================================

Public Function Create() As Pairs
  With New Pairs
    Set Create = .Init()
  End With
End Function

Friend Function Init() As Pairs
  With This
    Set .Collection = New Collection
  End With
  Set Init = Me
End Function

'===============================================================================

Public Sub Add(ByVal Item As Pair)
  If Not IsDuplicate(Item) Then This.Collection.Add Item
End Sub

Public Property Get Count() As Long
  Count = This.Collection.Count
End Property

Public Function Exists(ByVal Number As Long) As Boolean
  Exists = (Not IndexByNumber(Number) = 0)
End Function

Public Function GetPairByShape(ByVal PairableShape As Shape) As Pair
  Dim tempPair As Pair
  For Each tempPair In This.Collection
    If lib_elvin.IsSameShapes(tempPair.Shape1, PairableShape) Or _
       lib_elvin.IsSameShapes(tempPair.Shape2, PairableShape) Then
      Set GetPairByShape = tempPair
      Exit Function
    End If
  Next tempPair
End Function

'@DefaultMember
Public Property Get Item(ByVal Number As Long) As Pair
Attribute Item.VB_UserMemId = 0
  Set Item = This.Collection.Item(IndexByNumber(Number))
End Property

Public Function Range(ByVal FromNumber As Long, _
                      Optional ByVal ToNumber As Long _
                      ) As Pairs
  If ToNumber = 0 Then ToNumber = LastNumber
  Set Range = Create
  If FromNumber > ToNumber Then Exit Function
  Dim tempPair As Pair
  For Each tempPair In This.Collection
    If tempPair.Number >= FromNumber And _
       tempPair.Number <= ToNumber Then _
      Range.Add tempPair
  Next tempPair
End Function

Public Sub PushForward(ByVal Number As Long)
  Dim PushedNumber As Long
  PushedNumber = Number
  Dim Pool As Pairs
  Set Pool = Range(Number)
  If Pool.Count = 0 Then Exit Sub
  Dim tempPair As Pair
  For Each tempPair In Pool.NewEnum
    If tempPair.Number = PushedNumber Then
      tempPair.Number = tempPair.Number + 1
      PushedNumber = tempPair.Number
    End If
  Next tempPair
End Sub

Public Sub PullBack(ByVal Number As Long)
  Dim Pool As Pairs
  Set Pool = Range(Number)
  If Pool.Count = 0 Then Exit Sub
  Dim tempPair As Pair
  For Each tempPair In Pool.NewEnum
    tempPair.Number = tempPair.Number - 1
  Next tempPair
End Sub

Public Property Get FirstNumber() As Long
  FirstNumber = LastNumber
  Dim tempPair As Pair
  For Each tempPair In This.Collection
    If Not tempPair.Number = 0 Then
      If FirstNumber > tempPair.Number Then FirstNumber = tempPair.Number
    End If
  Next tempPair
End Property

Public Property Get LastNumber() As Long
  Dim tempPair As Pair
  For Each tempPair In This.Collection
    If LastNumber < tempPair.Number Then LastNumber = tempPair.Number
  Next tempPair
End Property

Public Property Get FreeNumber() As Long
  If This.Collection.Count = 0 Then Exit Property
  Dim i As Long
  For i = 1 To This.Collection.Count
    FreeNumber = FreeNumber + 1
    If Not Self.Exists(FreeNumber) Then Exit Property
  Next i
End Property

Public Function NewEnum() As EnumHelper
  With New EnumHelper
    Set .EnumVariant = This.Collection.[_NewEnum]
    Set NewEnum = .Self
  End With
End Function

Public Sub Remove(ByVal Number As Long)
  This.Collection.Remove IndexByNumber(Number)
End Sub

Public Property Get Self() As Pairs
  Set Self = Me
End Property

'===============================================================================

Private Function IsDuplicate(ByVal Pair As Pair) As Boolean
  Dim tempPair As Pair
  For Each tempPair In This.Collection
    If tempPair.IsSame(Pair) Then
      IsDuplicate = True
      Exit Function
    End If
  Next tempPair
End Function

Private Function IndexByNumber(ByVal Number As Long) As Long
  With This.Collection
    If .Count = 0 Then Exit Function
    Dim i As Long
    For i = 1 To .Count
      If .Item(i).Number = Number Then
        IndexByNumber = i
        Exit Function
      End If
    Next i
  End With
End Function

Private Sub ThrowInvalidRange()
  Err.Raise vbObjectError Or 32, , "Invalid range"
End Sub
