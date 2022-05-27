VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParsedPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

'===============================================================================

Private Type typeThis
  Page As Page
  PairableShapes As Collection
  Pairs As Pairs
  LastNumber As Long
End Type
Private This As typeThis

'===============================================================================

Public Function Create(ByVal Page As Page) As ParsedPage
  With New ParsedPage
    Set Create = .Init(Page)
  End With
End Function

Friend Function Init(ByVal Page As Page) As ParsedPage
  With This
    Set .Page = Page
  End With
  Parse
  Set Init = Me
End Function

'===============================================================================

Public Property Get GetPairs() As Pairs
  Set GetPairs = This.Pairs
End Property

Public Function GetPairsFromShapes(ByVal Shapes As ShapeRange) As Pairs
  Set GetPairsFromShapes = Pairs.Create
  Dim Shape As Shape
  Dim Pair As Pair
  For Each Shape In Shapes
    Set Pair = This.Pairs.GetPairByShape(Shape)
    If Not Pair Is Nothing Then
      If Not GetPairsFromShapes.Exists(Pair.Number) Then _
        GetPairsFromShapes.Add Pair
    End If
  Next Shape
End Function

Public Property Get LastNumber() As Long
  LastNumber = This.LastNumber
End Property

Public Property Get Self() As ParsedPage
  Set Self = Me
End Property

'===============================================================================

Sub Parse()
  With This
    CreatePairableShapes
    CreatePairsAndLastNumber
  End With
End Sub

Private Sub CreatePairableShapes()
  With This
    Set .PairableShapes = New Collection
    Dim Shape As Shape
    For Each Shape In .Page.Shapes
      If Shape.Type = cdrTextShape And VBA.IsNumeric(Shape.Name) Then
        .PairableShapes.Add Shape
      End If
    Next Shape
  End With
End Sub

Private Sub CreatePairsAndLastNumber()
  With This
    Set .Pairs = Pairs.Create
    Dim SourceShapes As Collection
    Set SourceShapes = New Collection
    lib_elvin.CopyCollection .PairableShapes, SourceShapes
    Dim Shape As Shape
    Dim FoundShape As Shape
    Dim NewPair As Pair
    Do Until SourceShapes.Count = 0
      Set Shape = SourceShapes(1)
      Set FoundShape = Common.FindPair(Shape, SourceShapes)
      If Not FoundShape Is Nothing Then
        Set NewPair = Pair.Create(Shape, FoundShape)
        If NewPair.Number > .LastNumber Then .LastNumber = NewPair.Number
        .Pairs.Add NewPair
        lib_elvin.RemoveElementFromCollection FoundShape, SourceShapes
      End If
      lib_elvin.RemoveElementFromCollection Shape, SourceShapes
    Loop
  End With
End Sub

Private Function IsPairInCollection(ByVal Pair As Pair, _
                                    ByVal Pairs As Collection _
                                    ) As Boolean
  Dim tempPair As Pair
  For Each tempPair In Pairs
    If Pair.Number = tempPair.Number Then
      IsPairInCollection = True
      Exit Function
    End If
  Next tempPair
End Function
