VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Singleton
'@PredeclaredId
Option Explicit

'===============================================================================

Private Sub Class_Initialize()
  If Not Me Is Common Then Err.Raise 425 'Invalid object use
End Sub

'===============================================================================

Public Function CollectTextShapes(ByVal Range As ShapeRange) As Collection
  Set CollectTextShapes = New Collection
  Dim Shape As Shape
  For Each Shape In Range
    If Shape.Type = cdrTextShape Then CollectTextShapes.Add Shape
  Next Shape
End Function

Public Sub UnlockAndDeleteSVGData()
  Dim Shape As Shape
  For Each Shape In ActivePage.Shapes.All
    Shape.Locked = False
    If Shape.Type = cdrNoShape Then Shape.Delete
  Next Shape
End Sub

Public Sub FindPairsAndSetIDs(ByVal TextShapes As Collection)
  Dim SourceShapes As Collection
  Set SourceShapes = New Collection
  lib_elvin.CopyCollection TextShapes, SourceShapes
  Dim Bar As IProgressBar
  Set Bar = ProgressBar.CreateNumeric(SourceShapes.Count / 2)
  Bar.Caption = "Поиск пар..."
  Dim Shape As Shape
  Dim FoundShape As Shape
  Do Until SourceShapes.Count = 0
    Set Shape = SourceShapes(1)
    Set FoundShape = FindPairByText(Shape, TextShapes)
    If Not FoundShape Is Nothing Then
      FoundShape.Name = VBA.CStr(Shape.StaticID)
      Shape.Name = VBA.CStr(FoundShape.StaticID)
      lib_elvin.RemoveElementFromCollection FoundShape, SourceShapes
    End If
    lib_elvin.RemoveElementFromCollection Shape, SourceShapes
    Set FoundShape = Nothing
    Bar.Update
  Loop
End Sub

Public Function FindPair(ByVal TextShapeInQuestion As Shape, _
                         ByVal TextShapes As Collection _
                         ) As Shape
  Dim Shape As Shape
  For Each Shape In TextShapes
    If Shape.StaticID = VBA.CLng(TextShapeInQuestion.Name) Then
      Set FindPair = Shape
      Exit Function
    End If
  Next Shape
End Function

Public Function FindPairByText(ByVal TextShapeInQuestion As Shape, _
                                ByVal TextShapes As Collection _
                                ) As Shape
  Dim Shape As Shape
  For Each Shape In TextShapes
    If Not TextShapeInQuestion Is Shape Then
      If TextShapeInQuestion.Text.Story.Text = Shape.Text.Story.Text Then
        Set FindPairByText = Shape
        Exit Function
      End If
    End If
  Next Shape
End Function

Public Sub SetTextShapesToPlaceholder(ByVal TextShapes As Collection)
  Dim Shape As Shape
  For Each Shape In TextShapes
    SetTextShapeToPlaceholder Shape
  Next Shape
End Sub

Public Sub SetTextShapeToPlaceholder(ByVal TextShape As Shape)
  TextShape.Text.Story.Text = PlaceholderText
End Sub

Public Sub RemoveNumbers(ByVal Pairs As Pairs, ByVal Page As ParsedPage)
  Dim tempPair As Pair
  For Each tempPair In Pairs.NewEnum
    Page.GetPairs.PullBack tempPair.Number
    tempPair.Number = 0
  Next tempPair
End Sub

Public Sub tryUnlockAndUngroup(ByVal Shape As Shape)
  If Shape Is Nothing Then Exit Sub
  Shape.Locked = False
  Shape.Ungroup
End Sub

Public Function CreateLongDashArrows(ByVal Layer As Layer) As Shape
    
  Dim SavedUnit As cdrUnit
  Dim Doc As Document
  Set Doc = Layer.Page.Parent.Parent
  SavedUnit = Doc.Unit
  Doc.Unit = cdrInch
    
  Dim crv As Curve
  Set crv = Doc.CreateCurve
  With crv.CreateSubPath(-0.038217, 0.097114)
    .AppendLineSegment 0.038319, 0.097071
    .AppendLineSegment 0.000008, 0.020555
    .AppendLineSegment -0.038217, 0.097114
    .Closed = True
  End With
  With crv.CreateSubPath(-0.000008, -0.020555)
    .AppendLineSegment 0.038224, -0.097114
    .AppendLineSegment -0.038319, -0.097071
    .AppendLineSegment -0.000008, -0.020555
    .Closed = True
  End With
  Set CreateLongDashArrows = Layer.CreateCurve(crv)
  CreateLongDashArrows.Outline.SetNoOutline
  CreateLongDashArrows.Fill.ApplyUniformFill CreateColor("CMYK,USER,0,0,0,30")
  
  Doc.Unit = SavedUnit
    
End Function

Public Function CreateTwoDashArrows(ByVal Layer As Layer) As Shape
    
  Dim SavedUnit As cdrUnit
  Dim Doc As Document
  Set Doc = Layer.Page.Parent.Parent
  SavedUnit = Doc.Unit
  Doc.Unit = cdrInch
  
  Dim crv As Curve
  Set crv = Doc.CreateCurve
  With crv.CreateSubPath(-0.038268, 0.020555)
    .AppendLineSegment 0.000004, 0.097091
    .AppendLineSegment 0.038268, 0.020551
    .AppendLineSegment -0.038268, 0.020555
    .Closed = True
  End With
  With crv.CreateSubPath(-0.038272, -0.020551)
    .AppendLineSegment 0.038264, -0.020559
    .AppendLineSegment -0.000004, -0.097091
    .AppendLineSegment -0.038272, -0.020551
    .Closed = True
  End With
  Set CreateTwoDashArrows = Layer.CreateCurve(crv)
  CreateTwoDashArrows.Outline.SetNoOutline
  CreateTwoDashArrows.Fill.ApplyUniformFill CreateColor("CMYK,USER,0,0,0,60")
  
  Doc.Unit = SavedUnit
    
End Function

Public Function IsShapeOverlapsOneOfRange(ByVal Shape As Shape, _
                                          ByVal Range As ShapeRange _
                                          ) As Boolean
  Dim Item As Shape
  For Each Item In Range
    If IsOverlapBox(Shape, Item) Then
      IsShapeOverlapsOneOfRange = True
      Exit Function
    End If
  Next Item
End Function

Public Function IsOverlapBox(ByVal Shape1 As Shape, _
                             ByVal Shape2 As Shape _
                             ) As Boolean
  Dim Rect As Rect
  Set Rect = Shape1.BoundingBox.Intersect(Shape2.BoundingBox)
  If Rect.Width = 0 And Rect.Height = 0 Then
    IsOverlapBox = False
  Else
    IsOverlapBox = True
  End If
End Function

Public Sub SetArrowsSymbolOnLine(ByVal Layer As Layer, _
                                 ByVal SymbolName As String, _
                                 ByVal Line As Shape, _
                                 ByVal ShapesToAvoid As ShapeRange)
  Dim SymbolWidth As Double
  Dim X As Double, Y As Double
  Dim ArrowsSymbol As Shape
  Dim Segment As Segment
  Set Segment = Line.Curve.Segments.First
  Segment.GetPointPositionAt X, Y
  Set ArrowsSymbol = ActiveLayer.CreateSymbol(X, Y, SymbolName)
  With ArrowsSymbol
    SymbolWidth = .SizeWidth
    If IsShapeOverlapsOneOfRange(ArrowsSymbol, ShapesToAvoid) Then
      Segment.GetPointPositionAt X, Y, _
        Segment.Length / 2 + SymbolWidth * OffsetToAvoidMult, _
        cdrAbsoluteSegmentOffset
      .SetPosition X, Y
    End If
    .RotationAngle = Segment.StartingControlPointAngle
    .Name = SymbolName
  End With
End Sub

'===============================================================================
' не пошли в дело
'===============================================================================

Private Function IsCloneOrCloned(ByVal Shape As Shape) As Boolean
  IsCloneOrCloned = True
  If Not Shape.CloneLink Is Nothing Then Exit Function
  If Not Shape.Clones.Count = 0 Then Exit Function
  IsCloneOrCloned = False
End Function

Private Sub ReplaceByClone(ByVal CloneMaster As Shape, ByVal ShapeToReplace As Shape)
  Dim Clone As Shape
  Set Clone = CloneMaster.Clone
  With Clone
    .PositionX = ShapeToReplace.PositionX
    .PositionY = ShapeToReplace.PositionY
    .RotationAngle = ShapeToReplace.RotationAngle
  End With
  ShapeToReplace.Delete
End Sub
