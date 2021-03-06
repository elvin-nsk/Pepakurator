Attribute VB_Name = "lib_elvin"
'===============================================================================
' ??????           : lib_elvin
' ??????           : 2021.11.15
' ?????            : elvin-nsk (me@elvin.nsk.ru)
' ??????????? ???  : dizzy (?? ??????? CtC), Alex Vakulenko
'                    ? ??.
' ????????         : ?????????? ??????? ??? ???????? ?? elvin-nsk
' ?????????????    :
' ???????????      : ???????????????
'===============================================================================

Option Explicit

'===============================================================================
' # ????????? ?????????? ??????
'===============================================================================

Private Type type_LayerProps
  Visible As Boolean
  Printable As Boolean
  Editable As Boolean
End Type

Private startTime#

'===============================================================================
' ????????? ??????????
'===============================================================================

Public Const CustomError = vbObjectError Or 32

Public Type type_Matrix
  d11 As Double
  d12 As Double
  d21 As Double
  d22 As Double
  tx As Double
  ty As Double
End Type

'===============================================================================
' ??????? ?????? ??????????
'===============================================================================

'-------------------------------------------------------------------------------
' ???????          : BoostStart, BoostFinish
' ??????           : 2020.04.30
' ??????           : dizzy, elvin-nsk
' ??????????       : ???????????? ???????????? ?? CtC
' ???????????      : ???????????????
'
' ?????????:
' ~~~~~~~~~~
'
'
' ?????????????:
' ~~~~~~~~~~~~~~
'
'-------------------------------------------------------------------------------
Public Sub BoostStart(Optional ByVal UnDo$ = "", Optional ByVal Optimize = True)
  If UnDo <> "" And Not (ActiveDocument Is Nothing) Then ActiveDocument.BeginCommandGroup UnDo
  If Optimize Then Optimization = True
  EventsEnabled = False
  If Not ActiveDocument Is Nothing Then
    With ActiveDocument
      .SaveSettings
      .PreserveSelection = False
      .Unit = cdrMillimeter
      .WorldScale = 1
      .ReferencePoint = cdrCenter
    End With
  End If
End Sub
Public Sub BoostFinish(Optional ByVal EndUndoGroup = True)
  EventsEnabled = True
  Optimization = False
  If Not ActiveDocument Is Nothing Then
    With ActiveDocument
      .RestoreSettings
      .PreserveSelection = True
      If EndUndoGroup Then .EndCommandGroup
    End With
    ActiveWindow.Refresh
  End If
  Application.Refresh
  Application.Windows.Refresh '?????????
End Sub

'===============================================================================
' ??????? ??????????? ? ????????? ??????
'===============================================================================

'??? ??????? ?? ???? ?????????, ??????? ??????-???????? - ?? ???? ????
'??? ???????? ???????????, ??? ??????? ?? ????? guides ???????????
Public Function FlattenPagesToLayer(ByVal LayerName$) As Layer

  Dim DL As Layer: Set DL = ActiveDocument.MasterPage.DesktopLayer
  Dim DLstate As Boolean: DLstate = DL.Editable
  Dim P As Page
  Dim L As Layer
  
  DL.Editable = False
  
  For Each P In ActiveDocument.Pages
    For Each L In P.Layers
      If L.IsSpecialLayer Then
        L.Shapes.All.Delete
      Else
        L.Activate
        L.Editable = True
        With L.Shapes.All
          .MoveToLayer DL
          .OrderToBack
        End With
        L.Delete
      End If
    Next
    If P.Index <> 1 Then P.Delete
  Next
  
  Set FlattenPagesToLayer = ActiveDocument.Pages.First.CreateLayer(LayerName)
  FlattenPagesToLayer.MoveBelow ActiveDocument.Pages.First.GuidesLayer
  
  For Each L In ActiveDocument.MasterPage.Layers
    If Not L.IsSpecialLayer Or L.IsDesktopLayer Then
      L.Activate
      L.Editable = True
      With L.Shapes.All
        .MoveToLayer FlattenPagesToLayer
        .OrderToBack
      End With
      If Not L.IsSpecialLayer Then L.Delete
    Else
      L.Shapes.All.Delete
    End If
  Next
  
  FlattenPagesToLayer.Activate
  DL.Editable = DLstate

End Function

'????????? ?????????? Shape ??? ShapeRange ?? ?????? ????
Public Function MoveToLayer(ShapeOrRange As Object, Layer As Layer)
  
  Dim tSrcLayer() As Layer
  Dim tProps() As type_LayerProps
  Dim tLayersCol As Collection
  Dim i&
  
  If TypeOf ShapeOrRange Is Shape Then
  
    Set tLayersCol = New Collection
    tLayersCol.Add ShapeOrRange.Layer
    
  ElseIf TypeOf ShapeOrRange Is ShapeRange Then
    
    If ShapeOrRange.Count < 1 Then Exit Function
    Set tLayersCol = ShapeRangeLayers(ShapeOrRange)
    
  Else
  
    Err.Raise 13, Source:="MoveToLayer", Description:="Type mismatch: ShapeOrRange ?????? ???? Shape ??? ShapeRange"
    Exit Function
  
  End If
  
  ReDim tSrcLayer(1 To tLayersCol.Count)
  ReDim tProps(1 To tLayersCol.Count)
  For i = 1 To tLayersCol.Count
    Set tSrcLayer(i) = tLayersCol(i)
    layerPropsPreserveAndReset tSrcLayer(i), tProps(i)
  Next i
  ShapeOrRange.MoveToLayer Layer
  For i = 1 To tLayersCol.Count
    layerPropsRestore tSrcLayer(i), tProps(i)
  Next i

End Function

'????????? ???????? Shape ??? ShapeRange ?? ?????? ????
Public Function CopyToLayer(ShapeOrRange As Object, Layer As Layer) As Object

  If Not TypeOf ShapeOrRange Is Shape And Not TypeOf ShapeOrRange Is ShapeRange Then
    Err.Raise 13, Source:="CopyToLayer", Description:="Type mismatch: ShapeOrRange ?????? ???? Shape ??? ShapeRange"
    Exit Function
  End If
  
  Set CopyToLayer = ShapeOrRange.Duplicate
  MoveToLayer CopyToLayer, Layer

End Function

'??????????? ???????? ???????? ?? ????? ?????? ? ?????????
Public Function DuplicateActivePage(ByVal NumberOfPages&, Optional ExcludeLayerName$ = "") As Page
  Dim tRange As ShapeRange
  Dim tShape As Shape, sDuplicate As Shape
  Dim tProps As type_LayerProps
  Dim i&
  For i = 1 To NumberOfPages
    Set tRange = FindShapesActivePageLayers
    Set DuplicateActivePage = ActiveDocument.InsertPages(1, False, ActivePage.Index)
    DuplicateActivePage.SizeHeight = ActivePage.SizeHeight
    DuplicateActivePage.SizeWidth = ActivePage.SizeWidth
    For Each tShape In tRange.ReverseRange
      If tShape.Layer.Name <> ExcludeLayerName Then
        layerPropsPreserveAndReset tShape.Layer, tProps
        Set sDuplicate = tShape.Duplicate
        sDuplicate.MoveToLayer FindLayerDuplicate(DuplicateActivePage, tShape.Layer)
        layerPropsRestore tShape.Layer, tProps
      End If
    Next tShape
  Next i
End Function

'????????????? ?????? ? ?????? ??? ????? ? ????? ?????,
'? ??????????? ?? ????????? ?????
'??????????
Public Function ContrastShape(Shape As Shape) As Shape
  With Shape.Fill
    Select Case .Type
      Case cdrUniformFill
        .UniformColor.ConvertToGray
        If .UniformColor.Gray < 128 Then .UniformColor.GrayAssign 0 Else .UniformColor.GrayAssign 255
      Case cdrFountainFill
        'todo
    End Select
  End With
  With Shape.Outline
    If .Type <> cdrNoOutline Then
      .Color.ConvertToGray
      If .Color.Gray < 128 Then .Color.GrayAssign 0 Else .Color.GrayAssign 255
    End If
  End With
  Set ContrastShape = Shape
End Function

'???????? ?????? ?? CropEnvelopeShape, ?? ??-??????, ??????? ??????? ?? EXPANDBY ???????? ????????
Public Function TrimBitmap(BitmapShape As Shape, CropEnvelopeShape As Shape, Optional ByVal LeaveCropEnvelope As Boolean = True) As Shape

  Const EXPANDBY& = 2 'px
  
  Dim tCrop As Shape
  Dim tPxW#, tPxH#
  Dim tSaveUnit As cdrUnit

  If BitmapShape.Type <> cdrBitmapShape Then Exit Function
  
  'save
  tSaveUnit = ActiveDocument.Unit
  
  ActiveDocument.Unit = cdrInch
  tPxW = 1 / BitmapShape.Bitmap.ResolutionX
  tPxH = 1 / BitmapShape.Bitmap.ResolutionY
  BitmapShape.Bitmap.ResetCropEnvelope
  Set tCrop = BitmapShape.Layer.CreateRectangle(CropEnvelopeShape.LeftX - tPxW * EXPANDBY, _
                                                CropEnvelopeShape.TopY + tPxH * EXPANDBY, _
                                                CropEnvelopeShape.RightX + tPxW * EXPANDBY, _
                                                CropEnvelopeShape.BottomY - tPxH * EXPANDBY)
  Set TrimBitmap = Intersect(tCrop, BitmapShape, False, False)
  If TrimBitmap Is Nothing Then
    tCrop.Delete
    GoTo Finally
  End If
  TrimBitmap.Bitmap.Crop
  Set TrimBitmap = Intersect(CropEnvelopeShape, TrimBitmap, LeaveCropEnvelope, False)
  
Finally:
  'restore
  ActiveDocument.Unit = tSaveUnit
  
End Function

'?????????? ?????????
Public Function Intersect(SourceShape As Shape, _
                   TargetShape As Shape, _
                   Optional ByVal LeaveSource As Boolean = True, _
                   Optional ByVal LeaveTarget As Boolean = True _
                   ) As Shape
                   
  Dim tPropsSource As type_LayerProps
  Dim tPropsTarget As type_LayerProps
  
  If Not SourceShape.Layer Is TargetShape.Layer Then _
    layerPropsPreserveAndReset SourceShape.Layer, tPropsSource
  layerPropsPreserveAndReset TargetShape.Layer, tPropsTarget
  
  Set Intersect = SourceShape.Intersect(TargetShape)
  
  If Not SourceShape.Layer Is TargetShape.Layer Then _
    layerPropsRestore SourceShape.Layer, tPropsSource
  layerPropsRestore TargetShape.Layer, tPropsTarget
  
  If Intersect Is Nothing Then Exit Function
  
  Intersect.OrderFrontOf TargetShape
  If Not LeaveSource Then SourceShape.Delete
  If Not LeaveTarget Then TargetShape.Delete

End Function

'???????? ????? ?? Shape ?? ??????? Knife, ?????????? ?????????? ?????
Public Function Dissect(ByRef Shape As Shape, ByRef Knife As Shape) As Shape
  Set Dissect = Intersect(Knife, Shape, True, True)
  Set Shape = Knife.Trim(Shape, True, False)
End Function

'?????????? Crop Tool
Public Function CropTool(ShapeOrRangeOrPage As Object, ByVal x1#, ByVal y1#, ByVal x2#, ByVal y2#, Optional ByVal Angle = 0) As ShapeRange
  If TypeOf ShapeOrRangeOrPage Is Shape Or _
     TypeOf ShapeOrRangeOrPage Is ShapeRange Or _
     TypeOf ShapeOrRangeOrPage Is Page Then
    Set CropTool = ShapeOrRangeOrPage.CustomCommand("Crop", "CropRectArea", x1, y1, x2, y2, Angle)
  Else
    Err.Raise 13, Source:="CropTool", Description:="Type mismatch: ShapeOrRangeOrPage ?????? ???? Shape, ShapeRange ??? Page"
    Exit Function
  End If
End Function

'?????????? Boundary
Public Function CreateBoundary(ShapeOrRange As Object) As Shape
  On Error GoTo Catch
  Dim tShape As Shape, tRange As ShapeRange
  '?????? ?????? ?? ???, ???? ?????????? ???
  If TypeOf ShapeOrRange Is Shape Then
    Set tShape = ShapeOrRange
    Set CreateBoundary = tShape.CustomCommand("Boundary", "CreateBoundary")
  ElseIf TypeOf ShapeOrRange Is ShapeRange Then
    Set tRange = ShapeOrRange
    Set CreateBoundary = tRange.CustomCommand("Boundary", "CreateBoundary")
  Else
    Err.Raise 13, Source:="CreateBoundary", Description:="Type mismatch: ShapeOrRange ?????? ???? Shape ??? ShapeRange"
    Exit Function
  End If
  Exit Function
Catch:
  Debug.Print Err.Number
End Function

'?????????? Join Curves
Public Function JoinCurves(SrcRange As ShapeRange, ByVal Tolerance#)
  SrcRange.CustomCommand "ConvertTo", "JoinCurves", Tolerance
End Function

'???????? ????????
'?????: Alex Vakulenko http://www.oberonplace.com/vba/drawmacros/delsegment.htm
Public Sub SegmentDelete(Segment As Segment)
  If Not Segment.EndNode.IsEnding Then
    Segment.EndNode.BreakApart
    Set Segment = Segment.SubPath.LastSegment
  End If
  Segment.EndNode.Delete
End Sub

'?? ???????? ? ???????????
Public Sub MatrixCopy(SourceShape As Shape, TargetShape As Shape)
  Dim tMatrix As type_Matrix
  With tMatrix
    SourceShape.GetMatrix .d11, .d12, .d21, .d22, .tx, .ty
    TargetShape.SetMatrix .d11, .d12, .d21, .d22, .tx, .ty
  End With
End Sub

'????????? ???? ?????? ??????
Public Sub SetOutlineColor(ByVal Shapes As ShapeRange, ByVal Color As Color)
  Dim Shape As Shape
  For Each Shape In Shapes
    Shape.Outline.Color.CopyAssign Color
  Next Shape
End Sub

'===============================================================================
' ??????? ?????? ? ????????? ?????????? ?? ???????? ??????
'===============================================================================

'????????? ?? ?????? ??????????? ??????
'??? ??????? ??????? ?????????, ?. ?. ??? Nothing ?????? ????? ???? ??? ?? ????? ??????
Public Function IsNothing(Object As Object) As Boolean
  Dim t As Variant
  If Object Is Nothing Then GoTo ExitTrue
  If TypeOf Object Is Document Then
    On Error GoTo ExitTrue
    t = Object.Name
  ElseIf TypeOf Object Is Page Then
    On Error GoTo ExitTrue
    t = Object.Name
  ElseIf TypeOf Object Is Layer Then
    On Error GoTo ExitTrue
    t = Object.Name
  ElseIf TypeOf Object Is Shape Then
    On Error GoTo ExitTrue
    t = Object.Name
  ElseIf TypeOf Object Is Curve Then
    On Error GoTo ExitTrue
    t = Object.Length
  ElseIf TypeOf Object Is SubPath Then
    On Error GoTo ExitTrue
    t = Object.Closed
  ElseIf TypeOf Object Is Segment Then
    On Error GoTo ExitTrue
    t = Object.AbsoluteIndex
  ElseIf TypeOf Object Is Node Then
    On Error GoTo ExitTrue
    t = Object.AbsoluteIndex
  End If
  Exit Function
ExitTrue:
  IsNothing = True
End Function

'??????? ??? ????? ? ?????? ??????, ??????? ????? ? ???????????, ? ?????????
Public Function FindShapesByName(ShapeRange As ShapeRange, ByVal Name$) As ShapeRange
  Set FindShapesByName = FindAllShapes(ShapeRange).Shapes.FindShapes(Name)
End Function

'??????? ??? ?????, ????? ????? ??????? ????????? ? NamePart,
'??????? ????? ? ???????????, ? ?????????
Public Function FindShapesByNamePart(ShapeRange As ShapeRange, ByVal NamePart$) As ShapeRange
  Set FindShapesByNamePart = FindAllShapes(ShapeRange).Shapes.FindShapes(Query:="@Name.Contains('" & NamePart & "')")
End Function

'??????? ??????????, ??? ????????
Public Function FindPowerClips(ShapeRange As ShapeRange) As ShapeRange
  Set FindPowerClips = CreateShapeRange
  'On Error Resume Next
  'FindPowerClips.AddRange ShapeRange.Shapes.FindShapes(Query:="!@com.PowerClip.IsNull")
  Dim Shape As Shape
  For Each Shape In ShapeRange
    If Not lib_elvin.IsNothing(Shape) Then _
      If Not Shape.PowerClip Is Nothing Then FindPowerClips.Add Shape
  Next Shape
End Function

'??????? ?????????? ???????????, ??? ????????
Public Function FindShapesInPowerClips(ShapeRange As ShapeRange) As ShapeRange
  Dim tShape As Shape
  Set FindShapesInPowerClips = CreateShapeRange
  For Each tShape In FindPowerClips(ShapeRange)
    FindShapesInPowerClips.AddRange tShape.PowerClip.Shapes.All
  Next tShape
End Function

'??????? ??? ?????, ??????? ????? ? ???????????, ? ?????????
Public Function FindAllShapes(ShapeRange As ShapeRange) As ShapeRange
  Dim tShape As Shape
  Set FindAllShapes = CreateShapeRange
  FindAllShapes.AddRange ShapeRange.Shapes.FindShapes
  For Each tShape In FindPowerClips(ShapeRange)
    FindAllShapes.AddRange FindAllShapes(tShape.PowerClip.Shapes.All)
  Next tShape
End Function

'?????????? ??? ????? ?? ???? ????? ??????? ????????, ?? ????????? - ??? ??????-????? ? ??? ??????
Public Function FindShapesActivePageLayers(Optional GuidesLayers As Boolean = False, _
                                    Optional MasterLayers As Boolean = False _
                                    ) As ShapeRange
  Dim tLayer As Layer
  Set FindShapesActivePageLayers = CreateShapeRange
  For Each tLayer In ActivePage.Layers
    If Not (tLayer.IsGuidesLayer And (GuidesLayers = False)) Then _
      FindShapesActivePageLayers.AddRange tLayer.Shapes.All
  Next
  If MasterLayers Then
    For Each tLayer In ActiveDocument.MasterPage.Layers
      If Not (tLayer.IsGuidesLayer And (GuidesLayers = False)) Then _
        FindShapesActivePageLayers.AddRange tLayer.Shapes.All
  Next
  End If
End Function

'?????????? ????????? ????? ? ??????? ????????, ????? ??????? ???????? NamePart
Public Function FindLayersActivePageByNamePart(ByVal NamePart$, Optional ByVal SearchMasters = True) As Collection
  Dim tLayer As Layer
  Dim tLayers As Layers
  If SearchMasters Then Set tLayers = ActivePage.AllLayers Else Set tLayers = ActivePage.Layers
  Set FindLayersActivePageByNamePart = New Collection
  For Each tLayer In tLayers
    If InStr(tLayer.Name, NamePart) > 0 Then FindLayersActivePageByNamePart.Add tLayer
  Next
End Function

'????? ???????? ???? ?? ???? ?????????? (???????????, ??? ????? ?? ?????)
Public Function FindLayerDuplicate(PageToSearch As Page, SrcLayer As Layer) As Layer
  For Each FindLayerDuplicate In PageToSearch.AllLayers
    With FindLayerDuplicate
      If (.Name = SrcLayer.Name) And _
         (.IsDesktopLayer = SrcLayer.IsDesktopLayer) And _
         (.Master = SrcLayer.Master) And _
         (.Color.IsSame(SrcLayer.Color)) Then _
         Exit Function
    End With
  Next
  Set FindLayerDuplicate = Nothing
End Function

'?????????? ????????? ?????, ?? ??????? ????? ????? ?? ??????
Public Function ShapeRangeLayers(ShapeRange As ShapeRange) As Collection
  
  Dim tShape As Shape
  Dim tLayer As Layer
  Dim inCol As Boolean
  
  If ShapeRange.Count = 0 Then Exit Function
  Set ShapeRangeLayers = New Collection
  If ShapeRange.Count = 1 Then
    ShapeRangeLayers.Add ShapeRange(1).Layer
    Exit Function
  End If
  
  For Each tShape In ShapeRange
    inCol = False
    For Each tLayer In ShapeRangeLayers
      If tLayer Is tShape.Layer Then
        inCol = True
        Exit For
      End If
    Next tLayer
    If inCol = False Then ShapeRangeLayers.Add tShape.Layer
  Next tShape

End Function

'?????????? ??????? ??????? ?????/???????/????????
Public Function GreaterDim(ShapeOrRangeOrPage As Object) As Double
  If Not TypeOf ShapeOrRangeOrPage Is Shape And Not TypeOf ShapeOrRangeOrPage Is ShapeRange And Not TypeOf ShapeOrRangeOrPage Is Page Then
    Err.Raise 13, Source:="GreaterDim", Description:="Type mismatch: ShapeOrRangeOrPage ?????? ???? Shape, ShapeRange ??? Page"
    Exit Function
  End If
  If ShapeOrRangeOrPage.SizeWidth > ShapeOrRangeOrPage.SizeHeight Then GreaterDim = ShapeOrRangeOrPage.SizeWidth Else GreaterDim = ShapeOrRangeOrPage.SizeHeight
End Function

'?????????? ??????? ??????? ?????/???????/????????
Public Function LesserDim(ShapeOrRangeOrPage As Object) As Double
  If Not TypeOf ShapeOrRangeOrPage Is Shape And Not TypeOf ShapeOrRangeOrPage Is ShapeRange And Not TypeOf ShapeOrRangeOrPage Is Page Then
    Err.Raise 13, Source:="LesserDim", Description:="Type mismatch: ShapeOrRangeOrPage ?????? ???? Shape, ShapeRange ??? Page"
    Exit Function
  End If
  If ShapeOrRangeOrPage.SizeWidth < ShapeOrRangeOrPage.SizeHeight Then LesserDim = ShapeOrRangeOrPage.SizeWidth Else LesserDim = ShapeOrRangeOrPage.SizeHeight
End Function

'?????????? ??????? ?????? ?????/???????/????????
Public Function AverageDim(ShapeOrRangeOrPage As Object) As Double
  If Not TypeOf ShapeOrRangeOrPage Is Shape And Not TypeOf ShapeOrRangeOrPage Is ShapeRange And Not TypeOf ShapeOrRangeOrPage Is Page Then
    Err.Raise 13, Source:="AverageDim", Description:="Type mismatch: ShapeOrRangeOrPage ?????? ???? Shape, ShapeRange ??? Page"
    Exit Function
  End If
  AverageDim = (ShapeOrRangeOrPage.SizeWidth + ShapeOrRangeOrPage.SizeHeight) / 2
End Function

'?????????? Rect, ?????? ????????? ??????? ???? Space ?? ???? ??????
Public Function SpaceBox(ShapeOrRange As Object, Space#) As Rect
 If Not TypeOf ShapeOrRange Is Shape And Not TypeOf ShapeOrRange Is ShapeRange Then
    Err.Raise 13, Source:="SpaceBox", Description:="Type mismatch: ShapeOrRange ?????? ???? Shape ??? ShapeRange"
    Exit Function
  End If
  Set SpaceBox = ShapeOrRange.BoundingBox.GetCopy
  SpaceBox.Inflate Space, Space, Space, Space
End Function

'???????? ?? ????/??????/???????? ?????????
Public Function IsLandscape(ShapeOrRangeOrPage As Object) As Boolean
  If Not TypeOf ShapeOrRangeOrPage Is Shape And Not TypeOf ShapeOrRangeOrPage Is ShapeRange And Not TypeOf ShapeOrRangeOrPage Is Page Then
    Err.Raise 13, Source:="IsLandscape", Description:="Type mismatch: ShapeOrRangeOrPage ?????? ???? Shape, ShapeRange ??? Page"
    Exit Function
  End If
  If ShapeOrRangeOrPage.SizeWidth > ShapeOrRangeOrPage.SizeHeight Then IsLandscape = True Else IsLandscape = False
End Function

'???????? ?? ?????? ???????????, ???????????? ???? ??? ?????? ? ????? ????? (underlying dubs)
Public Function IsDuplicateCurves(Curve1 As Curve, Curve2 As Curve) As Boolean
  Dim tNode As Node
  Dim tJitter#: tJitter = ConvertUnits(0.001, cdrMillimeter, ActiveDocument.Unit) '?????? = 0.001 ??
  If Curve1.Nodes.Count <> Curve2.Nodes.Count Then Exit Function
  If Abs(Curve1.Length - Curve2.Length) > tJitter Then Exit Function
  For Each tNode In Curve1.Nodes
    If Curve2.FindNodeAtPoint(tNode.PositionX, tNode.PositionY, tJitter * 2) Is Nothing Then Exit Function
  Next
  IsDuplicateCurves = True
End Function

'????????? ??? ???????
Public Function IsOverlap(FirstShape As Shape, SecondShape As Shape) As Boolean
  
  Dim tIS As Shape
  Dim tShape1 As Shape, tShape2 As Shape
  Dim tBound1 As Shape, tBound2 As Shape
  Dim tProps As type_LayerProps
  
  If FirstShape.Type = cdrConnectorShape Or SecondShape.Type = cdrConnectorShape Then Exit Function
  
  '?????????? ????? ???? ??? ????????
  Dim tLayer As Layer: Set tLayer = ActiveLayer
  '?????????? ????????? ??????? ????
  FirstShape.Layer.Activate
  layerPropsPreserveAndReset FirstShape.Layer, tProps
  
  If isIntersectReady(FirstShape) Then
    Set tShape1 = FirstShape
  Else
    Set tShape1 = CreateBoundary(FirstShape)
    Set tBound1 = tShape1
  End If
  
  If isIntersectReady(SecondShape) Then
    Set tShape2 = SecondShape
  Else
    Set tShape2 = CreateBoundary(SecondShape)
    Set tBound2 = tShape2
  End If
  
  Set tIS = tShape1.Intersect(tShape2)
  If tIS Is Nothing Then
    IsOverlap = False
  Else
    tIS.Delete
    IsOverlap = True
  End If
  
  On Error Resume Next
    tBound1.Delete
    tBound2.Delete
  On Error GoTo 0
  
  '?????????? ??? ?? ?????
  layerPropsRestore FirstShape.Layer, tProps
  tLayer.Activate

End Function

'IsOverlap ????????? ???????? - ?????? ?? ?????????,
'?? ???? ????????? ???????? ? ? ??????????? ??????? ??? ??????????
Public Function IsOverlapBox(FirstShape As Shape, SecondShape As Shape) As Boolean
  Dim tShape As Shape
  Dim tProps As type_LayerProps
  '?????????? ????? ???? ??? ????????
  Dim tLayer As Layer: Set tLayer = ActiveLayer
  '?????????? ????????? ??????? ????
  FirstShape.Layer.Activate
  layerPropsPreserveAndReset FirstShape.Layer, tProps
  Dim tRect As Rect
  Set tRect = FirstShape.BoundingBox.Intersect(SecondShape.BoundingBox)
  If tRect.Width = 0 And tRect.Height = 0 Then
    IsOverlapBox = False
  Else
    IsOverlapBox = True
  End If
  '?????????? ??? ?? ?????
  layerPropsRestore FirstShape.Layer, tProps
  tLayer.Activate
End Function

Public Function IsSameShapes(ByVal Shape1 As Shape, Shape2 As Shape) As Boolean
  IsSameShapes = (Shape1.StaticID = Shape2.StaticID)
End Function

'===============================================================================
' ??????? ?????? ? ???????
'===============================================================================

'??????? ????????? ?????
Public Function GetTempFolder() As String
  GetTempFolder = Environ$("TEMP")
  If GetTempFolder = "" Then
    GetTempFolder = Environ$("TMP")
    If GetTempFolder = "" Then
      If Dir("c:\", vbDirectory) <> "" Then GetTempFolder = "c:\"
    End If
  End If
End Function

'?????? ??? ?????????? ?????
Public Function GetTempFile() As String
  GetTempFile = GetTempFolder & GetTempFileName
End Function

'??? ?????????? ?????
Public Function GetTempFileName() As String
  GetTempFileName = "elvin_" & CreateGUID & ".tmp"
End Function

'????????? ?????? Content ? ????, ?????????????, ????? ? ???????? temp ????,
'? ???????? ?????, ???? ??????????
Public Sub SaveStrToFile(ByRef Content$, ByVal File$, Optional ByVal KeepBak As Boolean = False)

  Dim tFileNum&: tFileNum = FreeFile
  Dim tBak$: tBak = SetFileExt(File, "bak")
  Dim tTemp$
  
  If KeepBak Then
    If FileExist(File) Then FileCopy File, tBak
  Else
    If FileExist(File) Then
      tTemp = GetFilePath(File) & GetTempFileName
      FileCopy File, tTemp
    End If
  End If
    
  Open File For Output Access Write As #tFileNum
  Print #tFileNum, Content
  Close #tFileNum
  
  On Error Resume Next
    If Not KeepBak Then Kill tTemp
  On Error GoTo 0

End Sub

'????????? ???? ? ??????
Public Function LoadStrFromFile(ByVal File$) As String
  Dim tFileNum&: tFileNum = FreeFile
  Open File For Input As #tFileNum
  LoadStrFromFile = Input(LOF(tFileNum), tFileNum)
  Close #tFileNum
End Function

'???????? ?????????? ????? ?? ????????
Public Function SetFileExt(ByVal SourceFile$, ByVal NewExt$) As String
  If Right(SourceFile, 1) <> "\" And Len(SourceFile) > 0 Then
    SetFileExt = GetFileNameNoExt(SourceFile$) & "." & NewExt
  End If
End Function

'?????????? ??? ????? ??? ??????????
Public Function GetFileNameNoExt(ByVal FileName$) As String
  If Right(FileName, 1) <> "\" And Len(FileName) > 0 Then
    GetFileNameNoExt = Left(FileName, _
      Switch _
        (InStr(FileName, ".") = 0, _
          Len(FileName), _
        InStr(FileName, ".") > 0, _
          InStrRev(FileName, ".") - 1))
  End If
End Function

'??????? ?????, ???? ?? ????
'?????????? Path ??????? (??? inline-?????????????)
Public Function MakeDir(ByVal Path$) As String
  If Dir(Path, vbDirectory) = "" Then MkDir Path
  MakeDir = Path
End Function

'?????????? ?? ???? ??? ????? (????? ?????? ????????????? ?? "\")
Public Function FileExist(ByVal File As String) As Boolean
  If File = "" Then Exit Function
  If Len(Dir(File)) > 0 Then
    FileExist = True
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFileName
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Return the filename from a path\filename input
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sFile - string of a path and filename (ie: "c:\temp\Test.xls")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2008-Feb-06                 Initial Release
'---------------------------------------------------------------------------------------
Public Function GetFileName(sFile As String)
On Error GoTo Err_Handler
 
    GetFileName = Right(sFile, Len(sFile) - InStrRev(sFile, "\"))
 
Exit_Err_Handler:
    Exit Function
 
Err_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: GetFileName" & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFilePath
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Return the path from a path\filename input
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sFile - string of a path and filename (ie: "c:\temp\Test.xls")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2008-Feb-06                 Initial Release
'---------------------------------------------------------------------------------------
Public Function GetFilePath(sFile As String)
On Error GoTo Err_Handler
 
    GetFilePath = Left(sFile, InStrRev(sFile, "\"))
 
Exit_Err_Handler:
    Exit Function
 
Err_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: GetFilePath" & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function

'===============================================================================
' ?????? ???????
'===============================================================================

Public Sub AssignUnknown(ByRef Destination As Variant, ByRef Value As Variant)
  If VBA.IsObject(Value) Then
    Set Destination = Value
  Else
    Destination = Value
  End If
End Sub

Public Sub CopyCollection(ByVal Source As Collection, _
                          ByVal Target As Collection)
  Dim Element As Variant
  For Each Element In Source
    Target.Add Element
  Next Element
End Sub

'https://www.codegrepper.com/code-examples/vb/excel+vba+generate+guid+uuid
Public Function CreateGUID(Optional Lowercase As Boolean, _
                           Optional Parens As Boolean _
                           ) As String
  Dim k As Long, h As String
  CreateGUID = VBA.Space(36)
  For k = 1 To VBA.Len(CreateGUID)
    VBA.Randomize
    Select Case k
      Case 9, 14, 19, 24: h = "-"
      Case 15:            h = "4"
      Case 20:            h = VBA.Hex(VBA.Rnd * 3 + 8)
      Case Else:          h = VBA.Hex(VBA.Rnd * 15)
    End Select
    Mid(CreateGUID, k, 1) = h
  Next
  If Lowercase Then CreateGUID = VBA.LCase$(CreateGUID)
  If Parens Then CreateGUID = "{" & CreateGUID & "}"
End Function

Public Function FindMaxValue(ByRef Collection As Collection) As Variant
  Dim Item As Variant
  For Each Item In CollectionOrArray
    If VBA.IsNumeric(Item) Then
      If Item > FindMaxValue Then FindMaxValue = Item
    End If
  Next Item
End Function

Public Function FindMinValue(ByRef Collection As Collection) As Variant
  Dim Item As Variant
  For Each Item In CollectionOrArray
    If VBA.IsNumeric(Item) Then
      If Item < FindMinValue Then
        Debug.Print "aaa"
        FindMinValue = Item
      End If
    End If
  Next Item
End Function

Public Function FindMaxItemNum(ByRef Collection As Collection) As Long
  FindMaxItemNum = 1
  Dim i As Long
  For i = 1 To Collection.Count
    If VBA.IsNumeric(Collection(i)) Then
      If Collection(i) > Collection(FindMaxItemNum) Then FindMaxItemNum = i
    End If
  Next i
End Function

Public Function FindMinItemNum(ByRef Collection As Collection) As Long
  FindMinItemNum = 1
  Dim i As Long
  For i = 1 To Collection.Count
    If VBA.IsNumeric(Collection(i)) Then
      If Collection(i) < Collection(FindMinItemNum) Then FindMinItemNum = i
    End If
  Next i
End Function

Public Function MinOfTwo(ByVal Value1 As Variant, _
                         ByVal Value2 As Variant _
                         ) As Variant
  If Value1 < Value2 Then MinOfTwo = Value1 Else MinOfTwo = Value2
End Function

Public Function MaxOfTwo(ByVal Value1 As Variant, _
                         ByVal Value2 As Variant _
                         ) As Variant
  If Value1 > Value2 Then MaxOfTwo = Value1 Else MaxOfTwo = Value2
End Function

Public Function IsSame(ByRef Value1 As Variant, ByRef Value2 As Variant) As Boolean
  If VBA.IsObject(Value1) And VBA.IsObject(Value2) Then
    IsSame = Value1 Is Value2
  ElseIf Not VBA.IsObject(Value1) And Not VBA.IsObject(Value2) Then
    IsSame = (Value1 = Value2)
  End If
End Function

'??????? ??????: https://stackoverflow.com/questions/38267950/check-if-a-value-is-in-an-array-or-not-with-excel-vba
Public Function IsStrInArr(ByVal stringToBeFound$, Arr As Variant) As Boolean
    Dim i&
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i) = stringToBeFound Then
            IsStrInArr = True
            Exit Function
        End If
    Next i
    IsStrInArr = False
End Function

'???????? ?? ????? ?????? :) ??? ????? Even ? Odd ?????????? ????...
Public Function IsChet(ByVal X) As Boolean
  If X Mod 2 = 0 Then IsChet = True Else IsChet = False
End Function

'??????? ?? Number ?? Divider ??????
Public Function IsDivider(ByVal Number&, ByVal Divider&) As Boolean
  If Number Mod Divider = 0 Then IsDivider = True Else IsDivider = False
End Function

Public Sub RemoveElementFromCollection(ByVal Element As Variant, _
                                       ByVal Collection As Collection)
  If Collection.Count = 0 Then Exit Sub
  Dim i As Long
  For i = 1 To Collection.Count
    If IsSame(Element, Collection(i)) Then
      Collection.Remove i
      Exit Sub
    End If
  Next i
End Sub

'????????? ????? ?? LowerBound  ?? UpperBound
Public Function RndInt(LowerBound As Long, UpperBound As Long) As Long
  RndInt = Int((UpperBound - LowerBound + 1) * Rnd + LowerBound)
End Function

Public Function MeasureStart()
  startTime = Timer
End Function
Public Function MeasureFinish(Optional Message$ = "")
  Debug.Print Message & CStr(Round(Timer - startTime, 3)) & " ??????"
End Function

'===============================================================================
' # ????????? ??????? ??????
'===============================================================================

Private Sub layerPropsPreserve(L As Layer, ByRef Props As type_LayerProps)
  With Props
    .Visible = L.Visible
    .Printable = L.Printable
    .Editable = L.Editable
  End With
End Sub
Private Sub layerPropsReset(L As Layer)
  With L
    If Not .Visible Then .Visible = True
    If Not .Printable Then .Printable = True
    If Not .Editable Then .Editable = True
  End With
End Sub
Private Sub layerPropsRestore(L As Layer, ByRef Props As type_LayerProps)
  With Props
    If L.Visible <> .Visible Then L.Visible = .Visible
    If L.Printable <> .Printable Then L.Printable = .Printable
    If L.Editable <> .Editable Then L.Editable = .Editable
  End With
End Sub
Private Sub layerPropsPreserveAndReset(L As Layer, ByRef Props As type_LayerProps)
  layerPropsPreserve L, Props
  layerPropsReset L
End Sub

'??? IsOverlap
Private Function isIntersectReady(Shape As Shape) As Boolean
  With Shape
    If .Type = cdrCustomShape Or _
       .Type = cdrBlendGroupShape Or _
       .Type = cdrOLEObjectShape Or _
       .Type = cdrExtrudeGroupShape Or _
       .Type = cdrContourGroupShape Or _
       .Type = cdrBevelGroupShape Or _
       .Type = cdrConnectorShape Or _
       .Type = cdrMeshFillShape Or _
       .Type = cdrTextShape Then
      isIntersectReady = False
    Else
      isIntersectReady = True
    End If
  End With
End Function

Private Sub ThrowIfNotCollectionOrArray(ByRef CollectionOrArray As Variant)
  If VBA.IsObject(CollectionOrArray) Then _
    If TypeOf CollectionOrArray Is Collection Then Exit Sub
  If VBA.IsArray(CollectionOrArray) Then Exit Sub
  VBA.Err.Raise 13, Source:="lib_elvin", _
                Description:="Type mismatch: CollectionOrArray ?????? ???? Collection ??? Array"
End Sub
