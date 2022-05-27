Attribute VB_Name = "Pepakurator"
'===============================================================================
' Макрос           : Pepakurator
' Версия           : 2021.11.16
' Автор            : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Const RELEASE As Boolean = True

'===============================================================================

Public Const PlaceholderText As String = "xxx"
Public Const OffsetToAvoidMult As Double = 1.5
Public Const LongDashArrowsSymbolName As String = "LongDashArrows"
Public Const TwoDashArrowsSymbolName As String = "TwoDashArrows"

Private Const EdgeIDGroupName As String = "EdgeID"
Private Const LinesGroupName As String = "Lines"

'===============================================================================

Sub AddArrows()

  If RELEASE Then On Error GoTo Catch
  
  If ActiveDocument Is Nothing Then
    VBA.MsgBox "Нет активного документа"
    Exit Sub
  End If
  
  ActiveDocument.Unit = cdrMillimeter
  
  lib_elvin.BoostStart "Добавление стрелочек", RELEASE
    
  Common.CreateLongDashArrows(ActiveLayer) _
    .ConvertToSymbol(LongDashArrowsSymbolName).Delete
  Common.CreateTwoDashArrows(ActiveLayer) _
    .ConvertToSymbol(TwoDashArrowsSymbolName).Delete
  
  Dim Line As Shape
  With DashLinesParser.Create(ActivePage)
    For Each Line In .LongDashLines
      Common.SetArrowsSymbolOnLine _
        ActiveLayer, LongDashArrowsSymbolName, Line, .ShapesToAvoid
    Next Line
    For Each Line In .TwoDashLines
      Common.SetArrowsSymbolOnLine _
        ActiveLayer, TwoDashArrowsSymbolName, Line, .ShapesToAvoid
    Next Line
  End With
  
Finally:
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Ошибка"
  Resume Finally

End Sub

Sub Prepare()
  
  If RELEASE Then On Error GoTo Catch
  
  If ActiveDocument Is Nothing Then
    VBA.MsgBox "Нет активного документа"
    Exit Sub
  End If
  
  Dim EdgeIDGroup As Shape, LinesGroup As Shape
  Set EdgeIDGroup = ActivePage.FindShape(EdgeIDGroupName)
  Set LinesGroup = ActivePage.FindShape(LinesGroupName)
  
  If EdgeIDGroup Is Nothing And LinesGroup Is Nothing Then
    VBA.MsgBox "Страница уже подготовлена"
    Exit Sub
  End If
  
  lib_elvin.BoostStart "Подготовка SVG из Pepakura", RELEASE
  
  Common.tryUnlockAndUngroup EdgeIDGroup
  Common.tryUnlockAndUngroup LinesGroup
  Common.UnlockAndDeleteSVGData
   
  Dim TextShapes As Collection
  Set TextShapes = Common.CollectTextShapes(ActivePage.Shapes.All)
  
  Common.FindPairsAndSetIDs TextShapes
  Common.SetTextShapesToPlaceholder TextShapes
  
Finally:
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Ошибка"
  Resume Finally

End Sub

Sub SetNumber()

  If RELEASE Then On Error GoTo Catch
  
  If ActiveShape Is Nothing Then Exit Sub
  
  Dim Page As ParsedPage
  Set Page = ParsedPage.Create(ActivePage)
  
  Dim Pair As Pair
  Set Pair = Page.GetPairs.GetPairByShape(ActiveShape)
  If Pair Is Nothing Then Exit Sub
  
  Dim Number As Long
  With New NumberView
    .Show
    If .IsOK Then Number = .Number
  End With
  If Number = 0 Then Exit Sub
  
  lib_elvin.BoostStart "Присвоение номера", RELEASE
  
  Dim Selected As Shape
  Set Selected = ActiveShape
    
  Page.GetPairs.PushForward Number
  Pair.Number = Number
  
Finally:
  Selected.CreateSelection
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Ошибка"
  Resume Finally

End Sub

Sub SetNextNumber()

  If RELEASE Then On Error GoTo Catch
  
  If ActiveShape Is Nothing Then Exit Sub

  Dim Page As ParsedPage
  Set Page = ParsedPage.Create(ActivePage)
  
  Dim Pair As Pair
  Set Pair = Page.GetPairs.GetPairByShape(ActiveShape)
  If Pair Is Nothing Then Exit Sub
  If Not Pair.Number = 0 Then Exit Sub
  
  lib_elvin.BoostStart "Присвоение следующего номера", RELEASE
  
  Dim Selected As Shape
  Set Selected = ActiveShape
  
  Pair.Number = Page.GetPairs.FreeNumber
  
Finally:
  Selected.CreateSelection
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Ошибка"
  Resume Finally

End Sub

Sub RemoveNumbers()

  If RELEASE Then On Error GoTo Catch

  Dim Page As ParsedPage
  Set Page = ParsedPage.Create(ActivePage)
  
  Dim SelectedPairs As Pairs
  Set SelectedPairs = Page.GetPairsFromShapes(ActiveSelectionRange)
  If SelectedPairs.Count = 0 Then Exit Sub
  
  lib_elvin.BoostStart "Удаление номеров", RELEASE
  
  Dim Selected As ShapeRange
  Set Selected = ActiveSelectionRange
  
  Common.RemoveNumbers SelectedPairs, Page
  
Finally:
  Selected.CreateSelection
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Ошибка"
  Resume Finally

End Sub

'===============================================================================
' тесты
'===============================================================================

Private Sub testWhatID()
  MsgBox ActiveShape.StaticID
End Sub

Private Sub testFindPairByID()
  Common.FindPair(ActiveShape, ActivePage.Shapes.All).CreateSelection
End Sub

Private Sub testCheckPairs()
  BoostStart "testCheckPairs", False
  Dim Page As ParsedPage
  Set Page = ParsedPage.Create(ActivePage)
  Dim tempPair As Pair
  For Each tempPair In Page.Pairs
    tempPair.Number = 333
  Next tempPair
  BoostFinish
End Sub

Private Sub testPageParsing()
  Dim Page As ParsedPage
  Set Page = ParsedPage.Create(ActivePage)
End Sub

Private Sub testPairsAreSame()
  Dim Page As ParsedPage
  Set Page = ParsedPage.Create(ActivePage)
  With Page.GetPairsFromShapes(ActivePage.Shapes.All)
    Debug.Print .Item(1).IsSame(.Item(1))
  End With
End Sub

Private Sub testPairs()
  Dim Page As ParsedPage
  Set Page = ParsedPage.Create(ActivePage)
  With Page.GetPairsFromShapes(ActivePage.Shapes.All)
    'Debug.Print .Item(2).IsSame(.Item(2))
    Debug.Print .FirstNumber
    Debug.Print .LastNumber
    Debug.Print .FreeNumber
    'Debug.Print .Range(4, .FreeNumber).Count
  End With
End Sub

Private Sub testPairsInAction()
  Dim Page As ParsedPage
  Set Page = ParsedPage.Create(ActivePage)
  With Page.GetPairsFromShapes(ActivePage.Shapes.All)
    BoostStart "testPairs", False
    BoostFinish
  End With
End Sub

Private Sub testOutlineStyleDashCount()
  Debug.Print ActiveShape.Outline.Style.DashCount
End Sub

Private Sub testCreateLongDashArrows()
  Common.CreateLongDashArrows ActiveLayer
End Sub

Private Sub testCreateTwoDashArrows()
  Common.CreateTwoDashArrows ActiveLayer
End Sub

Private Sub testDashLinesParser()
  With DashLinesParser.Create(ActivePage)
    Debug.Print .LongDashLines.Count
    Debug.Print .TwoDashLines.Count
    Debug.Print .ShapesToAvoid.Count
  End With
End Sub

Private Sub testDashLinesParser2()
  With DashLinesParser.Create(ActivePage)
    .ShapesToAvoid.ApplyUniformFill CreateColor("CMYK,USER,100,0,0,0")
  End With
End Sub

Private Sub testShapeOverlapsOneOfRange()
  Debug.Print Common.IsShapeOverlapsOneOfRange _
                (ActiveSelectionRange.FirstShape, _
                 ActiveSelectionRange.LastShape.Shapes.All)
End Sub
