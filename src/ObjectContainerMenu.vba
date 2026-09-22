' MacroRunner integration: no reference to the runner project is required.
Private pMRObserver As Object
Private pMRToken As String

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdProcess_Click()

    Dim doc As Document
    Dim selectedShapes As shapeRange
    Dim shp As Shape
    Dim markerShapes As Collection
    Dim stage4Shapes As Collection
    Dim stage5Group As Shape
    Dim marker As Shape
    Dim zoneIndex As Long
    Dim previousUnit As cdrUnit
    Dim commandGroupOpen As Boolean

    On Error GoTo ProcessFailed

    Set doc = ActiveDocument
    previousUnit = doc.Unit
    doc.Unit = cdrMillimeter

    Set selectedShapes = ActiveSelectionRange
    If selectedShapes Is Nothing Or selectedShapes.Count = 0 Then
        MsgBox "Pilih objek terlebih dahulu.", vbExclamation, "Object Container"
        GoTo CleanExit
    End If

    doc.BeginCommandGroup "Object Container"
    commandGroupOpen = True

    SetOutlinePropertiesForRange selectedShapes

    If optDesign.Value Then
        selectedShapes.UngroupAll
        ActiveDocument.ClearSelection
    Else
        selectedShapes.BreakApart
    End If

    Set markerShapes = FindShapesBySize(299#, 399#, 301#, 401#)
    If markerShapes.Count = 0 Then
        MsgBox "Selection zone 299-301 x 399-401 mm tidak ditemukan.", vbExclamation, "Object Container"
        GoTo CleanExit
    End If

    For zoneIndex = 1 To markerShapes.Count
        Set stage4Shapes = FindStage4Shapes(markerShapes(zoneIndex))
        Set stage5Group = GroupZoneContents(markerShapes(zoneIndex), stage4Shapes)
        If stage4Shapes.Count = 2 Then Set stage5Group = GroupCollection(stage4Shapes)
    Next zoneIndex

    For zoneIndex = 1 To markerShapes.Count
        GroupAllZoneContents markerShapes(zoneIndex)
    Next zoneIndex

    For Each marker In markerShapes
        marker.Delete
    Next marker

CleanExit:
    On Error Resume Next
    If commandGroupOpen Then doc.EndCommandGroup
    doc.Unit = previousUnit
    On Error GoTo 0
    Unload Me
    Exit Sub

ProcessFailed:

    Dim errorNumber As Long
    Dim errorDescription As String

    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next
    If commandGroupOpen Then doc.EndCommandGroup
    If Not doc Is Nothing Then doc.Unit = previousUnit
    On Error GoTo 0

    MsgBox "Error " & errorNumber & vbCrLf & _
           "Description: [" & errorDescription & "]", _
           vbCritical, "Object Container"

    Unload Me

End Sub

Private Sub SetOutlinePropertiesForRange(ByVal shapesToProcess As ShapeRange)

    Dim shp As Shape

    If optDesign.Value Then

        For Each shp In shapesToProcess
            shp.Outline.SetNoOutline
        Next shp

    ElseIf optCutLine.Value Then

        ' Outline Justification sengaja didahulukan sebelum menerapkan properti lain pada outline objek, terutama sebelum pewarnaan.
        ' Style.StringAssign dapat mengubah warna outline menjadi RGB #000000.
        ' Karena itu, SetOutlineProperties dijalankan setelahnya untuk memulihkan warna yang ditetapkan.
        For Each shp In shapesToProcess
            shp.Style.StringAssign "{""outline"":{""justification"":""0""}}"
        Next shp

        shapesToProcess.SetOutlineProperties _
            Color:=CreateCMYKColor(0, 100, 100, 0), _
            MiterLimit:=5#, _
            LineCaps:=cdrOutlineLineCapSquare, _
            LineJoin:=cdrOutlineJoinMiter

    End If

End Sub

Private Function FindShapesBySize(ByVal minWidth As Double, ByVal minHeight As Double, ByVal maxWidth As Double, ByVal maxHeight As Double) As Collection
    Dim result As New Collection
    Dim pageShapes() As Shape
    Dim shp As Shape
    Dim shapeIndex As Long

    SnapshotPageShapes pageShapes

    For shapeIndex = 1 To PageShapeSnapshotCount(pageShapes)
        Set shp = pageShapes(shapeIndex)
        If IsWithinSize(shp, minWidth, minHeight, maxWidth, maxHeight) Then result.Add shp
    Next shapeIndex

    Set FindShapesBySize = result
End Function

Private Function FindStage4Shapes(ByVal marker As Shape) As Collection
    Dim result As New Collection
    Dim pageShapes() As Shape
    Dim shp As Shape
    Dim shapeIndex As Long
    Dim horizontalFound As Boolean
    Dim verticalFound As Boolean

    SnapshotPageShapes pageShapes

    For shapeIndex = 1 To PageShapeSnapshotCount(pageShapes)
        Set shp = pageShapes(shapeIndex)
        If Not horizontalFound And IsInsideMarker(shp, marker) And _
           IsWithinSize(shp, 249#, 2#, 251#, 4#) Then
            result.Add shp
            horizontalFound = True
        ElseIf Not verticalFound And IsInsideMarker(shp, marker) And _
               IsWithinSize(shp, 0#, 249#, 2#, 251#) Then
            result.Add shp
            verticalFound = True
        End If
    Next shapeIndex

    Set FindStage4Shapes = result
End Function

Private Function IsWithinSize(ByVal shp As Shape, ByVal minWidth As Double, ByVal minHeight As Double, _
                              ByVal maxWidth As Double, ByVal maxHeight As Double) As Boolean
    IsWithinSize = (shp.SizeWidth >= minWidth And shp.SizeWidth <= maxWidth And _
                    shp.SizeHeight >= minHeight And shp.SizeHeight <= maxHeight)
End Function

Private Function GroupCollection(ByVal shapesToGroup As Collection) As Shape
    Dim shapeRange As New shapeRange
    Dim item As Variant

    For Each item In shapesToGroup
        shapeRange.Add item
    Next item

    If shapeRange.Count > 0 Then Set GroupCollection = shapeRange.Group
End Function

Private Function GroupZoneContents(ByVal marker As Shape, ByVal excludedShapes As Collection) As Shape
    Dim shapeRange As New shapeRange
    Dim pageShapes() As Shape
    Dim shp As Shape
    Dim shapeIndex As Long

    SnapshotPageShapes pageShapes

    For shapeIndex = 1 To PageShapeSnapshotCount(pageShapes)
        Set shp = pageShapes(shapeIndex)
        If Not (shp Is marker) Then
            If IsInsideMarker(shp, marker) And Not IsInCollection(shp, excludedShapes) Then
                shapeRange.Add shp
            End If
        End If
    Next shapeIndex

    If shapeRange.Count > 0 Then Set GroupZoneContents = shapeRange.Group
End Function

Private Function GroupAllZoneContents(ByVal marker As Shape) As Shape
    Dim shapeRange As New shapeRange
    Dim pageShapes() As Shape
    Dim shp As Shape
    Dim shapeIndex As Long

    SnapshotPageShapes pageShapes

    For shapeIndex = 1 To PageShapeSnapshotCount(pageShapes)
        Set shp = pageShapes(shapeIndex)
        If Not (shp Is marker) Then
            If IsInsideMarker(shp, marker) Then shapeRange.Add shp
        End If
    Next shapeIndex

    If shapeRange.Count > 0 Then Set GroupAllZoneContents = shapeRange.Group
End Function

Private Sub SnapshotPageShapes(ByRef pageShapes() As Shape)
    Dim shp As Shape
    Dim shapeIndex As Long

    If ActivePage.Shapes.Count = 0 Then Exit Sub

    ReDim pageShapes(1 To ActivePage.Shapes.Count)
    For Each shp In ActivePage.Shapes
        shapeIndex = shapeIndex + 1
        Set pageShapes(shapeIndex) = shp
    Next shp
End Sub

Private Function PageShapeSnapshotCount(ByRef pageShapes() As Shape) As Long
    On Error GoTo EmptySnapshot
    PageShapeSnapshotCount = UBound(pageShapes)
    Exit Function

EmptySnapshot:
    PageShapeSnapshotCount = 0
End Function

Private Function IsInsideMarker(ByVal shp As Shape, ByVal marker As Shape) As Boolean
    IsInsideMarker = (shp.LeftX >= marker.LeftX And shp.RightX <= marker.RightX And _
                      shp.BottomY >= marker.BottomY And shp.TopY <= marker.TopY)
End Function

Private Function IsInCollection(ByVal target As Shape, ByVal shapesToCheck As Collection) As Boolean
    Dim item As Variant

    For Each item In shapesToCheck
        If item Is target Then
            IsInCollection = True
            Exit Function
        End If
    Next item
End Function

Private Sub optCutLine_Click()
    If optCutLine.Value Then optDesign.Value = False
End Sub

Private Sub optDesign_Click()
    If optDesign.Value Then optCutLine.Value = False
End Sub

Private Sub UserForm_Initialize()
    If Not optDesign.Value And Not optCutLine.Value Then
        optCutLine.Value = True
    End If
End Sub

' Called only by MRTargetBridge; normal menu entry points remain unchanged.
Public Sub MRBindRunner(ByVal observer As Object, ByVal token As String)
    Set pMRObserver = observer
    pMRToken = token
End Sub

Public Sub MRDetachRunner()
    Set pMRObserver = Nothing
    pMRToken = vbNullString
End Sub

Private Sub UserForm_Terminate()
    Dim observer As Object, token As String
    On Error GoTo NotifyFailed
    Set observer = pMRObserver
    token = pMRToken
    MRDetachRunner
    If Not observer Is Nothing Then CallByName observer, "MacroUnloaded", VbMethod, token
    Exit Sub
NotifyFailed:
    MsgBox "Gagal memberitahu Macro Runner bahwa form sudah ditutup (" & CStr(Err.Number) & "): " & _
        Err.Description, vbExclamation, "Macro Runner"
End Sub
