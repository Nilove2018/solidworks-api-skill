Attribute VB_Name = "CreateFlange"
' CreateFlange.bas - early-bound VBA macro: a 25 x 24 x 2.5 mm flange with four M2 holes.
'
' How to use (a .swp file is binary, so it cannot be shipped as text):
'   1. SOLIDWORKS > Tools > Macro > New... > save as CreateFlange.swp
'   2. In the VBA editor: File > Import File... > CreateFlange.bas, delete the empty Module1
'   3. Tools > References: SOLIDWORKS <version> Type Library and Constant type library are ticked
'   4. Run: F5 in the editor, or Tools > Macro > Run > CreateFlange.swp
'
' All API lengths are METRES. Constants come from the swconst reference - never raw integers.
Option Explicit

Const MM As Double = 0.001

Dim swApp As SldWorks.SldWorks

Sub main()
    Set swApp = Application.SldWorks

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.NewDocument(swApp.GetUserPreferenceStringValue(swDefaultTemplatePart), 0, 0, 0)
    If swModel Is Nothing Then
        MsgBox "No default part template - check Tools > Options > Default Templates"
        Exit Sub
    End If
    swModel.Extension.SetUserPreferenceInteger swUnitSystem, 0, swUnitSystem_MMGS

    Dim swSketchMgr As SldWorks.SketchManager
    Set swSketchMgr = swModel.SketchManager
    swSketchMgr.AddToDB = True                         ' exact geometry, no snapping

    ' --- sketch on the Front Plane (the first plane in the tree) ---
    Dim frontPlane As SldWorks.Feature
    Set frontPlane = FirstRefPlane(swModel)
    swModel.ClearSelection2 True
    frontPlane.Select2 False, 0
    swSketchMgr.InsertSketch True

    swSketchMgr.CreateCenterRectangle 0, 0, 0, 12.5 * MM, 12 * MM, 0     ' 25 x 24 outline
    Dim sx As Variant, sy As Variant
    For Each sx In Array(-10.5, 10.5)
        For Each sy In Array(-6.25, 6.25)
            swSketchMgr.CreateCircleByRadius sx * MM, sy * MM, 0, 1.2 * MM ' D2.4 holes, 21 x 12.5 pattern
        Next
    Next
    swSketchMgr.InsertSketch True                      ' close the sketch
    swSketchMgr.AddToDB = False

    Dim swSketch As SldWorks.Feature
    Set swSketch = swModel.FeatureByPositionReverse(0)
    swSketch.Name = "Flange Sketch"

    ' --- extrude 2.5 mm (FeatureExtrusion3, 23 arguments) ---
    swModel.ClearSelection2 True
    swSketch.Select2 False, 0
    Dim swFeat As SldWorks.Feature
    Set swFeat = swModel.FeatureManager.FeatureExtrusion3(True, False, False, swEndCondBlind, 0, _
        2.5 * MM, 0, False, False, False, False, 0, 0, False, False, False, False, True, True, True, _
        swStartSketchPlane, 0, False)
    If swFeat Is Nothing Then
        MsgBox "Extrude failed"
        Exit Sub
    End If
    swFeat.Name = "Flange"

    swModel.ClearSelection2 True
    swModel.ShowNamedView2 "*Isometric", 7
    swModel.ViewZoomtofit2
End Sub

' The first RefPlane feature is the Front Plane in every language edition of SOLIDWORKS.
Private Function FirstRefPlane(swModel As SldWorks.ModelDoc2) As SldWorks.Feature
    Dim f As SldWorks.Feature
    Set f = swModel.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "RefPlane" Then
            Set FirstRefPlane = f
            Exit Function
        End If
        Set f = f.GetNextFeature
    Loop
End Function
