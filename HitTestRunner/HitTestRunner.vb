Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.EditorExt
Imports System.Windows.Forms

Public Class HitTestRunner
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

#Region "properties"

    Private _pMxDoc As IMxDocument
    Public Property pMxDoc() As IMxDocument
        Get
            Return _pMxDoc
        End Get
        Set(ByVal value As IMxDocument)
            _pMxDoc = value
        End Set
    End Property


    Private _pMap As IMap
    Public Property pMap() As IMap
        Get
            Return _pMap
        End Get
        Set(ByVal value As IMap)
            _pMap = value
        End Set
    End Property


    Private _hittestRunner As HitTestForm
    Public Property UserFormHitLayers() As HitTestForm
        Get
            Return _hittestRunner
        End Get
        Set(ByVal value As HitTestForm)
            _hittestRunner = value
        End Set
    End Property

    Private _networkLayer As String
    Public Property NetworkLayer() As String
        Get
            Return _networkLayer
        End Get
        Set(ByVal value As String)
            _networkLayer = value
        End Set
    End Property

    Private _point1Layer As String
    Public Property Point1Layer() As String
        Get
            Return _point1Layer
        End Get
        Set(ByVal value As String)
            _point1Layer = value
        End Set
    End Property

    Private _searchRadius As Double
    Public Property SearchRadius() As Double
        Get
            Return _searchRadius
        End Get
        Set(ByVal value As Double)
            _searchRadius = value
        End Set
    End Property


#End Region

    Public Sub New()

      

    End Sub

    Protected Overrides Sub OnClick()

        pMxDoc = My.ArcMap.Application.Document
        pMap = pMxDoc.FocusMap
        UserFormHitLayers = New HitTestForm

        Dim pLayer As IFeatureLayer
        Dim pField As IField
        Dim pGeometryDef As IGeometryDef
        Dim C As Long

        For C = 0 To pMap.LayerCount - 1

            If TypeOf pMap.Layer(C) Is IFeatureLayer Then
                pLayer = pMap.Layer(C)
                If UserFormHitLayers.ListBox1.Text = "" Then
                    If pLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                        pField = pLayer.FeatureClass.Fields.Field _
                          (pLayer.FeatureClass.FindField(pLayer.FeatureClass.ShapeFieldName))
                        pGeometryDef = pField.GeometryDef
                        If pGeometryDef.HasM Then
                            UserFormHitLayers.ListBox1.Items.Add(pLayer.Name)
                        End If
                    End If
                End If
                If UserFormHitLayers.ListBox2.Text = "" Then
                    If pLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint Then
                        UserFormHitLayers.ListBox2.Items.Add(pLayer.Name)
                    End If
                End If

            End If

        Next C

        UserFormHitLayers.ShowDialog()

        If UserFormHitLayers.DialogResult = Windows.Forms.DialogResult.OK Then
            If Not UserFormHitLayers.TextBox1.Text = vbNullString Then
                MsgBox("DEBUG you entered " + UserFormHitLayers.TextBox1.Text)

                NetworkLayer = UserFormHitLayers.ListBox1.Text
                Point1Layer = UserFormHitLayers.ListBox2.Text
                Double.TryParse(UserFormHitLayers.TextBox1.Text, SearchRadius)
                RunHitTest(NetworkLayer, Point1Layer, SearchRadius)
            End If

        End If
        'RunAnalysis()
        'Why did I use a delegate here??
        ' Dim GoButton As System.Windows.Forms.Button = UserFormHitLayers.Button1
        ' AddHandler GoButton.Click, AddressOf RunAnalysis





        My.ArcMap.Application.CurrentTool = Nothing
    End Sub

    'dELETE THIS SUBROUTINE??
    Sub RunAnalysis()
        RunHitTest(NetworkLayer, Point1Layer, SearchRadius)
    End Sub



    Protected Overrides Sub OnUpdate()
        Enabled = My.ArcMap.Application IsNot Nothing
    End Sub

    Sub RunHitTest(strLineName As String, strPtName As String, dblSearchRadius As Double)

        Dim pApp As IApplication
        pMxDoc = My.ArcMap.Application.Document
        pMap = pMxDoc.FocusMap
        Dim pFLayer As IFeatureLayer
        Dim pFlayerPoly As IFeatureLayer
        Dim pSBar As IStatusBar
        Dim pStepProg As IStepProgressor
        Dim i As Long
        Dim intResponse As Integer

        For i = 0 To pMxDoc.FocusMap.LayerCount - 1
            If pMxDoc.FocusMap.Layer(i).Name = strLineName Then
                pFlayerPoly = pMxDoc.FocusMap.Layer(i)
            End If
            If pMxDoc.FocusMap.Layer(i).Name = strPtName Then
                pFLayer = pMxDoc.FocusMap.Layer(i)
                If pFLayer.FeatureClass.Fields.FindField("measure") = -1 Then
                    intResponse = MsgBox(pFLayer.Name & " has no measure field." & vbNewLine & "Do you want to add it now?", vbYesNo)
                    If intResponse = vbYes Then
                        Dim pFields As IFields
                        Dim pFieldsEdit As IFieldsEdit
                        Dim pFieldEdit As IFieldEdit2
                        Dim pField As IField
                        pField = New Field
                        pFieldEdit = pField
                        With pFieldEdit
                            .Name_2 = "measure"
                            .Type_2 = esriFieldType.esriFieldTypeDouble
                        End With
                        pFLayer.FeatureClass.AddField(pField)
                    End If
                    If intResponse = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
        Next i

        MsgBox("Point Layer is " & pFLayer.Name & vbNewLine & "Line Layer is " & pFlayerPoly.Name)
        Dim pFSel As IFeatureSelection
        pFSel = pFLayer

        Dim pSelectionSet As ISelectionSet
        pSelectionSet = pFSel.SelectionSet

        If pSelectionSet.Count = 0 Then
            pFSel.SelectFeatures(Nothing, esriSelectionResultEnum.esriSelectionResultNew, False)
            pFSel.SelectionChanged()
            pSelectionSet = pFSel.SelectionSet
        End If
        MsgBox(pSelectionSet.Count)

        Dim pFeature As IFeature
        Dim pFeaturePoly As IFeature
        Dim pFCursor As IFeatureCursor
        Dim pFeatureCursorPolyLine As IFeatureCursor
        Dim pLQueryFilter As IQueryFilter
        Dim sLlidval As String
        Dim sFieldName As String
        sFieldName = "LLID"

        Dim pHitTest As IHitTest
        Dim pGeomHitPartType As esriGeometryHitPartType
        pGeomHitPartType = esriGeometryHitPartType.esriGeometryPartBoundary
        Dim phitpnt As IPoint
        Dim pnt As IPoint

        Dim pFgeom As IGeometry
        Dim pNEWFeature As IFeature
        Dim pSpatialReference As ISpatialReference
        pSpatialReference = pMap.SpatialReference
        'the following 4 variables are not needed for my results, but are required for the processing
        Dim lPart As Long
        Dim lSeg As Long
        Dim bHitRt As Boolean
        Dim dHitDist As Double

        Dim pPolyline As IPolyline
        Dim pFeatGeom As IGeometry

        Dim pCurve As ICurve, dAlong As Double, pJunk As IPoint
        Dim pMSeg As IMSegmentation
        Dim pTopoOp As ITopologicalOperator
        Dim vMeas 'As Variant,
        Dim dFrom As Double, bRight As Boolean

        ' Setup and display the progress bar
        'pSBar = StatusBar
        'pStepProg = pSBar.ProgressBar
        'With pStepProg
        '    .MinRange = 0
        '    .MaxRange = pSelectionSet.Count
        '    .StepValue = 1
        '    .Position = 0
        '    .Show()
        'End With

        Dim vWenLeng As Object
        pSelectionSet.Search(Nothing, False, pFCursor)
        pFeature = pFCursor.NextFeature
        Do While Not pFeature Is Nothing
            'get the LLID value of the point
            sLlidval = pFeature.Value(pFeature.Fields.FindField(sFieldName))
            If sLlidval <> "" Then
                pLQueryFilter = New QueryFilter
                pLQueryFilter.WhereClause = "LLID = '" & sLlidval & "'"

                'select the line whose LLID matches the point
                pFeatureCursorPolyLine = pFlayerPoly.FeatureClass.Search(pLQueryFilter, False)
                pFeaturePoly = pFeatureCursorPolyLine.NextFeature
                If Not pFeaturePoly Is Nothing Then
                    pPolyline = New Polyline
                    pPolyline = pFeaturePoly.Shape
                    pnt = New Point
                    pnt = pFeature.Shape
                    pFgeom = pFeature.Shape
                    phitpnt = New Point
                    pnt.SpatialReference = pSpatialReference
                    pHitTest = pPolyline 'the hitTest is set to the Polyline we want to "snap" to
                    If pHitTest Is Nothing Then
                        MsgBox("Oops!")
                        On Error GoTo errorhandler
                    End If


                    Debug.Print("pntX = " + pnt.X.ToString + " and pntY =" + pnt.Y.ToString + "pntZ = " + pnt.Z.ToString + " where = " + pLQueryFilter.WhereClause.ToString)


                    'FROM SDK 'changed geometryCollection_Polyline to pPOlyline and polyLine.FromPoint to pnt
                    Dim polyline As ESRI.ArcGIS.Geometry.IPolyline = CType(pPolyline, ESRI.ArcGIS.Geometry.IPolyline)
                    Dim queryPoint As ESRI.ArcGIS.Geometry.IPoint = pnt

                    Dim hitPoint As ESRI.ArcGIS.Geometry.IPoint = New ESRI.ArcGIS.Geometry.PointClass

                    'Define and initialize the variables that will get populated from the .HitTest() method.
                    Dim hitDistance As System.Double = 0
                    Dim hitPartIndex As System.Int32 = 0
                    Dim hitSegmentIndex As System.Int32 = 0
                    Dim rightSide As System.Boolean = False

                    Dim hitTest As ESRI.ArcGIS.Geometry.IHitTest = CType(pPolyline, ESRI.ArcGIS.Geometry.IHitTest)
                    Dim foundGeometry As System.Boolean = hitTest.HitTest(queryPoint, SearchRadius, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartVertex, hitPoint, hitDistance, hitPartIndex, hitSegmentIndex, rightSide)
                    'END FROM SDK
                    queryPoint.PutCoords(hitPoint.X, hitPoint.Y)


                    Dim hitCheck As Boolean

                    hitCheck = pHitTest.HitTest(pnt, dblSearchRadius, pGeomHitPartType, phitpnt, dHitDist, lPart, lSeg, bHitRt)
                    'If pHitTest.HitTest(pnt, dblSearchRadius, pGeomHitPartType, phitpnt, dHitDist, lPart, lSeg, bHitRt) Then
                    pnt.PutCoords(phitpnt.X, phitpnt.Y)
                    phitpnt = pFeature.Shape
                    pnt.Project(phitpnt.SpatialReference)
                    pFeature.Value(pFeature.Fields.FindField("Shape")) = pnt
                    pTopoOp = pFeaturePoly.ShapeCopy
                    pMSeg = pTopoOp
                    pJunk = New Point
                    pCurve = pTopoOp
                    pCurve.QueryPointAndDistance(esriSegmentExtension.esriNoExtension, pFeature.Shape, True, pJunk, dAlong, dFrom, bRight)
                    vMeas = pMSeg.GetMsAtDistance(dAlong, True)
                    pFeature.Value(pFeature.Fields.FindField("measure")) = vMeas(0)
                    '***Add code here to adjust for true river mile
                    pFeature.Store()

                    ' Else
                    'You have a chance to make an exception report of some kind...
                    'Debug.Print ("TOO FAR!")
                    ' End If
                End If
            End If
            '     pStepProg.Step()
            pFeature = pFCursor.NextFeature

        Loop
        ' Clean up the progress bar
        'pStepProg.Hide()
        'pStepProg = Nothing
        'pSBar = Nothing
        MsgBox("Done!")
        Exit Sub
errorhandler:
        MsgBox("ERROR")
        Resume Next
    End Sub


End Class
