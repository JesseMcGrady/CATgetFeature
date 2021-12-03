Option Explicit On
Imports INFITF
Imports MECMOD
Imports PARTITF
Imports KnowledgewareTypeLib
Imports HybridShapeTypeLib
Imports SPATypeLib
Imports ProductStructureTypeLib
Imports KinTypeLib
Imports System
Imports System.IO

Public Class CATIA_Property
    Private _Documents As INFITF.Documents
    Property Documents As INFITF.Documents
        Get
            Return _Documents
        End Get
        Set(value As INFITF.Documents)
            If value Is Nothing Then
                Throw New ArgumentNullException("Documents", "Document can not be nothing")
            End If
            _Documents = value
        End Set
    End Property
    Private _ProductDocuments As INFITF.Document
    Property ProductDocument As INFITF.Document
        Get
            Return _ProductDocuments
        End Get
        Set(value As INFITF.Document)
            If value Is Nothing Then
                Throw New ArgumentNullException("Documents", "Document can not be nothing")
            End If
            _ProductDocuments = value
        End Set
    End Property
    Private _ShapeFactory As ShapeFactory
    Property ShapeFactory As ShapeFactory
        Get
            Return _ShapeFactory
        End Get
        Set(value As ShapeFactory)
            If value Is Nothing Then
                Throw New ArgumentNullException("ShapeFactory", "ShapeFactory can not be nothing")
            End If
            _ShapeFactory = value
        End Set
    End Property
    Private _HybridFactory As HybridShapeFactory
    Property HybridFactory As HybridShapeFactory
        Get
            Return _HybridFactory
        End Get
        Set(value As HybridShapeFactory)
            If value Is Nothing Then
                Throw New ArgumentNullException("HybridShapeFactory", "HybridShapeFactory can not be nothing")
            End If
            _HybridFactory = value
        End Set
    End Property
    Private _InstanceFactory As InstanceFactory
    Property InstanceFactory As InstanceFactory
        Get
            Return _InstanceFactory
        End Get
        Set(value As InstanceFactory)
            If value Is Nothing Then
                Throw New ArgumentNullException("InstanceFactory", "InstanceFactory can not be nothing")
            End If
            _InstanceFactory = value
        End Set
    End Property
    Private _CATIA As INFITF.Application
    Property MyCATIA As INFITF.Application
        Get
            Return _CATIA
        End Get
        Set(value As INFITF.Application)
            If value Is Nothing Then
                Throw New ArgumentNullException("CATIA", "CATIA can not be nothing")
            End If
            _CATIA = value
        End Set
    End Property
    Private _PartDocument As MECMOD.PartDocument
    Property PartDocument As MECMOD.PartDocument
        Get
            Return _PartDocument
        End Get
        Set(value As MECMOD.PartDocument)
            If value Is Nothing Then
                Throw New ArgumentNullException("PartDocuments", "PartDocuments can not be nothing")
            End If
            _PartDocument = value
        End Set
    End Property
    Private _Selection As INFITF.Selection
    Property Selection As INFITF.Selection
        Get
            Return _Selection
        End Get
        Set(value As INFITF.Selection)
            If value Is Nothing Then
                Throw New ArgumentNullException("Selection", "Selection can not be nothing")
            End If
            _Selection = value
        End Set
    End Property
    Public Shared Function SetInitialCATIA()
        'Dim XCATIA As CATIA_Property = New CATIA_Property
        Dim XCATIA As New CATIA_Property
        Try
            XCATIA.MyCATIA = GetObject("", "CATIA.Application")
        Catch
            XCATIA.MyCATIA = CreateObject("CATIA.Application")
        End Try

        XCATIA.MyCATIA.Visible = True
        XCATIA.MyCATIA.DisplayFileAlerts = True
        Try
            'With XCATIA
            '    Try
            '        .PartDocument = XCATIA.MyCATIA.ActiveDocument
            '        .Selection = .PartDocument.Selection
            '    Catch ex As Exception
            '        .ProductDocument = XCATIA.MyCATIA.ActiveDocument
            '        .Selection = .ProductDocument.Selection
            '    End Try
            '    .Documents = XCATIA.MyCATIA.Documents
            'End With
            Return XCATIA
        Catch ex As Exception
            Console.WriteLine("Didn't open CATIA File")
        End Try

        'myCATIA.DisplayFileAlerts = False
        Return Nothing
    End Function
    Public Shared Function SetInitialCATIA_batch()
        'Dim XCATIA As CATIA_Property = New CATIA_Property
        Dim XCATIA As New CATIA_Property
        Try
            XCATIA.MyCATIA = System.Activator.CreateInstance(System.Type.GetTypeFromProgID("Catia.Application"))
        Catch
            System.Diagnostics.Process.Start(System.AppDomain.CurrentDomain.BaseDirectory & "\ActiveCATIA.bat", vbNormalNoFocus)
            XCATIA.MyCATIA = System.Activator.CreateInstance(System.Type.GetTypeFromProgID("Catia.Application"))
        End Try

        XCATIA.MyCATIA.Visible = False
        XCATIA.MyCATIA.DisplayFileAlerts = False
        Try
            With XCATIA
                '.Documents = XCATIA.MyCATIA.Documents
            End With
            Return XCATIA
        Catch ex As Exception
            Console.WriteLine("Didn't open CATIA File")
        End Try
        Return Nothing
    End Function
    Public Shared Function C_Select(ByRef SelectionFrom As Selection, ByRef SelectType As Array, ByVal Comment As String) As String                 'Single Selection in CATIA
        Try
            Dim mStatus = SelectionFrom.SelectElement2(SelectType, Comment, False)        'Selection 單選指令
            Return mStatus
        Catch ex As Exception
            Console.WriteLine("Didn't Selected!!")
        End Try
    End Function
    Public Shared Function C_SelectMuti(ByRef SelectionFrom As Selection, ByRef SelectType As Array, ByVal Comment As String) As String                 'Muti- Selection in CATIA
        Try
            Dim mStatus = SelectionFrom.SelectElement3(SelectType, Comment, False, CATMultiSelectionMode.CATMultiSelTriggWhenUserValidatesSelection, False)        'Selection 複選指令
            Return mStatus
        Catch ex As Exception
            Console.WriteLine("Didn't Selected!!")
        End Try
    End Function
    Public Shared Function C_SelectdefaultAxis(ByRef PartDoc As PartDocument) As Selection                      'Select default Axis in Part
        Dim mPart As Part = PartDoc.Part
        Dim mSelection As Selection = PartDoc.Selection
        Dim sFilter(0)
        Dim Sel_Property As VisPropertySet
        Dim FirstResult As SelectedElement
        mSelection.Search("CatPrtSearch.AxisSystem,All")
        If mSelection.Count = 0 Then
            mSelection.Add(C_CreateAxis(PartDoc))
        Else
            mSelection.Item2(1)
            FirstResult = mSelection.Item(1)
            mSelection.Clear()
            mSelection.Add(FirstResult.Value)        'select element. value equal to object
        End If
        Return mSelection
    End Function
    Public Shared Function C_CreateAxis(ByRef PartDoc As PartDocument) As AxisSystem            'Create Axis in Part
        Dim mPart As Part = PartDoc.Part
        Dim mAxis As AxisSystems = mPart.AxisSystems
        Dim Axis1 As AxisSystem = mAxis.Add()

        Axis1.OriginType = CATAxisSystemOriginType.catAxisSystemOriginByCoordinates
        Dim ArrayOrigin(2)
        ArrayOrigin(0) = 0.0#
        ArrayOrigin(1) = 0.0#
        ArrayOrigin(2) = 0.0#
        Axis1.PutOrigin(ArrayOrigin)

        Axis1.XAxisType = CATAxisSystemAxisType.catAxisSystemAxisByCoordinates
        Dim ArrayXaxis(2)
        ArrayXaxis(0) = 1.0#
        ArrayXaxis(1) = 0.0#
        ArrayXaxis(2) = 0.0#
        Axis1.PutOrigin(ArrayXaxis)

        Axis1.YAxisType = CATAxisSystemAxisType.catAxisSystemAxisByCoordinates
        Dim ArrayYaxis(2)
        ArrayYaxis(0) = 0.0#
        ArrayYaxis(1) = 1.0#
        ArrayYaxis(2) = 0.0#
        Axis1.PutOrigin(ArrayYaxis)

        Axis1.ZAxisType = CATAxisSystemAxisType.catAxisSystemAxisByCoordinates
        Dim ArrayZaxis(2)
        ArrayZaxis(0) = 0.0#
        ArrayZaxis(1) = 0.0#
        ArrayZaxis(2) = 1.0#
        Axis1.PutOrigin(ArrayZaxis)

        mPart.UpdateObject(Axis1)
        Axis1.IsCurrent = True
        mPart.Update()
        Return Axis1
    End Function
    Public Shared Function ExceptSelectedItemHide(ByRef ProductDoc As ProductDocument, ByVal ItemCount As Integer)                          'Only the one which is selected show
        Dim PartCount As Integer = ProductDoc.Product.Products.Count
        Dim i As Integer
        Dim Component As Product
        For i = 1 To PartCount
            Component = ProductDoc.Product.Products.Item(i)
            If i = ItemCount Then
                Call ShowObjectFromProduct(ProductDoc, Component.Name)
            Else
                Call HideObjectFromProduct(ProductDoc, Component.Name)
            End If
        Next
    End Function
    Public Shared Function GetISO_ViewPoint(ByRef mProduct As Product, ByRef CATIA As INFITF.Application, ByVal FileLocation As String) As String       '輸出ISO視角圖片
        Dim Space As SpecsAndGeomWindow = CATIA.ActiveWindow
        Dim ActWin As Window = CATIA.ActiveWindow
        Dim ActView As Viewer3D = ActWin.ActiveViewer
        'CATIA.StartCommand("Compass")
        Space.Layout = CatSpecsAndGeomWindowLayout.catWindowGeomOnly
        ActView.FullScreen = True
        ActView.Reframe()
        ActView.Viewpoint3D = CATIA.ActiveDocument.Cameras.Item(1)
        ActView.ZoomIn()
        Dim Color(2)
        ActView.GetBackgroundColor(Color)
        Dim BlackArray(2)
        BlackArray(0) = 1
        BlackArray(1) = 1
        BlackArray(2) = 1
        ActView.PutBackgroundColor(BlackArray)
        Dim imageFilePath As String = FileLocation + "\" + mProduct.Name + ".JPG"
        ActView.CaptureToFile(CatCaptureFormat.catCaptureFormatJPEG, imageFilePath)
        ActView.PutBackgroundColor(Color)

        Space.Layout = CatSpecsAndGeomWindowLayout.catWindowSpecsAndGeom
        ActView.FullScreen = False
        'CATIA.StartCommand("Compass")
        Return imageFilePath
    End Function
    Public Shared Sub HideObject(ByRef PartDoc As PartDocument, ByRef GeometricalSetName As Object)                        '隱藏Part物件
        Dim visPropertySet1 As VisPropertySet
        Dim mSelection As Selection = PartDoc.Selection
        mSelection.Clear()
        PartDoc.Selection.Search("Name:*" & GeometricalSetName & "*,All")
        visPropertySet1 = mSelection.VisProperties
        visPropertySet1.SetShow(1)
        mSelection.Clear()
    End Sub
    Public Shared Sub HideObjectFromProduct(ByRef ProDoc As ProductStructureTypeLib.ProductDocument, ByRef GeometricalSetName As Object)                        '隱藏Product物件
        Dim visPropertySet1 As VisPropertySet
        Dim mSelection As Selection = ProDoc.Selection
        mSelection.Clear()
        ProDoc.Selection.Search("Name:*" & GeometricalSetName & "*,All")
        visPropertySet1 = mSelection.VisProperties
        visPropertySet1.SetShow(1)
        mSelection.Clear()
    End Sub
    Public Shared Sub ShowObjectFromProduct(ByRef ProDoc As ProductStructureTypeLib.ProductDocument, ByRef GeometricalSetName As Object)                        '顯示Product物件
        Dim visPropertySet1 As VisPropertySet
        Dim mSelection As Selection = ProDoc.Selection
        mSelection.Clear()
        ProDoc.Selection.Search("Name:*" & GeometricalSetName & "*,All")
        visPropertySet1 = mSelection.VisProperties
        visPropertySet1.SetShow(0)
        mSelection.Clear()
    End Sub
    Public Shared Function CreateRefFromObj(ByRef mPart As Part, ByRef mObject As Object) As Reference                     '抓取Reference
        Dim oObject As Reference
        Try
            oObject = mPart.CreateReferenceFromObject(mObject)
            Return oObject
        Catch ex As Exception
            MsgBox("Can't Get Reference!" & vbCrLf & ex.Message)
            'Throw New SystemException("Can't Get Reference!")
            Return Nothing
        Finally

        End Try
    End Function

    Public Shared Function C_HybridShapeExtremum(ByRef mHybridShapeFactory As HybridShapeFactory, hybridShapeDX As HybridShapeDirection, Reference1 As Reference, ByVal D_MinMax As Integer)              '建立極端點
        Dim HybridShapeExtremum1 As HybridShapeExtremum = mHybridShapeFactory.AddNewExtremum(Reference1, hybridShapeDX, D_MinMax)
        Return HybridShapeExtremum1
    End Function

    Public Shared Function C_HybridShapeLinePtDir(mHybridShapeFactory As HybridShapeFactory, Originpoint As HybridShapePointCoord, hybridShapeDX As HybridShapeDirection)          '建立線(點和方向)
        Dim Plane_line_1 As HybridShapeLinePtDir = mHybridShapeFactory.AddNewLinePtDir(Originpoint, hybridShapeDX, 0, 0, False)
        Return Plane_line_1
    End Function
    Public Shared Function OpenToNewWindow(ByRef CatiaFactory As CATIA_Property, ByRef mProduct As Product) '打開子件檔案
        Dim StiEngine As CATSmarTeamInteg.StiEngine               'Smart Team Libery
        StiEngine = CatiaFactory.MyCATIA.GetItem("CAIEngine")
        Dim StiDBItem As CATSmarTeamInteg.StiDBItem = StiEngine.GetStiDBItemFromAnyObject(mProduct.ReferenceProduct.Parent)
        Dim FileFullName As String = StiDBItem.GetDocumentFullPath
        CatiaFactory.Documents.Open(FileFullName)

    End Function
End Class