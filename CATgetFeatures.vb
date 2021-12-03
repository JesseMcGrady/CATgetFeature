Imports Excel = Microsoft.Office.Interop.Excel
Module CATgetFeatures
    Private xExcelapp As New Excel.Application
    Private CATIAFactory As CATIA_Property = New CATIA_Property
    Private oPart As MECMOD.Part
    Private oPartDoc As MECMOD.PartDocument
    Private FilesPath As String
    Private Files_Catia() As String
    Private Act As Excel.Worksheet
    Private Wb As Excel.Workbook
    Sub Main()
        'Dim oCat As System.Type = System.Type.GetTypeFromProgID("Catia.Application")
        'Dim CATIA As Object = System.Activator.CreateInstance(oCat)
        'Console.WriteLine(CATIA.Caption)
        Call ReadFile()
        Dim fileName As String = System.AppDomain.CurrentDomain.BaseDirectory + "Result.xlsx"
        Wb = xExcelapp.Workbooks.Add
        Act = Wb.Sheets(1)
        CATIAFactory = CATIA_Property.SetInitialCATIA_batch
        Files_Catia = IO.Directory.GetFiles(FilesPath, "*.CATPart")
        Call GetMultifileFeatures()
        Console.WriteLine("Finish")
        'Call CompareParts()
        Act.UsedRange.EntireColumn.AutoFit()
        Act.SaveAs(fileName)
        Wb.Close()
        xExcelapp.Quit()
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xExcelapp)
        Act = Nothing
        Wb = Nothing
        xExcelapp = Nothing
        CATIAFactory.MyCATIA.Quit()
        CATIAFactory = Nothing
    End Sub
    Private Sub GetMultifileFeatures()
        Dim i, j, m As Integer
        Dim PartBody As MECMOD.Body
        Dim PartName, Fname As String
        Dim Fnames() As String
        Dim Feautures_Part() As String
        Dim oSelection As INFITF.Selection
        Dim SearchCount, SurFlangeCount, FlangeCount, CirStampCount As Integer
        'Dim sel As INFITF.Selection
        Act.Cells(1, 1) = "Category"
        Act.Cells(1, 2) = "CATIA File Name"
        Act.Cells(1, 3) = "Material"
        Act.Cells(1, 4) = "Flange"
        Act.Cells(1, 5) = "Surfacic Flange"
        Act.Cells(1, 6) = "Circular Stamp"
        ReDim Feautures_Part(0)
        For i = 0 To UBound(Files_Catia)
            oPartDoc = CATIAFactory.MyCATIA.Documents.Open(Files_Catia(i))
            PartName = oPartDoc.Name
            Console.WriteLine(PartName)
            Act.Cells(i + 2, 2) = PartName
            'Console.WriteLine(oPartDoc.Name)
            oPartDoc.Activate()
            oPart = oPartDoc.Part
            oSelection = oPartDoc.Selection
            oSelection.Search("Name=Surfacic Flange* & (((FreeStyle.'PartDesign Feature' + 'Part Design'.'PartDesign Feature') + 'Generative Shape Design'.'PartDesign Feature')),all")
            Console.WriteLine("Surfacic Flange Count : ")
            SurFlangeCount = Int(oSelection.Count)
            Console.WriteLine(SurFlangeCount)
            oSelection.Clear()
            oSelection.Search("Name=Flange* & (((FreeStyle.'PartDesign Feature' + 'Part Design'.'PartDesign Feature') + 'Generative Shape Design'.'PartDesign Feature')),all")
            Console.WriteLine("Flange Count : ")
            FlangeCount = Int(oSelection.Count)
            Console.WriteLine(FlangeCount)
            oSelection.Clear()
            oSelection.Search("Name=Circular Stamp* & (((FreeStyle.'PartDesign Feature' + 'Part Design'.'PartDesign Feature') + 'Generative Shape Design'.'PartDesign Feature')),all")
            Console.WriteLine("Circular Stamp Count : ")
            CirStampCount = Int(oSelection.Count)
            Console.WriteLine(CirStampCount)
            oSelection.Clear()
            oSelection.Search("Name=Joggle* & (((FreeStyle.'PartDesign Feature' + 'Part Design'.'PartDesign Feature') + 'Generative Shape Design'.'PartDesign Feature')),all")
            'Console.WriteLine("Joggle Count : ")
            SearchCount = Int(oSelection.Count)
            oSelection.Clear()
            'Console.WriteLine(SearchCount)
            oSelection = Nothing
            'Console.WriteLine(oPart.Bodies.Count)
            PartBody = oPart.MainBody
            Act.Cells(i + 2, 3) = GetMaterial(oPart)
            For j = 1 To PartBody.Shapes.Count
                'Console.WriteLine(PartBody.Shapes.Item(j).Name)
                Fname = PartBody.Shapes.Item(j).Name
                ReDim Preserve Feautures_Part(UBound(Feautures_Part) + 1)
                Feautures_Part(UBound(Feautures_Part)) = Fname
                If Left(Fname, 14) = "Circular Stamp" Then
                    Act.Cells(i + 2, 6) = "V"
                ElseIf Left(Fname, 15) = "Surfacic Flange" Then
                    Act.Cells(i + 2, 5) = "V"
                ElseIf Left(Fname, 6) = "Flange" Then
                    Act.Cells(i + 2, 4) = "V"
                Else

                    For m = 4 To Act.UsedRange.Columns.Count

                        Fnames = Split(Fname, ".")
                        If Act.Cells(1, m).Value = Fnames(0) Then
                            Act.Cells(i + 2, m) = "V"
                            Exit For
                        End If
                    Next
                    If m = Act.UsedRange.Columns.Count + 1 Then
                        Act.Cells(1, m) = Fnames(0)
                        Act.Cells(i + 2, m) = "V"
                    End If
                End If

                '增加Joggle特徵判別

                'Call ExportToExcel(i + 2, j, PartName, Fname)
            Next
            If SurFlangeCount > 0 Then
                If CheckSheetMetal(Feautures_Part, SearchCount) = True And SearchCount <> 0 Then
                    Act.Cells(i + 2, 1) = "SH"
                Else
                    Act.Cells(i + 2, 1) = "SB"
                End If
            ElseIf CirStampCount > 0 Then
                Act.Cells(i + 2, 1) = "SH"
            Else
                If FlangeCount > 0 Then
                    Act.Cells(i + 2, 1) = "SB"
                ElseIf CirStampCount = 0 And FlangeCount = 0 Then
                    Act.Cells(i + 2, 1) = "SM"
                End If
            End If

            oPartDoc.Close()
        Next
    End Sub
    Private Function CheckSheetMetal(ByRef Feature_Array() As String, ByRef JoggleCount As Integer) As Boolean
        Dim FirstlayerCount As Integer
        FirstlayerCount = 0
        For i = 1 To UBound(Feature_Array)
            If Left(Feature_Array(i), 6) = "Joggle" Then
                FirstlayerCount = FirstlayerCount + 1
            End If
        Next
        Console.WriteLine(FirstlayerCount)
        If FirstlayerCount = JoggleCount Then
            CheckSheetMetal = False
        Else
            CheckSheetMetal = True
        End If
    End Function
    Private Sub ExportToExcel(Rows As Integer, Columns As Integer, DocName As String, FeaturesName As String)

        Act.Cells(Rows, Columns + 3) = FeaturesName

    End Sub
    Private Sub CompareParts()      '在Excel分辨鈑金件種類
        Dim Row, Column As Integer
        Dim CheckSM, CheckSB, CheckSH, CheckSF As Boolean
        For Row = 2 To Act.UsedRange.Rows.Count
            CheckSB = False
            CheckSM = False
            CheckSH = False
            CheckSF = False
            For j = 4 To 6
                If Act.Cells(Row, j).Value = "V" Then
                    CheckSH = True
                    Act.Cells(Row, 1) = "SH"
                ElseIf Act.Cells(Row, j).Value = "V" Then
                    CheckSF = True
                    'Act.Cells(Row, 1) = "Please Check Surfacic Features!!"
                ElseIf Act.Cells(Row, j).Value = "V" Then
                    CheckSB = True
                Else
                    CheckSM = True
                End If
            Next
            If CheckSM = True And CheckSH = False And CheckSB = False And CheckSF = False Then
                Act.Cells(Row, 1) = "SM"
            ElseIf CheckSM = True And CheckSH = False And CheckSB = True And CheckSF = False Then
                Act.Cells(Row, 1) = "SB"
            End If
        Next

    End Sub
    Private Sub ReadFile()              '讀取批次檔內設定的CATIA設計檔路徑
        Dim fileReader As System.IO.StreamReader
        fileReader = My.Computer.FileSystem.OpenTextFileReader(System.AppDomain.CurrentDomain.BaseDirectory & "\ActiveCATIA.bat")
        Dim stringReader, path As String
        stringReader = fileReader.ReadLine()
        path = Replace(stringReader, "rem ", "")
        Console.WriteLine("File path is : " & path)
        FilesPath = path
    End Sub
    Private Function GetMaterial(P As MECMOD.Part) As String
        Dim oManger As Object = P.GetItem("CATMatManagerVBExt")
        Dim mBody As MECMOD.Body = P.MainBody
        Dim MaterialinP
        Call oManger.GetMaterialOnBody(mBody, MaterialinP)
        If MaterialinP IsNot Nothing Then
            GetMaterial = MaterialinP.Name
        Else
            GetMaterial = "None"
        End If
    End Function
End Module
