'Dimentions for create and modify Regal 3d model
Dim swApp As Object
Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Private Sub cmdCalculate_Click()
Dim Dim_mm As Double
Dim Dim_m As Double
    'Regal Box Capacity Calculation
    If opt4Fach.Value = True Then
        Dim_mm = txbxRegLvl.Value * txbxRegFld.Value * 16
        lblRegCap.Caption = Dim_mm & " Boxs"
    End If

    If opt3Fach.Value = True Then
        Dim_mm = txbxRegLvl.Value * txbxRegFld.Value * 12
        lblRegCap.Caption = Dim_mm & " Boxs"
    End If

    'Regal Width Calculation
    If opt4Fach.Value = True Then
        Dim_mm = (txbxBoxWid.Value * 8) + txbxShuWid.Value + 186
        Dim_m = Dim_mm / 1000
        lblRegWid.Caption = Dim_mm & " mm [" & Dim_m & " m]"
    End If
    
    If opt3Fach.Value = True Then
        Dim_mm = (txbxBoxWid.Value * 6) + txbxShuWid.Value + 146
        Dim_m = Dim_mm / 1000
        lblRegWid.Caption = Dim_mm & " mm [" & Dim_m & " m]"
    End If

    'Regal Length Calculation
        Dim_mm = (((txbxBoxLen.Value * 2) + 142) * txbxRegFld.Value) + 1060
        Dim_m = Dim_mm / 1000
        lblRegLen.Caption = Dim_mm & " mm [" & Dim_m & " m]"
    
    'Regal Height Calculation
        Dim_mm = txbxFirLvl.Value + ((txbxBoxHig.Value + 140) * txbxRegLvl.Value) + 160
        Dim_m = Dim_mm / 1000
        lblRegHig.Caption = Dim_mm & " mm [" & Dim_m & " m]"
End Sub

Private Sub cmdCloseDoc_Click()
Set swApp = Application.SldWorks
boolstatus = swApp.CloseAllDocuments(True)
Debug.Print "All documents, including dirty documents, closed: " & boolstatus
cmdSave.Enabled = False
cmdCreate.Caption = "Create Regal"
cmdCreate.Enabled = True
End Sub

Private Sub cmdCreate_Click()

'Connect Solidworks Active Model
Set swApp = Application.SldWorks
Set Part = swApp.ActiveDoc

Dim myDimension As Object

' Open Regal 3D model
'Set Part = swApp.OpenDoc6("C:\Users\MohammadrezaFathi\Documents\Arbeiten\Macro\RSO Regal\RSO_Regal.SLDASM", 2, 0, "", longstatus, longwarnings)
Set Part = swApp.OpenDoc6(Left(swApp.GetCurrentMacroPathName, InStrRev(swApp.GetCurrentMacroPathName, "\")) & "RSO_Regal.SLDASM", 2, 0, "", longstatus, longwarnings)
Set Part = swApp.OpenDoc6("RSO_Regal.SLDASM", 2, 0, "", longstatus, longwarnings)
Dim swAssembly As Object
Set swAssembly = Part
swApp.ActivateDoc2 "RSO_Regal.SLDASM", False, longstatus
Set Part = swApp.ActiveDoc

'Hide/Unhid the FeatureManager design tree
    Dim swModel As SldWorks.ModelDoc2
    Dim swModDocExt As SldWorks.ModelDocExtension
    Dim bRet As Boolean

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swModDocExt = swModel.Extension
    bRet = swModDocExt.HideFeatureManager(True)

'Change to single viewports and set view to clear screen
Part.ShowNamedView2 "Clear Screen View", -1
Part.ModelViewManager.ViewportDisplay = swViewportDisplay_e.swViewportSingle

'Change Button Caption
cmdCreate.Caption = "Apply Change"

'Modify Regal first Levels
boolstatus = Part.Extension.SelectByID2("D1@First_Level_Plane@RSO_Regal.SLDASM", "DIMENSION", 0#, 0#, 0#, True, 0, Nothing, 0)
Set myDimension = Part.Parameter("D1@First_Level_Plane@RSO_Regal.SLDASM")
myDimension.SystemValue = txbxFirLvl.Value / 1000

'Modify the Number of Regal Levels
If txbxRegLvl.Value > 1 Then
    boolstatus = Part.Extension.SelectByID2("D3@Height_LPattern@RSO_Regal.SLDASM", "DIMENSION", 0#, 0#, 0#, True, 0, Nothing, 0)
    Set myDimension = Part.Parameter("D3@Height_LPattern@RSO_Regal.SLDASM")
    myDimension.SystemValue = (txbxBoxHig.Value + 140) / 1000
    
    boolstatus = Part.Extension.SelectByID2("D1@Height_LPattern@RSO_Regal.SLDASM", "DIMENSION", 0#, 0#, 0#, True, 0, Nothing, 0)
    Set myDimension = Part.Parameter("D1@Height_LPattern@RSO_Regal.SLDASM")
    myDimension.SystemValue = txbxRegLvl.Value
    
    boolstatus = Part.Extension.SelectByID2("D1@Regal_Height_Plane@RSO_Regal", "DIMENSION", 0#, 0#, 0#, False, 0, Nothing, 0)
    Set myDimension = Part.Parameter("D1@Regal_Height_Plane@RSO_Regal.SLDASM")
    myDimension.SystemValue = (((txbxBoxHig.Value + 140) / 1000) * txbxRegLvl.Value) + 0.1
Else
    MsgBox ("The number of Regal Levels should be more than two!")
End If

'Modify the Number of Regal Fields und box lenght
If txbxRegFld.Value > 0 Then
    boolstatus = Part.Extension.SelectByID2("D1@Fields_LPattern@RSO_Regal.SLDASM", "DIMENSION", 0#, 0#, 0#, True, 0, Nothing, 0)
    Set myDimension = Part.Parameter("D1@Fields_LPattern@RSO_Regal.SLDASM")
    myDimension.SystemValue = txbxRegFld.Value

    boolstatus = Part.Extension.SelectByID2("D2@Height_LPattern@RSO_Regal.SLDASM", "DIMENSION", 0#, 0#, 0#, True, 0, Nothing, 0)
    Set myDimension = Part.Parameter("D2@Height_LPattern@RSO_Regal.SLDASM")
    myDimension.SystemValue = txbxRegFld.Value
    
    boolstatus = Part.Extension.SelectByID2("D1@Regal_Fields_Plane@RSO_Regal", "DIMENSION", 0#, 0#, 0#, False, 0, Nothing, 0)
    Set myDimension = Part.Parameter("D1@Regal_Fields_Plane@RSO_Regal.SLDASM")
    myDimension.SystemValue = ((txbxBoxLen.Value / 1000) * 2) + 0.142
Else
MsgBox ("The number of Fields should be more than one!")
End If

'Modify Gasseanfang_Mirror Position
boolstatus = Part.Extension.SelectByID2("D1@Gasseanfang_Mirror_Plane@RSO_Regal", "DIMENSION", 0#, 0#, 0#, False, 0, Nothing, 0)
Set myDimension = Part.Parameter("D1@Gasseanfang_Mirror_Plane@RSO_Regal.SLDASM")
myDimension.SystemValue = (((txbxBoxLen.Value / 1000) * 2) + 0.142) * (txbxRegFld.Value / 2)

'Modify Fahrschiene levels and lenght
boolstatus = Part.Extension.SelectByID2("D1@Fahrschiene_LPattern@RSO_Regal.SLDASM", "DIMENSION", 0#, 0#, 0#, True, 0, Nothing, 0)
Set myDimension = Part.Parameter("D1@Fahrschiene_LPattern@RSO_Regal.SLDASM")
myDimension.SystemValue = txbxRegLvl.Value

boolstatus = Part.Extension.SelectByID2("D3@Fahrschiene_LPattern@RSO_Regal.SLDASM", "DIMENSION", 0#, 0#, 0#, True, 0, Nothing, 0)
Set myDimension = Part.Parameter("D3@Fahrschiene_LPattern@RSO_Regal.SLDASM")
myDimension.SystemValue = (txbxBoxHig.Value + 140) / 1000

boolstatus = Part.Extension.SelectByID2("Boss-Extrude1@Fahrschiene^RSO_Regal-1@RSO_Regal", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
Set myDimension = Part.Parameter("D1@Boss-Extrude1@Fahrschiene^RSO_Regal-1@RSO_Regal.SLDASM")
myDimension.SystemValue = ((((txbxBoxLen.Value * 2) + 142) * txbxRegFld.Value) + 1060) / 1000

'Modify Shuttle Width
boolstatus = Part.Extension.SelectByID2("D1@Shuttle_Width_Plane@RSO_Regal", "DIMENSION", 0#, 0#, 0#, False, 0, Nothing, 0)
Set myDimension = Part.Parameter("D1@Shuttle_Width_Plane@RSO_Regal.SLDASM")
myDimension.SystemValue = (txbxShuWid.Value / 2) / 1000

'Modify Box Width
If opt4Fach.Value = True Then
    boolstatus = Part.Extension.SelectByID2("D1@Gasse_Width_Plane@RSO_Regal", "DIMENSION", 0#, 0#, 0#, False, 0, Nothing, 0)
    Set myDimension = Part.Parameter("D1@Gasse_Width_Plane@RSO_Regal.SLDASM")
    myDimension.SystemValue = ((txbxBoxWid.Value * 4) + 93) / 1000

    boolstatus = Part.Extension.SelectByID2("3Fach_Box^RSO_Regal-1@RSO_Regal", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
    Part.HideComponent2
    boolstatus = Part.Extension.SelectByID2("4Fach_Box^RSO_Regal-1@RSO_Regal", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
    Part.ShowComponent2
End If

If opt3Fach.Value = True Then
    boolstatus = Part.Extension.SelectByID2("D1@Gasse_Width_Plane@RSO_Regal", "DIMENSION", 0#, 0#, 0#, False, 0, Nothing, 0)
    Set myDimension = Part.Parameter("D1@Gasse_Width_Plane@RSO_Regal.SLDASM")
    myDimension.SystemValue = ((txbxBoxWid.Value * 3) + 73) / 1000
    
    boolstatus = Part.Extension.SelectByID2("4Fach_Box^RSO_Regal-1@RSO_Regal", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
    Part.HideComponent2
    boolstatus = Part.Extension.SelectByID2("3Fach_Box^RSO_Regal-1@RSO_Regal", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
    Part.ShowComponent2
End If

'Rebuild 3D-Model and Zoomtofit and change to four viewports
boolstatus = Part.EditRebuild3()
'Part.ShowNamedView2 "*Front", 1
'Part.ViewZoomtofit2
'Part.ViewZoomtofit2
'Part.ModelViewManager.ViewportDisplay = swViewportDisplay_e.swViewportFourView
'Part.ViewZoomtofit2
'Part.ShowNamedView2 "BOM", -1
'Part.ViewZoomtofit2
Part.ModelViewManager.ViewportDisplay = swViewportDisplay_e.swViewportTwoViewVertical
Part.ShowNamedView2 "BOM01", -1


'Enable Save Regal Button
cmdSave.Enabled = True

End Sub

Private Sub cmdSave_Click()

'Dissolve Component Pattern
boolstatus = Part.Extension.SelectByID2("Fields_LPattern", "COMPPATTERN", 0, 0, 0, False, 0, Nothing, 0)
Part.DissolveComponentPattern
boolstatus = Part.Extension.SelectByID2("Height_LPattern", "COMPPATTERN", 0, 0, 0, False, 0, Nothing, 0)
Part.DissolveComponentPattern
boolstatus = Part.Extension.SelectByID2("Fahrschiene_LPattern", "COMPPATTERN", 0, 0, 0, False, 0, Nothing, 0)
Part.DissolveComponentPattern
boolstatus = Part.Extension.SelectByID2("Gasse_Mirror", "COMPPATTERN", 0, 0, 0, False, 0, Nothing, 0)
Part.DissolveComponentPattern
boolstatus = Part.Extension.SelectByID2("Gasseanfang_Mirror", "COMPPATTERN", 0, 0, 0, False, 0, Nothing, 0)
Part.DissolveComponentPattern
boolstatus = Part.Extension.SelectByID2("Fachwerkstrebe_LPattern", "COMPPATTERN", 0, 0, 0, False, 0, Nothing, 0)
Part.DissolveComponentPattern
   
'Remove all mates and fia all Components using ClsMate CLASS MODULE
Dim RemoveMate As New ClsMate
RemoveMate.RemoveMateAndFixComponents

'Disable Create Regal/Aplly Change buttom
cmdCreate.Enabled = False

'Hide/Unhid the FeatureManager design tree
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swModDocExt = swModel.Extension
    bRet = swModDocExt.HideFeatureManager(False)
    
'End of Design and close program
MsgBox ("Thank you for using Regal Automation. Please save as your 3D-Model, bofor close the document.")
cmdSave.Enabled = True
        If cmdSave.Enabled = True Then
            End
        End If

End Sub

Private Sub UserForm_Activate()
MsgBox "Please close all open files before using this software!"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Prevent user from closing with the Close box in the title bar.
    If CloseMode <> 1 Then Cancel = 1
    MsgBox ("The Close box won't work! Please click on the 'Create Regal' button first and then on the 'Save' button!")
End Sub

