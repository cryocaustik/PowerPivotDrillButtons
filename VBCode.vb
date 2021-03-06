Sub DrillDown()
On Error GoTo ErrorHandler
'This code was developed by stackoverflow user CITYINBETWEEN and was posted on the stackoverflow 
'forums for everyone to use free of charge and is not to be sold to others.   
'
' Drill Down Macro
'
    Dim HrchyPreFix, HrchyLstLvl, MyCurrLocation, MyPivTblName, MyDrillTo
    '---------- User Entry Needed ----------'
    ' prefix used for hierarchy levels
    HrchyPreFix = "Lvl"
    ' set hierarchy last drill down level
    HrchyLstLvl = "4"
    '---------- End of User Entry ----------'

    ' set pivot table name of active cell
    MyPivTblName = ActiveCell.PivotTable
    ' set pivot field selected of active cell
    MyCurrLocation = ActiveCell.PivotCell.PivotField
    ' set what hierarchy lvl to drill down to
    MyDrillTo = ActiveCell.PivotCell.PivotItem

    ' find current hierarchy lvl of active cell. if at the last lvl, if statement goes to BottomOfDrillDownHandler
    HrchyCurrLvl = (Right(Left(Mid(ActiveCell.PivotField, InStr(1, ActiveCell.PivotField, HrchyPreFix)), 4), 1))
    ' If at last hierarchy lvl, go to BottomOfDrillDownHandler
    If HrchyCurrLvl = HrchyLstLvl Then
        GoTo BottomOfDrillDownHandler
    End If

    ' drill down code
    ActiveSheet.PivotTables(MyPivTblName).DrillDown ActiveSheet.PivotTables( _
        MyPivTblName).PivotFields(MyCurrLocation).PivotItems(MyDrillTo), _
        ActiveSheet.PivotTables(MyPivTblName).PivotRowAxis.PivotLines(1)
    Exit Sub

' Error handler for when you cannot drill down any further
BottomOfDrillDownHandler:
    Dim ErrMsg1, ErrTitle1
    ErrMsg1 = "Unable to Drill Down any further as you're at the bottom of the Drill Down"
    ErrTitle1 = "Drill Down Error"
    MsgBox ErrMsg1, , ErrTitle1
    Exit Sub

' general error handler
ErrorHandler:
    Dim ErrMsg2, ErrTitle2, ErrMsg3, ErrTitle3
    If Err.Number = 1004 Then
        ErrMsg2 = "Please select a drillable item"
        ErrTitle2 = "Drill Down Error"
        MsgBox ErrMsg2, , ErrTitle2
    ElseIf Err.Number <> 0 Then
        ErrMsg3 = "Error # " & Str(Err.Number) & " was generated by " _
            & Err.Source & Chr(13) & "Error Line: " & Erl & Chr(13) & Err.Description
        ErrTitle3 = "Error"
    MsgBox ErrMsg3, , ErrTitle3, Err.HelpFile, Err.HelpContext
    End If
End Sub

'--------------------------------------------------------------------
Sub DrillUp()
On Error GoTo ErrorHandler
'This code was developed by stackoverflow user CITYINBETWEEN and was posted on the stackoverflow 
'forums for everyone to use free of charge and is not to be sold to others.   
'
' Drill Up 1 level Macro
'
    Dim PwrPivTblNm, HrchyNm, HrchyPreFix, HrchyTopLvl, MyCurrLocation, MyPivTblName, MyDrillTo, MyCurrLvl, HrchyPrevLvl As Integer

    '---------- User Entry Needed ----------'
    ' Name of table in powerpivot where the hierarchy exists
    PwrPivTblNm = "vEmployeeDepartment"
    ' name given to hierarchy in powerpivot
    HrchyNm = "Hierarchy1"
    ' prefix used for hierarchy levels
    HrchyPreFix = "Lvl"
    ' set top hierarchy level
    HrchyTopLvl = "1"
    '---------- End of User Entry ----------'

    ' set pivot table name of active cell
    MyPivTblName = ActiveCell.PivotTable
    ' set pivot field selected of active cell
    MyCurrLocation = ActiveCell.PivotCell.PivotField
    ' set from what hierarchy lvl to drill up from
    MyDrillUpFrom = ActiveCell.PivotCell.PivotItem
    ' find prev. hierarchy lvl of active cell
    HrchyPrevLvl = (Right(Left(Mid(ActiveCell.PivotField, InStr(1, ActiveCell.PivotField, HrchyPreFix)), 4), 1) - 1)

    ' find current hierarchy lvl of active cell. if at the top lvl, if statement goes to TopOfDrillUpHandler
    HrchyCurrLvl = (Right(Left(Mid(ActiveCell.PivotField, InStr(1, ActiveCell.PivotField, HrchyPreFix)), 4), 1))
    ' If at last hierarchy lvl, go to TopOfDrillUpHandler
    If HrchyCurrLvl = HrchyTopLvl Then
        GoTo TopOfDrillUpHandler
    End If

    ' set hierarchy level to drill up to
    HrchyLvlDrillTo = "[" & PwrPivTblNm & "].[" & HrchyNm & "].[" & _
                    Mid(ActiveCell.PivotField, InStr(1, ActiveCell.PivotField, HrchyPreFix), 3) & HrchyPrevLvl _
                    & "]"

    ' drill up code
    ActiveSheet.PivotTables(MyPivTblName).DrillUp ActiveSheet.PivotTables( _
        MyPivTblName).PivotFields(MyCurrLocation).PivotItems(MyDrillUpFrom), _
        ActiveSheet.PivotTables(MyPivTblName).PivotRowAxis.PivotLines(1), HrchyLvlDrillTo
    Exit Sub

' Error handler for when you cannot drill up any further
TopOfDrillUpHandler:
    Dim ErrMsg1, ErrTitle1
    ErrMsg1 = "Unable to Drill Up any further as you're at the top of the Drill Up"
    ErrTitle1 = "Drill Up Error"
    MsgBox ErrMsg1, , ErrTitle1
    Exit Sub

' General Error handler
ErrorHandler:
    Dim ErrMsg2, ErrTitle2, ErrMsg3, ErrTitle3
    If Err.Number = 1004 Then
        ErrMsg2 = "Please select a drillable item"
        ErrTitle2 = "Drill Up Error"
        MsgBox ErrMsg2, , ErrTitle2
    ElseIf Err.Number <> 0 Then
        ErrMsg3 = "Error # " & Str(Err.Number) & " was generated by " _
            & Err.Source & Chr(13) & "Error Line: " & Erl & Chr(13) & Err.Description
        ErrTitle3 = "Error"
    MsgBox ErrMsg3, , ErrTitle3, Err.HelpFile, Err.HelpContext
    End If
End Sub


'--------------------------------------------------------------------
Sub DrillToTop()
On Error GoTo ErrorHandler
'This code was developed by stackoverflow user CITYINBETWEEN and was posted on the stackoverflow 
'forums for everyone to use free of charge and is not to be sold to others.   
'
' Dill To Top Macro Macro
'
    Dim PwrPivTblNm, HrchyNm, HrchyPreFix, HrchyTopLvl, MyCurrLocation, MyPivTblName, MyDrillTo

    '---------- User Entry Needed ----------'
    ' Name of table in powerpivot where the hierarchy exists
    PwrPivTblNm = "vEmployeeDepartment"
    ' name given to hierarchy in powerpivot
    HrchyNm = "Hierarchy1"
    ' prefix used for hierarchy levels
    HrchyPreFix = "Lvl"
    ' set top hierarchy level
    HrchyTopLvl = "1"
    '---------- End of User Entry ----------'

    ' set pivot table name of active cell
    MyPivTblName = ActiveCell.PivotTable
    ' set pivot field selected of active cell
    MyCurrLocation = ActiveCell.PivotCell.PivotField
    ' set from what hierarchy lvl to drill up from
    MyDrillUpFrom = ActiveCell.PivotCell.PivotItem

    ' find prev. hierarchy lvl of active cell. if already at top lvl, if statement goes to AlreadyAtTopHandler
    HrchyPrevLvl = (Right(Left(Mid(ActiveCell.PivotField, InStr(1, ActiveCell.PivotField, "Lvl")), 4), 1) - 1)
    ' If at hierarchy lvl 1, go to TopOfDrillUpHandler
    If HrchyPrevLvl = "0" Then
        GoTo AlreadyAtTopHandler
    End If

    ' set top hierarchy level to drill up to
    HrchyLvlDrillTo = "[" & PwrPivTblNm & "].[" & HrchyNm & "].[" & _
                    Mid(ActiveCell.PivotField, InStr(1, ActiveCell.PivotField, HrchyPreFix), 3) & HrchyTopLvl _
                    & "]"

    ' drill to top code
    ActiveSheet.PivotTables(MyPivTblName).DrillUp ActiveSheet.PivotTables( _
        MyPivTblName).PivotFields(MyCurrLocation).PivotItems(MyDrillUpFrom), _
        ActiveSheet.PivotTables(MyPivTblName).PivotRowAxis.PivotLines(1), _
        HrchyLvlDrillTo
    Exit Sub

' Error handler for when user is already at the top level
AlreadyAtTopHandler:
    Dim ErrMsg1, ErrTitle1
    ErrMsg1 = "Unable to Drill to Top as you're already at the top level"
    ErrTitle1 = "Drill to Top Error"
    MsgBox ErrMsg1, , ErrTitle1
    Exit Sub

' General Error handler
ErrorHandler:
    Dim ErrMsg2, ErrTitle2, ErrMsg3, ErrTitle3
    If Err.Number = 1004 Then
        ErrMsg2 = "Please select a drillable item"
        ErrTitle2 = "Drill to Top Error"
        MsgBox ErrMsg2, , ErrTitle2
    ElseIf Err.Number <> 0 Then
        ErrMsg3 = "Error # " & Str(Err.Number) & " was generated by " _
            & Err.Source & Chr(13) & "Error Line: " & Erl & Chr(13) & Err.Description
        ErrTitle3 = "Error"
    MsgBox ErrMsg3, , ErrTitle3, Err.HelpFile, Err.HelpContext
    End If
End Sub

