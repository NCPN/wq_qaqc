﻿Operation =1
Option =0
Where ="(((tblProjects.ProjectID)=\"NCPN_UTE\") AND ((tblVisits.START_DATE) Between #10/"
    "1/2018# And #9/30/2019#) AND ((tblCharacteristics.DISPLAY_NAME)=\"Specific condu"
    "ctance\") AND ((tblResults.DETECTION_CONDITION)<>\"*Not Reported\") AND ((tblRes"
    "ults.VALUE_STATUS)=\"P\") AND ((IIf(IsNumeric([RESULT_TEXT]),CDbl([RESULT_TEXT])"
    ",-999))<100 Or (IIf(IsNumeric([RESULT_TEXT]),CDbl([RESULT_TEXT]),-999))>2100))"
Begin InputTables
    Name ="tblProjects"
    Name ="tblLocations"
    Name ="tblActivities"
    Name ="tblVisits"
    Name ="tblResults"
    Name ="tblCharacteristics"
End
Begin OutputColumns
    Expression ="tblProjects.ProjectID"
    Expression ="tblLocations.StationID"
    Expression ="tblLocations.StationName"
    Expression ="tblVisits.START_DATE"
    Expression ="tblCharacteristics.DISPLAY_NAME"
    Expression ="tblResults.DETECTION_CONDITION"
    Expression ="tblResults.RESULT_TEXT"
    Expression ="tblResults.LAB_REMARKS"
    Expression ="tblResults.RESULT_COMMENT"
    Expression ="tblVisits.VISIT_COMMENT"
    Expression ="tblResults.VALUE_STATUS"
End
Begin Joins
    LeftTable ="tblActivities"
    RightTable ="tblResults"
    Expression ="tblActivities.LocFdAct_IS_NUMBER = tblResults.LocFdAct_IS_NUMBER"
    Flag =1
    LeftTable ="tblActivities"
    RightTable ="tblResults"
    Expression ="tblActivities.LocFdAct_ORG_ID = tblResults.LocFdAct_Org_ID"
    Flag =1
    LeftTable ="tblCharacteristics"
    RightTable ="tblResults"
    Expression ="tblCharacteristics.LocCHDEF_IS_NUMBER = tblResults.LocChDef_IS_NUMBER"
    Flag =1
    LeftTable ="tblCharacteristics"
    RightTable ="tblResults"
    Expression ="tblCharacteristics.LocCHDEF_ORG_ID = tblResults.LocChDef_Org_ID"
    Flag =1
    LeftTable ="tblLocations"
    RightTable ="tblVisits"
    Expression ="tblLocations.LocSTATN_IS_NUMBER = tblVisits.LocSTATN_IS_NUMBER"
    Flag =1
    LeftTable ="tblLocations"
    RightTable ="tblVisits"
    Expression ="tblLocations.LocSTATN_ORG_ID = tblVisits.LocSTATN_ORG_ID"
    Flag =1
    LeftTable ="tblProjects"
    RightTable ="tblVisits"
    Expression ="tblProjects.LocProj_IS_NUMBER = tblVisits.LocProj_IS_NUMBER"
    Flag =1
    LeftTable ="tblProjects"
    RightTable ="tblVisits"
    Expression ="tblProjects.LocProj_ORG_ID = tblVisits.LocProj_ORG_ID"
    Flag =1
    LeftTable ="tblVisits"
    RightTable ="tblActivities"
    Expression ="tblVisits.LocStVst_IS_NUMBER = tblActivities.LocStVst_IS_NUMBER"
    Flag =1
    LeftTable ="tblVisits"
    RightTable ="tblActivities"
    Expression ="tblVisits.LocStVst_ORG_ID = tblActivities.LocStVst_ORG_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tblLocations.StationID"
    Flag =0
    Expression ="tblVisits.START_DATE"
    Flag =0
    Expression ="tblCharacteristics.DISPLAY_NAME"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblProjects.ProjectID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLocations.StationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVisits.START_DATE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCharacteristics.DISPLAY_NAME"
        dbInteger "ColumnWidth" ="3705"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblResults.LAB_REMARKS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblResults.RESULT_COMMENT"
        dbInteger "ColumnWidth" ="3000"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblResults.DETECTION_CONDITION"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblResults.RESULT_TEXT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVisits.VISIT_COMMENT"
        dbInteger "ColumnWidth" ="2850"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblResults.VALUE_STATUS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLocations.StationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLocations.[Station Name]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =12
    Top =19
    Right =1026
    Bottom =383
    Left =-1
    Top =-1
    Right =990
    Bottom =130
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =16
        Top =6
        Right =160
        Bottom =131
        Top =0
        Name ="tblProjects"
        Name =""
    End
    Begin
        Left =18
        Top =139
        Right =162
        Bottom =266
        Top =0
        Name ="tblLocations"
        Name =""
    End
    Begin
        Left =417
        Top =30
        Right =561
        Bottom =174
        Top =0
        Name ="tblActivities"
        Name =""
    End
    Begin
        Left =194
        Top =27
        Right =384
        Bottom =171
        Top =0
        Name ="tblVisits"
        Name =""
    End
    Begin
        Left =588
        Top =34
        Right =748
        Bottom =176
        Top =0
        Name ="tblResults"
        Name =""
    End
    Begin
        Left =792
        Top =40
        Right =936
        Bottom =184
        Top =0
        Name ="tblCharacteristics"
        Name =""
    End
End
