﻿Operation =1
Option =0
Where ="(((tblProjects.ProjectID)=\"NCPN_UTE\") AND ((tblVisits.START_DATE) Between #10/"
    "1/2009# And #9/30/2012#) AND ((tblCharacteristics.DISPLAY_NAME)=\"Temperature, w"
    "ater\"))"
Begin InputTables
    Name ="tblProjects"
    Name ="tblLocationProjectAssignment"
    Name ="tblLocations"
    Name ="tblVisits"
    Name ="tblActivities"
    Name ="tblResults"
    Name ="tblCharacteristics"
End
Begin OutputColumns
    Expression ="tblProjects.ProjectID"
    Expression ="tblLocations.StationID"
    Expression ="tblLocations.StationName"
    Expression ="tblVisits.START_DATE"
    Expression ="tblCharacteristics.DISPLAY_NAME"
    Expression ="tblResults.RESULT_TEXT"
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
    RightTable ="tblLocationProjectAssignment"
    Expression ="tblLocations.LocSTATN_IS_NUMBER = tblLocationProjectAssignment.LocSTATN_IS_NUMBE"
        "R"
    Flag =1
    LeftTable ="tblLocations"
    RightTable ="tblLocationProjectAssignment"
    Expression ="tblLocations.LocSTATN_ORG_ID = tblLocationProjectAssignment.LocSTATN_ORG_ID"
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
    RightTable ="tblLocationProjectAssignment"
    Expression ="tblProjects.LocProj_IS_NUMBER = tblLocationProjectAssignment.LocProj_IS_NUMBER"
    Flag =1
    LeftTable ="tblProjects"
    RightTable ="tblLocationProjectAssignment"
    Expression ="tblProjects.LocProj_ORG_ID = tblLocationProjectAssignment.LocProj_ORG_ID"
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
        dbText "Name" ="tblResults.RESULT_TEXT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblResults.VALUE_STATUS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCharacteristics.DISPLAY_NAME"
        dbInteger "ColumnWidth" ="2595"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLocations.StationName"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =17
    Top =39
    Right =1185
    Bottom =527
    Left =-1
    Top =-1
    Right =1140
    Bottom =277
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =13
        Top =13
        Right =157
        Bottom =157
        Top =0
        Name ="tblProjects"
        Name =""
    End
    Begin
        Left =9
        Top =164
        Right =153
        Bottom =308
        Top =0
        Name ="tblLocationProjectAssignment"
        Name =""
    End
    Begin
        Left =197
        Top =16
        Right =341
        Bottom =160
        Top =0
        Name ="tblLocations"
        Name =""
    End
    Begin
        Left =358
        Top =17
        Right =502
        Bottom =161
        Top =0
        Name ="tblVisits"
        Name =""
    End
    Begin
        Left =530
        Top =19
        Right =674
        Bottom =163
        Top =0
        Name ="tblActivities"
        Name =""
    End
    Begin
        Left =712
        Top =17
        Right =856
        Bottom =161
        Top =0
        Name ="tblResults"
        Name =""
    End
    Begin
        Left =880
        Top =20
        Right =1024
        Bottom =164
        Top =0
        Name ="tblCharacteristics"
        Name =""
    End
End
