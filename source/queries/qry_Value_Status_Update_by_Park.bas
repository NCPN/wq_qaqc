﻿Operation =4
Option =0
Where ="(((tblResults.VALUE_STATUS)=\"P\") AND ((tblProjects.ProjectID)=\"NCPN_UTE\" Or "
    "(tblProjects.ProjectID)=\"NCPN_UTM\") AND ((tblVisits.START_DATE) Between #10/1/"
    "2018# And #9/30/2019#))"
Begin InputTables
    Name ="tblProjects"
    Name ="tblLocationProjectAssignment"
    Name ="tblLocations"
    Name ="tblLocationStationGroupAssignment"
    Name ="tblLocationStationGroups"
    Name ="tblVisits"
    Name ="tblActivities"
    Name ="tblResults"
    Name ="tblCharacteristics"
End
Begin OutputColumns
    Name ="tblResults.VALUE_STATUS"
    Expression ="\"F\""
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
    LeftTable ="tblLocationStationGroups"
    RightTable ="tblLocationStationGroupAssignment"
    Expression ="tblLocationStationGroups.LocStatnGRP_IS_NUMBER = tblLocationStationGroupAssignme"
        "nt.LocStatnGrp_IS_NUMBER"
    Flag =1
    LeftTable ="tblLocationStationGroups"
    RightTable ="tblLocationStationGroupAssignment"
    Expression ="tblLocationStationGroups.LocStatnGRP_ORG_ID = tblLocationStationGroupAssignment."
        "LocStatnGrp_ORG_ID"
    Flag =1
    LeftTable ="tblLocations"
    RightTable ="tblLocationStationGroupAssignment"
    Expression ="tblLocations.LocSTATN_IS_NUMBER = tblLocationStationGroupAssignment.LocStatn_IS_"
        "NUMBER"
    Flag =1
    LeftTable ="tblLocations"
    RightTable ="tblLocationStationGroupAssignment"
    Expression ="tblLocations.LocSTATN_ORG_ID = tblLocationStationGroupAssignment.LocStatn_ORG_ID"
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
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
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
        dbText "Name" ="tblLocations.[Station Name]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVisits.START_DATE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblResults.VALUE_TYPE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblResults.LAB_REMARKS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblResults.RESULT_COMMENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCharacteristics.LocCharNameCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCharacteristics.FIELD_LAB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblResults.VALUE_STATUS"
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
        dbText "Name" ="tblLocationStationGroups.ID_CODE"
        dbInteger "ColumnWidth" ="1380"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =17
    Top =22
    Right =973
    Bottom =500
    Left =-1
    Top =-1
    Right =932
    Bottom =262
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblProjects"
        Name =""
    End
    Begin
        Left =47
        Top =167
        Right =191
        Bottom =311
        Top =0
        Name ="tblLocationProjectAssignment"
        Name =""
    End
    Begin
        Left =243
        Top =10
        Right =387
        Bottom =154
        Top =0
        Name ="tblLocations"
        Name =""
    End
    Begin
        Left =422
        Top =6
        Right =580
        Bottom =150
        Top =0
        Name ="tblLocationStationGroupAssignment"
        Name =""
    End
    Begin
        Left =612
        Top =7
        Right =763
        Bottom =151
        Top =0
        Name ="tblLocationStationGroups"
        Name =""
    End
    Begin
        Left =232
        Top =170
        Right =376
        Bottom =314
        Top =0
        Name ="tblVisits"
        Name =""
    End
    Begin
        Left =401
        Top =166
        Right =545
        Bottom =310
        Top =0
        Name ="tblActivities"
        Name =""
    End
    Begin
        Left =564
        Top =159
        Right =724
        Bottom =303
        Top =0
        Name ="tblResults"
        Name =""
    End
    Begin
        Left =750
        Top =157
        Right =894
        Bottom =301
        Top =0
        Name ="tblCharacteristics"
        Name =""
    End
End
