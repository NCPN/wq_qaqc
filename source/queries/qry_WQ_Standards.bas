﻿Operation =1
Option =0
Begin InputTables
    Name ="tblLocations"
    Name ="tblLocationStationGroupAssignment"
    Name ="tblLocationStationGroups"
    Name ="tblLocationWQStandardAssignment"
    Name ="tblWQStandards"
End
Begin OutputColumns
    Alias ="Park"
    Expression ="tblLocationStationGroups.ID_CODE"
    Expression ="tblLocations.State"
    Expression ="tblLocations.StationID"
    Expression ="tblLocations.StationName"
    Expression ="tblWQStandards.StandardID"
    Expression ="tblWQStandards.StandardName"
End
Begin Joins
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
    RightTable ="tblLocationWQStandardAssignment"
    Expression ="tblLocations.LocSTATN_IS_NUMBER = tblLocationWQStandardAssignment.LocStatn_IS_NU"
        "MBER"
    Flag =1
    LeftTable ="tblLocations"
    RightTable ="tblLocationWQStandardAssignment"
    Expression ="tblLocations.LocSTATN_ORG_ID = tblLocationWQStandardAssignment.LocStatn_ORG_ID"
    Flag =1
    LeftTable ="tblWQStandards"
    RightTable ="tblLocationWQStandardAssignment"
    Expression ="tblWQStandards.Standard_IS_NUMBER = tblLocationWQStandardAssignment.Standard_IS_"
        "NUMBER"
    Flag =1
    LeftTable ="tblWQStandards"
    RightTable ="tblLocationWQStandardAssignment"
    Expression ="tblWQStandards.Standard_Org_ID = tblLocationWQStandardAssignment.Standard_Org_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tblLocationStationGroups.ID_CODE"
    Flag =0
    Expression ="tblLocations.StationID"
    Flag =0
    Expression ="tblWQStandards.StandardID"
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
        dbText "Name" ="tblLocations.StationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLocations.[Station Name]"
        dbInteger "ColumnWidth" ="5325"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLocations.State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Park"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWQStandards.StandardID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblWQStandards.StandardName"
        dbInteger "ColumnWidth" ="3450"
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
    Left =42
    Top =28
    Right =906
    Bottom =434
    Left =-1
    Top =-1
    Right =840
    Bottom =186
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =31
        Top =11
        Right =175
        Bottom =155
        Top =0
        Name ="tblLocations"
        Name =""
    End
    Begin
        Left =22
        Top =166
        Right =220
        Bottom =310
        Top =0
        Name ="tblLocationStationGroupAssignment"
        Name =""
    End
    Begin
        Left =263
        Top =166
        Right =442
        Bottom =310
        Top =0
        Name ="tblLocationStationGroups"
        Name =""
    End
    Begin
        Left =261
        Top =8
        Right =405
        Bottom =152
        Top =0
        Name ="tblLocationWQStandardAssignment"
        Name =""
    End
    Begin
        Left =468
        Top =5
        Right =612
        Bottom =149
        Top =0
        Name ="tblWQStandards"
        Name =""
    End
End
