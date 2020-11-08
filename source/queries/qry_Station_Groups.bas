Operation =1
Option =0
Begin InputTables
    Name ="tblProjects"
    Name ="tblLocations"
    Name ="tblLocationStationGroups"
    Name ="tblLocationStationGroupAssignment"
    Name ="tblLocationProjectAssignment"
End
Begin OutputColumns
    Expression ="tblProjects.ProjectID"
    Expression ="tblLocationStationGroups.ID_CODE"
    Expression ="tblLocationStationGroups.NAME"
    Expression ="tblLocationStationGroups.DESCRIPTION_TEXT"
    Expression ="tblLocations.StationID"
    Expression ="tblLocations.[Station Name]"
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
    RightTable ="tblLocationProjectAssignment"
    Expression ="tblLocations.LocSTATN_IS_NUMBER = tblLocationProjectAssignment.LocSTATN_IS_NUMBE"
        "R"
    Flag =1
    LeftTable ="tblLocations"
    RightTable ="tblLocationProjectAssignment"
    Expression ="tblLocations.LocSTATN_ORG_ID = tblLocationProjectAssignment.LocSTATN_ORG_ID"
    Flag =1
    LeftTable ="tblProjects"
    RightTable ="tblLocationProjectAssignment"
    Expression ="tblProjects.LocProj_IS_NUMBER = tblLocationProjectAssignment.LocProj_IS_NUMBER"
    Flag =1
    LeftTable ="tblProjects"
    RightTable ="tblLocationProjectAssignment"
    Expression ="tblProjects.LocProj_ORG_ID = tblLocationProjectAssignment.LocProj_ORG_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tblProjects.ProjectID"
    Flag =0
    Expression ="tblLocationStationGroups.ID_CODE"
    Flag =0
    Expression ="tblLocations.StationID"
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
        dbText "Name" ="tblLocations.[Station Name]"
        dbInteger "ColumnWidth" ="5865"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLocationStationGroups.ID_CODE"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLocationStationGroups.NAME"
        dbInteger "ColumnWidth" ="2655"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLocationStationGroups.DESCRIPTION_TEXT"
        dbInteger "ColumnWidth" ="4350"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =39
    Top =317
    Right =1400
    Bottom =697
    Left =-1
    Top =-1
    Right =1329
    Bottom =169
    Left =0
    Top =0
    ColumnsShown =539
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
        Left =406
        Top =14
        Right =550
        Bottom =158
        Top =0
        Name ="tblLocations"
        Name =""
    End
    Begin
        Left =854
        Top =13
        Right =1055
        Bottom =157
        Top =0
        Name ="tblLocationStationGroups"
        Name =""
    End
    Begin
        Left =596
        Top =14
        Right =806
        Bottom =158
        Top =0
        Name ="tblLocationStationGroupAssignment"
        Name =""
    End
    Begin
        Left =226
        Top =22
        Right =370
        Bottom =166
        Top =0
        Name ="tblLocationProjectAssignment"
        Name =""
    End
End
