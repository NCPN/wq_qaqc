Operation =1
Option =0
Having ="(((tblProjects.ProjectID)=\"NCPN_UTE\" Or (tblProjects.ProjectID)=\"NCPN_UTM\") "
    "AND ((tblVisits.START_DATE) Between #10/1/2005# And #9/30/2019#) AND ((tblCharac"
    "teristics.FIELD_LAB)=\"Field\"))"
Begin InputTables
    Name ="tblProjects"
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
    Expression ="tblCharacteristics.FIELD_LAB"
    Alias ="Result_Count"
    Expression ="Count(tblResults.LocRSULT_IS_NUMBER)"
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
    Expression ="tblProjects.ProjectID"
    Flag =0
    Expression ="tblLocations.StationID"
    Flag =0
    Expression ="tblVisits.START_DATE"
    Flag =0
End
Begin Groups
    Expression ="tblProjects.ProjectID"
    GroupLevel =0
    Expression ="tblLocations.StationID"
    GroupLevel =0
    Expression ="tblLocations.StationName"
    GroupLevel =0
    Expression ="tblVisits.START_DATE"
    GroupLevel =0
    Expression ="tblCharacteristics.FIELD_LAB"
    GroupLevel =0
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
        dbText "Name" ="tblCharacteristics.FIELD_LAB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Result_Count"
        dbInteger "ColumnWidth" ="1545"
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
    Left =20
    Top =10
    Right =1007
    Bottom =406
    Left =-1
    Top =-1
    Right =963
    Bottom =181
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =14
        Top =9
        Right =158
        Bottom =153
        Top =0
        Name ="tblProjects"
        Name =""
    End
    Begin
        Left =22
        Top =159
        Right =166
        Bottom =303
        Top =0
        Name ="tblLocations"
        Name =""
    End
    Begin
        Left =202
        Top =18
        Right =346
        Bottom =152
        Top =0
        Name ="tblVisits"
        Name =""
    End
    Begin
        Left =384
        Top =15
        Right =528
        Bottom =159
        Top =0
        Name ="tblActivities"
        Name =""
    End
    Begin
        Left =578
        Top =17
        Right =722
        Bottom =161
        Top =0
        Name ="tblResults"
        Name =""
    End
    Begin
        Left =755
        Top =20
        Right =933
        Bottom =163
        Top =0
        Name ="tblCharacteristics"
        Name =""
    End
End
