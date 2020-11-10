Operation =1
Option =0
Where ="(((tblProjects.ProjectID)=\"NCPN_UTE\") And ((tblVisits.START_DATE) Between #10/"
    "1/2018# And #9/30/2019#) And ((IIf(IsNumeric(tblResults!PRECISION),CDbl(tblResul"
    "ts!PRECISION),0))>5) And ((tblResults.VALUE_STATUS)=\"P\") And ((tblCharacterist"
    "ics.FIELD_LAB)=\"Field\") And ((tblCharacteristics.LocCharNameCode)=\"NCPN_flow_"
    "meter_002\"))"
Begin InputTables
    Name ="tblProjects"
    Name ="tblLocations"
    Name ="tblVisits"
    Name ="tblCharacteristics"
    Name ="tblActivities"
    Name ="tblResults"
End
Begin OutputColumns
    Expression ="tblProjects.ProjectID"
    Expression ="tblLocations.StationID"
    Expression ="tblVisits.START_DATE"
    Expression ="tblLocations.StationName"
    Expression ="tblCharacteristics.DISPLAY_NAME"
    Expression ="tblResults.DETECTION_CONDITION"
    Expression ="tblResults.RESULT_TEXT"
    Alias ="PRECISION"
    Expression ="IIf(IsNumeric([tblResults]![PRECISION]),CDbl([tblResults]![PRECISION]),0)"
    Expression ="tblResults.VALUE_STATUS"
    Expression ="tblResults.LAB_REMARKS"
    Expression ="tblResults.RESULT_COMMENT"
    Expression ="tblVisits.VISIT_COMMENT"
End
Begin Joins
    LeftTable ="tblLocations"
    RightTable ="tblVisits"
    Expression ="tblLocations.LocSTATN_ORG_ID = tblVisits.LocSTATN_ORG_ID"
    Flag =1
    LeftTable ="tblLocations"
    RightTable ="tblVisits"
    Expression ="tblLocations.LocSTATN_IS_NUMBER = tblVisits.LocSTATN_IS_NUMBER"
    Flag =1
    LeftTable ="tblProjects"
    RightTable ="tblVisits"
    Expression ="tblProjects.LocProj_ORG_ID = tblVisits.LocProj_ORG_ID"
    Flag =1
    LeftTable ="tblProjects"
    RightTable ="tblVisits"
    Expression ="tblProjects.LocProj_IS_NUMBER = tblVisits.LocProj_IS_NUMBER"
    Flag =1
    LeftTable ="tblActivities"
    RightTable ="tblResults"
    Expression ="tblActivities.LocFdAct_ORG_ID = tblResults.LocFdAct_Org_ID"
    Flag =1
    LeftTable ="tblActivities"
    RightTable ="tblResults"
    Expression ="tblActivities.LocFdAct_IS_NUMBER = tblResults.LocFdAct_IS_NUMBER"
    Flag =1
    LeftTable ="tblCharacteristics"
    RightTable ="tblResults"
    Expression ="tblCharacteristics.LocCHDEF_ORG_ID = tblResults.LocChDef_Org_ID"
    Flag =1
    LeftTable ="tblCharacteristics"
    RightTable ="tblResults"
    Expression ="tblCharacteristics.LocCHDEF_IS_NUMBER = tblResults.LocChDef_IS_NUMBER"
    Flag =1
    LeftTable ="tblVisits"
    RightTable ="tblActivities"
    Expression ="tblVisits.LocStVst_ORG_ID = tblActivities.LocStVst_ORG_ID"
    Flag =1
    LeftTable ="tblVisits"
    RightTable ="tblActivities"
    Expression ="tblVisits.LocStVst_IS_NUMBER = tblActivities.LocStVst_IS_NUMBER"
    Flag =1
End
Begin OrderBy
    Expression ="tblVisits.START_DATE"
    Flag =0
    Expression ="tblLocations.StationName"
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
        dbText "Name" ="tblResults.VALUE_STATUS"
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
        dbText "Name" ="PRECISION"
        dbInteger "ColumnWidth" ="1410"
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
    Left =33
    Top =17
    Right =1130
    Bottom =423
    Left =-1
    Top =-1
    Right =1069
    Bottom =180
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
        Left =45
        Top =164
        Right =189
        Bottom =308
        Top =0
        Name ="tblLocations"
        Name =""
    End
    Begin
        Left =227
        Top =32
        Right =417
        Bottom =176
        Top =0
        Name ="tblVisits"
        Name =""
    End
    Begin
        Left =824
        Top =39
        Right =968
        Bottom =183
        Top =0
        Name ="tblCharacteristics"
        Name =""
    End
    Begin
        Left =458
        Top =32
        Right =602
        Bottom =176
        Top =0
        Name ="tblActivities"
        Name =""
    End
    Begin
        Left =634
        Top =35
        Right =793
        Bottom =177
        Top =0
        Name ="tblResults"
        Name =""
    End
End
