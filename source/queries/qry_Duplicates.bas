Operation =1
Option =0
Where ="(((tblLocations.StationID) In (\"4951260\",\"4951261\",\"4951265\",\"4951266\"))"
    " AND ((tblResults.VALUE_STATUS)='P') AND ((IIf(IsNumeric([RESULT_TEXT]),CDbl([RE"
    "SULT_TEXT]),-9999))<>-9999) AND ((tblCharacteristics.FIELD_LAB)=\"Lab\"))"
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
    Expression ="tblVisits.START_DATE"
    Expression ="tblLocations.StationID"
    Expression ="tblCharacteristics.DISPLAY_NAME"
    Expression ="tblLocations.StationName"
    Expression ="tblResults.DETECTION_CONDITION"
    Expression ="tblResults.RESULT_TEXT"
    Expression ="tblResults.LAB_REMARKS"
    Expression ="tblResults.RESULT_COMMENT"
    Expression ="tblVisits.VISIT_COMMENT"
    Expression ="tblResults.VALUE_STATUS"
    Alias ="Numeric_Result"
    Expression ="IIf(IsNumeric([RESULT_TEXT]),CDbl([RESULT_TEXT]),-9999)"
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
    Expression ="tblVisits.START_DATE"
    Flag =0
    Expression ="tblLocations.StationID"
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
        dbText "Name" ="Numeric_Result"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblLocations.StationName"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =15
    Top =19
    Right =1010
    Bottom =497
    Left =-1
    Top =-1
    Right =971
    Bottom =140
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =11
        Top =10
        Right =155
        Bottom =154
        Top =0
        Name ="tblProjects"
        Name =""
    End
    Begin
        Left =10
        Top =160
        Right =154
        Bottom =304
        Top =0
        Name ="tblLocations"
        Name =""
    End
    Begin
        Left =436
        Top =8
        Right =580
        Bottom =152
        Top =0
        Name ="tblActivities"
        Name =""
    End
    Begin
        Left =198
        Top =12
        Right =388
        Bottom =156
        Top =0
        Name ="tblVisits"
        Name =""
    End
    Begin
        Left =612
        Top =9
        Right =771
        Bottom =151
        Top =0
        Name ="tblResults"
        Name =""
    End
    Begin
        Left =811
        Top =10
        Right =955
        Bottom =154
        Top =0
        Name ="tblCharacteristics"
        Name =""
    End
End
