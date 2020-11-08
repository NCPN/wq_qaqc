﻿Operation =1
Option =0
Where ="(((tblLocations.StationID)=\"4950800\") AND ((tblCharacteristics.DISPLAY_NAME)=\""
    "flow\") AND ((tblResults.RESULT_TEXT) Is Not Null) AND ((tblResults.DETECTION_CO"
    "NDITION)<>\"*Non-detect\") AND ((tblResults.VALUE_STATUS)=\"P\") AND ((tblCharac"
    "teristics.FIELD_LAB)=\"Field\") AND ((IIf(IsNumeric([Result_text]),1,0))=1))"
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
    Expression ="tblCharacteristics.DISPLAY_NAME"
    Alias ="NumericValue"
    Expression ="CDbl([Result_Text])"
    Expression ="tblVisits.START_DATE"
    Expression ="tblResults.RESULT_TEXT"
    Expression ="tblResults.DETECTION_CONDITION"
    Expression ="tblResults.LAB_REMARKS"
    Expression ="tblResults.RESULT_COMMENT"
    Expression ="tblVisits.VISIT_COMMENT"
    Expression ="tblLocations.[Station Name]"
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
    Expression ="tblProjects.ProjectID"
    Flag =0
    Expression ="tblLocations.StationID"
    Flag =0
    Expression ="tblCharacteristics.DISPLAY_NAME"
    Flag =0
    Expression ="CDbl([Result_Text])"
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
        dbInteger "ColumnWidth" ="2445"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="NumericValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblResults.VALUE_STATUS"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-16
    Top =0
    Right =1323
    Bottom =748
    Left =-1
    Top =-1
    Right =1307
    Bottom =32
    Left =6
    Top =0
    ColumnsShown =539
    Begin
        Left =42
        Top =12
        Right =186
        Bottom =156
        Top =0
        Name ="tblProjects"
        Name =""
    End
    Begin
        Left =234
        Top =12
        Right =378
        Bottom =156
        Top =0
        Name ="tblLocations"
        Name =""
    End
    Begin
        Left =732
        Top =14
        Right =876
        Bottom =158
        Top =0
        Name ="tblActivities"
        Name =""
    End
    Begin
        Left =472
        Top =14
        Right =662
        Bottom =158
        Top =0
        Name ="tblVisits"
        Name =""
    End
    Begin
        Left =939
        Top =13
        Right =1098
        Bottom =155
        Top =0
        Name ="tblResults"
        Name =""
    End
    Begin
        Left =1154
        Top =13
        Right =1356
        Bottom =157
        Top =0
        Name ="tblCharacteristics"
        Name =""
    End
End
