Operation =1
Option =0
Begin InputTables
    Name ="tbl_wrk_Outliers"
End
Begin OutputColumns
    Expression ="tbl_wrk_Outliers.ProjectID"
    Expression ="tbl_wrk_Outliers.StationID"
    Expression ="tbl_wrk_Outliers.StationName"
    Expression ="tbl_wrk_Outliers.Start_Date"
    Expression ="tbl_wrk_Outliers.CharacteristicName"
    Expression ="tbl_wrk_Outliers.DetectionCondition"
    Expression ="tbl_wrk_Outliers.ResultValue"
    Expression ="tbl_wrk_Outliers.RemarkCode"
    Expression ="tbl_wrk_Outliers.ResultComment"
    Expression ="tbl_wrk_Outliers.VisitComment"
    Expression ="tbl_wrk_Outliers.Cutoff_5"
    Expression ="tbl_wrk_Outliers.Cutoff_95"
    Expression ="tbl_wrk_Outliers.Sample_Size"
End
Begin OrderBy
    Expression ="tbl_wrk_Outliers.ProjectID"
    Flag =0
    Expression ="tbl_wrk_Outliers.StationID"
    Flag =0
    Expression ="tbl_wrk_Outliers.Start_Date"
    Flag =0
    Expression ="tbl_wrk_Outliers.CharacteristicName"
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
        dbText "Name" ="tbl_wrk_Outliers.VisitComment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_wrk_Outliers.ProjectID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_wrk_Outliers.StationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_wrk_Outliers.Start_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_wrk_Outliers.CharacteristicName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_wrk_Outliers.StationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_wrk_Outliers.DetectionCondition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_wrk_Outliers.ResultValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_wrk_Outliers.RemarkCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_wrk_Outliers.ResultComment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_wrk_Outliers.Sample_Size"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_wrk_Outliers.Cutoff_5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_wrk_Outliers.Cutoff_95"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =42
    Top =157
    Right =1155
    Bottom =670
    Left =-1
    Top =-1
    Right =1081
    Bottom =203
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_wrk_Outliers"
        Name =""
    End
End
