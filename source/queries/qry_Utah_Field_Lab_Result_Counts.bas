Operation =1
Option =0
Begin InputTables
    Name ="qry_Utah_Field_Results"
    Name ="qry_Utah_Lab_Results"
End
Begin OutputColumns
    Expression ="qry_Utah_Field_Results.ProjectID"
    Expression ="qry_Utah_Field_Results.StationID"
    Expression ="qry_Utah_Field_Results.[Station Name]"
    Expression ="qry_Utah_Field_Results.START_DATE"
    Alias ="Field_Result_Count"
    Expression ="qry_Utah_Field_Results.Result_Count"
    Alias ="Lab_Result_Count"
    Expression ="qry_Utah_Lab_Results.Result_Count"
End
Begin Joins
    LeftTable ="qry_Utah_Field_Results"
    RightTable ="qry_Utah_Lab_Results"
    Expression ="qry_Utah_Field_Results.ProjectID = qry_Utah_Lab_Results.ProjectID"
    Flag =2
    LeftTable ="qry_Utah_Field_Results"
    RightTable ="qry_Utah_Lab_Results"
    Expression ="qry_Utah_Field_Results.StationID = qry_Utah_Lab_Results.StationID"
    Flag =2
    LeftTable ="qry_Utah_Field_Results"
    RightTable ="qry_Utah_Lab_Results"
    Expression ="qry_Utah_Field_Results.START_DATE = qry_Utah_Lab_Results.START_DATE"
    Flag =2
End
Begin OrderBy
    Expression ="qry_Utah_Field_Results.ProjectID"
    Flag =0
    Expression ="qry_Utah_Field_Results.StationID"
    Flag =0
    Expression ="qry_Utah_Field_Results.START_DATE"
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
        dbText "Name" ="qry_Utah_Field_Results.ProjectID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Utah_Field_Results.StationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Utah_Field_Results.[Station Name]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Utah_Field_Results.START_DATE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Field_Result_Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lab_Result_Count"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =8
    Top =334
    Right =1189
    Bottom =767
    Left =-1
    Top =-1
    Right =1149
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qry_Utah_Field_Results"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="qry_Utah_Lab_Results"
        Name =""
    End
End
