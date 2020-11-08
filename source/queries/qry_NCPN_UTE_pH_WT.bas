Operation =1
Option =0
Begin InputTables
    Name ="qry_NCPN_UTE_Water_Temp"
    Name ="qry_NCPN_UTE_pH"
End
Begin OutputColumns
    Expression ="qry_NCPN_UTE_Water_Temp.ProjectID"
    Expression ="qry_NCPN_UTE_Water_Temp.StationID"
    Expression ="qry_NCPN_UTE_Water_Temp.[Station Name]"
    Expression ="qry_NCPN_UTE_Water_Temp.START_DATE"
    Alias ="pH_Result"
    Expression ="qry_NCPN_UTE_pH.RESULT_TEXT"
    Alias ="pH_Status"
    Expression ="qry_NCPN_UTE_pH.VALUE_STATUS"
    Alias ="Water_Temp_Result"
    Expression ="qry_NCPN_UTE_Water_Temp.RESULT_TEXT"
    Alias ="Water_Temp_Status"
    Expression ="qry_NCPN_UTE_Water_Temp.VALUE_STATUS"
End
Begin Joins
    LeftTable ="qry_NCPN_UTE_Water_Temp"
    RightTable ="qry_NCPN_UTE_pH"
    Expression ="qry_NCPN_UTE_Water_Temp.ProjectID = qry_NCPN_UTE_pH.ProjectID"
    Flag =1
    LeftTable ="qry_NCPN_UTE_Water_Temp"
    RightTable ="qry_NCPN_UTE_pH"
    Expression ="qry_NCPN_UTE_Water_Temp.StationID = qry_NCPN_UTE_pH.StationID"
    Flag =1
    LeftTable ="qry_NCPN_UTE_Water_Temp"
    RightTable ="qry_NCPN_UTE_pH"
    Expression ="qry_NCPN_UTE_Water_Temp.START_DATE = qry_NCPN_UTE_pH.START_DATE"
    Flag =1
End
Begin OrderBy
    Expression ="qry_NCPN_UTE_Water_Temp.StationID"
    Flag =0
    Expression ="qry_NCPN_UTE_Water_Temp.START_DATE"
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
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="qry_NCPN_UTE_Water_Temp.ProjectID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_NCPN_UTE_Water_Temp.StationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_NCPN_UTE_Water_Temp.[Station Name]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_NCPN_UTE_Water_Temp.START_DATE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Water_Temp_Result"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Water_Temp_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pH_Result"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pH_Status"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =3
    Top =17
    Right =1184
    Bottom =735
    Left =-1
    Top =-1
    Right =1149
    Bottom =350
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qry_NCPN_UTE_Water_Temp"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="qry_NCPN_UTE_pH"
        Name =""
    End
End
