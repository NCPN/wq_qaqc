Operation =1
Option =0
Where ="(((qry_NCPN_UTE_Ammonia.ProjectID) Is Null))"
Begin InputTables
    Name ="qry_NCPN_UTE_pH_WT"
    Name ="qry_NCPN_UTE_Ammonia"
End
Begin OutputColumns
    Expression ="qry_NCPN_UTE_pH_WT.ProjectID"
    Expression ="qry_NCPN_UTE_pH_WT.StationID"
    Expression ="qry_NCPN_UTE_pH_WT.[Station Name]"
    Expression ="qry_NCPN_UTE_pH_WT.START_DATE"
    Expression ="qry_NCPN_UTE_pH_WT.pH_Result"
    Expression ="qry_NCPN_UTE_pH_WT.pH_Status"
    Expression ="qry_NCPN_UTE_pH_WT.Water_Temp_Result"
    Expression ="qry_NCPN_UTE_pH_WT.Water_Temp_Status"
End
Begin Joins
    LeftTable ="qry_NCPN_UTE_pH_WT"
    RightTable ="qry_NCPN_UTE_Ammonia"
    Expression ="qry_NCPN_UTE_pH_WT.ProjectID = qry_NCPN_UTE_Ammonia.ProjectID"
    Flag =2
    LeftTable ="qry_NCPN_UTE_pH_WT"
    RightTable ="qry_NCPN_UTE_Ammonia"
    Expression ="qry_NCPN_UTE_pH_WT.StationID = qry_NCPN_UTE_Ammonia.StationID"
    Flag =2
    LeftTable ="qry_NCPN_UTE_pH_WT"
    RightTable ="qry_NCPN_UTE_Ammonia"
    Expression ="qry_NCPN_UTE_pH_WT.START_DATE = qry_NCPN_UTE_Ammonia.START_DATE"
    Flag =2
End
Begin OrderBy
    Expression ="qry_NCPN_UTE_pH_WT.StationID"
    Flag =0
    Expression ="qry_NCPN_UTE_pH_WT.START_DATE"
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
        dbText "Name" ="qry_NCPN_UTE_pH_WT.Water_Temp_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_NCPN_UTE_pH_WT.StationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_NCPN_UTE_pH_WT.START_DATE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_NCPN_UTE_pH_WT.ProjectID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_NCPN_UTE_pH_WT.[Station Name]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_NCPN_UTE_pH_WT.pH_Result"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_NCPN_UTE_pH_WT.pH_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_NCPN_UTE_pH_WT.Water_Temp_Result"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1196
    Bottom =758
    Left =-1
    Top =-1
    Right =1164
    Bottom =401
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =203
        Bottom =156
        Top =0
        Name ="qry_NCPN_UTE_pH_WT"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =411
        Bottom =156
        Top =0
        Name ="qry_NCPN_UTE_Ammonia"
        Name =""
    End
End
