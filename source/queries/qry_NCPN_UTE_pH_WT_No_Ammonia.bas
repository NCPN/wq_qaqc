dbMemo "SQL" ="SELECT qry_NCPN_UTE_pH_WT.ProjectID, qry_NCPN_UTE_pH_WT.StationID, qry_NCPN_UTE_"
    "pH_WT.StationName, qry_NCPN_UTE_pH_WT.START_DATE, qry_NCPN_UTE_pH_WT.pH_Result, "
    "qry_NCPN_UTE_pH_WT.pH_Status, qry_NCPN_UTE_pH_WT.Water_Temp_Result, qry_NCPN_UTE"
    "_pH_WT.Water_Temp_Status\015\012FROM qry_NCPN_UTE_pH_WT LEFT JOIN qry_NCPN_UTE_A"
    "mmonia ON (qry_NCPN_UTE_pH_WT.ProjectID = qry_NCPN_UTE_Ammonia.ProjectID) AND (q"
    "ry_NCPN_UTE_pH_WT.StationID = qry_NCPN_UTE_Ammonia.StationID) AND (qry_NCPN_UTE_"
    "pH_WT.START_DATE = qry_NCPN_UTE_Ammonia.START_DATE)\015\012WHERE (((qry_NCPN_UTE"
    "_Ammonia.ProjectID) Is Null))\015\012ORDER BY qry_NCPN_UTE_pH_WT.StationID, qry_"
    "NCPN_UTE_pH_WT.START_DATE;\015\012"
dbMemo "Connect" =""
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
