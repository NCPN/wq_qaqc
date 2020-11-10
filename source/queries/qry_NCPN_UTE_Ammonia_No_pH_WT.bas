dbMemo "SQL" ="SELECT qry_NCPN_UTE_Ammonia.ProjectID, qry_NCPN_UTE_Ammonia.StationID, qry_NCPN_"
    "UTE_Ammonia.StationName, qry_NCPN_UTE_Ammonia.START_DATE, qry_NCPN_UTE_Ammonia.R"
    "ESULT_TEXT, qry_NCPN_UTE_Ammonia.VALUE_STATUS, qry_NCPN_UTE_Ammonia.LOWER_QUANT_"
    "LIMIT\015\012FROM qry_NCPN_UTE_Ammonia LEFT JOIN qry_NCPN_UTE_pH_WT ON (qry_NCPN"
    "_UTE_Ammonia.START_DATE = qry_NCPN_UTE_pH_WT.START_DATE) AND (qry_NCPN_UTE_Ammon"
    "ia.StationID = qry_NCPN_UTE_pH_WT.StationID) AND (qry_NCPN_UTE_Ammonia.ProjectID"
    " = qry_NCPN_UTE_pH_WT.ProjectID)\015\012WHERE (((qry_NCPN_UTE_pH_WT.ProjectID) I"
    "s Null))\015\012ORDER BY qry_NCPN_UTE_Ammonia.StationID, qry_NCPN_UTE_Ammonia.ST"
    "ART_DATE;\015\012"
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
        dbText "Name" ="qry_NCPN_UTE_Ammonia.LOWER_QUANT_LIMIT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_NCPN_UTE_Ammonia.RESULT_TEXT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_NCPN_UTE_Ammonia.StationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_NCPN_UTE_Ammonia.START_DATE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_NCPN_UTE_Ammonia.ProjectID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_NCPN_UTE_Ammonia.VALUE_STATUS"
        dbLong "AggregateType" ="-1"
    End
End
