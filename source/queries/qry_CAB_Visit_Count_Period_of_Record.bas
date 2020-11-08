dbMemo "SQL" ="SELECT Count(*) AS Visit_Count\015\012FROM (SELECT DISTINCT [qry_CAB]![StationID"
    "] & \"|\" & [qry_CAB]![START_DATE] AS Visit_Count FROM qry_CAB WHERE ProjectID ="
    " [Enter ProjectID:])  AS Station_Visits;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Visit_Count"
        dbLong "AggregateType" ="-1"
    End
End
