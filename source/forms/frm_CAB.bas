Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =9
    ItemSuffix =41
    Left =6510
    Top =2190
    Right =13710
    Bottom =6495
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1385341e7574e340
    End
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =4320
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =4020
                    Top =2340
                    Width =1739
                    Height =300
                    Name ="Button_Close"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =2340
                    LayoutCachedWidth =5759
                    LayoutCachedHeight =2640
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1230
                    Top =180
                    Width =4785
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label6"
                    Caption ="Cation-anion Balance"
                    LayoutCachedLeft =1230
                    LayoutCachedTop =180
                    LayoutCachedWidth =6015
                    LayoutCachedHeight =600
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1440
                    Top =2340
                    Width =1740
                    Height =300
                    TabIndex =1
                    Name ="ButtonCanopyGap"
                    Caption ="CAB Report"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1440
                    LayoutCachedTop =2340
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =2640
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4500
                    Top =1320
                    Width =1200
                    Height =239
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="StartDate"
                    Format ="Short Date"
                    InputMask ="99/99/0000;0;_"
                    GridlineColor =10921638

                    LayoutCachedLeft =4500
                    LayoutCachedTop =1320
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =1559
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1260
                            Top =1320
                            Width =3120
                            Height =239
                            FontWeight =700
                            Name ="Label38"
                            Caption ="Enter a date range if desired - From"
                            LayoutCachedLeft =1260
                            LayoutCachedTop =1320
                            LayoutCachedWidth =4380
                            LayoutCachedHeight =1559
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4500
                    Top =1800
                    Width =1200
                    Height =239
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="EndDate"
                    Format ="Short Date"
                    InputMask ="99/99/0000;0;_"
                    GridlineColor =10921638

                    LayoutCachedLeft =4500
                    LayoutCachedTop =1800
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =2039
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4020
                            Top =1800
                            Width =360
                            Height =239
                            FontWeight =700
                            Name ="Label40"
                            Caption ="To"
                            LayoutCachedLeft =4020
                            LayoutCachedTop =1800
                            LayoutCachedWidth =4380
                            LayoutCachedHeight =2039
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =720
                    Left =4500
                    Top =840
                    Width =1200
                    Height =239
                    TabIndex =4
                    Name ="ProjectID"
                    RowSourceType ="Value List"
                    RowSource ="\"NCPN_UTE\";\"NCPN_UTM\""
                    ColumnWidths ="720"

                    LayoutCachedLeft =4500
                    LayoutCachedTop =840
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =1079
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1260
                            Top =840
                            Width =2760
                            Height =239
                            FontWeight =700
                            Name ="ProjectID_Label"
                            Caption ="Select a project ID if desired"
                            LayoutCachedLeft =1260
                            LayoutCachedTop =840
                            LayoutCachedWidth =4020
                            LayoutCachedHeight =1079
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Button_Close_Click()

    DoCmd.Close

End Sub

Private Sub ButtonCanopyGap_Click()
On Error GoTo Err_CanopyGap_Click

  Dim strSQL As String
  Dim db As DAO.Database
  Dim WorkOutput As DAO.Recordset
  Dim Obsvalues As DAO.Recordset
  Dim DateSave As String
  Dim ProjectSave As String
  Dim StationSave As String
  Dim NameSave As String
  Dim TotalCation As Double
  Dim TotalAnion As Double
  Dim CAB As Double
  Dim ArrayIndex As Integer
  Dim ConstArray(8, 4) As Variant ' Array for constituant values
  ' Parameter values array
  ' Column x,0 is constituant name
  ' Column x,1 is atomic weight
  ' Column x,2 is valence
  ' Column x,3 is measured value

  DoCmd.SetWarnings False
  DoCmd.OpenQuery "qry_Clear_CAB"
  DoCmd.SetWarnings True

'  Build SQL statement
  strSQL = "SELECT * FROM qry_CAB WHERE 0 = 0"
  If Not IsNull(Me!ProjectID) Then
    strSQL = strSQL & " AND ProjectID = '" & Me!ProjectID & "'"
  End If
  If Not IsNull(Me!StartDate) Then
    If IsNull(Me!EndDate) Then
      MsgBox "You must enter an end date.", vbOKOnly, "CAB Report"
      Exit Sub
    Else
      strSQL = strSQL & " AND [START_DATE] BETWEEN #" & Me!StartDate & "# AND #" & Me!EndDate & "#"
    End If
  End If
  ' strSQL = strSQL & "AND StationID = '9127000'"
  strSQL = strSQL & " ORDER BY StationID, START_DATE"
  ' MsgBox strSQL
  DoCmd.Hourglass True
  Set db = CurrentDb

   Set Obsvalues = db.OpenRecordset(strSQL)
   If Obsvalues.EOF Then
     MsgBox "No valid Result records found.", vbOKOnly, "Outliers Report"
     Obsvalues.Close
     Set Obsvalues = Nothing
     GoTo Exit_CanopyGap_Click
   End If
   ConstArray(0, 0) = "Calcium"
   ConstArray(0, 1) = 40.08
   ConstArray(0, 2) = 2
   ConstArray(1, 0) = "Magnesium"
   ConstArray(1, 1) = 24.31
   ConstArray(1, 2) = 2
   ConstArray(2, 0) = "Sodium"
   ConstArray(2, 1) = 23
   ConstArray(2, 2) = 1
   ConstArray(3, 0) = "Potassium"
   ConstArray(3, 1) = 39.1
   ConstArray(3, 2) = 1
   ConstArray(4, 0) = "Bicarbonate"
   ConstArray(4, 1) = 61
   ConstArray(4, 2) = 1
   ConstArray(5, 0) = "Sulfur, sulfate (SO4) as SO4"
   ConstArray(5, 1) = 96
   ConstArray(5, 2) = 2
   ConstArray(6, 0) = "Chloride"
   ConstArray(6, 1) = 35.5
   ConstArray(6, 2) = 1
   ConstArray(7, 0) = "Nitrogen, Nitrite (NO2) + Nitrate (NO3) as N"
   ConstArray(7, 1) = 62
   ConstArray(7, 2) = 1
   ArrayIndex = 0
   Do Until ArrayIndex > 7           ' Initialize array values
     ConstArray(ArrayIndex, 3) = 0
     ArrayIndex = ArrayIndex + 1
   Loop
   IDSave = Obsvalues!StationID    ' Save necessary fields
   ProjectSave = Obsvalues!ProjectID
   NameSave = Obsvalues!StationName
   DateSave = Obsvalues![START_DATE]
   ArrayIndex = 0
   Set WorkOutput = db.OpenRecordset("tbl_wrk_CAB")
   Do Until Obsvalues.EOF
     If IDSave <> Obsvalues!StationID Or DateSave <> Obsvalues![START_DATE] Then  ' New station?
       ' Calculate CAB
       TotalAnion = 0
       TotalCation = 0
       ArrayIndex = 0
       Do Until ArrayIndex > 3
         TotalCation = TotalCation + ((ConstArray(ArrayIndex, 3) / ConstArray(ArrayIndex, 1)) * ConstArray(ArrayIndex, 2))  ' calculate total cation
         ArrayIndex = ArrayIndex + 1
       Loop
       Do Until ArrayIndex > 7
         TotalAnion = TotalAnion + ((ConstArray(ArrayIndex, 3) / ConstArray(ArrayIndex, 1)) * ConstArray(ArrayIndex, 2))  ' calculate total cation
         ArrayIndex = ArrayIndex + 1
       Loop
       If (TotalCation + TotalAnion) <> 0 Then
         CAB = ((TotalCation - TotalAnion) / (TotalCation + TotalAnion)) * 100
       Else
         CAB = 0
       End If
       If CAB > 5 Then
         ' Write outlier record
         WorkOutput.AddNew
         WorkOutput!ProjectID = ProjectSave  '
         WorkOutput!StationID = IDSave  '
         WorkOutput!StationName = NameSave
         WorkOutput!START_DATE = DateSave  '
         WorkOutput!CAB = CAB  '
         WorkOutput.Update  ' Write outlier record
       End If
       ArrayIndex = 0
       Do Until ArrayIndex > 7           ' Clear array values
         ConstArray(ArrayIndex, 3) = 0
         ArrayIndex = ArrayIndex + 1
       Loop
       IDSave = Obsvalues!StationID    ' Save necessary fields
       NameSave = Obsvalues!StationName
       DateSave = Obsvalues!START_DATE
       ProjectSave = Obsvalues!ProjectID
     End If
     ArrayIndex = 0
     Do Until ArrayIndex > 7           ' Initialize array values
       If ConstArray(ArrayIndex, 0) = Obsvalues!DISPLAY_NAME Then
         ConstArray(ArrayIndex, 3) = Obsvalues!Concentration  ' Save observed concentration
         GoTo NextObs
       End If
       ArrayIndex = ArrayIndex + 1
     Loop
     MsgBox "Constituant name " & Obsvalues!StoretDisplayName & " not found in array.", vbOKOnly, "Loading parameter array."
     GoTo Exit_CanopyGap_Click
NextObs:
     Obsvalues.MoveNext
   Loop
       ' Do last visit
       TotalAnion = 0
       TotalCation = 0
       ArrayIndex = 0
       Do Until ArrayIndex > 3
         TotalCation = TotalCation + ((ConstArray(ArrayIndex, 3) / ConstArray(ArrayIndex, 1)) * ConstArray(ArrayIndex, 2))  ' calculate total cation
         ArrayIndex = ArrayIndex + 1
       Loop
       Do Until ArrayIndex > 7
         TotalAnion = TotalAnion + ((ConstArray(ArrayIndex, 3) / ConstArray(ArrayIndex, 1)) * ConstArray(ArrayIndex, 2))  ' calculate total cation
         ArrayIndex = ArrayIndex + 1
       Loop
       CAB = ((TotalCation - TotalAnion) / (TotalCation + TotalAnion)) * 100
       If CAB > 5 Then
         ' Write outlier record
         WorkOutput.AddNew
         WorkOutput!ProjectID = ProjectSave  '
         WorkOutput!StationID = IDSave  '
         WorkOutput!StationName = NameSave
         WorkOutput!START_DATE = DateSave  '
         WorkOutput!CAB = CAB  '
         WorkOutput.Update  ' Write outlier record
       End If
     WorkOutput.Close
     Set WorkOutput = Nothing
     Obsvalues.Close
     Set Obsvalues = Nothing
     
Exit_CanopyGap_Click:
    DoCmd.Hourglass False
    MsgBox "Finished - results are in tbl_wrk_CAB."
    Exit Sub

Err_CanopyGap_Click:
    MsgBox Err.Description
    Resume Exit_CanopyGap_Click
End Sub


Private Sub Form_Open(Cancel As Integer)
  DoCmd.Restore
End Sub
