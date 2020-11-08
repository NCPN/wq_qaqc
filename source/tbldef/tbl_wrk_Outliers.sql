CREATE TABLE [tbl_wrk_Outliers] (
  [Outllier_Key] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [ProjectID] VARCHAR (50),
  [StationID] VARCHAR (50),
  [StationName] VARCHAR (255),
  [Start_Date] DATETIME ,
  [CharacteristicName] VARCHAR (50),
  [DetectionCondition] VARCHAR (255),
  [ResultValue] VARCHAR (50),
  [RemarkCode] VARCHAR (50),
  [ResultComment] LONGTEXT ,
  [VisitComment] LONGTEXT ,
  [Cutoff_5] DOUBLE ,
  [Cutoff_95] DOUBLE ,
  [Sample_Size] SHORT 
)
