if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[userLog]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[userLog]
GO

CREATE TABLE [dbo].[userLog] (
	[uid] [int] NOT NULL ,
	[username] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[userLevel] [int] NULL ,
	[actionDes] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[actionTime] [smalldatetime] NULL 
) ON [PRIMARY]
GO

