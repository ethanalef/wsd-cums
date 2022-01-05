if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DIVIDEND]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DIVIDEND]
GO

CREATE TABLE [dbo].[DIVIDEND] (
	[MEMNO] [float] NULL ,
	[DIVIDEND] [float] NULL ,
	[BANK] [nvarchar] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[SHARE] [nvarchar] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[DELETED] [bit] NOT NULL 
) ON [PRIMARY]
GO

