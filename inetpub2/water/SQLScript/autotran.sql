if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Autotran]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Autotran]
GO

CREATE TABLE [dbo].[Autotran] (
	[Lnnum] [char] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memno] [int] NULL ,
	[memName] [varbinary] (35)  COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memcName] [varbinary] (10)  COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[autodate] [smalldatetime] NULL ,
	[Bankin] [money] NULL ,
	[TTlsum] [money] NULL ,
	[interest] [money] NULL ,
	[pamtin] [money] NULL ,
	[pamt] [money] NULL ,
	[samtin] [money] NULL ,
	[samt] [money] NULL 
) ON [PRIMARY]
GO

