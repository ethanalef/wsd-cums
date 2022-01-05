if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sadttran]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[sadttran]
GO

CREATE TABLE [dbo].[sadttran] (
	[memno] [int] NULL ,
	[memName] [varbinary] (35) NULL ,
	[memcName] [varbinary] (10) NULL ,
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

