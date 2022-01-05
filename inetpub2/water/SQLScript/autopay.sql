if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[autopay]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[autopay]
GO

CREATE TABLE [dbo].[autopay] (
	[memno] [int] NULL ,
	[adate] [smalldatetime] NULL ,
	[lnnum] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[code] [char] (2) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[bankin] [money] NULL ,
	[curamt] [money] NULL ,
	[updamt] [money] NULL ,
	[status] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[flag] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[pdate] [smalldatetime] NULL ,
	[delyflag] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[deleted] [bit] NULL ,
	[mstatus] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[pflag] [bit] NULL 
) ON [PRIMARY]
GO

