if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[share]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[share]
GO

CREATE TABLE [dbo].[share] (
	[memno] [int] NULL ,
	[Ldate] [smalldatetime] NULL ,
	[code] [char] (2) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[amount] [money] NULL ,
	[pflag] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[bal] [money] NULL ,
	[lnflag] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[sdesc] [char] (100) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[pdate] [smalldatetime] NULL 
) ON [PRIMARY]
GO

