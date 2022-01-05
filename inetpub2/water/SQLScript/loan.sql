if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[loan]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[loan]
GO

CREATE TABLE [dbo].[loan] (
	[memno] [int] NULL ,
	[lnnum] [char] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[ldate] [smalldatetime] NULL ,
	[code] [char] (2) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[amount] [money] NULL ,
	[newlnnum] [char] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[pflag] [bit] NULL ,
	[bal] [money] NULL ,
	[pdate] [smalldatetime] NULL 
) ON [PRIMARY]
GO

