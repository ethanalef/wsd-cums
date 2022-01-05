if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Loanrec]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Loanrec]
GO

CREATE TABLE [dbo].[Loanrec] (
	[memno] [int] NULL ,
	[lnnum] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[lndate] [smalldatetime] NULL ,
	[reqamt] [money] NULL ,
	[appamt] [money] NULL ,
	[install] [int] NULL ,
	[monthrepay] [money] NULL ,
	[cleardate] [smalldatetime] NULL ,
	[repaystat] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[bal] [money] NULL 
	
) ON [PRIMARY]
GO

