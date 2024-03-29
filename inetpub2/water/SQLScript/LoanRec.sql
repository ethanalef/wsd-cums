if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Loanrec]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Loanrec]
GO

CREATE TABLE [dbo].[Loanrec] (
	[memno] [int] NULL ,
	[lnnum] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NOT NULL ,
	[lndate] [smalldatetime] NULL ,
	[appamt] [money] NULL ,
	[install] [int] NULL ,
	[monthrepay] [money] NULL ,
	[cleardate] [smalldatetime] NULL ,
	[repaystat] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[bal] [money] NULL ,
	[lnflag] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[chequeamt] [money] NULL ,
	[oldlnnum] [char] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[calflag] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[delyflag] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[months] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[chkmon] [int] NULL ,
	[delydate] [smalldatetime] NULL ,
	[applyamt] [int] NULL ,
	[loantype] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL 
) ON [PRIMARY]
GO

