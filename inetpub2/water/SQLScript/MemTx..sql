if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[memTx]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[memTx]
GO

CREATE TABLE [dbo].[memTx] (
	[memTxNo] [int] NOT NULL ,
	[memNo] [int] NULL ,
	[txDate] [smalldatetime] NULL ,
	[treNo] [nvarchar] (2) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[sharePaid] [money] NULL ,
	[shareWithdrawn] [money] NULL ,
	[amtLoan] [money] NULL ,
	[Lcode] [char] (2) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[monthlyRepaid] [money] NULL ,
	[interestPaid] [money] NULL ,
	[loanPaid] [money] NULL ,
	[txAmt] [money] NULL ,
	[deleted] [bit] NULL ,
	[LNNUM] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[Newlnnum] [char] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL 
) ON [PRIMARY]
GO

