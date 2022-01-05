if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[guarantor]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[guarantor]
GO

CREATE TABLE [dbo].[guarantor] (
	[memno] [int] NULL ,
	[lnnum] [char] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[date] [smalldatetime] NULL ,
	[guarantorID] [int] NULL ,
	[guarantorName] [nvarchar] (35) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[guarantorCname] [nvarchar] (35) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL 
) ON [PRIMARY]
GO

