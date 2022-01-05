if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[bank]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[bank]
GO

CREATE TABLE [dbo].[bank] (
	[BNCODE] [nvarchar] (3) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[BANK] [nvarchar] (47) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL 
) ON [PRIMARY]
GO

