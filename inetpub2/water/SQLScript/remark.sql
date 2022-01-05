if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[remark]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[remark]
GO

CREATE TABLE [dbo].[remark] (
	[term] [char] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[remak] [ntext] COLLATE Chinese_Taiwan_Stroke_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

