if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[loanApp]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[loanApp]
GO

CREATE TABLE [dbo].[loanApp] (
	[uid] [int] NOT NULL ,
	[loanType] [nvarchar] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memNo] [int] NULL ,
	[memName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[appDate] [smalldatetime] NULL ,
	[memGrade] [nvarchar] (4) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[age] [int] NULL ,
	[firstAppointDate] [smalldatetime] NULL ,
	[employCond] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[netSalary] [money] NULL ,
	[loanAmt] [money] NULL ,
	[installment] [int] NULL ,
	[monthrepay] [money] NULL ,
	[monthint] [money] NULL ,
	[status] [bit] NULL ,
	[chequeDate] [smalldatetime] NULL ,
	[guarantorID] [int] NULL ,
	[guarantorName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[guarantorGrade] [nvarchar] (4) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[guarantorSalary] [money] NULL ,
	[interviewDate] [smalldatetime] NULL ,
	[interviewDetail] [ntext] COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[firstApprovalDate] [smalldatetime] NULL ,
	[firstApproval] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[secondApprovalDate] [smalldatetime] NULL ,
	[secondApproval] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[rejectReason] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[loanPlanID] [int] NULL ,
	[SpecialPlanID] [int] NULL ,
	[interest] [bit] NULL ,
	[otherReason1] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[otherReason2] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[remarks] [ntext] COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[deleted] [bit] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

