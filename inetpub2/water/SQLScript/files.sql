if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Autotran]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Autotran]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Loanrec]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Loanrec]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MemMaster]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MemMaster]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[loan]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[loan]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[loanApp]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[loanApp]
GO

CREATE TABLE [dbo].[Autotran] (
	[Lnnum] [char] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memno] [int] NULL ,
	[autodate] [smalldatetime] NULL ,
	[trefno] [char] (2) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[amount] [money] NULL 
) ON [PRIMARY]
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
	[autopamt] [money] NULL ,
	[salarydeduct] [money] NULL ,
	[lnflag] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[chequeamt] [money] NULL ,
	[oldlnnum] [char] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MemMaster] (
	[memNo] [int] NULL ,
	[memname] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memCName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memAddr1] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memAddr2] [nvarchar] (255) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memAddr3] [nvarchar] (255) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memContactTel] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memOfficeTel] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memrank] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[mempos] [nchar] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memMobile] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memEmail] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memHKID] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memGender] [nvarchar] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memBday] [smalldatetime] NULL ,
	[memGrade] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memSection] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[cuDesignation] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memG1] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memG2] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memG3] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memG4O1] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memG4O2] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memG4O3] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[treasRefNo] [nvarchar] (8) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[employCond] [nvarchar] (8) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[firstAppointDate] [smalldatetime] NULL ,
	[memDate] [smalldatetime] NULL ,
	[EdComm] [int] NULL ,
	[Wdate] [datetime] NULL ,
	[B1] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[B2] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[B1ID] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[B2ID] [nvarchar] (20) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[B1relation] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[B2relation] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[B1Add] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[B2Add] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[remark] [nvarchar] (100) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[creditStatus] [nvarchar] (2) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memPhoto] [image] NULL ,
	[monthsave] [money] NULL ,
	[monthssave] [money] NULL ,
	[tpayamt] [money] NULL ,
	[loanreplay] [money] NULL ,
	[calcInst] [bit] NULL ,
	[Osinst] [money] NULL ,
	[inst] [money] NULL ,
	[LeagueDue] [bit] NULL ,
	[bnk] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[bch] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[bacct] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[bnklmt] [money] NULL ,
	[bnkamt] [money] NULL ,
	[tryamt] [money] NULL ,
	[personalEnt] [nvarchar] (30) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[share_tx_date] [smalldatetime] NULL ,
	[Share_amt_last] [money] NULL ,
	[loan_tx_date] [smalldatetime] NULL ,
	[loan_amt_last] [money] NULL ,
	[loanamt] [money] NULL ,
	[loan_app_date] [smalldatetime] NULL ,
	[term] [int] NULL ,
	[repaylst] [smalldatetime] NULL ,
	[payby] [nvarchar] (4) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[loanbeg] [smalldatetime] NULL ,
	[loanend] [smalldatetime] NULL ,
	[overdue] [money] NULL ,
	[ttlshare] [money] NULL ,
	[ttllastshare] [money] NULL ,
	[dividend] [money] NULL ,
	[status] [nvarchar] (4) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[mstatus] [nvarchar] (4) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[deleted] [bit] NOT NULL ,
	[dtupdat] [smalldatetime] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[loan] (
	[memno] [int] NULL ,
	[lnnum] [char] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[ldate] [smalldatetime] NULL ,
	[code] [char] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[amount] [money] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[loanApp] (
	[uid] [int] NULL ,
	[loanType] [nvarchar] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[memNo] [int] NULL ,
	[memName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[appDate] [smalldatetime] NULL ,
	[memGrade] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[age] [int] NULL ,
	[firstAppointDate] [smalldatetime] NULL ,
	[employCond] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[netSalary] [money] NULL ,
	[loanAmt] [money] NULL ,
	[installment] [int] NULL ,
	[chequeDate] [smalldatetime] NULL ,
	[guarantorID] [int] NULL ,
	[guarantorName] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[guarantorGrade] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
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
	[interest] [bit] NOT NULL ,
	[otherReason1] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[otherReason2] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[remarks] [ntext] COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[deleted] [bit] NOT NULL ,
	[savettl] [money] NULL ,
	[guarantor2ID] [int] NULL ,
	[guarantor2Name] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[guarantor2Grade] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[guarantor2Salary] [money] NULL ,
	[interview2Date] [smalldatetime] NULL ,
	[interview2Detail] [ntext] COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[guarantor3ID] [int] NULL ,
	[guarantor3Name] [nvarchar] (50) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[guarantor3Grade] [nvarchar] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[guarantor3Salary] [money] NULL ,
	[interview3Date] [smalldatetime] NULL ,
	[interview3Detail] [ntext] COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[bnklny] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[bnklnn] [char] (1) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL ,
	[bnklnamt] [money] NULL ,
	[chequeamt] [money] NULL ,
	[oldlnnum] [char] (10) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

