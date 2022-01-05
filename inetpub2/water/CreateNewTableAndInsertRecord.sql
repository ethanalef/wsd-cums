USE [wsdscu]
GO
/****** 物件:  Table [dbo].[userRights]    指令碼日期: 06/20/2007 18:20:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[userRights](
	[PID] [bigint] IDENTITY(1,1) NOT NULL,
	[Member1] [bit] NULL CONSTRAINT [DF_userRights_A1]  DEFAULT (0),
	[Member2] [bit] NULL CONSTRAINT [DF_Table_1_A11]  DEFAULT (0),
	[Member3] [bit] NULL CONSTRAINT [DF_Table_1_A12]  DEFAULT (0),
	[Member4] [bit] NULL CONSTRAINT [DF_Table_1_A13]  DEFAULT (0),
	[Loan1] [bit] NULL CONSTRAINT [DF_Table_1_A11_1]  DEFAULT (0),
	[Loan2] [bit] NULL CONSTRAINT [DF_Table_1_A21]  DEFAULT (0),
	[Loan3] [bit] NULL CONSTRAINT [DF_Table_1_A31]  DEFAULT (0),
	[Loan4] [bit] NULL CONSTRAINT [DF_Table_1_A41]  DEFAULT (0),
	[Loan5] [bit] NULL CONSTRAINT [DF_Table_1_B41]  DEFAULT (0),
	[Loan6] [bit] NULL CONSTRAINT [DF_Table_1_B51]  DEFAULT (0),
	[Loan7] [bit] NULL CONSTRAINT [DF_Table_1_B52]  DEFAULT (0),
	[Loan8] [bit] NULL CONSTRAINT [DF_Table_1_B53]  DEFAULT (0),
	[Loan9] [bit] NULL CONSTRAINT [DF_Table_1_B54]  DEFAULT (0),
	[cLoan1] [bit] NULL CONSTRAINT [DF_Table_1_B11]  DEFAULT (0),
	[cLoan2] [bit] NULL CONSTRAINT [DF_Table_1_B21]  DEFAULT (0),
	[cLoan3] [bit] NULL CONSTRAINT [DF_Table_1_B31]  DEFAULT (0),
	[AutoPay1] [bit] NULL CONSTRAINT [DF_Table_1_B11_1]  DEFAULT (0),
	[AutoPay2] [bit] NULL CONSTRAINT [DF_Table_1_B21_1]  DEFAULT (0),
	[AutoPay3] [bit] NULL CONSTRAINT [DF_Table_1_B31_1]  DEFAULT (0),
	[AutoPay4] [bit] NULL CONSTRAINT [DF_Table_1_B41_1]  DEFAULT (0),
	[AutoPay5] [bit] NULL CONSTRAINT [DF_Table_1_B51_1]  DEFAULT (0),
	[AutoPay6] [bit] NULL CONSTRAINT [DF_Table_1_B61]  DEFAULT (0),
	[AutoPay7] [bit] NULL CONSTRAINT [DF_Table_1_B71]  DEFAULT (0),
	[AutoPay8] [bit] NULL CONSTRAINT [DF_Table_1_B81]  DEFAULT (0),
	[AutoPay9] [bit] NULL CONSTRAINT [DF_Table_1_B91]  DEFAULT (0),
	[AutoPay11] [bit] NULL CONSTRAINT [DF_Table_1_D91]  DEFAULT (0),
	[AutoPay12] [bit] NULL CONSTRAINT [DF_Table_1_D92]  DEFAULT (0),
	[Saving1] [bit] NULL CONSTRAINT [DF_Table_1_D121]  DEFAULT (0),
	[Saving2] [bit] NULL CONSTRAINT [DF_Table_1_E18]  DEFAULT (0),
	[Saving3] [bit] NULL CONSTRAINT [DF_Table_1_E17]  DEFAULT (0),
	[Saving4] [bit] NULL CONSTRAINT [DF_Table_1_E16]  DEFAULT (0),
	[Saving5] [bit] NULL CONSTRAINT [DF_Table_1_E15]  DEFAULT (0),
	[Saving6] [bit] NULL CONSTRAINT [DF_Table_1_E14]  DEFAULT (0),
	[Saving7] [bit] NULL CONSTRAINT [DF_Table_1_E13]  DEFAULT (0),
	[Saving8] [bit] NULL CONSTRAINT [DF_Table_1_E12]  DEFAULT (0),
	[Saving9] [bit] NULL CONSTRAINT [DF_Table_1_E11]  DEFAULT (0),
	[MemAcct1] [bit] NULL CONSTRAINT [DF_userRights_E91]  DEFAULT (0),
	[Reporting1] [bit] NULL CONSTRAINT [DF_userRights_F11]  DEFAULT (0),
	[Reporting2] [bit] NULL CONSTRAINT [DF_userRights_G11]  DEFAULT (0),
	[Reporting3] [bit] NULL CONSTRAINT [DF_userRights_G17]  DEFAULT (0),
	[Reporting4] [bit] NULL CONSTRAINT [DF_userRights_G16]  DEFAULT (0),
	[Reporting5] [bit] NULL CONSTRAINT [DF_userRights_G15]  DEFAULT (0),
	[Reporting6] [bit] NULL CONSTRAINT [DF_userRights_G14]  DEFAULT (0),
	[Reporting7] [bit] NULL CONSTRAINT [DF_userRights_G13]  DEFAULT (0),
	[Reporting8] [bit] NULL CONSTRAINT [DF_userRights_G12]  DEFAULT (0),
	[Reporting9] [bit] NULL CONSTRAINT [DF_userRights_G81]  DEFAULT (0),
	[Reporting10] [bit] NULL CONSTRAINT [DF_userRights_G92]  DEFAULT (0),
	[Reporting11] [bit] NULL CONSTRAINT [DF_userRights_G91]  DEFAULT (0),
	[Other4] [bit] NULL CONSTRAINT [DF_userRights_Other4]  DEFAULT (0),
	[Other3] [bit] NULL CONSTRAINT [DF_userRights_Other3]  DEFAULT (0),
	[Other2] [bit] NULL CONSTRAINT [DF_userRights_Other2]  DEFAULT (0),
	[Other1] [bit] NULL CONSTRAINT [DF_userRights_Other1]  DEFAULT (0),
	[User_Fk] [bigint] NULL,
 CONSTRAINT [PK_userRights] PRIMARY KEY CLUSTERED 
(
	[PID] ASC
) ON [PRIMARY]
) ON [PRIMARY]



USE [wsdscu]
GO
/****** 物件:  Table [dbo].[userRights]    指令碼日期: 06/20/2007 18:20:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
INSERT INTO [dbo].[userRights] (User_FK) Values (1)
INSERT INTO [dbo].[userRights] (User_FK) Values (2)
INSERT INTO [dbo].[userRights] (User_FK) Values (3)
INSERT INTO [dbo].[userRights] (User_FK) Values (4)
INSERT INTO [dbo].[userRights] (User_FK) Values (5)
INSERT INTO [dbo].[userRights] (User_FK) Values (6)

