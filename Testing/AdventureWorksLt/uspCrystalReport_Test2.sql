/****** Object:  StoredProcedure [dbo].[uspCrystalReport_Test2]    Script Date: 9/4/2020 11:45:06 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create procedure [dbo].[uspCrystalReport_Test2]

@Size nvarchar(5)

as

select * from SalesLT.Product where size = @Size;
GO

