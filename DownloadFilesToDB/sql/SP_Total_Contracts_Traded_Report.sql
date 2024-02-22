-- ================================================
-- Template generated from Template Explorer using:
-- Create Procedure (New Menu).SQL
--
-- Use the Specify Values for Template Parameters 
-- command (Ctrl-Shift-M) to fill in the parameter 
-- values below.
--
-- This block of comments will not be included in
-- the definition of the procedure.
-- ================================================
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		t zeeman
-- Create date: <Create Date,,>
-- Description: accepts Date-From and Date-To inputs 
-- returns a table with the headers [File Date], [Contract], [ContractsTraded], [% Of Total Contracts Traded].
-- =============================================
CREATE PROCEDURE SP_Total_Contracts_Traded_Report
	-- Add the parameters for the stored procedure here
@DateFrom DateTime, 
@DateTo DateTime
AS
BEGIN
	
	SET NOCOUNT ON;
	declare @totalcontracts float 
	set  @totalcontracts = (select  sum(contractstraded) from DailyMTM)
	
	declare @returnTable table(FileDate date, [Contract] nvarchar(50), [ContractsTraded] float, PercTotalContractsTraded decimal(10,2))

    -- Insert statements for procedure here
	insert @returnTable select FileDate, contract,sum(ContractsTraded) as totalcontracts, sum(ContractsTraded)/@totalcontracts *100 from DailyMTM
	group by FileDate, [Contract]
	having sum(ContractsTraded) > 0
		and FileDate < @DateTo 
		and FileDate > @DateFrom
	select * from @returnTable
END
GO
