-- =====================================================================
-- Script 04: Table-Valued Function - Categorized Transactions by Period
-- =====================================================================
-- Returns all transactions for a period with their category
-- Uses the fn_GetValueCategory function
-- =====================================================================

USE FinancialTransactionDB;
GO

IF OBJECT_ID('dbo.tvf_GetCategorizedTransactionsByPeriod', 'TF') IS NOT NULL
BEGIN
    DROP FUNCTION dbo.tvf_GetCategorizedTransactionsByPeriod;
    PRINT 'Function tvf_GetCategorizedTransactionsByPeriod dropped for recreation.';
END
GO

CREATE FUNCTION dbo.tvf_GetCategorizedTransactionsByPeriod(
    @StartDate DATETIME,
    @EndDate   DATETIME
)
RETURNS TABLE
WITH SCHEMABINDING
AS
RETURN
SELECT
    t.TransactionId,
    t.CardNumber,
    t.TransactionAmount,
    t.TransactionDate,
    t.Description,
    t.TransactionStatus,
    dbo.fn_GetValueCategory(t.TransactionAmount) AS Category
FROM dbo.Transactions t
WHERE t.TransactionDate BETWEEN @StartDate AND @EndDate;
GO

PRINT 'Table-Valued Function tvf_GetCategorizedTransactionsByPeriod created successfully.';
GO

-- =====================================================================
-- Usage Examples:
-- =====================================================================
-- SELECT * FROM dbo.tvf_GetCategorizedTransactionsByPeriod('2026-01-01', '2026-01-31');
-- SELECT * FROM dbo.tvf_GetCategorizedTransactionsByPeriod('2026-01-01', GETDATE())
-- WHERE Category = 'Premium'
-- ORDER BY TransactionAmount DESC;