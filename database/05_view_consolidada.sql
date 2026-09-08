-- =====================================================================
-- Script 05: Consolidated View for Financial Queries
-- =====================================================================
-- View that facilitates queries by adding category and date fields
-- =====================================================================

USE TransactionManagemer;
GO

IF OBJECT_ID('dbo.vw_ConsolidatedTransactions', 'V') IS NOT NULL
BEGIN
    DROP VIEW dbo.vw_ConsolidatedTransactions;
    PRINT 'View vw_ConsolidatedTransactions dropped for recreation.';
END
GO

CREATE VIEW dbo.vw_ConsolidatedTransactions
WITH SCHEMABINDING
AS
SELECT
    t.TransactionId,
    t.CardNumber,
    t.TransactionAmount,
    t.TransactionDate,
    t.Description,
    t.TransactionStatus,
    dbo.fn_GetValueCategory(t.TransactionAmount) AS Category,
    YEAR(t.TransactionDate) AS Year,
    MONTH(t.TransactionDate) AS Month,
    DAY(t.TransactionDate) AS Day,
    DATENAME(WEEKDAY, t.TransactionDate) AS DayOfWeek
FROM dbo.Transactions t;
GO

PRINT 'View vw_ConsolidatedTransactions created successfully.';
GO

-- =====================================================================
-- Usage Examples:
-- =====================================================================
-- -- All transactions from last month with category
-- SELECT * FROM dbo.vw_ConsolidatedTransactions
-- WHERE Year = YEAR(DATEADD(MONTH, -1, GETDATE()))
--   AND Month = MONTH(DATEADD(MONTH, -1, GETDATE()));
--
-- -- Total by category in current year
-- SELECT Category, SUM(TransactionAmount) AS Total, COUNT(*) AS Count
-- FROM dbo.vw_ConsolidatedTransactions
-- WHERE Year = YEAR(GETDATE())
-- GROUP BY Category
-- ORDER BY Total DESC;
--
-- -- Premium transactions from last 30 days
-- SELECT * FROM dbo.vw_ConsolidatedTransactions
-- WHERE Category = 'Premium'
--   AND TransactionDate >= DATEADD(DAY, -30, GETDATE());