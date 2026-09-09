USE TransactionManagement;
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
