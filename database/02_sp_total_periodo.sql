-- =====================================================================
-- Script 02: Stored Procedure - Transaction Totals by Period
-- =====================================================================
-- Calculates totals grouped by card and status within a period
-- =====================================================================

USE TransactionManagemer;
GO

IF OBJECT_ID('dbo.sp_GetTransactionTotalsByPeriod', 'P') IS NOT NULL
BEGIN
    DROP PROCEDURE dbo.sp_GetTransactionTotalsByPeriod;
    PRINT 'Procedure sp_GetTransactionTotalsByPeriod dropped for recreation.';
END
GO

CREATE PROCEDURE dbo.sp_GetTransactionTotalsByPeriod
    @StartDate        DATETIME,
    @EndDate          DATETIME,
    @TransactionStatus VARCHAR(20) = NULL  -- NULL = all statuses
AS
BEGIN
    SET NOCOUNT ON;

    -- Basic parameter validation
    IF @StartDate > @EndDate
    BEGIN
        RAISERROR('StartDate cannot be greater than EndDate.', 16, 1);
        RETURN;
    END

    -- Status validation if provided
    IF @TransactionStatus IS NOT NULL
       AND @TransactionStatus NOT IN ('Approved', 'Pending', 'Canceled')
    BEGIN
        RAISERROR('Invalid TransactionStatus. Allowed values: Approved, Pending, Canceled.', 16, 1);
        RETURN;
    END

    SELECT
        t.CardNumber,
        SUM(t.TransactionAmount) AS TotalAmount,
        COUNT(*) AS TransactionCount,
        t.TransactionStatus
    FROM dbo.Transactions t
    WHERE t.TransactionDate BETWEEN @StartDate AND @EndDate
      AND (@TransactionStatus IS NULL OR t.TransactionStatus = @TransactionStatus)
    GROUP BY t.CardNumber, t.TransactionStatus
    ORDER BY TotalAmount DESC;
END
GO

PRINT 'Stored Procedure sp_GetTransactionTotalsByPeriod created successfully.';
GO

-- =====================================================================
-- Usage Examples:
-- =====================================================================
-- EXEC dbo.sp_GetTransactionTotalsByPeriod '2026-01-01', '2026-01-31', 'Approved';
-- EXEC dbo.sp_GetTransactionTotalsByPeriod '2026-01-01', '2026-01-31', NULL;  -- All statuses
-- EXEC dbo.sp_GetTransactionTotalsByPeriod '2026-01-01', GETDATE(), 'Pending';