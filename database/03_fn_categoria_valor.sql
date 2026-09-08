-- =====================================================================
-- Script 03: Scalar Function - Value Categorization
-- =====================================================================
-- Returns category based on transaction amount
-- Value Range (USD)    Category
-- > 2000               Premium
-- 1000 - 2000          High
-- 500 - 1000           Medium
-- < 500                Low
-- =====================================================================

USE TransactionManagemer;
GO

IF OBJECT_ID('dbo.fn_GetValueCategory', 'FN') IS NOT NULL
BEGIN
    DROP FUNCTION dbo.fn_GetValueCategory;
    PRINT 'Function fn_GetValueCategory dropped for recreation.';
END
GO

CREATE FUNCTION dbo.fn_GetValueCategory(@Amount DECIMAL(18,2))
RETURNS VARCHAR(20)
WITH SCHEMABINDING
AS
BEGIN
    RETURN CASE
        WHEN @Amount > 2000 THEN 'Premium'
        WHEN @Amount >= 1000 THEN 'High'
        WHEN @Amount >= 500 THEN 'Medium'
        ELSE 'Low'
    END;
END
GO

PRINT 'Scalar Function fn_GetValueCategory created successfully.';
GO

-- =====================================================================
-- Usage Examples:
-- =====================================================================
-- SELECT dbo.fn_GetValueCategory(2500);   -- Returns 'Premium'
-- SELECT dbo.fn_GetValueCategory(1500);   -- Returns 'High'
-- SELECT dbo.fn_GetValueCategory(750);    -- Returns 'Medium'
-- SELECT dbo.fn_GetValueCategory(300);    -- Returns 'Low'
-- SELECT dbo.fn_GetValueCategory(1000);   -- Returns 'High' (inclusive lower bound)
-- SELECT dbo.fn_GetValueCategory(2000);   -- Returns 'High' (inclusive upper bound)
-- SELECT dbo.fn_GetValueCategory(500);    -- Returns 'Medium' (inclusive lower bound)