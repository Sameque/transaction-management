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
