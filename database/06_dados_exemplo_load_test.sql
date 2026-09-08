-- =====================================================================
-- Script 06: Sample Data for Load Testing (200 Records)
-- =====================================================================
-- Inserts exactly 200 varied transactions covering:
-- - 10 different cards
-- - All statuses (Approved, Pending, Canceled)
-- - All categories (Low, Medium, High, Premium)
-- - Data from last 6 months (Apr 2026 - Sep 2026)
-- =====================================================================

USE FinancialTransactionDB;
GO

SET NOCOUNT ON;

-- Clear existing data (if any)
DELETE FROM dbo.Transactions;
DBCC CHECKIDENT('dbo.Transactions', RESEED, 0);
PRINT 'Previous data removed. Inserting 200 sample transactions...';
GO

-- =====================================================================
-- Test Cards (10 different cards)
-- =====================================================================
DECLARE @CardNumbers TABLE (CardNumber VARCHAR(20), CardType VARCHAR(20));
INSERT INTO @CardNumbers VALUES
('4111111111111111', 'Visa'),
('4000056655665556', 'Visa'),
('5555555555554444', 'Mastercard'),
('5111111111111111', 'Mastercard'),
('3782822463100050', 'Amex'),
('3714496353984310', 'Amex'),
('6011111111111117', 'Discover'),
('6011000990139424', 'Discover'),
('3566002020360505', 'JCB'),
('6222222222222222', 'UnionPay');

-- =====================================================================
-- Transaction descriptions with amount ranges
-- =====================================================================
DECLARE @Descriptions TABLE (
    Description VARCHAR(100),
    MinAmount DECIMAL(18,2),
    MaxAmount DECIMAL(18,2)
);

INSERT INTO @Descriptions VALUES
-- Low (< 100) - 50 transactions
('Coffee Shop', 3.00, 15.00),
('Bakery', 5.00, 25.00),
('Convenience Store', 8.00, 30.00),
('Fast Food', 10.00, 35.00),
('Parking Meter', 2.00, 10.00),
('Newsstand', 3.00, 20.00),
('Pharmacy Small', 10.00, 50.00),
('Vending Machine', 1.50, 5.00),
('Laundry', 8.00, 25.00),
('Toll Road', 3.00, 15.00),
-- Medium (100 - 500) - 80 transactions
('Supermarket', 80.00, 300.00),
('Restaurant', 60.00, 250.00),
('Gas Station', 50.00, 200.00),
('Clothing Store', 100.00, 400.00),
('Electronics Accessories', 50.00, 300.00),
('Online Shopping', 30.00, 400.00),
('Pharmacy', 40.00, 300.00),
('Streaming Services', 15.00, 80.00),
('Gym Membership', 50.00, 150.00),
('Ride Sharing', 20.00, 100.00),
('Grocery Delivery', 60.00, 250.00),
('Department Store', 100.00, 450.00),
('Beauty Salon', 80.00, 300.00),
('Pet Store', 50.00, 200.00),
('Bookstore', 30.00, 150.00),
-- High (500 - 2000) - 50 transactions
('Electronics - Laptop', 800.00, 3000.00),
('Electronics - Phone', 500.00, 2000.00),
('Furniture', 300.00, 2000.00),
('Appliances', 400.00, 2500.00),
('Airline Ticket', 400.00, 2000.00),
('Hotel Stay', 200.00, 1500.00),
('Car Repair', 300.00, 2000.00),
('Home Renovation', 500.00, 3000.00),
('Online Course', 200.00, 1500.00),
('Travel Package', 800.00, 2500.00),
('Jewelry', 500.00, 2000.00),
('Sporting Goods', 200.00, 1500.00),
-- Premium (> 2000) - 20 transactions
('Luxury Electronics', 2000.00, 5000.00),
('Luxury Jewelry', 2500.00, 10000.00),
('High-End Appliances', 2000.00, 5000.00),
('International Travel', 3000.00, 12000.00),
('Designer Fashion', 2000.00, 8000.00);

-- =====================================================================
-- Generate exactly 200 transactions
-- =====================================================================
DECLARE @TxnCount INT = 0;
DECLARE @TargetCount INT = 200;
DECLARE @CardIndex INT = 1;
DECLARE @DescCount INT = (SELECT COUNT(*) FROM @Descriptions);

WHILE @TxnCount < @TargetCount
BEGIN
    -- Round-robin through cards (20 transactions per card = 200 total)
    DECLARE @CardNumber VARCHAR(20);
    SELECT @CardNumber = CardNumber FROM (
        SELECT CardNumber, ROW_NUMBER() OVER (ORDER BY CardNumber) as rn
        FROM @CardNumbers
    ) c WHERE rn = ((@TxnCount % 10) + 1);

    -- Pick description (cycle through all descriptions)
    DECLARE @DescIdx INT = (@TxnCount % @DescCount) + 1;
    DECLARE @Description VARCHAR(100);
    DECLARE @MinAmt DECIMAL(18,2);
    DECLARE @MaxAmt DECIMAL(18,2);

    SELECT @Description = Description, @MinAmt = MinAmount, @MaxAmt = MaxAmount
    FROM (
        SELECT Description, MinAmount, MaxAmount, ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) as rn
        FROM @Descriptions
    ) d WHERE rn = @DescIdx;

    -- Generate amount
    DECLARE @Amount DECIMAL(18,2) = ROUND(@MinAmt + RAND() * (@MaxAmt - @MinAmt), 2);

    -- Date: spread across 6 months (Apr-Sep 2026)
    DECLARE @DaysOffset INT = CAST(RAND() * 183 AS INT); -- ~6 months
    DECLARE @TxnDate DATETIME = DATEADD(DAY, -@DaysOffset, '2026-09-07');

    -- Add random time (business hours)
    DECLARE @Hour INT = 8 + CAST(RAND() * 12 AS INT); -- 8am-8pm
    DECLARE @Minute INT = CAST(RAND() * 60 AS INT);
    DECLARE @Second INT = CAST(RAND() * 60 AS INT);
    SET @TxnDate = DATEADD(HOUR, @Hour, DATEADD(MINUTE, @Minute, DATEADD(SECOND, @Second, CAST(CAST(@TxnDate AS DATE) AS DATETIME))));

    -- Status: 75% Approved, 15% Pending, 10% Canceled
    DECLARE @Status VARCHAR(20);
    DECLARE @StatusRand FLOAT = RAND();
    IF @StatusRand < 0.75 SET @Status = 'Approved';
    ELSE IF @StatusRand < 0.90 SET @Status = 'Pending';
    ELSE SET @Status = 'Canceled';

    INSERT INTO dbo.Transactions (CardNumber, TransactionAmount, TransactionDate, Description, TransactionStatus)
    VALUES (@CardNumber, @Amount, @TxnDate, @Description, @Status);

    SET @TxnCount = @TxnCount + 1;
END

PRINT '200 transactions inserted successfully!';
GO

-- =====================================================================
-- Data Verification
-- =====================================================================
USE FinancialTransactionDB;
GO

SELECT
    COUNT(*) AS TotalTransactions,
    COUNT(DISTINCT CardNumber) AS UniqueCards,
    MIN(TransactionDate) AS MinDate,
    MAX(TransactionDate) AS MaxDate,
    SUM(CASE WHEN TransactionStatus = 'Approved' THEN 1 ELSE 0 END) AS Approved,
    SUM(CASE WHEN TransactionStatus = 'Pending' THEN 1 ELSE 0 END) AS Pending,
    SUM(CASE WHEN TransactionStatus = 'Canceled' THEN 1 ELSE 0 END) AS Canceled
FROM dbo.Transactions;

SELECT
    dbo.fn_GetValueCategory(TransactionAmount) AS Category,
    COUNT(*) AS Count,
    SUM(TransactionAmount) AS TotalAmount,
    AVG(TransactionAmount) AS AvgAmount
FROM dbo.Transactions
GROUP BY dbo.fn_GetValueCategory(TransactionAmount)
ORDER BY
    CASE dbo.fn_GetValueCategory(TransactionAmount)
        WHEN 'Low' THEN 1 WHEN 'Medium' THEN 2 WHEN 'High' THEN 3 WHEN 'Premium' THEN 4
    END;