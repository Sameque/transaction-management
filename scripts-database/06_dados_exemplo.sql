USE TransactionManagement;
GO

SET NOCOUNT ON;

-- Clear existing data (if any)
DELETE FROM dbo.Transactions;
DBCC CHECKIDENT('dbo.Transactions', RESEED, 0);
PRINT 'Previous data removed. Inserting new sample data...';
GO

INSERT INTO dbo.Transactions (CardNumber, TransactionAmount, TransactionDate, Description, TransactionStatus) VALUES
-- Card 1 - September 2026
('4111111111111111', 150.00,  '2026-09-01 08:30:00', 'Supermarket', 'Approved'),
('4111111111111111', 3200.00, '2026-09-02 14:15:00', 'Electronics - Laptop', 'Approved'),
('4111111111111111', 85.50,   '2026-09-03 10:00:00', 'Pharmacy', 'Pending'),
('4111111111111111', 1200.00, '2026-09-05 16:45:00', 'Airline Ticket', 'Approved'),
('4111111111111111', 45.00,   '2026-09-07 12:20:00', 'Ride Sharing App', 'Canceled'),

-- Card 2 - September 2026
('5555555555554444', 750.00,  '2026-09-01 09:00:00', 'Restaurant', 'Approved'),
('5555555555554444', 250.00,  '2026-09-04 11:30:00', 'Clothing Store', 'Approved'),
('5555555555554444', 5500.00, '2026-09-06 15:00:00', 'Jewelry', 'Pending'),
('5555555555554444', 180.00,  '2026-09-08 13:10:00', 'Streaming Services', 'Approved'),

-- Card 3 - September 2026
('3782822463100050', 320.00,  '2026-09-02 10:00:00', 'Gas Station', 'Approved'),
('3782822463100050', 1500.00, '2026-09-05 18:20:00', 'Hotel', 'Approved'),
('3782822463100050', 95.00,   '2026-09-09 08:45:00', 'Toll', 'Canceled'),

-- Card 4 - September 2026
('6011111111111117', 4200.00, '2026-09-03 11:00:00', 'Appliances', 'Approved'),
('6011111111111117', 650.00,  '2026-09-07 14:30:00', 'Online Course', 'Pending'),
('6011111111111117', 230.00,  '2026-09-10 09:15:00', 'Grocery', 'Approved'),

-- Card 5 - September 2026
('4000056655665556', 1800.00, '2026-09-01 12:00:00', 'Furniture', 'Approved'),
('4000056655665556', 350.00,  '2026-09-04 16:00:00', 'Gym', 'Approved'),
('4000056655665556', 2700.00, '2026-09-08 10:30:00', 'Home Renovation', 'Pending');

INSERT INTO dbo.Transactions (CardNumber, TransactionAmount, TransactionDate, Description, TransactionStatus) VALUES
-- Card 1 - August 2026
('4111111111111111', 250.00,  '2026-08-02 09:00:00', 'Supermarket', 'Approved'),
('4111111111111111', 450.00,  '2026-08-05 14:00:00', 'Restaurant', 'Approved'),
('4111111111111111', 850.00,  '2026-08-10 11:30:00', 'Department Store', 'Approved'),
('4111111111111111', 2100.00, '2026-08-15 16:45:00', 'Travel', 'Approved'),
('4111111111111111', 120.00,  '2026-08-20 10:00:00', 'Pharmacy', 'Canceled'),
('4111111111111111', 95.00,   '2026-08-25 13:20:00', 'Delivery App', 'Pending'),

-- Card 2 - August 2026
('5555555555554444', 3000.00, '2026-08-01 10:00:00', 'Electronics', 'Approved'),
('5555555555554444', 180.00,  '2026-08-08 12:00:00', 'Cinema', 'Approved'),
('5555555555554444', 650.00,  '2026-08-12 15:30:00', 'Concert', 'Approved'),
('5555555555554444', 420.00,  '2026-08-18 19:00:00', 'Dinner', 'Pending'),
('5555555555554444', 2800.00, '2026-08-22 11:00:00', 'Furniture', 'Approved'),

-- Card 3 - August 2026
('3782822463100050', 120.00,  '2026-08-03 08:00:00', 'Coffee', 'Approved'),
('3782822463100050', 750.00,  '2026-08-09 14:00:00', 'Clothing', 'Approved'),
('3782822463100050', 3500.00, '2026-08-14 17:00:00', 'Laptop', 'Approved'),
('3782822463100050', 85.00,   '2026-08-20 10:00:00', 'Snack Bar', 'Canceled'),

-- Card 4 - August 2026
('6011111111111117', 150.00,  '2026-08-04 12:00:00', 'Grocery', 'Approved'),
('6011111111111117', 2200.00, '2026-08-11 16:00:00', 'Course', 'Approved'),
('6011111111111117', 400.00,  '2026-08-19 11:30:00', 'Gift', 'Pending'),
('6011111111111117', 65.00,   '2026-08-26 09:00:00', 'Pharmacy', 'Approved'),

-- Card 5 - August 2026
('4000056655665556', 900.00,  '2026-08-02 10:00:00', 'Construction Materials', 'Approved'),
('4000056655665556', 3200.00, '2026-08-07 14:00:00', 'Appliances', 'Approved'),
('4000056655665556', 150.00,  '2026-08-13 15:00:00', 'Supermarket', 'Canceled'),
('4000056655665556', 550.00,  '2026-08-21 10:30:00', 'Restaurant', 'Approved');

INSERT INTO dbo.Transactions (CardNumber, TransactionAmount, TransactionDate, Description, TransactionStatus) VALUES
-- Card 1 - July 2026
('4111111111111111', 180.00,  '2026-07-03 09:00:00', 'Bakery', 'Approved'),
('4111111111111111', 2500.00, '2026-07-10 11:00:00', 'Smartphone', 'Approved'),
('4111111111111111', 75.00,   '2026-07-15 14:00:00', 'Uber', 'Approved'),
('4111111111111111', 420.00,  '2026-07-22 16:00:00', 'Pharmacy', 'Approved'),
('4111111111111111', 3100.00, '2026-07-28 10:00:00', 'Flights', 'Pending'),

-- Card 2 - July 2026
('5555555555554444', 650.00,  '2026-07-05 10:00:00', 'Clothing', 'Approved'),
('5555555555554444', 120.00,  '2026-07-12 13:00:00', 'Snack', 'Canceled'),
('5555555555554444', 800.00,  '2026-07-19 15:00:00', 'Electronics', 'Approved'),
('5555555555554444', 2300.00, '2026-07-26 11:00:00', 'Travel', 'Approved'),

-- Card 3 - July 2026
('3782822463100050', 1500.00, '2026-07-01 09:00:00', 'Hotel', 'Approved'),
('3782822463100050', 350.00,  '2026-07-08 12:00:00', 'Lunch', 'Approved'),
('3782822463100050', 4800.00, '2026-07-15 14:00:00', 'Jewelry', 'Approved'),
('3782822463100050', 90.00,   '2026-07-20 16:00:00', 'Taxi', 'Pending'),

-- Card 4 - July 2026
('6011111111111117', 200.00,  '2026-07-02 11:00:00', 'Grocery', 'Approved'),
('6011111111111117', 1800.00, '2026-07-09 14:00:00', 'Laptop', 'Approved'),
('6011111111111117', 450.00,  '2026-07-16 10:00:00', 'Course', 'Canceled'),
('6011111111111117', 380.00,  '2026-07-23 12:00:00', 'Restaurant', 'Approved'),

-- Card 5 - July 2026
('4000056655665556', 2700.00, '2026-07-04 10:00:00', 'Furniture', 'Approved'),
('4000056655665556', 120.00,  '2026-07-11 13:00:00', 'Supermarket', 'Approved'),
('4000056655665556', 1900.00, '2026-07-18 15:00:00', 'Electronics', 'Pending'),
('4000056655665556', 85.00,   '2026-07-25 17:00:00', 'Pharmacy', 'Approved');
GO

PRINT 'Sample data inserted successfully!';
PRINT 'Total transactions: ' + CAST(@@ROWCOUNT AS VARCHAR(10));

USE TransactionManagement;
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

-- Distribution by Category
SELECT
    dbo.fn_GetValueCategory(TransactionAmount) AS Category,
    COUNT(*) AS Count,
    SUM(TransactionAmount) AS TotalAmount
FROM dbo.Transactions
GROUP BY dbo.fn_GetValueCategory(TransactionAmount)
ORDER BY TotalAmount DESC;