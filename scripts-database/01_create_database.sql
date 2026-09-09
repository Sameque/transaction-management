
-- Create database (if not exists)
IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = 'TransactionManagemer')
BEGIN
    CREATE DATABASE TransactionManagemer;
    PRINT 'Database TransactionManagemer created successfully.';
END
ELSE
BEGIN
    PRINT 'Database TransactionManagemer already exists.';
END
GO

USE TransactionManagemer;
GO

IF OBJECT_ID('dbo.Transactions', 'U') IS NOT NULL
BEGIN
    DROP TABLE dbo.Transactions;
    PRINT 'Transactions table dropped for recreation.';
END
GO

CREATE TABLE dbo.Transactions
(
    TransactionId     INT IDENTITY(1,1) PRIMARY KEY,
    CardNumber        CHAR(16) NOT NULL,
    TransactionAmount DECIMAL(18,2) NOT NULL CHECK (TransactionAmount > 0),
    TransactionDate   DATETIME NOT NULL DEFAULT GETDATE(),
    Description       VARCHAR(255) NULL,
    TransactionStatus VARCHAR(20) NOT NULL
        CHECK (TransactionStatus IN ('Approved', 'Pending', 'Canceled'))
        DEFAULT 'Pending'
);
GO

CREATE INDEX IX_Transactions_CardNumber
ON dbo.Transactions (CardNumber);
GO

CREATE INDEX IX_Transactions_TransactionDate
ON dbo.Transactions (TransactionDate DESC);
GO

CREATE INDEX IX_Transactions_Status
ON dbo.Transactions (TransactionStatus);
GO

CREATE INDEX IX_Transactions_DateStatus
ON dbo.Transactions (TransactionDate, TransactionStatus);
GO

PRINT 'Transactions table and indexes created successfully.';
GO