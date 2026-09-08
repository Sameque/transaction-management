CREATE DATABASE TransactionManagemer;

USE TransactionManagemer;
GO

-- =====================================================================
-- Table: Transactions
-- =====================================================================
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

-- =====================================================================
-- Indexes for Performance
-- =====================================================================

-- Index for card number filter
CREATE INDEX IX_Transactions_CardNumber
ON dbo.Transactions (CardNumber);
GO

-- Index for date filter and ordering
CREATE INDEX IX_Transactions_TransactionDate
ON dbo.Transactions (TransactionDate DESC);
GO

-- Index for status filter
CREATE INDEX IX_Transactions_Status
ON dbo.Transactions (TransactionStatus);
GO

-- Composite index for period + status queries (used in SP)
CREATE INDEX IX_Transactions_DateStatus
ON dbo.Transactions (TransactionDate, TransactionStatus);
GO

PRINT 'Transactions table and indexes created successfully.';
GO