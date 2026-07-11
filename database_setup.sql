-- SQL Server Database Setup Script for QR Code Generator Project
-- Targets production database: Es_Comp0001_2026

USE Es_Comp0001_2026;
GO

-- 1. Create In_CustomerPDF Table to track imported PDF documents
IF OBJECT_ID('In_CustomerPDF', 'U') IS NULL
BEGIN
    CREATE TABLE In_CustomerPDF (
        PDFID INT IDENTITY(1,1) PRIMARY KEY,
        ItemCode VARCHAR(100) NOT NULL,
        FileName VARCHAR(255) NOT NULL UNIQUE,
        FilePath VARCHAR(500) NOT NULL,
        TotalQty INT NOT NULL,
        ImportedOn DATETIME DEFAULT GETDATE(),
        Printed BIT DEFAULT 0,
        PrintedOn DATETIME NULL,
        ReprintCount INT DEFAULT 0
    );
    PRINT 'Created In_CustomerPDF table.';
END
ELSE
BEGIN
    PRINT 'In_CustomerPDF table already exists.';
END

-- 2. Alter existing In_QRmaster table if needed (adding new columns conditionally)
IF OBJECT_ID('In_QRmaster', 'U') IS NULL
BEGIN
    -- Fallback in case In_QRmaster is missing in some test environment
    CREATE TABLE In_QRmaster (
        QRID INT IDENTITY(1,1) PRIMARY KEY,
        ItemCode VARCHAR(100) NOT NULL,
        QRMonth VARCHAR(6) NOT NULL,
        SerialNo INT NOT NULL,
        QRValue VARCHAR(500) NOT NULL UNIQUE,
        GeneratedOn DATETIME DEFAULT GETDATE(),
        Printed BIT DEFAULT 0,
        IsCustomerQR BIT DEFAULT 0,
        CustomerPDFID INT NULL
    );
    PRINT 'Created In_QRmaster table.';
END
ELSE
BEGIN
    -- Add columns to existing table if they don't exist
    IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('In_QRmaster') AND name = 'Printed')
    BEGIN
        ALTER TABLE In_QRmaster ADD Printed BIT DEFAULT 0 WITH VALUES;
        PRINT 'Added Printed column to In_QRmaster.';
    END
    
    IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('In_QRmaster') AND name = 'IsCustomerQR')
    BEGIN
        ALTER TABLE In_QRmaster ADD IsCustomerQR BIT DEFAULT 0 WITH VALUES;
        PRINT 'Added IsCustomerQR column to In_QRmaster.';
    END
    
    IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('In_QRmaster') AND name = 'CustomerPDFID')
    BEGIN
        ALTER TABLE In_QRmaster ADD CustomerPDFID INT NULL;
        PRINT 'Added CustomerPDFID column to In_QRmaster.';
    END
    
    IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('In_QRmaster') AND name = 'ReprintCount')
    BEGIN
        ALTER TABLE In_QRmaster ADD ReprintCount INT DEFAULT 0 WITH VALUES;
        PRINT 'Added ReprintCount column to In_QRmaster.';
    END
END

-- 3. Create UserMaster Table for Role-Based Access Control
IF OBJECT_ID('UserMaster', 'U') IS NULL
BEGIN
    CREATE TABLE UserMaster (
        UserID VARCHAR(50) PRIMARY KEY,
        PasswordHash VARCHAR(100) NOT NULL,
        UserRole VARCHAR(20) NOT NULL -- 'Operator', 'Supervisor', 'Admin'
    );
    
    -- Insert Default Accounts (Passwords are stored as plain text for simplicity in this setup)
    INSERT INTO UserMaster (UserID, PasswordHash, UserRole) VALUES ('operator1', 'op123', 'Operator');
    INSERT INTO UserMaster (UserID, PasswordHash, UserRole) VALUES ('supervisor1', 'sup123', 'Supervisor');
    INSERT INTO UserMaster (UserID, PasswordHash, UserRole) VALUES ('admin', 'admin123', 'Admin');
    
    PRINT 'Created UserMaster table and inserted default users.';
END
ELSE
BEGIN
    PRINT 'UserMaster table already exists.';
END

-- 4. Create QRReprintLog Table for Auditing Supervisor-Approved Reprinting
IF OBJECT_ID('QRReprintLog', 'U') IS NULL
BEGIN
    CREATE TABLE QRReprintLog (
        LogID INT IDENTITY(1,1) PRIMARY KEY,
        QRValue VARCHAR(500) NOT NULL,
        AuthorizedBy VARCHAR(50) NOT NULL,
        ReprintedOn DATETIME DEFAULT GETDATE(),
        Reason VARCHAR(250) NULL
    );
    PRINT 'Created QRReprintLog table.';
END
ELSE
BEGIN
    PRINT 'QRReprintLog table already exists.';
END

-- 5. Create dummy esjobscandata table for testing if it doesn't exist
IF OBJECT_ID('esjobscandata', 'U') IS NULL
BEGIN
    CREATE TABLE esjobscandata (
        ScanID INT IDENTITY(1,1) PRIMARY KEY,
        QRText VARCHAR(500) NOT NULL UNIQUE,
        ItemCode VARCHAR(100) NOT NULL,
        ScanDateTime DATETIME DEFAULT GETDATE()
    );
    PRINT 'Created dummy esjobscandata table for testing.';
END
ELSE
BEGIN
    PRINT 'esjobscandata table already exists.';
END
GO
