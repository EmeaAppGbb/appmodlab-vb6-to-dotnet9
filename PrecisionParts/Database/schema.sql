-- ============================================================================
-- Precision Parts Manufacturing System - Database Schema
-- Microsoft Access Database (.mdb) Schema Definition
-- Created: January 2001
-- Last Modified: March 2024
-- ============================================================================

-- Table: Users (Authentication - Plain text passwords - security anti-pattern)
CREATE TABLE Users (
    UserID COUNTER PRIMARY KEY,
    Username VARCHAR(50) NOT NULL,
    Password VARCHAR(50) NOT NULL,
    Role VARCHAR(50) NOT NULL,
    FullName VARCHAR(100),
    Email VARCHAR(100),
    Active BIT DEFAULT 1,
    CreatedDate DATETIME DEFAULT Now(),
    LastLogin DATETIME
);

-- Table: Parts (Part Master Data)
CREATE TABLE Parts (
    PartID COUNTER PRIMARY KEY,
    PartNumber VARCHAR(50) NOT NULL UNIQUE,
    Description VARCHAR(255) NOT NULL,
    Material VARCHAR(100),
    UnitOfMeasure VARCHAR(20) DEFAULT 'EA',
    Weight DOUBLE,
    DrawingNumber VARCHAR(50),
    RevisionLevel VARCHAR(10) DEFAULT 'A',
    UnitCost CURRENCY,
    Active BIT DEFAULT 1,
    CreatedDate DATETIME DEFAULT Now(),
    ModifiedDate DATETIME
);

-- Table: RawMaterials (Inventory)
CREATE TABLE RawMaterials (
    MaterialID COUNTER PRIMARY KEY,
    MaterialCode VARCHAR(50) NOT NULL UNIQUE,
    Description VARCHAR(255) NOT NULL,
    Supplier VARCHAR(100),
    QuantityOnHand DOUBLE DEFAULT 0,
    ReorderPoint DOUBLE DEFAULT 100,
    ReorderQuantity DOUBLE DEFAULT 500,
    UnitCost CURRENCY,
    UnitOfMeasure VARCHAR(20) DEFAULT 'EA',
    Location VARCHAR(50),
    LastOrderDate DATETIME,
    Active BIT DEFAULT 1
);

-- Table: Customers
CREATE TABLE Customers (
    CustomerID COUNTER PRIMARY KEY,
    CustomerCode VARCHAR(50) NOT NULL UNIQUE,
    CompanyName VARCHAR(100) NOT NULL,
    ContactName VARCHAR(100),
    Phone VARCHAR(20),
    Email VARCHAR(100),
    Address VARCHAR(255),
    City VARCHAR(50),
    State VARCHAR(2),
    ZipCode VARCHAR(10),
    Country VARCHAR(50) DEFAULT 'USA',
    PaymentTerms VARCHAR(50) DEFAULT 'Net 30',
    Active BIT DEFAULT 1
);

-- Table: WorkOrders
CREATE TABLE WorkOrders (
    WorkOrderID COUNTER PRIMARY KEY,
    WorkOrderNumber VARCHAR(50) NOT NULL UNIQUE,
    PartNumber VARCHAR(50) NOT NULL,
    Quantity DOUBLE NOT NULL,
    QuantityCompleted DOUBLE DEFAULT 0,
    Status VARCHAR(50) DEFAULT 'New',
    Priority VARCHAR(20) DEFAULT 'Normal',
    CustomerID LONG,
    CustomerPO VARCHAR(50),
    DueDate DATETIME,
    StartDate DATETIME,
    CompletedDate DATETIME,
    AssignedTo VARCHAR(50),
    EstimatedCost CURRENCY,
    ActualCost CURRENCY,
    Notes MEMO,
    CreatedBy VARCHAR(50),
    CreatedDate DATETIME DEFAULT Now(),
    ModifiedDate DATETIME,
    FOREIGN KEY (CustomerID) REFERENCES Customers(CustomerID)
);

-- Table: QualityChecks
CREATE TABLE QualityChecks (
    CheckID COUNTER PRIMARY KEY,
    WorkOrderID LONG NOT NULL,
    Inspector VARCHAR(100) NOT NULL,
    CheckDate DATETIME DEFAULT Now(),
    Result VARCHAR(50) DEFAULT 'Pass',
    Measurements VARCHAR(255),
    Notes MEMO,
    DefectCount INTEGER DEFAULT 0,
    SampleSize INTEGER,
    Temperature DOUBLE,
    Humidity DOUBLE,
    CreatedBy VARCHAR(50),
    CreatedDate DATETIME DEFAULT Now(),
    FOREIGN KEY (WorkOrderID) REFERENCES WorkOrders(WorkOrderID)
);

-- Table: ShippingManifests
CREATE TABLE ShippingManifests (
    ManifestID COUNTER PRIMARY KEY,
    ManifestNumber VARCHAR(50) NOT NULL UNIQUE,
    WorkOrderID LONG NOT NULL,
    ShipDate DATETIME DEFAULT Now(),
    Carrier VARCHAR(50),
    TrackingNumber VARCHAR(100),
    ShipToName VARCHAR(100),
    ShipToAddress VARCHAR(255),
    ShipToCity VARCHAR(50),
    ShipToState VARCHAR(2),
    ShipToZipCode VARCHAR(10),
    Weight DOUBLE,
    Boxes INTEGER DEFAULT 1,
    FreightCost CURRENCY,
    Notes MEMO,
    CreatedBy VARCHAR(50),
    CreatedDate DATETIME DEFAULT Now(),
    FOREIGN KEY (WorkOrderID) REFERENCES WorkOrders(WorkOrderID)
);

-- Table: ProductionLog (Real-time monitoring data)
CREATE TABLE ProductionLog (
    LogID COUNTER PRIMARY KEY,
    WorkOrderID LONG NOT NULL,
    LogDate DATETIME DEFAULT Now(),
    Shift VARCHAR(20),
    Operator VARCHAR(100),
    MachineID VARCHAR(50),
    QuantityProduced DOUBLE,
    QuantityRejected DOUBLE,
    DowntimeMinutes INTEGER DEFAULT 0,
    Notes MEMO,
    FOREIGN KEY (WorkOrderID) REFERENCES WorkOrders(WorkOrderID)
);

-- Table: MachineStatus (For real-time monitoring simulation)
CREATE TABLE MachineStatus (
    MachineID VARCHAR(50) PRIMARY KEY,
    MachineName VARCHAR(100) NOT NULL,
    Status VARCHAR(50) DEFAULT 'Idle',
    CurrentWorkOrderID LONG,
    CurrentOperator VARCHAR(100),
    LastCycleTime DOUBLE,
    TotalCyclesCompleted LONG DEFAULT 0,
    LastMaintenanceDate DATETIME,
    NextMaintenanceDate DATETIME,
    Location VARCHAR(50),
    Active BIT DEFAULT 1
);

-- Table: ShiftSchedule
CREATE TABLE ShiftSchedule (
    ScheduleID COUNTER PRIMARY KEY,
    ShiftName VARCHAR(50) NOT NULL,
    StartTime DATETIME NOT NULL,
    EndTime DATETIME NOT NULL,
    DayOfWeek INTEGER,
    Supervisor VARCHAR(100),
    Active BIT DEFAULT 1
);

-- Indexes for performance
CREATE INDEX idx_WorkOrders_Status ON WorkOrders(Status);
CREATE INDEX idx_WorkOrders_DueDate ON WorkOrders(DueDate);
CREATE INDEX idx_WorkOrders_PartNumber ON WorkOrders(PartNumber);
CREATE INDEX idx_QualityChecks_WorkOrderID ON QualityChecks(WorkOrderID);
CREATE INDEX idx_QualityChecks_Result ON QualityChecks(Result);
CREATE INDEX idx_ShippingManifests_WorkOrderID ON ShippingManifests(WorkOrderID);
CREATE INDEX idx_ProductionLog_WorkOrderID ON ProductionLog(WorkOrderID);
CREATE INDEX idx_RawMaterials_Supplier ON RawMaterials(Supplier);
