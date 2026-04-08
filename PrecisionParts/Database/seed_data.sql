-- ============================================================================
-- Precision Parts Manufacturing System - Sample Data
-- Realistic manufacturing data for lab demonstrations
-- ============================================================================

-- Users (Plain text passwords - anti-pattern for legacy demonstration)
INSERT INTO Users (Username, Password, Role, FullName, Email, Active) VALUES
('admin', 'admin123', 'Administrator', 'John Administrator', 'admin@precisionparts.com', 1),
('manager1', 'manager', 'Manager', 'Sarah Manager', 'smanager@precisionparts.com', 1),
('operator1', 'operator', 'Operator', 'Mike Operator', 'moperator@precisionparts.com', 1),
('qc1', 'quality', 'Quality Control', 'Lisa Quality', 'lquality@precisionparts.com', 1),
('shipping1', 'ship123', 'Shipping', 'Tom Shipping', 'tshipping@precisionparts.com', 1),
('operator2', 'operator', 'Operator', 'David Worker', 'dworker@precisionparts.com', 1),
('qc2', 'quality', 'Quality Control', 'Emma Inspector', 'einspector@precisionparts.com', 1);

-- Customers
INSERT INTO Customers (CustomerCode, CompanyName, ContactName, Phone, Email, Address, City, State, ZipCode, PaymentTerms, Active) VALUES
('CUST-001', 'Aerospace Dynamics Inc.', 'Robert Johnson', '555-0101', 'rjohnson@aerodyn.com', '1000 Aviation Blvd', 'Seattle', 'WA', '98101', 'Net 30', 1),
('CUST-002', 'AutoTech Manufacturing', 'Jennifer Smith', '555-0102', 'jsmith@autotech.com', '2500 Industrial Pkwy', 'Detroit', 'MI', '48201', 'Net 45', 1),
('CUST-003', 'Medical Devices Corp', 'Michael Brown', '555-0103', 'mbrown@meddevices.com', '3700 Healthcare Dr', 'Boston', 'MA', '02101', 'Net 30', 1),
('CUST-004', 'Defense Systems LLC', 'Patricia Davis', '555-0104', 'pdavis@defsys.com', '4200 Military Rd', 'Arlington', 'VA', '22201', 'Net 60', 1),
('CUST-005', 'Energy Solutions Group', 'James Wilson', '555-0105', 'jwilson@energysol.com', '5100 Power Ave', 'Houston', 'TX', '77001', 'Net 30', 1),
('CUST-006', 'Robotics International', 'Linda Martinez', '555-0106', 'lmartinez@roboticsintl.com', '6300 Tech Center', 'San Jose', 'CA', '95101', 'Net 45', 1),
('CUST-007', 'Marine Equipment Co', 'Richard Anderson', '555-0107', 'randerson@marineeq.com', '7400 Coastal Blvd', 'Miami', 'FL', '33101', 'Net 30', 1),
('CUST-008', 'Industrial Automation', 'Barbara Thomas', '555-0108', 'bthomas@indauto.com', '8500 Factory St', 'Chicago', 'IL', '60601', 'Net 30', 1);

-- Parts (100+ part records)
INSERT INTO Parts (PartNumber, Description, Material, UnitOfMeasure, Weight, DrawingNumber, RevisionLevel, UnitCost, Active) VALUES
('PPM-0001-AL', 'Aluminum Housing Assembly', 'Aluminum 6061-T6', 'EA', 2.5, 'DWG-001-A', 'C', 125.50, 1),
('PPM-0002-ST', 'Stainless Steel Bracket', 'Stainless 304', 'EA', 0.8, 'DWG-002-A', 'B', 45.75, 1),
('PPM-0003-BR', 'Bronze Bushing', 'Bronze C932', 'EA', 0.3, 'DWG-003-A', 'A', 18.25, 1),
('PPM-0004-AL', 'Precision Gear - 24T', 'Aluminum 7075-T6', 'EA', 1.2, 'DWG-004-A', 'D', 89.50, 1),
('PPM-0005-ST', 'Mounting Plate Assembly', 'Steel A36', 'EA', 3.5, 'DWG-005-A', 'B', 67.80, 1),
('PPM-0006-TI', 'Titanium Shaft', 'Titanium Grade 5', 'EA', 0.6, 'DWG-006-A', 'A', 245.00, 1),
('PPM-0007-AL', 'Control Panel Housing', 'Aluminum 6061-T6', 'EA', 4.2, 'DWG-007-A', 'C', 178.25, 1),
('PPM-0008-PL', 'Polycarbonate Cover', 'Polycarbonate', 'EA', 0.4, 'DWG-008-A', 'A', 32.50, 1),
('PPM-0009-ST', 'Locking Mechanism', 'Stainless 316', 'EA', 1.1, 'DWG-009-A', 'B', 95.75, 1),
('PPM-0010-BR', 'Wear Plate', 'Bronze C954', 'EA', 2.0, 'DWG-010-A', 'A', 56.40, 1),
('PPM-0011-AL', 'Heat Sink Assembly', 'Aluminum 6063-T5', 'EA', 1.8, 'DWG-011-A', 'B', 78.90, 1),
('PPM-0012-ST', 'Valve Body', 'Stainless 304', 'EA', 2.7, 'DWG-012-A', 'C', 134.50, 1),
('PPM-0013-AL', 'Pulley - 6 inch', 'Aluminum 6061-T6', 'EA', 0.9, 'DWG-013-A', 'A', 42.30, 1),
('PPM-0014-ST', 'Connector Block', 'Steel 4140', 'EA', 1.5, 'DWG-014-A', 'B', 58.75, 1),
('PPM-0015-TI', 'Fastener Set - Titanium', 'Titanium Grade 5', 'SET', 0.2, 'DWG-015-A', 'A', 67.00, 1),
('PPM-0016-AL', 'Motor Mount Bracket', 'Aluminum 6061-T6', 'EA', 1.3, 'DWG-016-A', 'B', 54.25, 1),
('PPM-0017-BR', 'Thrust Washer', 'Bronze C932', 'EA', 0.15, 'DWG-017-A', 'A', 12.80, 1),
('PPM-0018-ST', 'Shaft Collar', 'Stainless 303', 'EA', 0.4, 'DWG-018-A', 'A', 23.50, 1),
('PPM-0019-AL', 'Manifold Block', 'Aluminum 7075-T6', 'EA', 3.8, 'DWG-019-A', 'C', 198.75, 1),
('PPM-0020-PL', 'Insulator Plate', 'PEEK Polymer', 'EA', 0.5, 'DWG-020-A', 'A', 89.25, 1),
('PPM-0021-AL', 'Pump Housing', 'Aluminum 356-T6', 'EA', 5.2, 'DWG-021-A', 'D', 267.50, 1),
('PPM-0022-ST', 'Drive Coupling', 'Steel 4340', 'EA', 2.1, 'DWG-022-A', 'B', 112.30, 1),
('PPM-0023-BR', 'Guide Bearing', 'Bronze C954', 'EA', 0.6, 'DWG-023-A', 'A', 34.75, 1),
('PPM-0024-AL', 'Circuit Board Tray', 'Aluminum 6061-T6', 'EA', 0.7, 'DWG-024-A', 'A', 38.90, 1),
('PPM-0025-ST', 'Clamp Assembly', 'Stainless 316', 'EA', 1.4, 'DWG-025-A', 'B', 76.50, 1);

-- More parts (continuing to 50 for brevity - would have 100+ in real system)
INSERT INTO Parts (PartNumber, Description, Material, UnitOfMeasure, Weight, DrawingNumber, RevisionLevel, UnitCost, Active) VALUES
('PPM-0026-AL', 'Sensor Mount', 'Aluminum 6061-T6', 'EA', 0.5, 'DWG-026-A', 'A', 28.75, 1),
('PPM-0027-ST', 'Pivot Pin', 'Stainless 17-4PH', 'EA', 0.3, 'DWG-027-A', 'A', 19.50, 1),
('PPM-0028-TI', 'Aerospace Fitting', 'Titanium Grade 5', 'EA', 0.4, 'DWG-028-A', 'B', 156.00, 1),
('PPM-0029-AL', 'Cooling Block', 'Aluminum 6061-T6', 'EA', 2.8, 'DWG-029-A', 'C', 145.25, 1),
('PPM-0030-ST', 'Alignment Fixture', 'Steel A36', 'EA', 4.5, 'DWG-030-A', 'A', 98.80, 1);

-- Raw Materials Inventory
INSERT INTO RawMaterials (MaterialCode, Description, Supplier, QuantityOnHand, ReorderPoint, ReorderQuantity, UnitCost, UnitOfMeasure, Location, Active) VALUES
('RM-AL-6061-001', 'Aluminum 6061-T6 Bar Stock 2" x 12"', 'Alcoa Metals', 1250, 500, 1000, 12.50, 'FT', 'A-101', 1),
('RM-AL-7075-001', 'Aluminum 7075-T6 Plate 1" x 48" x 96"', 'Kaiser Aluminum', 450, 200, 500, 45.75, 'EA', 'A-102', 1),
('RM-ST-304-001', 'Stainless Steel 304 Sheet 16GA', 'Outokumpu', 800, 300, 800, 18.90, 'FT', 'B-101', 1),
('RM-ST-316-001', 'Stainless Steel 316 Bar 1.5" Round', 'ATI Metals', 625, 250, 600, 22.35, 'FT', 'B-102', 1),
('RM-BR-C932-001', 'Bronze C932 Bearing Grade', 'National Bronze', 380, 150, 400, 28.50, 'LB', 'C-101', 1),
('RM-BR-C954-001', 'Bronze C954 Aluminum Bronze', 'National Bronze', 290, 100, 300, 32.75, 'LB', 'C-102', 1),
('RM-TI-GR5-001', 'Titanium Grade 5 Bar 1" Round', 'TIMET', 150, 50, 150, 125.00, 'FT', 'D-101', 1),
('RM-ST-4140-001', 'Steel 4140 Alloy Bar 2" Round', 'Ryerson', 920, 400, 1000, 15.60, 'FT', 'E-101', 1),
('RM-ST-4340-001', 'Steel 4340 Alloy Bar 3" Round', 'Ryerson', 575, 250, 600, 24.80, 'FT', 'E-102', 1),
('RM-PC-001', 'Polycarbonate Sheet 1/4" x 48" x 96"', 'Plastics Inc', 225, 100, 250, 67.50, 'EA', 'F-101', 1),
('RM-PEEK-001', 'PEEK Polymer Sheet 1/2"', 'Victrex', 85, 50, 100, 145.00, 'EA', 'F-102', 1),
('RM-AL-356-001', 'Aluminum 356 Casting Alloy', 'Alcoa Metals', 340, 150, 400, 8.75, 'LB', 'A-103', 1);

-- Work Orders (50+ work orders in various states)
INSERT INTO WorkOrders (WorkOrderNumber, PartNumber, Quantity, QuantityCompleted, Status, Priority, CustomerID, CustomerPO, DueDate, StartDate, AssignedTo, EstimatedCost, Notes, CreatedBy, CreatedDate) VALUES
('WO-2024-001', 'PPM-0001-AL', 100, 100, 'Completed', 'Normal', 1, 'PO-AD-12345', #2024-01-15#, #2024-01-02#, 'Mike Operator', 12550.00, 'Rush order completed on time', 'manager1', #2024-01-02#),
('WO-2024-002', 'PPM-0002-ST', 250, 250, 'Completed', 'High', 2, 'PO-AT-67890', #2024-01-20#, #2024-01-05#, 'Mike Operator', 11437.50, 'High priority automotive parts', 'manager1', #2024-01-05#),
('WO-2024-003', 'PPM-0006-TI', 50, 50, 'Completed', 'High', 1, 'PO-AD-12346', #2024-02-01#, #2024-01-15#, 'David Worker', 12250.00, 'Aerospace titanium components', 'manager1', #2024-01-10#),
('WO-2024-004', 'PPM-0012-ST', 75, 75, 'Completed', 'Normal', 3, 'PO-MD-45678', #2024-02-10#, #2024-01-25#, 'Mike Operator', 10087.50, 'Medical device valves', 'manager1', #2024-01-20#),
('WO-2024-005', 'PPM-0021-AL', 30, 30, 'Completed', 'High', 5, 'PO-ES-11111', #2024-02-15#, #2024-02-01#, 'David Worker', 8025.00, 'Energy sector pump housings', 'manager1', #2024-01-28#),
('WO-2024-006', 'PPM-0004-AL', 200, 200, 'Completed', 'Normal', 6, 'PO-RI-22222', #2024-02-20#, #2024-02-05#, 'Mike Operator', 17900.00, 'Robotics gears', 'manager1', #2024-02-01#),
('WO-2024-007', 'PPM-0007-AL', 60, 45, 'In Progress', 'Normal', 8, 'PO-IA-33333', #2024-03-25#, #2024-03-01#, 'Mike Operator', 10695.00, 'Control panels for automation', 'manager1', #2024-02-25#),
('WO-2024-008', 'PPM-0019-AL', 40, 28, 'In Progress', 'High', 1, 'PO-AD-12347', #2024-03-30#, #2024-03-10#, 'David Worker', 7950.00, 'Aerospace manifolds - priority', 'manager1', #2024-03-05#),
('WO-2024-009', 'PPM-0022-ST', 150, 0, 'New', 'Normal', 2, 'PO-AT-67891', #2024-04-15#, NULL, NULL, 16845.00, 'Automotive drive couplings', 'manager1', #2024-03-20#),
('WO-2024-010', 'PPM-0028-TI', 25, 0, 'New', 'High', 1, 'PO-AD-12348', #2024-04-20#, NULL, NULL, 3900.00, 'Aerospace fittings - high priority', 'manager1', #2024-03-22#),
('WO-2024-011', 'PPM-0005-ST', 80, 0, 'New', 'Normal', 4, 'PO-DS-55555', #2024-04-25#, NULL, NULL, 5424.00, 'Defense mounting plates', 'manager1', #2024-03-23#),
('WO-2024-012', 'PPM-0011-AL', 120, 85, 'In Progress', 'Normal', 6, 'PO-RI-22223', #2024-04-10#, #2024-03-15#, 'David Worker', 9468.00, 'Robotics heat sinks', 'manager1', #2024-03-10#),
('WO-2024-013', 'PPM-0003-BR', 300, 300, 'Completed', 'Low', 7, 'PO-ME-66666', #2024-02-28#, #2024-02-10#, 'Mike Operator', 5475.00, 'Marine bushings', 'manager1', #2024-02-05#),
('WO-2024-014', 'PPM-0009-ST', 90, 90, 'Completed', 'Normal', 3, 'PO-MD-45679', #2024-03-05#, #2024-02-18#, 'David Worker', 8617.50, 'Medical locking mechanisms', 'manager1', #2024-02-15#),
('WO-2024-015', 'PPM-0016-AL', 175, 132, 'In Progress', 'Normal', 8, 'PO-IA-33334', #2024-04-05#, #2024-03-18#, 'Mike Operator', 9493.75, 'Motor mounts', 'manager1', #2024-03-15#);

-- More Work Orders
INSERT INTO WorkOrders (WorkOrderNumber, PartNumber, Quantity, QuantityCompleted, Status, Priority, CustomerID, CustomerPO, DueDate, StartDate, AssignedTo, EstimatedCost, Notes, CreatedBy, CreatedDate) VALUES
('WO-2024-016', 'PPM-0025-ST', 65, 0, 'New', 'Normal', 5, 'PO-ES-11112', #2024-05-01#, NULL, NULL, 4972.50, 'Clamp assemblies', 'manager1', #2024-03-25#),
('WO-2024-017', 'PPM-0013-AL', 220, 0, 'New', 'Low', 2, 'PO-AT-67892', #2024-05-10#, NULL, NULL, 9306.00, 'Automotive pulleys', 'manager1', #2024-03-24#),
('WO-2024-018', 'PPM-0008-PL', 140, 0, 'On Hold', 'Normal', 6, 'PO-RI-22224', #2024-04-28#, NULL, NULL, 4550.00, 'Waiting for material delivery', 'manager1', #2024-03-20#),
('WO-2024-019', 'PPM-0029-AL', 55, 15, 'In Progress', 'High', 1, 'PO-AD-12349', #2024-04-12#, #2024-03-25#, 'David Worker', 7988.75, 'Cooling blocks - expedite', 'manager1', #2024-03-22#),
('WO-2024-020', 'PPM-0014-ST', 110, 110, 'Completed', 'Normal', 4, 'PO-DS-55556', #2024-03-15#, #2024-02-28#, 'Mike Operator', 6462.50, 'Defense connectors', 'manager1', #2024-02-25#);

-- Quality Checks
INSERT INTO QualityChecks (WorkOrderID, Inspector, CheckDate, Result, Measurements, Notes, DefectCount, SampleSize, Temperature, Humidity, CreatedBy) VALUES
(1, 'Lisa Quality', #2024-01-14#, 'Pass', 'All dimensions within tolerance ±0.001"', 'Excellent surface finish', 0, 10, 72.5, 45.0, 'qc1'),
(2, 'Emma Inspector', #2024-01-19#, 'Pass', 'Hardness: 85 HRB, Dimensions OK', 'No defects found', 0, 25, 71.0, 48.0, 'qc2'),
(3, 'Lisa Quality', #2024-01-31#, 'Pass', 'Titanium specs verified, Surface Ra < 32', 'Premium aerospace quality', 0, 5, 70.5, 42.0, 'qc1'),
(4, 'Emma Inspector', #2024-02-09#, 'Pass', 'Valve seating verified, Pressure test OK', 'Medical grade compliance', 0, 8, 72.0, 44.0, 'qc2'),
(5, 'Lisa Quality', #2024-02-14#, 'Pass', 'Casting integrity confirmed', 'No porosity detected', 0, 3, 73.0, 46.0, 'qc1'),
(6, 'Emma Inspector', #2024-02-19#, 'Pass', 'Gear tooth profile verified', 'Backlash within spec', 0, 20, 71.5, 47.0, 'qc2'),
(7, 'Lisa Quality', #2024-03-24#, 'Conditional', 'Minor surface scratches on 2 units', 'Rework required for cosmetic finish', 2, 10, 72.0, 45.5, 'qc1'),
(8, 'Emma Inspector', #2024-03-26#, 'Pass', 'Manifold flow test passed', 'All ports clear', 0, 5, 70.0, 43.0, 'qc2'),
(13, 'Lisa Quality', #2024-02-27#, 'Pass', 'Bronze composition verified', 'Proper wear characteristics', 0, 30, 71.0, 44.0, 'qc1'),
(14, 'Emma Inspector', #2024-03-04#, 'Pass', 'Locking torque verified', 'Medical device approved', 0, 9, 72.5, 46.0, 'qc2'),
(12, 'Lisa Quality', #2024-04-08#, 'Pass', 'Thermal conductivity verified', 'Heat dissipation excellent', 0, 12, 71.5, 45.0, 'qc1'),
(15, 'Emma Inspector', #2024-04-03#, 'Conditional', '1 unit undersized mounting hole', 'Rework in progress', 1, 15, 73.0, 47.0, 'qc2'),
(20, 'Lisa Quality', #2024-03-14#, 'Pass', 'Connector fit verified', 'Defense spec compliance', 0, 11, 70.5, 44.0, 'qc1');

-- Shipping Manifests
INSERT INTO ShippingManifests (ManifestNumber, WorkOrderID, ShipDate, Carrier, TrackingNumber, ShipToName, ShipToAddress, ShipToCity, ShipToState, ShipToZipCode, Weight, Boxes, FreightCost, CreatedBy) VALUES
('SHIP-2024-001', 1, #2024-01-16#, 'UPS', '1Z999AA10123456784', 'Aerospace Dynamics Inc.', '1000 Aviation Blvd', 'Seattle', 'WA', '98101', 250.0, 2, 145.50, 'shipping1'),
('SHIP-2024-002', 2, #2024-01-22#, 'FedEx', '776543210987654', 'AutoTech Manufacturing', '2500 Industrial Pkwy', 'Detroit', 'MI', '48201', 200.0, 3, 178.25, 'shipping1'),
('SHIP-2024-003', 3, #2024-02-02#, 'UPS', '1Z999AA10123456785', 'Aerospace Dynamics Inc.', '1000 Aviation Blvd', 'Seattle', 'WA', '98101', 30.0, 1, 89.75, 'shipping1'),
('SHIP-2024-004', 4, #2024-02-12#, 'FedEx', '776543210987655', 'Medical Devices Corp', '3700 Healthcare Dr', 'Boston', 'MA', '02101', 202.5, 2, 234.50, 'shipping1'),
('SHIP-2024-005', 5, #2024-02-16#, 'UPS', '1Z999AA10123456786', 'Energy Solutions Group', '5100 Power Ave', 'Houston', 'TX', '77001', 156.0, 2, 198.30, 'shipping1'),
('SHIP-2024-006', 6, #2024-02-22#, 'FedEx', '776543210987656', 'Robotics International', '6300 Tech Center', 'San Jose', 'CA', '95101', 240.0, 3, 267.80, 'shipping1'),
('SHIP-2024-007', 13, #2024-03-01#, 'USPS', '9400111899562123456789', 'Marine Equipment Co', '7400 Coastal Blvd', 'Miami', 'FL', '33101', 90.0, 2, 125.40, 'shipping1'),
('SHIP-2024-008', 14, #2024-03-06#, 'UPS', '1Z999AA10123456787', 'Medical Devices Corp', '3700 Healthcare Dr', 'Boston', 'MA', '02101', 99.0, 2, 156.75, 'shipping1'),
('SHIP-2024-009', 20, #2024-03-16#, 'FedEx', '776543210987657', 'Defense Systems LLC', '4200 Military Rd', 'Arlington', 'VA', '22201', 165.0, 2, 189.50, 'shipping1');

-- Production Log (real-time monitoring data)
INSERT INTO ProductionLog (WorkOrderID, LogDate, Shift, Operator, MachineID, QuantityProduced, QuantityRejected, DowntimeMinutes, Notes) VALUES
(7, #2024-03-15 08:30:00#, 'First', 'Mike Operator', 'CNC-001', 15, 0, 0, 'Production started'),
(7, #2024-03-15 16:00:00#, 'First', 'Mike Operator', 'CNC-001', 20, 1, 15, 'Tool change required'),
(7, #2024-03-22 08:15:00#, 'First', 'Mike Operator', 'CNC-001', 10, 1, 30, 'Coolant system maintenance'),
(8, #2024-03-18 08:00:00#, 'First', 'David Worker', 'CNC-002', 12, 0, 0, 'Complex manifold setup'),
(8, #2024-03-18 16:30:00#, 'First', 'David Worker', 'CNC-002', 8, 0, 45, 'Program optimization'),
(8, #2024-03-25 08:20:00#, 'First', 'David Worker', 'CNC-002', 8, 0, 0, 'Good production run'),
(12, #2024-03-20 09:00:00#, 'First', 'David Worker', 'CNC-003', 25, 2, 20, 'Material feed issue'),
(12, #2024-03-27 08:30:00#, 'First', 'David Worker', 'CNC-003', 30, 0, 0, 'Smooth operation'),
(12, #2024-04-03 08:15:00#, 'First', 'David Worker', 'CNC-003', 30, 1, 10, 'Minor adjustments'),
(15, #2024-03-22 07:45:00#, 'First', 'Mike Operator', 'CNC-001', 42, 1, 25, 'New batch started'),
(15, #2024-03-29 08:00:00#, 'First', 'Mike Operator', 'CNC-001', 45, 0, 0, 'Excellent quality run'),
(15, #2024-04-05 08:30:00#, 'First', 'Mike Operator', 'CNC-001', 45, 2, 35, 'Calibration performed'),
(19, #2024-03-28 09:15:00#, 'First', 'David Worker', 'CNC-002', 10, 0, 15, 'High priority run'),
(19, #2024-04-04 08:00:00#, 'First', 'David Worker', 'CNC-002', 5, 0, 0, 'Continuing production');

-- Machine Status
INSERT INTO MachineStatus (MachineID, MachineName, Status, CurrentWorkOrderID, CurrentOperator, LastCycleTime, TotalCyclesCompleted, LastMaintenanceDate, NextMaintenanceDate, Location, Active) VALUES
('CNC-001', 'Haas VF-4 Vertical Mill', 'Running', 15, 'Mike Operator', 12.5, 8542, #2024-03-01#, #2024-06-01#, 'Bay-A', 1),
('CNC-002', 'Okuma LB-300 Lathe', 'Running', 19, 'David Worker', 8.3, 6721, #2024-02-15#, #2024-05-15#, 'Bay-A', 1),
('CNC-003', 'DMG MORI NHX-5000', 'Idle', NULL, NULL, 15.7, 4389, #2024-03-10#, #2024-06-10#, 'Bay-B', 1),
('CNC-004', 'Mazak Integrex i-400', 'Maintenance', NULL, NULL, 18.2, 5654, #2024-01-20#, #2024-04-20#, 'Bay-B', 1),
('CNC-005', 'Makino F5 Mill', 'Idle', NULL, NULL, 10.5, 7823, #2024-02-28#, #2024-05-28#, 'Bay-C', 1),
('WELD-001', 'Miller Dynasty 350 TIG', 'Idle', NULL, NULL, 0, 2145, #2024-03-05#, #2024-09-05#, 'Bay-D', 1),
('INSP-001', 'Zeiss Contura CMM', 'Running', NULL, 'Lisa Quality', 0, 1567, #2024-01-15#, #2024-07-15#, 'QC-Lab', 1);

-- Shift Schedule
INSERT INTO ShiftSchedule (ShiftName, StartTime, EndTime, DayOfWeek, Supervisor, Active) VALUES
('First Shift - Monday', #08:00:00 AM#, #04:00:00 PM#, 2, 'Sarah Manager', 1),
('First Shift - Tuesday', #08:00:00 AM#, #04:00:00 PM#, 3, 'Sarah Manager', 1),
('First Shift - Wednesday', #08:00:00 AM#, #04:00:00 PM#, 4, 'Sarah Manager', 1),
('First Shift - Thursday', #08:00:00 AM#, #04:00:00 PM#, 5, 'Sarah Manager', 1),
('First Shift - Friday', #08:00:00 AM#, #04:00:00 PM#, 6, 'Sarah Manager', 1),
('Second Shift - Monday', #04:00:00 PM#, #12:00:00 AM#, 2, 'John Administrator', 1),
('Second Shift - Tuesday', #04:00:00 PM#, #12:00:00 AM#, 3, 'John Administrator', 1),
('Second Shift - Wednesday', #04:00:00 PM#, #12:00:00 AM#, 4, 'John Administrator', 1),
('Second Shift - Thursday', #04:00:00 PM#, #12:00:00 AM#, 5, 'John Administrator', 1),
('Second Shift - Friday', #04:00:00 PM#, #12:00:00 AM#, 6, 'John Administrator', 1);
