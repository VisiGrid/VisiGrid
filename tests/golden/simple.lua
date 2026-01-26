-- api=v1
-- VisiGrid Provenance Script
-- A simple test script
-- Expected fingerprint: v1:6:b38b632d7f38dedf50ebea7b52785554

-- #1 Set header
grid.set{ sheet=1, cell="A1", value="Name" }

-- #2 Set header
grid.set{ sheet=1, cell="B1", value="Value" }

-- #3 Set data
grid.set{ sheet=1, cell="A2", value="Alpha" }

-- #4 Set data
grid.set{ sheet=1, cell="B2", value="100" }

-- #5 Set data
grid.set{ sheet=1, cell="A3", value="Beta" }

-- #6 Set data
grid.set{ sheet=1, cell="B3", value="200" }
