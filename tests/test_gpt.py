import unittest
import sheets

class TestWorkbook(unittest.TestCase):
    def setUp(self):
        """Set up a new Workbook instance before each test."""
        self.workbook = sheets.Workbook()

    def test_num_sheets_empty(self):
        """Test that a new workbook has zero sheets."""
        # Arrange (already done in setUp)
        # Act
        num_sheets = self.workbook.num_sheets()
        # Assert
        self.assertEqual(num_sheets, 0)

    def test_list_sheets_empty(self):
        """Test that a new workbook returns an empty list of sheets."""
        # Act
        sheets = self.workbook.list_sheets()
        # Assert
        self.assertEqual(sheets, [])

    def test_new_sheet_auto_name(self):
        """Test that adding a sheet with no name assigns a unique auto-generated name."""
        # Act
        index, name = self.workbook.new_sheet()
        # Assert
        self.assertEqual(index, 0)
        self.assertEqual(name, "Sheet1")
        self.assertIn("sheet1", self.workbook.sheets)

    def test_new_sheet_with_name(self):
        """Test that adding a sheet with a specific name works correctly."""
        # Act
        index, name = self.workbook.new_sheet("TestSheet")
        # Assert
        self.assertEqual(index, 0)
        self.assertEqual(name, "TestSheet")
        self.assertIn("testsheet", self.workbook.sheets)

    def test_new_sheet_name_conflict(self):
        """Test that adding a sheet with a conflicting name raises a ValueError."""
        # Arrange
        self.workbook.new_sheet("Duplicate")
        # Act and Assert
        with self.assertRaises(ValueError):
            self.workbook.new_sheet("duplicate")

    def test_new_sheet_invalid_name(self):
        """Test that adding a sheet with an invalid name raises a ValueError."""
        invalid_names = ["", " Invalid", "Invalid "]
        for name in invalid_names:
            with self.assertRaises(ValueError):
                self.workbook.new_sheet(name)

    # def test_del_sheet(self):
    #     """Test that deleting a sheet removes it from the workbook."""
    #     # Arrange
    #     self.workbook.new_sheet("ToDelete")
    #     # Act
    #     self.workbook.del_sheet("todelete")
    #     # Assert
    #     self.assertNotIn("todelete", self.workbook.sheets)

    # def test_del_sheet_not_found(self):
    #     """Test that deleting a non-existent sheet raises a KeyError."""
    #     # Act and Assert
    #     with self.assertRaises(KeyError):
    #         self.workbook.del_sheet("NonExistent")

    def test_get_sheet_extent(self):
        """Test that getting the extent of a sheet returns correct dimensions."""
        # Arrange
        self.workbook.new_sheet("ExtentTest")
        self.workbook.sheets["extenttest"].num_cols = 5
        self.workbook.sheets["extenttest"].num_rows = 10
        # Act
        extent = self.workbook.get_sheet_extent("ExtentTest")
        # Assert
        self.assertEqual(extent, (5, 10))

    def test_get_sheet_extent_not_found(self):
        """Test that getting the extent of a non-existent sheet raises a KeyError."""
        with self.assertRaises(KeyError):
            self.workbook.get_sheet_extent("NonExistent")

    def test_is_valid_location(self):
        """Test that valid and invalid cell locations are identified correctly."""
        valid_locations = ["A1", "Z9999", "AB12"]
        invalid_locations = ["", "123A", "AAAAA1", "A0"]
        for loc in valid_locations:
            self.assertTrue(self.workbook.is_valid_location(loc))
        for loc in invalid_locations:
            self.assertFalse(self.workbook.is_valid_location(loc))

    def test_set_cell_contents(self):
        """Test that setting cell contents updates the correct cell."""
        # Arrange
        self.workbook.new_sheet("Sheet1")
        # Act
        self.workbook.set_cell_contents("Sheet1", "A1", "Test")
        # Assert
        contents = self.workbook.get_cell_contents("Sheet1", "A1")
        self.assertEqual(contents, "Test")

    def test_set_cell_contents_invalid_sheet(self):
        """Test that setting cell contents for a non-existent sheet raises a KeyError."""
        with self.assertRaises(KeyError):
            self.workbook.set_cell_contents("InvalidSheet", "A1", "Test")

    def test_set_cell_contents_invalid_location(self):
        """Test that setting cell contents for an invalid location raises a ValueError."""
        self.workbook.new_sheet("Sheet1")
        with self.assertRaises(ValueError):
            self.workbook.set_cell_contents("Sheet1", "Invalid", "Test")

    def test_get_cell_contents(self):
        """Test that getting cell contents returns the correct value."""
        # Arrange
        self.workbook.new_sheet("Sheet1")
        self.workbook.set_cell_contents("Sheet1", "A1", "Hello")
        # Act
        contents = self.workbook.get_cell_contents("Sheet1", "A1")
        # Assert
        self.assertEqual(contents, "Hello")

    def test_get_cell_contents_not_found(self):
        """Test that getting cell contents for a non-existent sheet raises a KeyError."""
        with self.assertRaises(KeyError):
            self.workbook.get_cell_contents("NonExistent", "A1")

    def test_get_cell_contents_invalid_location(self):
        """Test that getting cell contents for an invalid location raises a ValueError."""
        self.workbook.new_sheet("Sheet1")
        with self.assertRaises(ValueError):
            self.workbook.get_cell_contents("Sheet1", "Invalid")

if __name__ == "__main__":
    unittest.main()