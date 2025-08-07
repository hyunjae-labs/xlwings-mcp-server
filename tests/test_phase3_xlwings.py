"""
Phase 3 xlwings implementation tests for range operations.
Tests merge/unmerge, copy, delete range operations.
"""

import os
import sys
import tempfile
import unittest
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

# Force xlwings mode for testing
os.environ['USE_XLWINGS'] = 'true'

from src.excel_mcp.xlwings_impl.range_xlw import (
    merge_cells_xlw,
    unmerge_cells_xlw,
    get_merged_cells_xlw,
    copy_range_xlw,
    delete_range_xlw,
    batch_range_operations_xlw
)
from src.excel_mcp.xlwings_impl.workbook_xlw import create_workbook_xlw
from src.excel_mcp.xlwings_impl.data_xlw import write_data_to_excel_xlw, read_data_from_excel_xlw


class TestPhase3XlwingsImplementation(unittest.TestCase):
    """Test Phase 3 xlwings implementations."""
    
    def setUp(self):
        """Create a temporary Excel file for testing."""
        self.test_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.test_dir, "test_phase3.xlsx")
        
        # Create test workbook with sample data
        result = create_workbook_xlw(self.test_file)
        self.assertIn("message", result)
        
        # Add test data
        test_data = [
            ["Name", "Age", "City"],
            ["Alice", 30, "Seoul"],
            ["Bob", 25, "Busan"],
            ["Charlie", 35, "Daegu"],
            ["David", 28, "Incheon"]
        ]
        write_result = write_data_to_excel_xlw(self.test_file, "Sheet1", test_data, "A1")
        self.assertIn("message", write_result)
    
    def tearDown(self):
        """Clean up test files."""
        import shutil
        if os.path.exists(self.test_dir):
            shutil.rmtree(self.test_dir, ignore_errors=True)
    
    def test_merge_cells(self):
        """Test merging cells."""
        print("\n=== Testing merge_cells ===")
        
        # Test successful merge
        result = merge_cells_xlw(self.test_file, "Sheet1", "D1", "F3")
        self.assertIn("message", result)
        self.assertIn("Successfully merged", result["message"])
        print(f"✅ Merge successful: {result['message']}")
        
        # Test merging already merged cells (should fail)
        result = merge_cells_xlw(self.test_file, "Sheet1", "D1", "F3")
        self.assertIn("error", result)
        self.assertIn("already merged", result["error"])
        print(f"✅ Duplicate merge prevented: {result['error']}")
        
        # Test invalid sheet
        result = merge_cells_xlw(self.test_file, "InvalidSheet", "A1", "B2")
        self.assertIn("error", result)
        print(f"✅ Invalid sheet handled: {result['error']}")
    
    def test_unmerge_cells(self):
        """Test unmerging cells."""
        print("\n=== Testing unmerge_cells ===")
        
        # First merge cells
        merge_result = merge_cells_xlw(self.test_file, "Sheet1", "G1", "I3")
        self.assertIn("message", merge_result)
        
        # Test successful unmerge
        result = unmerge_cells_xlw(self.test_file, "Sheet1", "G1", "I3")
        self.assertIn("message", result)
        self.assertIn("Successfully unmerged", result["message"])
        print(f"✅ Unmerge successful: {result['message']}")
        
        # Test unmerging non-merged cells (should fail)
        result = unmerge_cells_xlw(self.test_file, "Sheet1", "G1", "I3")
        self.assertIn("error", result)
        self.assertIn("not merged", result["error"])
        print(f"✅ Non-merged cells handled: {result['error']}")
    
    def test_get_merged_cells(self):
        """Test getting merged cells information."""
        print("\n=== Testing get_merged_cells ===")
        
        # Merge multiple ranges
        merge_cells_xlw(self.test_file, "Sheet1", "J1", "K2")
        merge_cells_xlw(self.test_file, "Sheet1", "L3", "M4")
        
        # Get merged cells
        result = get_merged_cells_xlw(self.test_file, "Sheet1")
        self.assertIn("merged_ranges", result)
        self.assertIn("count", result)
        self.assertIsInstance(result["merged_ranges"], list)
        print(f"✅ Found {result['count']} merged ranges")
        for r in result["merged_ranges"]:
            print(f"   - Range: {r['range']} ({r['rows']}x{r['columns']} cells)")
    
    def test_copy_range(self):
        """Test copying range of cells."""
        print("\n=== Testing copy_range ===")
        
        # Test copy within same sheet
        result = copy_range_xlw(
            self.test_file, "Sheet1",
            "A1", "C3",  # Source range
            "E1"         # Target start
        )
        self.assertIn("message", result)
        self.assertIn("Successfully copied", result["message"])
        print(f"✅ Copy successful: {result['message']}")
        
        # Verify copied data
        read_result = read_data_from_excel_xlw(self.test_file, "Sheet1", "E1", "G3")
        self.assertIn("data", read_result)
        original_data = read_data_from_excel_xlw(self.test_file, "Sheet1", "A1", "C3")
        
        # Compare values
        for i, row in enumerate(read_result["data"]):
            for j, cell in enumerate(row):
                self.assertEqual(cell["value"], original_data["data"][i][j]["value"])
        print(f"✅ Copied data verified")
        
        # Test copy to different sheet
        from src.excel_mcp.xlwings_impl.sheet_xlw import create_worksheet_xlw
        create_worksheet_xlw(self.test_file, "Sheet2")
        
        result = copy_range_xlw(
            self.test_file, "Sheet1",
            "A1", "C5",  # Source range
            "B2",        # Target start
            "Sheet2"     # Target sheet
        )
        self.assertIn("message", result)
        print(f"✅ Cross-sheet copy: {result['message']}")
    
    def test_delete_range(self):
        """Test deleting range of cells."""
        print("\n=== Testing delete_range ===")
        
        # Test delete with shift up
        result = delete_range_xlw(
            self.test_file, "Sheet1",
            "B2", "C3",
            "up"
        )
        self.assertIn("message", result)
        self.assertIn("shifted cells up", result["message"])
        print(f"✅ Delete with shift up: {result['message']}")
        
        # Test delete with shift left
        result = delete_range_xlw(
            self.test_file, "Sheet1",
            "D2", "E2",
            "left"
        )
        self.assertIn("message", result)
        self.assertIn("shifted cells left", result["message"])
        print(f"✅ Delete with shift left: {result['message']}")
        
        # Test invalid shift direction
        result = delete_range_xlw(
            self.test_file, "Sheet1",
            "A1", "B2",
            "invalid"
        )
        self.assertIn("error", result)
        print(f"✅ Invalid shift handled: {result['error']}")
    
    def test_batch_operations(self):
        """Test batch range operations for efficiency."""
        print("\n=== Testing batch_range_operations ===")
        
        operations = [
            {
                "type": "merge",
                "sheet_name": "Sheet1",
                "start_cell": "H5",
                "end_cell": "I6"
            },
            {
                "type": "copy",
                "source_sheet": "Sheet1",
                "source_start": "A1",
                "source_end": "B2",
                "target_start": "K1"
            },
            {
                "type": "merge",
                "sheet_name": "Sheet1",
                "start_cell": "L5",
                "end_cell": "M6"
            },
            {
                "type": "delete",
                "sheet_name": "Sheet1",
                "start_cell": "N1",
                "end_cell": "N2",
                "shift_direction": "up"
            }
        ]
        
        result = batch_range_operations_xlw(self.test_file, operations)
        self.assertIn("total_operations", result)
        self.assertIn("successes", result)
        self.assertIn("failures", result)
        self.assertIn("results", result)
        
        print(f"✅ Batch operations: {result['successes']}/{result['total_operations']} succeeded")
        for op_result in result["results"]:
            status = "✅" if op_result["status"] == "success" else "❌"
            print(f"   {status} Operation {op_result['operation']}: {op_result['message']}")
    
    def test_korean_data_handling(self):
        """Test handling of Korean data in range operations."""
        print("\n=== Testing Korean data handling ===")
        
        # Write Korean data
        korean_data = [
            ["이름", "나이", "도시"],
            ["김철수", 30, "서울"],
            ["이영희", 25, "부산"]
        ]
        write_data_to_excel_xlw(self.test_file, "Sheet1", korean_data, "A10")
        
        # Merge cells with Korean data
        result = merge_cells_xlw(self.test_file, "Sheet1", "A10", "C10")
        self.assertIn("message", result)
        print(f"✅ Korean data merge: {result['message']}")
        
        # Copy Korean data
        result = copy_range_xlw(
            self.test_file, "Sheet1",
            "A10", "C12",
            "E10"
        )
        self.assertIn("message", result)
        print(f"✅ Korean data copy: {result['message']}")
        
        # Verify Korean data preserved
        read_result = read_data_from_excel_xlw(self.test_file, "Sheet1", "E10", "G12")
        self.assertEqual(read_result["data"][0][0]["value"], "이름")
        print(f"✅ Korean data preserved after operations")
    
    def test_edge_cases(self):
        """Test edge cases and error handling."""
        print("\n=== Testing edge cases ===")
        
        # Test non-existent file
        result = merge_cells_xlw("non_existent.xlsx", "Sheet1", "A1", "B2")
        self.assertIn("error", result)
        print(f"✅ Non-existent file handled: {result['error']}")
        
        # Test single cell merge (should work but be pointless)
        result = merge_cells_xlw(self.test_file, "Sheet1", "Z1", "Z1")
        self.assertIn("message", result)
        print(f"✅ Single cell merge allowed: {result['message']}")
        
        # Test very large range
        result = copy_range_xlw(
            self.test_file, "Sheet1",
            "A1", "C5",
            "AA1"
        )
        self.assertIn("message", result)
        print(f"✅ Large range copy: {result['message']}")


def run_tests():
    """Run all Phase 3 tests."""
    print("=" * 60)
    print("Phase 3 xlwings Range Operations Tests")
    print("=" * 60)
    
    # Create test suite
    suite = unittest.TestLoader().loadTestsFromTestCase(TestPhase3XlwingsImplementation)
    
    # Run tests with detailed output
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    # Summary
    print("\n" + "=" * 60)
    if result.wasSuccessful():
        print("✅ All Phase 3 tests passed!")
    else:
        print(f"❌ {len(result.failures)} test(s) failed")
        print(f"❌ {len(result.errors)} test(s) had errors")
    print("=" * 60)
    
    return result.wasSuccessful()


if __name__ == "__main__":
    success = run_tests()
    sys.exit(0 if success else 1)