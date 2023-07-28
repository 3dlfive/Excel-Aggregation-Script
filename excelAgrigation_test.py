import unittest
from ExcelAgrigationOFAproval import scan_file_names

class TestScanningFileNames(unittest.TestCase):
    """
    Basic Test class
    """
    def test_scan_file_names(self):
        """
        The actual test
        """
        res = scan_file_names("D:\work\inertiallabs\Scripts\ExcelAgrigationOFAproval")
        self.assertEqual(res, 1)

if __name__ == '__main__':
    unittest.main()