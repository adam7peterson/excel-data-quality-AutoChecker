import pandas as pd
import numpy as np
from typing import Dict, List, Tuple
import openpyxl
from openpyxl.utils import get_column_letter


class ExcelQualityChecker:
    """A class to perform quality checks on Excel files."""
    
    def __init__(self, file_path: str):
        """Initialize the checker with an Excel file path."""
        self.file_path = file_path
        self.df = pd.read_excel(file_path)
        self.wb = openpyxl.load_workbook(file_path)
        self.results = {}

    def check_null_values(self) -> Dict[str, float]:
        """Check for null values in each column."""
        null_percentages = (self.df.isnull().sum() / len(self.df) * 100).to_dict()
        self.results['null_values'] = null_percentages
        return null_percentages

    def check_duplicates(self) -> Dict[str, int]:
        """Check for duplicate rows and values in each column."""
        duplicate_rows = len(self.df[self.df.duplicated()])
        duplicate_counts = {col: len(self.df[self.df[col].duplicated()]) 
                          for col in self.df.columns}
        
        self.results['duplicates'] = {
            'duplicate_rows': duplicate_rows,
            'duplicate_values': duplicate_counts
        }
        return self.results['duplicates']

    def check_data_types(self) -> Dict[str, str]:
        """Check data types of each column."""
        data_types = self.df.dtypes.astype(str).to_dict()
        self.results['data_types'] = data_types
        return data_types

    def generate_report(self) -> Dict:
        """Generate a comprehensive quality report."""
        self.check_null_values()
        self.check_duplicates()
        self.check_data_types()
        return self.results


def main():
    """Main function to demonstrate usage."""
    try:
        # Example usage
        checker = ExcelQualityChecker("example.xlsx")
        report = checker.generate_report()
        
        print("=== Excel Data Quality Report ===")
        print("\nNull Values (%):")
        for col, pct in report['null_values'].items():
            print(f"{col}: {pct:.2f}%")
        
        print("\nDuplicates:")
        print(f"Duplicate rows: {report['duplicates']['duplicate_rows']}")
        print("\nDuplicate values by column:")
        for col, count in report['duplicates']['duplicate_values'].items():
            print(f"{col}: {count}")
        
        print("\nData Types:")
        for col, dtype in report['data_types'].items():
            print(f"{col}: {dtype}")
            
    except Exception as e:
        print(f"Error: {str(e)}")


if __name__ == "__main__":
    main()
