"""
Example Usage of Preprocessing Pipeline
This script demonstrates how to use individual functions from the preprocessing pipeline.
"""

import json
from preprocess_pipeline import (
    preprocess_document,
    extract_text_from_docx,
    extract_tables_from_excel,
    normalize_to_json,
    enrich_with_metadata
)


def example_single_document():
    """Example: Process a single document"""
    print("Example 1: Processing a single document")
    print("-" * 50)
    
    # Process a Word document
    if __import__('os').path.exists('sample_document.docx'):
        result = preprocess_document('sample_document.docx')
        print(json.dumps(result, indent=2))
    else:
        print("sample_document.docx not found")
    print()


def example_custom_metadata():
    """Example: Extract text and add custom metadata"""
    print("Example 2: Using custom metadata")
    print("-" * 50)
    
    if __import__('os').path.exists('sample_document.docx'):
        # Extract text
        text = extract_text_from_docx('sample_document.docx')
        
        # Create normalized JSON
        data = normalize_to_json("Word", text, {})
        
        # Enrich with custom metadata
        result = enrich_with_metadata(
            data,
            author="John Doe",
            date="2024-02-18",
            source="sample_document.docx"
        )
        
        print(json.dumps(result, indent=2))
    else:
        print("sample_document.docx not found")
    print()


def example_excel_extraction():
    """Example: Extract data from Excel file"""
    print("Example 3: Extracting Excel tables")
    print("-" * 50)
    
    if __import__('os').path.exists('sample_data.xlsx'):
        tables = extract_tables_from_excel('sample_data.xlsx')
        print(json.dumps(tables, indent=2))
    else:
        print("sample_data.xlsx not found")
    print()


if __name__ == "__main__":
    print("=" * 50)
    print("Preprocessing Pipeline - Usage Examples")
    print("=" * 50)
    print()
    
    example_single_document()
    example_custom_metadata()
    example_excel_extraction()
    
    print("=" * 50)
    print("For batch processing, run: python preprocess_pipeline.py")
    print("=" * 50)
