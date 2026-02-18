# AI-based-programming-tutorial1

## Team Members
- [Abdallah Zain](https://github.com/Abdallah4Z)
- [Ahmed Islam](https://github.com/ahmedislamfarouk)
- [Alaaelddin Ibrahim](https://github.com/alaaelddin11)
- [Belal Fathy](https://github.com/belal212)
- [Rana Mousa](https://github.com/RanaMousa1)

## Project Description
This project implements a preprocessing pipeline for unstructured data to be used in LLM applications. The pipeline can extract and normalize data from various document types (PDF, Word, Excel, PowerPoint, EPUB) into a standardized JSON format with metadata enrichment.

### Features
- Extract text from multiple document formats (PDF, DOCX, XLSX, PPTX, EPUB)
- Normalize extracted data into JSON format
- Enrich data with metadata (author, date, source)
- Process multiple documents in batch
- Save normalized output as JSON files

## Instructions for Running the Code

### Prerequisites
Install required dependencies:
```bash
pip install -r requirements.txt
```

### Usage
1. Place your documents (PDF, DOCX, XLSX, PPTX, or EPUB files) in the repository directory
2. Run the preprocessing script:
```bash
python preprocess_pipeline.py
```
3. The normalized JSON files will be saved in the `output/` folder

### Running with Custom Files
To process specific files, modify the file paths in the `preprocess_pipeline.py` script or use the functions directly:
```python
from preprocess_pipeline import preprocess_document
import json

# Process a single document
result = preprocess_document('your_file.pdf')
print(json.dumps(result, indent=4))
```

### Example Usage
Run the example script to see different usage patterns:
```bash
python example_usage.py
```
