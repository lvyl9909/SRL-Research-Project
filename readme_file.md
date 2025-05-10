# SRL Data Analysis Toolkit

This toolkit contains multiple Python scripts for analyzing and processing Self-Regulated Learning (SRL) datasets. Below are the functionalities and usage instructions for each script.

## 1. transition.py - Transition Matrix Analysis

### Functionality
This script analyzes action transitions in Excel datasets, creating a transition matrix that displays frequencies and percentages of each behavior code appearing as a subsequent action.

### Usage
```bash
python transition.py <excel_file_path> [-c CASE] [-a ACTION]
```

### Parameters
- `<excel_file_path>`: Path to the Excel file to analyze
- `-c, --case`: Column name identifying cases (default: 'case_id')
- `-a, --action`: Column name for action codes (default: 'SRL_code')

### Output
- Generates an Excel file containing the transition matrix
- The matrix includes frequency (N) and percentage (%) rows at the top
- Output file is named as the original filename plus "_column_transition_matrix.xlsx"

### Features
- Automatically filters out 'R.SL *' codes
- Arranges behavior codes in predefined order
- Generates detailed transition frequency analysis

## 2. compare_phases.py - Phase Proportion Comparison

### Functionality
This script compares the proportions of each Phase between two datasets, calculating Mann-Whitney U test statistics.

### Usage
```bash
python compare_phases.py <file1> <file2> [-p PHASE] [-c CASE]
```

### Parameters
- `<file1>`: Path to the first Excel file
- `<file2>`: Path to the second Excel file
- `-p, --phase`: Column name for Phase (default: 'SRL_Phase')
- `-c, --case`: Column name identifying cases (default: 'case_id')

### Output
- Generates an Excel file "phase_comparison_results.xlsx"
- Contains mean proportions, mean ranks, Z-values, effect sizes, and significance levels for each Phase

### Features
- Automatically identifies stugptviz and recipe4u datasets
- Calculates proportion differences for each Phase across datasets
- Analyzes significance using Mann-Whitney U tests

## 3. code_comparison.py - Code Proportion Comparison

### Functionality
This script compares the proportions of each SRL_code between two datasets, calculating Mann-Whitney U test statistics.

### Usage
```bash
python code_comparison.py <file1> <file2> [-c CODE] [-i ID]
```

### Parameters
- `<file1>`: Path to the first Excel file
- `<file2>`: Path to the second Excel file
- `-c, --code`: Column name for codes (default: 'SRL_code')
- `-i, --id`: Column name identifying cases (default: 'case_id')

### Output
- Generates an Excel file "code_comparison_results.xlsx"
- Contains mean proportions, mean ranks, Z-values, effect sizes, and significance levels for each code

### Features
- Automatically identifies stugptviz and recipe4u datasets
- Excludes 'R.SL *' codes from analysis
- Arranges SRL codes in predefined order
- Analyzes statistical differences in code usage proportions

## 4. process_user.py - User Conversation Data Processing

### Functionality
This script processes student-GPT conversation data, extracting dialogue content from JSON or text files and organizing it into a structured format.

### Usage
```bash
python process_user.py
```

The script looks for conversation data in the "Student_GPT_conversation_by_question" folder and its subfolders in the current directory by default.

### Output
- Generates an Excel file "processed_users_data.xlsx"
- Contains organized conversation data with columns: case_id, week, idx, timestamp, chatgpt_before, user, and chatgpt_after

### Features
- Extracts conversation content from JSON format conversation history
- Adjusts timestamps based on week and conversation sequence
- Optimizes Excel formatting with auto-adjusting row heights

## Data Requirements

All scripts require input data that meets the following conditions:
- Excel format (.xlsx) files
- Contains necessary column names (e.g., case_id, SRL_code, SRL_Phase)
- For comparison scripts, filenames should contain "stugptviz" or "recipe4u" for automatic dataset identification

## Prerequisites

Before running the scripts, ensure you have installed the required Python libraries:

```bash
pip install pandas numpy scipy openpyxl matplotlib
```
