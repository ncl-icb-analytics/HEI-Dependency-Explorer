<<<<<<< HEAD

# HEI-Datasets-Exporter

=======

# HEI-Datasets-Exporter

**HealtheIntent Transformation Exporter to .docx with Syntax Highlighting and Dependency Analysis**

## Description:

This script processes SQL transformations from HealtheIntent and organises them into folders. Each folder represents a Workflow, containing a Word document for each dataset.

These documents detail the transformations, with SQL syntax highlighted, and include version and date modified information.

Transformations that do not contain SQL are excluded

 e.g. cerner-defined workflows and file uploads.

Its main purpose is to format SQL queries from the platform for sharing and to track dependencies between datasets.

## Functionality:

    - Converts SQL transformations to .docx files with syntax highlighting.

    - Identifies dependencies between datasets, exporting them to Excel and JSON, and includes recursive dependency tracing.

## Usage:

    1. Run 'HealtheIntent Transformation Exporter - All Workflows' from the 'Python Utilities' collection from HealtheIntent's queries page.

    2. Export as CSV, rename to 'transformations.csv'.

    3. Run 'Table Names' from the same collection, export, and rename to 'table_names.csv'.

    4. Place 'transformations.csv' and 'table_names.csv' in the project directory.

    5. Execute the script.

    6. Output will be saved in the 'output' directory.

### Notes:

    - Requires libraries: pandas, python-docx, sqlparse, and openpyxl.

    - Install libraries with 'pip install -r requirements.txt'.

Author: Eddie Davison

Modified: Nov 2023
