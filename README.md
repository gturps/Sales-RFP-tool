# Excel Update ACS Script

## Overview
This script integrates Azure Cognitive Search (ACS) and Azure OpenAI services with an Excel workbook to update cells based on semantic search and GPT-3.5 Turbo model responses.
It's designed for sales engineers and sales teams who need to respond to RFP questionnaires. The answers will be crafted from previous RFP documents available in the ACS index.

## Features
- Integration with Azure Cognitive Search for semantic queries.
- Utilization of Azure OpenAI's GPT-3.5 Turbo model for generating responses.
- Updating Excel cells with responses based on contents of specified columns.

## Prerequisites
- Python 3.8 or higher
- `openpyxl`, `requests`, and `logging` libraries
- Azure Cognitive Search and Azure OpenAI services set up.

## Setup
1. Clone the repository and navigate to the script's directory.
2. Install the required Python packages.
3. Configure the ACS and AOAI integration settings in the script:
   - Azure Search Service, Index, Key, etc.
   - Azure OpenAI resource, model, endpoint, key, etc.
   - Use environment variables instead of keys in the script if you share it with other people.

## Usage
1. Run the script: `python excel_update_ACS.py`.
2. Enter the required information when prompted: file name, sheet name, source column, and destination column.
3. The script updates the Excel workbook's specified cells based on the ACS and Azure OpenAI responses.
4. You might need to update the model and API options as it's been designed in Sep 2023.

## Contributing
Feel free to fork the repository and submit pull requests.

## License
This project is licensed under the MIT License - see the LICENSE file for details.
