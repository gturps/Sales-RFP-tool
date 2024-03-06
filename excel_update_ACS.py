import json
import os
import logging
import requests
import openai
import openpyxl
import time
from openpyxl.utils import column_index_from_string

# ACS Integration Settings
AZURE_SEARCH_SERVICE = "YourAZureSearchService"
AZURE_SEARCH_INDEX = "YourSearchIndex"
AZURE_SEARCH_KEY = "YourAzureSearchKey"
AZURE_SEARCH_USE_SEMANTIC_SEARCH = "true"
AZURE_SEARCH_SEMANTIC_SEARCH_CONFIG = "default"
AZURE_SEARCH_TOP_K = 5
AZURE_SEARCH_ENABLE_IN_DOMAIN = "true"
AZURE_SEARCH_CONTENT_COLUMNS = ""
AZURE_SEARCH_FILENAME_COLUMN = ""
AZURE_SEARCH_TITLE_COLUMN = ""
AZURE_SEARCH_URL_COLUMN = ""
AZURE_SEARCH_VECTOR_COLUMNS = ""
AZURE_SEARCH_QUERY_TYPE = ""
AZURE_SEARCH_PERMITTED_GROUPS_COLUMN = ""

# AOAI Integration Settings
AZURE_OPENAI_RESOURCE = 'YourAzureOpenAIresource'
AZURE_OPENAI_MODEL = 'gpt-35-turbo'
AZURE_OPENAI_ENDPOINT = 'YourAzureOpenAIresource'
AZURE_OPENAI_KEY = 'YourAzureOpenAIkey'
AZURE_OPENAI_TEMPERATURE = 0
AZURE_OPENAI_TOP_P = 1.0
AZURE_OPENAI_MAX_TOKENS = 200
AZURE_OPENAI_STOP_SEQUENCE = ""
AZURE_OPENAI_SYSTEM_MESSAGE = "You are a sales engineer, working for XXX, a YYY provider. Keep the answer focused on the question, statement. Reformulate the wording with correct British English if needed."
AZURE_OPENAI_PREVIEW_API_VERSION = "2023-06-01-preview"
AZURE_OPENAI_STREAM = "false"
AZURE_OPENAI_MODEL_NAME = 'gpt-35-turbo'
AZURE_OPENAI_EMBEDDING_ENDPOINT = ""
AZURE_OPENAI_EMBEDDING_KEY = ""

SHOULD_STREAM = True if AZURE_OPENAI_STREAM.lower() == "true" else False

# Set query type
query_type = "simple"
if AZURE_SEARCH_QUERY_TYPE:
    query_type = AZURE_SEARCH_QUERY_TYPE
elif AZURE_SEARCH_USE_SEMANTIC_SEARCH.lower() == "true" and AZURE_SEARCH_SEMANTIC_SEARCH_CONFIG:
    query_type = "semantic"

# Function to be called repeatedly, taking variables as input


def my_function(my_sheet, column, destination):
    my_column = my_sheet[column]

    for row_number, cell in enumerate(my_column, start=1):
        value_in_my_column = cell.value
        if value_in_my_column:
            body = {
                "temperature": float(AZURE_OPENAI_TEMPERATURE),
                "max_tokens": int(AZURE_OPENAI_MAX_TOKENS),
                "top_p": float(AZURE_OPENAI_TOP_P),
                "stop": AZURE_OPENAI_STOP_SEQUENCE.split("|") if AZURE_OPENAI_STOP_SEQUENCE else None,
                "stream": SHOULD_STREAM,
                "messages": [
                    {
                        "role": "user",
                        "content": value_in_my_column
                    }
                ],
                "dataSources": [
                    {
                        "type": "AzureCognitiveSearch",
                        "parameters": {
                            "endpoint": f"https://{AZURE_SEARCH_SERVICE}.search.windows.net",
                            "key": AZURE_SEARCH_KEY,
                            "indexName": AZURE_SEARCH_INDEX,
                            "fieldsMapping": {
                                "contentFields": AZURE_SEARCH_CONTENT_COLUMNS.split("|") if AZURE_SEARCH_CONTENT_COLUMNS else [],
                                "titleField": AZURE_SEARCH_TITLE_COLUMN if AZURE_SEARCH_TITLE_COLUMN else None,
                                "urlField": AZURE_SEARCH_URL_COLUMN if AZURE_SEARCH_URL_COLUMN else None,
                                "filepathField": AZURE_SEARCH_FILENAME_COLUMN if AZURE_SEARCH_FILENAME_COLUMN else None,
                                "vectorFields": AZURE_SEARCH_VECTOR_COLUMNS.split("|") if AZURE_SEARCH_VECTOR_COLUMNS else []
                            },
                            "inScope": True if AZURE_SEARCH_ENABLE_IN_DOMAIN.lower() == "true" else False,
                            "topNDocuments": AZURE_SEARCH_TOP_K,
                            "queryType": query_type,
                            "semanticConfiguration": AZURE_SEARCH_SEMANTIC_SEARCH_CONFIG if AZURE_SEARCH_SEMANTIC_SEARCH_CONFIG else "",
                            "roleInformation": AZURE_OPENAI_SYSTEM_MESSAGE,
                            "embeddingEndpoint": AZURE_OPENAI_EMBEDDING_ENDPOINT,
                            "embeddingKey": AZURE_OPENAI_EMBEDDING_KEY
                        }
                    }
                ]
            }
            headers = {
                'Content-Type': 'application/json',
                'api-key': AZURE_OPENAI_KEY,
                "x-ms-useragent": "GitHubSampleWebApp/PublicAPI/2.0.0"
            }

            base_url = AZURE_OPENAI_ENDPOINT if AZURE_OPENAI_ENDPOINT else f"https://{AZURE_OPENAI_RESOURCE}.openai.azure.com/"
            endpoint = f"{base_url}openai/deployments/{AZURE_OPENAI_MODEL}/extensions/chat/completions?api-version={AZURE_OPENAI_PREVIEW_API_VERSION}"

            try:
                feedback = requests.post(endpoint, headers=headers, json=body, timeout=10)
                feedback.raise_for_status()  # will raise an HTTPError if the HTTP request returned an unsuccessful status code
                data = feedback.json()
                last_message_content = data['choices'][-1]['messages'][-1]['content']
                my_sheet.cell(row=row_number, column=column_index_from_string(destination), value=last_message_content)
                time.sleep(1)
            except requests.Timeout:
                print('The request timed out')
            except requests.RequestException as error:
                print(f'An error occurred: {error}')

# Main program


def main():
    filepath = input("What is the name of the file? ")
    worksheet = input("What is the name of the sheet? ")
    source = input("What is the column where the questionnaire is? ")
    destination = input("What is the column where the responses should be? ")

    # Load the workbook and access the desired sheet
    my_workbook = openpyxl.load_workbook(filepath)
    my_sheet = my_workbook[worksheet]

    # Call the function with the opened workbook and sheet
    my_function(my_sheet, source, destination)

    # Save the changes to the workbook
    my_workbook.save(filepath)

    # Close the workbook after finishing the operation
    my_workbook.close()


if __name__ == "__main__":
    main()






