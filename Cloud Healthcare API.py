import os
import requests
import json
import pandas as pd

# Replace with your actual project ID and location
project_id = 'data-transformer-396716'
location = 'asia-south1'

# Path to the downloaded JSON key file
credentials_path = 'C:\\Users\\zaytona\\Downloads\\data-transformer-396716-5e63dc8f7f6d.json'

# Get the access token
access_token = os.popen('gcloud auth application-default print-access-token').read().strip()

# Set the API endpoint
api_url = f"https://healthcare.googleapis.com/v1/projects/{project_id}/locations/{location}/services/nlp:analyzeEntities"

# Request headers
headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json; charset=utf-8"
}

# Read the Excel file with the columns 'PatientID' and 'PhysicianNotes'
excel_file_path = 'C:\\Users\\zaytona\\my_project\\Sample For AI Bootcamp.xlsx'
document_contents = pd.read_excel(excel_file_path, usecols=['PatientID', 'PhysicianNotes'])

# Create a directory to save JSON files
json_output_dir = 'json_output'
os.makedirs(json_output_dir, exist_ok=True)

# Create a Pandas Excel writer
excel_output_path = 'output.xlsx'
excel_writer = pd.ExcelWriter(excel_output_path, engine='xlsxwriter')

for idx, row in document_contents.iterrows():
    patient_id = row['PatientID']
    document_content = row['PhysicianNotes']
    
    # Request payload
    payload = {
        "nlpService": f"projects/{project_id}/locations/{location}/services/nlp",
        "documentContent": document_content
    }

    # Make the API request
    response = requests.post(api_url, json=payload, headers=headers)

    # Parse the JSON response
    response_data = response.json()

    # Save the response data as a JSON file
    json_output_path = os.path.join(json_output_dir, f"document_{patient_id}_response.json")
    with open(json_output_path, 'w') as json_file:
        json.dump(response_data, json_file, indent=2)

    print(f"Response Data for Document {patient_id} saved to '{json_output_path}'")

    # Extract entity mentions from the response
    entity_mentions = response_data.get('entityMentions', [])
    entity_data = []
    # Append each entity mention to the entity data list
    for mention in entity_mentions:
        entity_data.append({
            "MentionId": mention.get('mentionId'),
            "Type": mention.get('type'),
            "Content": mention.get('text', {}).get('content'),
            "BeginOffset": mention.get('text', {}).get('beginOffset'),
            "LinkedEntities": mention.get('linkedEntities', []),
            "Confidence": mention.get('confidence')
        })

    # Create a DataFrame from the JSON response
    df_json = pd.json_normalize(entity_data)

    # Write the DataFrame to a new sheet in the Excel file
    sheet_name = f'Patient_{patient_id}'
    df_json.to_excel(excel_writer, sheet_name=sheet_name, index=False)

    print(f"JSON Data for Document {patient_id} saved to Excel sheet '{sheet_name}'")

# Close the ExcelWriter to save the Excel file
excel_writer.close()
print(f"JSON Data saved to '{excel_output_path}'")
