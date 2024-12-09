import json
import pandas as pd
import os
from pathlib import Path

# Load .har file
def convertHarToExcel(harFile):
    # Use the harFile argument to open the actual file
    with open(harFile, 'r', encoding='utf-8') as f:
        harData = json.load(f)

    # Map for entry type
    typeMap = {
        'text/html': 'document',
        'text/css': 'stylesheet',
        'text/javascript': 'script',
        'image/png': 'png',
        'application/octet-stream': 'font',
        'application/json': 'xhr',
        'image/svg+xml': 'fetch',
        'application/javascript': 'javascript',
        'font/woff2': 'font'
    }

    # Extract relevant data
    entries = harData['log']['entries']
    data = []
    for entry in entries:
        name = entry['request']['url'].split('/')[-1]
        if not name:
            name = entry['request']['url'].split('/')[-2] + '/'
        elif name.startswith('?'):
            name = entry['request']['url'].split('/')[-2] + '/' + entry['request']['url'].split('/')[-1]

        status = entry['response']['status']
        if status == 0:
            status = 'Finished'

        size = entry['response'].get('_transferSize')
        if size > 1000000:
            size = str(round(size/1000000, 1)) + ' MB'
        elif size > 1000:
            size = str(round(size/1000, 1)) + ' kB'
        else:
            size = str(size) + ' B'

        mimeType = entry['response']['content'].get('mimeType') 
        if mimeType == 'x-unknown':
            mimeType = entry['_resourceType']
        mime = typeMap.get(mimeType, mimeType)

        initiator = 'Other'
        if '_initiator' in entry:
            initiatorObj = entry['_initiator']
            if 'url' in initiatorObj:
                initiator = initiatorObj['url']
            elif 'stack' in initiatorObj:
                callFrames = initiatorObj['stack'].get('callFrames', [])
                if callFrames:
                    initiator = callFrames[0].get('url', 'Other')

        time = round(entry['time'] - entry['timings'].get('_blocked_queueing'))
        if time > 10000000:
            time = 'Pending'
        elif time >= 60000:
            time = str(round(time/60000, 1)) + ' min'
        elif time > 1000 and time < 60000:
            time = str(round(time/1000, 2)) + ' s'
        else:
            time = str(time) + ' ms'

        data.append({
            'Name': name,
            'Status': status,
            'Type': mime,
            'Initiator': initiator,
            'Size': size,
            'Time': time
        })

    # Create a DataFrame for better visualization
    df = pd.DataFrame(data)

    # Create new folder
    outputDir = Path('convertToExcel.py').resolve().parent / "CONVERTED FILES"
    outputDir.mkdir(exist_ok=True) 

    # Save to Excel with the same name as the HAR file
    outputFile = outputDir / f"{Path(harFile).stem}.xlsx"
    df.to_excel(outputFile, index=False)

def processHarFiles():
    # Define the path for the .har files
    harFolder = Path('convertToExcel.py').resolve().parent / "DROP HAR FILES"
    
    # Get all .har files in the folder
    harFiles = [f for f in os.listdir(harFolder) if f.endswith('.har')]
    
    # Process each .har file
    for harFile in harFiles:
        harFilePath = os.path.join(harFolder, harFile)
        print(f"Processing {harFilePath}...")
        convertHarToExcel(harFilePath)

if __name__ == "__main__":
    processHarFiles()