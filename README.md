# Geocoding_and_Distance_Matrix


This Python project focuses on manipulating geographical data stored in Excel files, 
with functionalities including geocoding addresses to obtain latitude and longitude coordinates, 
calculating distances between pairs of coordinates using an API, and reorganizing data into a more structured format within the same Excel file.
## 1. Geocoding.py

Python script "Geocoding.py" performs geocoding of addresses stored in an Excel file and adds their corresponding latitude and longitude coordinates to the same file in separate columns.

### Prerequisites

- Python 3.x
- Required Python libraries: openpyxl, geopy, dotenv

### Setup

1. Install the required Python libraries
2. Create a `.env` file in the same directory as the script with the following contents:

- Replace `USER_AGENT` with a unique identifier for your application (e.g., "GeocodingApp").
- Replace `USER_PATH` with the path to your Excel file containing addresses.

### Usage

Run the script using Python

## 2. Geocoding_HERE.py

Python script "Geocoding_HERE.py" geocodes addresses stored in an Excel file using the HERE Geocoding API. It retrieves latitude and longitude coordinates for each address and adds them to the same Excel file in separate columns.

### Prerequisites

- Python 3.x
- Required Python libraries: openpyxl, requests, dotenv

### Setup

1. Install the required Python libraries
2. Obtain an API key from HERE Developer (https://www.here.com/docs/bundle/geocoding-and-search-api-developer-guide/page/topics/quick-start.html) for accessing the HERE Geocoding API.
3. Create a .env file in the same directory as the script with the following contents:
- USER_PATH=path_to_your_excel_file
- HERE_API_KEY=your_here_api_key
- Replace "path_to_your_excel_file" with the path to your Excel file containing addresses.
- Replace "your_here_api_key" with the API key obtained from HERE Developer.

### Usage

Run the script using Python

### Funcionality
1. The script defines a function coordinate(address) that takes an address as input and returns its latitude and longitude coordinates using the HERE Geocoding API.
2. It loads addresses from an Excel file specified in the .env file.
3. For each address, it calls the coordinate() function to obtain the coordinates and stores them in separate lists.
4. It iterates over the lists of coordinates and writes them to the corresponding columns in the Excel worksheet.
5. Finally, it saves the changes to the Excel file and closes it.

## 3. Distance_Matrix.py

Python script "Distance_Matrix.py" calculates the distances between pairs of coordinates stored in an Excel file using the OSRM API and adds the distances to the same file in a separate column.

### Prerequisites

- Python 3.x
- Required Python libraries: openpyxl, requests, dotenv

### Setup

1. Install the required Python libraries
2. Create a `.env` file in the same directory as the script. Replace `USER_PATH` with the path to your Excel file containing addresses.

### Usage

Run the script using Python

## 4. Distance_Matrix_Tabular_Form.py

Python script "Distance_Matrix_Tabular_Form.py" transforms data from one worksheet ("Distance_Matrix") to another ("Distance_Matrix_Tabular_Form") in an Excel workbook.

### Prerequisites

- Python 3.x
- Required Python libraries: openpyxl, dotenv

### Setup

1. Install the required Python libraries:
2. Create a `.env` file in the same directory as the script. Replace `USER_PATH` with the path to your Excel file containing addresses.

### Usage

Run the script using Python
