## Manufacturing CLI Tool (internal use)

Run this commmand on terminal:  
```
python cli_tool.py  
```

Or double click on __`cli_tool.py`__  
Choose script from the menu  
Drag/drop or paste the file path  
Enjoy !  
  
  

### Scripts Menu 

<img src="web%20images/menu.PNG" alt="Cli Menu" width="300" />

## Script 1 - Download Images from URL 

Used for downloading images from an Excel sheet containing URLs  
Images can be found inside folder __*'downloaded_images'*__

## Script 2 - Analyze JSON/ZIP 

Used for analyzing test data for PoE EELoad or Adapter Edac from JSON files  
Accepts single JSON/TXT or multiple files inside a ZIP archive  
Results can be found inside folder __*'extracted'*__

## Script 3 - Convert CSV to Excel

Used for converting PoE EELoad data CSV file to an Excel file with custom headers  
Results can be found inside folder __*'extracted'*__

## Script 4 - Split CSV Tests  

Used for splitting a CSV file that contains more than 5000 tests  
This is done in order to upload data to Factory Web succesfully  
Results can be found inside folder __*'extracted'*__

## Script 5 - File Parser  

Used for parsing testing data from raw PoE Network (.zip) or Adapter Edac (.zip) files  
Results can be found inside folder __*'extracted'*__  