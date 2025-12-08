# Sequenceschreiber VBA Project

## Overview

This Excel VBA project automates the management of equipment sequences, including method selection, dropdown population, rack assignment, and standard weight calculations. The code is designed to simplify user interaction while maintaining data consistency from an external workbook (Daten für Sequenceschreiber.xlsx).

## Requirements

- Microsoft Excel with VBA support
- Access to the data file: L:\Makros\Sequenceschreiber\Daten für Sequenceschreiber.xlsx
- Write access to C:\TempTest (temporary storage folder)
- Write access to Batchflow in MS-Teams channel TZH ECO Lab [Orga_int] -> Operation

## Usage

- Open the workbook. The Workbook_Open event initializes dropdown lists and device settings. Select a method in the dropdown (cell I3). The worksheet will automatically:
    - Populate rack information
    - Update standard weights for samples
- Select product in the dropdown (cell J3).
- Import batch with button "Datei importieren".
- Changes sample weightings in column C will trigger recalculation of weights.
- Export the batch with button "Sequence exportieren", which automatically send an export to the Batchflow.
- Copy the sequence for the equipment with the button "Sequence Kopieren". 
    - It will save the date in the Clipboard or saves an export in the folder stated in "Daten für Sequenceschreiber.xlsx"
- To reset Sequenceschreiber click the Button "Sequenceschreiber bereinigen".
- Errors are displayed via message boxes and suggest contacting the Digital Laboratory Expert.

## Naming Conventions / Coding Style

To ensure consistency and readability throughout the project, the following naming conventions are used:

Prefix conventions for variables and objects:

| Prefix | Type / Usage                         | Example                     |
|--------|--------------------------------------|-----------------------------|
| ws     | Worksheet objects                    | wsMainPage                  |
| col    | Collection objects                   | colSample, colFinalSequence |
| obj    | Class instances or object references | objBlank                    |
| dict   | Dictionary objects                   | dictMetaData, dictBatchData |
| prv    | Private class member variables       | prvKonzentration            |
| int    | Integer variables                    | intPosition                 |
| lng    | Long integer variables               | lngIndex                    |
| dbl    | Double / floating-point variables    | dblKonzentration            |
| str    | String variables                     | strMethode                  |
| bln    | Boolean variables                    | blnDoProcess                |
| var    | Variant type variables               | varValue                    |
| arr    | Arrays                               | arrValues()                 |
| rng    | Range objects                        | rngSelection                |

Property and method naming:
- camelCase for class properties and public methods (e.g., .setMethodendaten, .AcquisitionMethode)
- camelCase for local variables within procedures (e.g., intPosition, blnDoProcess)

Module naming:
- mod prefix for standard modules (e.g., modUtilities, modQuickSort)
- cls prefix for classes (e.g., clsValues, clsMethodeLoader)

Other rules:
- Constants in PascaleCase (e.g., Sample, Blank)
- Enumerations clearly named and documented in modUtilities

## Files Procedures

# ThisWorkbook:

Handles initialization on workbook open:

- Creates a temporary copy of the workbook in C:\TempTest
- Populates dropdown lists for methods and recent dates
- Reads computer-specific device info from the external data workbook
- Locks and unlocks the main sheet as needed

# wsMainPage (Worksheet Module):

Handles dynamic updates when the user changes data:

- Updates rack information based on the selected method
- Calculates standard weight per sample
- Synchronizes changes with the external data workbook

## Modules

### Main Module

Responsible for the high-level workflow:

- Import: Opens and reads batch files, extracts relevant measurement data, and prepares sequence arrays.
- Sequence: Builds the measurement sequence including blanks, calibrations, special samples.
- Ausdruck (Print/Export): Generates a printable/exportable sequence in Excel and triggers the BatchFlow via HTTP request.
- Kill: Resets all sheets and clears previous data to prepare for a new run.

### modQuickSort

Handles sorting of collections and arrays used in sequence generation. It uses the Quick Sort method.

- defSortCollectionByIndex(col As Collection): Sorts a collection of objects based on the .Index property.
- defQuickSort(arr() As Variant, low As Long, high As Long, Optional isStringArray As Boolean = True): Recursive quicksort for arrays.
- funcGetValueForSorting(item As Variant, isStringArray As Boolean): Returns a comparable numeric value for sorting either string or object arrays.

### modUtilities

Provides utility functions and enumerations for handling sequence properties and measurement types.

- funcGetPropertyName(prp As Properties): Returns the property name as a string.
- funcGetMeasurementType(prp As MeasurementTypes): Returns the measurement type as a string.
- funcGetPosition(Probe As Collection, Collectionindex As Integer, MetaData As Object): Determines the next position for a sample in the sequence.
- defSetValue(...): Sets the value of a measurement object depending on its type (Sample, Calibration, Blank, etc.) and property.
- funcGetMethode(strTopic As Variant, Metadaten As Object):  Returns method of requests product.
- funcGetValue(prp As Properties, Optional ByVal Messung As clsValues = Nothing, Optional ByVal Ganzspalten As Object = Nothing): Returns requested value of given object.
- funcIsFileOpen(filename As String): Returns a boolean value if given workbook ist open or not.
- funcGetMaxPosition(colFinalSequence): Returns highes position value in given sequence.
- funcHasIntermediateCalibration(colSequence): Returns boolean value if sequence has an intermediate calibration.
- funcCloneObject(orig As clsValues, Index As Integer): Clone given object.
- funcIsArrayEmpty(arr As Variant): returns boolean value if given array is empty.
- funcIsOperatorPresent(arr As Variant, strName As String): Check if operator value ist empty.

## Classes

### clsMethodeLoader

Purpose:
This class manages all method-related data for measurement sequences, including batch data, method definitions, sample sequences, calibration, and blank measurements. It organizes and stores data from Excel worksheets and prepares it for processing and sequence generation.

Key Properties / Methods:

- wsMainPage: Reference to the main worksheet containing method data.
- .setDatenWorkbook: Initializes the workbook containing method information.
- .setMethodenZeile: Sets the row containing method definitions.
- .setWertePosition: Determines the start position for values within the method.
- .setExportordner: Sets the folder for exporting data.
- .setMethodendaten: Loads and stores all method-related information into the class instance.

#### Data Structure:
```
dictMetaData
├─ dictBatchData
│  ├─ Gerät
│  ├─ Methode
│  ├─ Topic
│  ├─ Operator
│  ├─ Rack
│  ├─ Position
│  ├─ AnzahlMessungen
│  ├─ Datum
│  └─ TotalProben
├─ wbDaten
├─ wsDaten
├─ Kolonnenposition
│  ├─ AcquisitionMethode 
│  ├─ Quantmethode 
│  ├─ Beschriftung 
│  ├─ Einwaage 
│  ├─ Exctraktionsvolumen 
│  ├─ Injektionsvolumen 
│  ├─ Kommentar 
│  ├─ Konzentration 
│  ├─ Position 
│  ├─ Produktklasse 
│  ├─ Rack 
│  ├─ Typ 
│  ├─ Verdünnung 
│  ├─ Level 
│  ├─ Sequencename 
│  ├─ Info1 
│  ├─ Info2 
│  ├─ Info3 
│  ├─ Info4 
│  ├─ Wert1 
│  ├─ Wert2 
│  ├─ Wert3 
│  └─ Wert4
├─ Methodedaten
│  ├─ Kalibrationsanzahl
│  ├─ Spezialbrobenanzahl
│  ├─ Team
│  ├─ MethodeSTD100
│  ├─ MethodeLeder
│  ├─ MethodeECO
│  ├─ MethodeKalibration
│  ├─ Standardeinwaage
│  ├─ Exctraktionsvolumen
│  ├─ Injektionsvolumen
│  ├─ ProbenTyp
│  ├─ Rackname
│  ├─ RackMin
│  ├─ RackMax
│  ├─ RackPositionen
│  ├─ ZwischenkaliEinzel_Volle
│  ├─ ZwischenkaliQC_Cal
│  ├─ BlankWechsel
│  ├─ KalWechsel
│  ├─ ZwischenBlankTrigger
│  ├─ ZwischenKalibartionTrigger
│  └─ ZwischenKalibartionModus
└─ Trigger
    ├─ MaxKalibration
    ├─ AnzahlProbenZwischenKalibrationen
    └─ AnzahlProbenZwischenBlank
colSample, colSpecialsample, colCalibration and/or objBlank
└─ object instances from clsValues
    ├─ AcquisitionMethode 
    ├─ Quantmethode 
    ├─ Beschriftung 
    ├─ Einwaage 
    ├─ Exctraktionsvolumen 
    ├─ Injektionsvolumen 
    ├─ Kommentar 
    ├─ Konzentration 
    ├─ Position 
    ├─ Produktklasse 
    ├─ Rack 
    ├─ Typ 
    ├─ Verdünnung 
    ├─ Level 
    ├─ Sequencename 
    ├─ Info1 
    ├─ Info2 
    ├─ Info3 
    ├─ Info4 
    ├─ Wert1 
    ├─ Wert2 
    ├─ Wert3 
    ├─ Wert4 
    ├─ Messkategorie 
    └─ Index
colRawSequence
└─ objects from colSample, colSpecialsample, colCalibration and/or objBlank
colFinalSequence
└─ cloned Objects from colRawSequence
 ```

Notes:
The order of method calls matters; for example, .setWertePosition will not work correctly if .setMethodenZeile has not been executed first.
Encapsulates all batch, method, and sequence data for easier handling in the main routine.

### clsValues

Purpose:
Represents a single measurement or value with all related metadata. This class is used to store and retrieve detailed information for each measurement in a structured way.

Notes:
Designed for storing detailed measurement metadata in a structured and reusable way.
Can be used in collections for sequencing, sorting, and processing measurement data.

## Workflow

Procedure: Import()
1. Check Required Information
    - Verify that all required information is present.
    - If information is missing, end the procedure and display an error message.
    - If information is present, continue.
2. Populate Variables
    - Read and populate variables from Sequenceschreiber and the external data sheet.
3. Search for Requested Exports
    - Locate the exports corresponding to the requested data.
4. Handle Multiple Operators
    - Check if more than one operator is listed in the found exports.
    - If multiple operators exist:
        - Ask the user to select their operator.
        - Remove all exports that do not belong to the selected operator.
5. Validate Exports
    - Check if any exports were found.
    - If no exports are found, end the procedure with an error message.
    - If exports are found, continue.
6. Prepare Export Data
    - Determine the operator abbreviation.
    - Write the operator abbreviation to the main page.
    - Sort the exports as needed.
7. Process Each Export
    - For each export:
        - Open the export.
        - Copy sample names to the main page in Sequenceschreiber.
        - Copy product classes to the main page in Sequenceschreiber.
8. Process Each Sample in Export
    - For each sample:
        - Replace commas with dots in the weightings.
        - If the weighing is greater than 50, divide by 1000.
        - Check if the sample is a leader sample:
            - If True:
                - Open Trockenmasse-Original.xlsm.
                - Search for the sample in the workbook.
                    - If sample is found, write the corrected sample weight w * (1 - v%) to the main page.
                    - If sample is not found, write 0.001 to the main page.
                - Close Trockenmasse-Original.xlsm.
            - If False:
                - Sum all weightings for the sample.
                - Write the sum to the main page.
9. Finalize Export
    - Close the current export.
    - Write the export name to the Sequenceschreiber export sheet.
                    
Procedure: Sequence()
1. Initialization
    - Create variables for messages, counters, objects, and database handling.
    - Initialize a dictionary to store the maximum usage of each category.
    - Create a new instance of the clsMethodeLoader class and initialize it with the worksheets and method information.
2. Read Batch Data
    - Load batch information from the database.
    - Check if a method has been selected:
        - If not, display an error message and end the procedure.
        - If yes, continue.
3. Prepare Method and Measurement Values
    - Open the workbook containing method information.
    - Read the rows and columns for the selected method.
    - Load values for blanks, calibrations, special samples, and regular samples into the database.
    - Set triggers and full columns for the sequence.
4. Build Initial Sequence
    - Add starting blanks and initial calibration.
    - Add special samples and another blank if any special samples exist.
    - For each regular sample:
        - Increment the sequence position.
        - Add the sample to the raw sequence.
        - Update triggers for calibration and blank measurements.
        - Check if an intermediate calibration is needed:
            - If yes, add a blank, intermediate calibration, and another blank.
        - Check if an intermediate blank is needed:
            - If yes, add a blank.
    - Add ending blanks and final calibration measurements.
5. Calculate Positions for Categories
    - Create a dictionary to define the maximum usage for each category.
    - Set the starting position based on batch data.
    - For each category from Blank to Sample:
    - Update the positions for each category.
    - Save position information for a position-help message.
6. Sort the Sequence
    - Sort the final sequence (colFinalSequence) by index or the desired property.
7. Write Sequence to Excel
    - Make the target worksheet visible and clear previous contents.
    - Write each measurement from the final sequence into the correct columns.
    - Fill full-batch columns (e.g., sequence name) if applicable.
    - Export the sequence:
        - If no export folder is specified, copy to the clipboard.
        - If an export folder is specified and the file is not open, save as CSV.
        - If the file is open, display an error message.
8. Cleanup
    - Close the method workbook.
    - Release database and dictionary objects (Set ... = Nothing).
    - Restore Excel settings for events, screen updating, and alerts.     

Procedure: Ausdruck()
1. Initialization
    - Disable Excel events, alerts, and screen updating to avoid interruptions during execution.
2. Sequence Preparation
    - Call the Sequence() procedure to generate the measurement sequence.
3. Read Metadata
    - Read equipment name, method, operator, and comment from the wsData sheet.
    - Look up the operator in the wsUser sheet:
        - If found, get the email from the sheet.
        - If not found, ask the user for their email via input box and add it to the sheet.
4. Prepare Worksheet for Export
    - Make the wsAusdruck sheet visible, activate it, and unprotect it.
    - Open the method data workbook and populate the database with batch and method data, including blanks, calibrations, special samples, and regular samples.
5. Copy Sequence Data
    - Make the wsSequence sheet visible.
    - Determine the starting row for the printout.
    - Clear existing content in the target range.
    - Write equipment name, method, operator, timestamp, and save folder to specific cells in the wsAusdruck sheet.
    - For each field in the defined array:
        - If a column exists for the field, copy the corresponding column values from wsSequence to wsAusdruck.
6. Format the Output
    - Loop through each row in the printout range:
        - Apply alternating row background colors for readability.
        - Determine font color based on the sample type:
            - Blue for blanks, red for calibrations, green for special samples, black otherwise.
7. Export the Sequence
    - Protect the wsAusdruck sheet.
    - Copy the sheet to a new workbook.
    - Save the workbook to SharePoint
8. Send to Batchflow
    - Prepare a JSON body containing title, team, comment, and user email.
    - Use an HTTP POST request to send the sequence file to the Batchflow workflow with an PowerAutomate flow.
    - If the request fails (status not 200 or 202), display an error message.
9. Cleanup
    - Close the method data workbook without saving.
    - Hide the wsAusdruck, wsUser, and wsSequence sheets.
    - Display a message confirming the export.
    - Re-enable Excel events, alerts, and screen updating.
                            
Procedure: Kill()
1. Initialization
    - Disable Excel events, alerts, and screen updating to avoid interruptions during the reset.
2. Clear Main Page
    - Clear the main input area (B3:F432).
    - Reset default values for method and standard columns:
        - Set cell I3 to "Methode".
        - Set cell J3 to "STD".
    - Clear additional method-related cells (I4:I5).
    - Set Position in J5 to 1.
    - Update the current date in K9.
3. Clear Ausdruck Sheet
    - Make the Ausdruck sheet visible and unprotect it.
    - Delete all rows starting from row 10 until the first empty row is reached.
    - If additional data exists below the "Name" column header, delete it.
    - Protect the sheet again and hide it.
4. Clear Sequence Sheet
    - Make the Sequence sheet visible.
    - Clear all content in the sheet.
    - Hide the sheet.
5. Cleanup
    - Re-enable Excel events, alerts, and screen updating.

## Keynotes

- Make sure the external workbook is not opened by another process during updates.
- Cells that trigger automatic updates are protected/unprotected dynamically to prevent accidental modification.
- Dropdowns for methods and dates are dynamically generated from the external workbook and system date.

## Troubleshooting

- Error opening external workbook: Ensure Daten für Sequenceschreiber.xlsx exists at the specified path.
- Computer not recognized: The device name is not found in the external data sheet. Contact the Digital Laboratory Expert or added it the data sheet.
- Unexpected behavior: Close all instances of Excel and reopen the workbook.

## Change Log

For detailed change history, see the Git repository.

## Contact

For technical assistance, contact: Digital Laboratory Expert