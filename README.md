# SplitStreet-AddressProcessor
SplitStreet-AddressProcessor is a console application for processing Excel files, splitting addresses into ADDRESS 1 and ADDRESS 2 using configurable headers and patterns. It supports .xlsx and .xls formats and provides robust logging for traceability.

## Features
- Configurable Directories: Define input and output paths in appsettings.json.
- Dynamic Address Splitting: Customize headers and regular expressions.
- Excel Support: Process .xlsx files with ClosedXML and .xls files with NPOI.
- Detailed Logging: Logs actions and errors for better traceability.
## Setup
### Prerequisites
- .NET SDK

## Configuration
Update appsettings.json: { "FileProcessingSettings": { "InputDirectory": "C:/Code/Custom/SplitStreet/Input", "OutputDirectory": "C:/Code/Custom/SplitStreet/Output", "PossibleHeaders": ["ADDRESS 1", "Street Address 1"], "SplitAddressPattern": "\b(Ste|Unit|Apt|Suite|Spc|Bldg|FL|PPW)\s?(\d+[A-Za-z]?)?\b" } }

## Usage
- Place input Excel files in the configured InputDirectory.
- Run the application: dotnet run
- Processed files will be saved in OutputDirectory with the prefix SplitStreet_.
## Logging
The application logs all key actions, including:

- Start and completion times.
- Files processed and skipped.
- Errors encountered during execution.
