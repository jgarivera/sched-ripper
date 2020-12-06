# SchedRipper2
It is a successor to the SchedRipper tool written in Java. SR2 is primarily written in Python 3.7.4 and XLSXWriter.

# Features
* It can consume and parse a subject offering JSON file acquired from the APC Masterlist Subject Offerings page to populate SR2's schedule entries
* It can parse an officers JSON file to serve as metadata for Excella.
* It can render an Excel spreadsheet of the officers' schedules from the provided officers JSON file.

# Usage
* Download zip
* Provide the necessary JSON files
* Edit `run.py` and execute it using `python run.py`
* View your generated Excel spreadsheet with the provided `export_name`
* Voila!