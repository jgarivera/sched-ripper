# SchedRipper2
It is a successor to the SchedRipper CLI tool previously written in Java. SchedRipper2 is primarily written in Python 3.7.4 with the XLSXWriter library.
![Screenshot](docs/sr2.png)

## Features
* It can consume and parse a subject offering JSON file acquired from the APC Masterlist Subject Offerings page to populate SR2's schedule entries.
* It can parse an officers JSON file to serve as metadata for the Excella renderer.
* It can render an Excel spreadsheet of the officers' schedules from the provided officers JSON file.

## Usage
* Download the code and install necessary Python dependencies using `pip`
* Provide the necessary JSON files such as the subject offering response and officers JSON file.
* Edit `run.py` and execute it using `python run.py`
* View your generated Excel spreadsheet with the provided `export_name`.xlsx
* Voila!

## What can be improved and added
* Maybe the schedules can be rendered into something more efficient than Excel. Native UI using `tkinter`?
* Implement algorithm to determine a free schedule slot so the organization knows when every member is free for a particular day.
