# Mortgage Process Automation

This project demonstrates automation of mortgage-related business workflows with Python desktop tools.

## Business task

Manual processing of bank exports and registry updates takes time, requires repetitive operations, and increases the risk of human error. The goal of this solution is to automate file processing, report preparation, and payment distribution across Excel registries.

## Solution

The project includes two main scripts:

### `ipoteka_obrabotka.py`
- reads bank `.txt` exports
- transforms raw data into a structured Excel output
- matches contracts with the main registry
- searches payment schedule files in large network folders
- logs the processing flow

### `ipoteka_raznesenie.py`
- distributes processed payment data across 4 Excel registries
- preserves change history through formulas and comments
- creates backup copies before updates
- supports recovery from backups
- can be packaged into a standalone `.exe`

## Technology stack

`Python` `pandas` `openpyxl` `xlrd` `customtkinter` `logging` `threading` `PyInstaller`

## Highlights

- desktop GUI built with `customtkinter`
- support for network folders and large file structures
- backup and restore logic for safer operations
- audit-friendly logging
- practical focus on financial and back-office workflows

## Files

- `ipoteka_obrabotka.py` - preparation and enrichment of mortgage payment data
- `ipoteka_raznesenie.py` - distribution of payment data across registries
- `ipoteka_obrabotka.PNG` - interface screenshot
- `ipoteka_raznesenie.PNG` - interface screenshot

## Why this project matters

This is not just a study project. It represents a practical business automation case where Python is used to reduce manual work, improve consistency, and support operational reporting.
