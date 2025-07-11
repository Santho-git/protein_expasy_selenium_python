# Protein Physicochemical Property Automation using Python & Selenium

This project automates the extraction and analysis of physicochemical properties of proteins by integrating UniProt data with Expasy's ProtParam tool using Python, Selenium, and Excel.

## Overview

The workflow starts with a manually downloaded UniProt Excel file containing protein accession numbers and metadata. The script automates submitting each protein accession number to Expasy ProtParam, retrieves key physicochemical parameters, and organizes the data into a structured Excel workbook with multiple analysis sheets.

Key analyses include:

- Molecular weight  
- Theoretical isoelectric point (pI) and classification (acidic, neutral, basic)  
- Amino acid composition  
- Grand average of hydropathicity (GRAVY) and hydrophobicity classification  
- Instability index and stability classification  
- Aliphatic index and thermostability classification  
- Signal peptide detection by comparing UniProt and Expasy amino acid counts  

## Features

- Reads input protein data directly from an Excel file  
- Automates Expasy ProtParam web tool interactions using Selenium  
- Parses and filters relevant data dynamically  
- Saves results in an Excel workbook with multiple informative sheets  
- Includes automated classification based on biochemical properties  
- Detects presence of signal peptides by sequence length comparison  
- Dynamically selects the latest input file for flexible workflows  

## Requirements

- Python 3.7+  
- Google Chrome browser  
- ChromeDriver (compatible with your Chrome version)  
- Python packages:  
  - selenium  
  - openpyxl  


