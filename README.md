# Projector Finder
## Description
Use this script to generate the data for [BENQ Find Your Perfect Project website](https://www.benq.com/en-us/projector/find-your-perfect-projector.html).  

<img src="docs/projector-finder-website.png" alt="projector-finder-website" width="600"/>

## Installation
Packages:
- python3.6
- pandas

## Usage
Generate `projector_database.json`  
```bash
python build_projector_database.py "path/to/projector_database.xlsx"
```

Generate `result.xlsx`  
```bash
python projector_finder.py "path/to/projector_matches.xlsx"
```
