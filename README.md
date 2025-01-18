# update_OCNs

Script to update OCLC numbers in Alma records after a Streamlined Holdings Update project

Updates all 035 fields with an (OCoLC) prefix with the OCNs from the input file, and changes any e-book unique identifiers in 035 fields with the prefix (ViHarT-EM) to also use the new OCN.


## Requirements
Created and tested with Python 3.6; see ```environment.yml``` for complete requirements.

Requires an Alma Bibs API key with read/write permissions. A config file (```local_settings.ini```) with this key should be located in the same directory as the script and input file.

Example of ```local_settings.ini```:

```
[Alma Bibs R/W]
key:apikey
```


## Input and Output
```input.xlsx``` is a spreadsheet listing row numbers in column A, MMS IDs in column B, and the new OCLC number in column C.
The output spreadsheet shows contents of all 035 fields in the records before and after changes.


## Usage
```python update_OCNs.py input.xlsx```


## Contact
Rebecca B. French - <https://github.com/frenchrb>
