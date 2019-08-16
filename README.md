**Status:** Expect regular updates and bug fixes.

# Utilities for interacting with OpenDSS

OpenDSS is a distribution system simulation platform. This repository provides a few methods to interact with OpenDSS using Python through the COM interface.

## Usage
You can clone the repository from GitHub with following commands:
```
git clone https://github.com/sibyjackgrove/OpenDSS-utilities.git
```
Place the *OpenDSS_basics.py* within the desired folder and import the *DSS* class as shown :

```python
from OpenDSS_utilities import DSS
```

Dependencies:  [OpenDSS](https://sourceforge.net/projects/electricdss/files/), [pywin32](https://pypi.org/project/pywin32/)

***
***Note:*** I could not get pywin32 to work on Python 3.5 or higher. So I could only test the code on Python 2.7 (32 bit).

***

## Examples
Usage of the class is descirbed in this Jupyter notebook.

[Basic usage](OpenDSS_with_Python_basic_usage.ipynb)

## Issues
Please feel free to raise an issue for bugs or feature requests.

## Who is responsible?

**Core developer:**
- Siby Jose Plathottam splathottam@anl.gov
