# SOLIDWORKS PDM Part Collector

**Author:** Ben Greenleaf

---

## Overview

This module provides functionality to **recursively collect approved part files (PDF and STEP)** from a SOLIDWORKS PDM vault, starting from a specified sub-assembly.

- Connects to the vault
- Traverses the reference tree of assemblies and sub-assemblies
- Copies relevant files to a destination folder
- Zips the result

---

## File Filtering

The code functions by only selecting parts which are:
- Approved
- Have a PDF and STP file generated
- The PDF & STP files are the same version as the part

---

## Usage
1. Ensure `config.json` in dist/config.json is present in the same folder, with correct vault name. The DLL location should be correct so unlikely you will need to change this. Contact Ben Greenleaf if it is not and erroring.
2. **Run the script or executable in dist/main.exe**
3. **Provide the sub-assembly name** when prompted with or without file extension (only .SLDASM files)


---

## Example `config.json`

```json
{
    "vault_name": "SSPDM",
    "dll_path": "C:\\Program Files (x86)\\SOLIDWORKS PDM\\EdmInterface.dll"
}
```

---

## Notes

- Requires the SOLIDWORKS PDM client and access to the vault.
- Tested on Windows with Python and PyInstaller.
- For issues, please contact **Ben Greenleaf**.
