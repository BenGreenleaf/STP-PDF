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

## Functions

### `connect_to_vault(vault_name, lib='comtypes')`
Connects to the SOLIDWORKS PDM vault using either `comtypes` or `win32com` libraries.

### `recursive_get(sub_assembly_obj, parent_folder_id, destination_folder, existing_parts)`
Recursively traverses the reference tree of a sub-assembly, collects approved part files, and copies them to the destination folder.

### `main(sub_assembly_name, destination_folder)`
Finds the sub-assembly file object in the vault folder, initiates recursive file collection, zips the results, and cleans up temporary files.

---

## Usage
1. Ensure `config.json` is present in the same folder, with correct vault and folder paths.
2. **Run the script or executable in dist/other.exe**
3. **Provide the sub-assembly name** when prompted with or without file extension (only .SLDASM files)


---

## Example `config.json`

```json
{
    "vault_name": "SSPDM",
    "folder_path": "C:\\SSPDM\\Projects\\P0011-NSIP\\Design\\SA - Sub Assembly",
    "dll_path": "C:\\Program Files (x86)\\SOLIDWORKS PDM\\EdmInterface.dll"
}
```

---

## Notes

- Requires the SOLIDWORKS PDM client and access to the vault.
- Tested on Windows with Python and PyInstaller.
- For issues, please contact **Ben Greenleaf**.

## Notes

- Requires the SOLIDWORKS PDM client and access to the vault.
- Tested on Windows with Python and PyInstaller.
- For issues, please contact **Ben Greenleaf**.
