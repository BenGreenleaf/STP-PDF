Author: Ben Greenleaf
This module provides functionality to recursively collect approved part files (PDF and STEP) from a SOLIDWORKS PDM vault, starting from a specified sub-assembly. It connects to the vault, traverses the reference tree of assemblies and sub-assemblies, and copies relevant files to a destination folder. The script checks the state of each part (Approved, Pending Approval, Work In Progress), warns if the part was modified after its PDF/STEP export, and allows user confirmation before proceeding. After collecting files, it zips the destination folder and cleans up temporary files.
Functions:
    connect_to_vault(vault_name, lib='comtypes'):
        Connects to the SOLIDWORKS PDM vault using either comtypes or win32com libraries.
    recursive_get(sub_assembly_obj, parent_folder_id, destination_folder, existing_parts):
        Recursively traverses the reference tree of a sub-assembly, collects approved part files, and copies them to the destination folder.
    main(sub_assembly_name, destination_folder):
        Finds the sub-assembly file object in the vault folder, initiates recursive file collection, zips the results, and cleans up temporary files.
Usage:
    Run the script, provide the sub-assembly name when prompted, and ensure config.json is present with correct vault and folder paths.
