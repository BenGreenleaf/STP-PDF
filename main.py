import win32com.client
import os
import shutil
import zipfile
import json
import comtypes.client as cc
import sys
from datetime import datetime



def connect_to_vault(vault_name, lib = 'comtypes'):
    if lib == 'comtypes':
        vault = cc.CreateObject('ConisioLib.EdmVault.1')
        vault.LoginAuto(vault_name, 0)
    else: 
        vault = win32com.client.dynamic.Dispatch('ConisioLib.EdmVault.1')
        vault.LoginAuto(vault_name, 0)
    return vault


def recursive_get(sub_assembly_obj, parent_folder_id, destination_folder, existing_parts, pending_parts, work_in_progress_parts, parts_missing_pdf, parts_missing_step):

    ref_tree = sub_assembly_obj.GetReferenceTree(parent_folder_id)
    project_name, ref_pos = ref_tree.GetFirstChildPosition("", True, True, 0)
    print(f"Referenced parts in sub-assembly {sub_assembly_obj.Name}:")

    destination_folder = os.path.join(destination_folder, sub_assembly_obj.Name.rsplit('.', 1)[0])
    os.makedirs(destination_folder, exist_ok=True)

    while not ref_pos.IsNull:
        ref_item = ref_tree.GetNextChild(ref_pos)
        part_file = ref_item.File
        part_name = part_file.Name.rsplit('.', 1)[0]  # Get base name without extension
        part_type = part_file.Name.rsplit('.', 1)[-1].lower()  # Get file type (e.g., sldprt, sldasm)
        state = getattr(part_file.CurrentState, "Name", "Unknown")  # Get the state name

        if part_type == 'sldprt' and part_name not in existing_parts:

            if state == "Pending Approval":
                pending_parts.append(part_name)
            elif state == "Work In Progress":
                work_in_progress_parts.append(part_name)
            elif state == "Approved":
                
                folder_pos = part_file.GetFirstFolderPosition()
                folder_obj = part_file.GetNextFolder(folder_pos)
                part_folder_id = folder_obj.ID

                try:
                    pdf = folder_obj.GetFile("{}.pdf".format(part_name))
                except:
                    print(f"      {part_name} - PDF not found, skipping")
                    continue

                try:
                    stp = folder_obj.GetFile("{}.step".format(part_name))
                except:
                    print(f"      {part_name} - STEP not found, skipping")
                    continue

                pdf.GetFileCopy(0, "", 0, 0, "")
                stp.GetFileCopy(0, "", 0, 0, "")

                part_modified_date = part_file.GetLocalFileDate(part_folder_id)
                pdf_modified_date = pdf.GetLocalFileDate(part_folder_id) if pdf else None
                stp_modified_date = stp.GetLocalFileDate(part_folder_id) if stp else None

                if part_modified_date is None:
                    print(f"      {part_name} - Part modified date not found")
                    cont = input(f"      Do you want to continue with {part_name}? (y/n): ")
                    if cont.lower() != 'n':
                        continue

                # Check if part was modified after PDF
                if part_modified_date > pdf_modified_date:
                    print(f"      Warning: {part_name} was modified after PDF")
                    parts_missing_pdf.append(part_name)

                if part_modified_date > stp_modified_date:
                    print(f"      Warning: {part_name} was modified after STEP")
                    parts_missing_step.append(part_name)

                cont = 'y'
                if part_modified_date > pdf_modified_date or part_modified_date > stp_modified_date:
                    cont = input(f"      Do you want to continue with {part_name}? (y/n): ")

                if cont.lower() == 'y':
                    print(f"      {ref_item.Name} - State: {state}")
                    # Get the local path of the file in your cache
                    local_pdf_path = pdf.GetLocalPath(part_folder_id)
                    local_stp_path = stp.GetLocalPath(part_folder_id)

                    # Copy from cache to your destination folder
                    shutil.copy(local_pdf_path, destination_folder)
                    shutil.copy(local_stp_path, destination_folder)

                    existing_parts.append(part_name)
            
            else:
                print(f"      {ref_item.Name} - State: {state} (skipped)") 

        elif part_type == 'sldasm':
            # Recursively process sub-assemblies
            folder_pos = part_file.GetFirstFolderPosition()
            folder_obj = part_file.GetNextFolder(folder_pos)
            part_folder_id = folder_obj.ID
            existing_parts, work_in_progress_parts, pending_parts, parts_missing_pdf, parts_missing_step = recursive_get(part_file, part_folder_id, destination_folder, existing_parts, work_in_progress_parts, pending_parts, parts_missing_pdf, parts_missing_step)

    return existing_parts, work_in_progress_parts, pending_parts, parts_missing_pdf, parts_missing_step


# Find the sub-assembly file object
def main(sub_assembly_name, destination_folder):
    enum_pos = folder.GetFirstFilePosition()
    sub_assembly_obj = None
    existing_parts = []
    work_in_progress_parts = []
    pending_parts = []
    parts_missing_pdf = []
    parts_missing_step = []

    while not enum_pos.IsNull:
        file = folder.GetNextFile(enum_pos)
        if file.Name.lower() == sub_assembly_name.lower():
            sub_assembly_obj = file
            break

    if sub_assembly_obj:
        existing_parts, work_in_progress_parts, pending_parts, parts_missing_pdf, parts_missing_step = recursive_get(sub_assembly_obj, folder.ID, destination_folder, [], [], [], [], [])
        
        
        if len(work_in_progress_parts) != 0:
            print(f"Warning: {sub_assembly_obj.Name} Contains work in progress parts which were skipped: {work_in_progress_parts}")
        if len(pending_parts) != 0:
            print(f"Warning: {sub_assembly_obj.Name} Contains pending parts which were skipped: {pending_parts}")
        if len(parts_missing_pdf) != 0:
            print(f"Warning: {sub_assembly_obj.Name} Contains parts missing PDF files: {parts_missing_pdf}")
        if len(parts_missing_step) != 0:
            print(f"Warning: {sub_assembly_obj.Name} Contains parts missing STEP files: {parts_missing_step}")
        

        # Zip the contents of the destination folder
        zip_filename = os.path.join(os.path.dirname(destination_folder), f"{os.path.splitext(sub_assembly_name)[0]}_parts.zip")
        with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(destination_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, destination_folder)
                    zipf.write(file_path, arcname)

        # Clean up by deleting the tmp folder and its contents
        shutil.rmtree(tmp_folder)

    else:
        print("Sub-assembly file not found in folder. Have you set the correct 'assembly_path' in config.json: \"{}\"".format(folder_path))


def find_assembly_file(curr_folder, sub_assembly_name):
    """
    Find the sub-assembly file object in the specified folder.
    """
    enum_pos = curr_folder.GetFirstFilePosition()
    while not enum_pos.IsNull:
        file = curr_folder.GetNextFile(enum_pos)
        if file.Name.lower() == sub_assembly_name.lower():
            folder_pos = file.GetFirstFolderPosition()
            folder_obj = file.GetNextFolder(folder_pos)

            part_folder_id = folder_obj.ID

            # Return the local path of the found file
            local_path = file.GetLocalPath(part_folder_id)
            return local_path
        
    # Iterate through child folders recursively
    subfolder_pos = curr_folder.GetFirstSubFolderPosition()
    while not subfolder_pos.IsNull:
        subfolder = curr_folder.GetNextSubFolder(subfolder_pos)
        found = find_assembly_file(subfolder, sub_assembly_name)
        if found:
            return found
    return None


#P0013-GA-0001
# p0011-sa-0029.sldasm


if __name__ == "__main__":
    try:
        # Create a new 'tmp' folder inside the destination folder
        script_dir = os.path.dirname(sys.executable)
        os.chdir(script_dir)  # Set working directory
        tmp_folder = os.path.join(script_dir, "temp")
        os.makedirs(tmp_folder, exist_ok=True)
        destination_folder = tmp_folder

        config_path = os.path.join(script_dir, 'config.json')
        print(f"Looking for config.json at: {config_path}")
        try:
            with open(config_path, 'r') as f:
                config = json.load(f)
        except Exception as e:
            print(f"ERROR: Could not read config.json: {e}")
            config = {}

        vault_name = config.get('vault_name', 'SSPDM')
        dll_path = config.get('dll_path', r'C:\Program Files (x86)\SOLIDWORKS PDM\EdmInterface.dll')
        root_folder_loc = config.get('root_folder', r'C:\SSPDM')

        cc.GetModule(dll_path)


        vault = connect_to_vault(vault_name)
        root_folder = vault.GetFolderFromPath(root_folder_loc)
        while True:
            sub_assembly_name = input("Enter sub-assembly name: ").lower()
            sub_assembly_name = sub_assembly_name + ".sldasm" if not sub_assembly_name.endswith('.sldasm') else sub_assembly_name
            #main(sub_assembly_name, destination_folder)
            assembly_file = find_assembly_file(root_folder, sub_assembly_name)
            if assembly_file == None:
                print(f"Sub-assembly {sub_assembly_name} not found in the vault.")
            else:
                print(f"Sub-assembly {sub_assembly_name} found at: {assembly_file}")
                folder_path = os.path.dirname(assembly_file)
                folder = vault.GetFolderFromPath(folder_path)
                main(sub_assembly_name, destination_folder)
    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        input("\nPlease report to Ben Greenleaf. Press Enter to exit...")