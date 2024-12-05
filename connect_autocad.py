import win32com.client
import sys

def connect_to_autocad():
    try:
        # Get the AutoCAD application
        acad = win32com.client.Dispatch("AutoCAD.Application")
        
        # Get the active document (drawing)
        doc = acad.ActiveDocument
        
        print("Successfully connected to AutoCAD!")
        print(f"AutoCAD Version: {acad.Version}")
        print(f"Active Document: {doc.Name}")
        
        return acad, doc
    
    except Exception as e:
        print("Error connecting to AutoCAD:")
        print(str(e))
        print("\nMake sure AutoCAD is running before executing this script.")
        sys.exit(1)

def get_all_blocks(doc):
    try:
        # Get the blocks collection
        blocks = doc.Blocks
        block_list = []
        
        # Iterate through all blocks
        for i in range(blocks.Count):
            block = blocks.Item(i)
            block_info = {
                'name': block.Name,
                'is_layout': block.IsLayout,
                'is_xref': block.IsXRef
            }
            # Count the number of attributes in the block
            try:
                attrs = block.GetAttributes()
                block_info['has_attributes'] = len(attrs) > 0
            except:
                block_info['has_attributes'] = False
            
            block_list.append(block_info)
            
        return block_list
        
    except Exception as e:
        print("Error getting blocks:")
        print(str(e))
        return []

def get_active_blocks(doc):
    try:
        # Get the model space
        modelspace = doc.ModelSpace
        block_refs = {}
        
        # Iterate through all objects in model space
        for i in range(modelspace.Count):
            item = modelspace.Item(i)
            # Check if the item is a block reference
            if item.ObjectName == "AcDbBlockReference":
                block_name = item.EffectiveName
                if block_name in block_refs:
                    block_refs[block_name] += 1
                else:
                    block_refs[block_name] = 1
        
        return block_refs
            
    except Exception as e:
        print("Error getting active blocks:")
        print(str(e))
        return {}

def get_inactive_blocks(doc):
    """Returns a list of block names that are not used in the drawing"""
    # Get both lists
    all_blocks = get_all_blocks(doc)
    active_blocks = get_active_blocks(doc)
    
    # Find blocks that are not active
    inactive_blocks = [block['name'] for block in all_blocks 
                      if block['name'] not in active_blocks.keys() 
                      and not block['is_layout']  # Exclude layout blocks
                      and not block['is_xref']]   # Exclude xrefs
    
    return inactive_blocks

def delete_blocks(doc, block_names):
    """Delete specified blocks from the drawing
    Args:
        doc: AutoCAD document object
        block_names: List of block names to delete
    Returns:
        tuple: (number of blocks deleted, list of blocks that couldn't be deleted)
    """
    try:
        blocks = doc.Blocks
        deleted_count = 0
        failed_blocks = []
        total_blocks = len(block_names)
        
        print(f"\nStarting deletion of {total_blocks} blocks...")
        
        for i, block_name in enumerate(block_names, 1):
            try:
                # Try to get the block
                block = blocks.Item(block_name)
                # Check if it's safe to delete
                if not block.IsLayout and not block.IsXRef:
                    print(f"Purging block {i}/{total_blocks}: {block_name}")
                    # Use PURGE command to remove the block
                    doc.SendCommand(f'._PURGE\nB\n{block_name}\nN\nY\n')
                    # Verify if block was actually deleted
                    try:
                        blocks.Item(block_name)
                        print(f"✗ Block {i}/{total_blocks}: {block_name} (Block still exists after purge)")
                        failed_blocks.append(f"{block_name} (Purge failed)")
                    except:
                        deleted_count += 1
                        print(f"✓ Block {block_name} purged successfully")
                else:
                    print(f"✗ Block {i}/{total_blocks}: {block_name} (Layout or XRef - cannot delete)")
                    failed_blocks.append(f"{block_name} (Layout or XRef)")
            except Exception as e:
                print(f"✗ Block {i}/{total_blocks}: {block_name} (Error during deletion)")
                failed_blocks.append(f"{block_name} (Error: {str(e)})")
        
        return deleted_count, failed_blocks
        
    except Exception as e:
        print(f"Error during block deletion: {str(e)}")
        return 0, block_names

if __name__ == "__main__":
    acad, doc = connect_to_autocad()
    
    # Example usage:
    inactive_blocks = get_inactive_blocks(doc)
    if inactive_blocks:
        print("\nFound inactive blocks:", inactive_blocks)
        response = input("\nDo you want to delete these blocks? (yes/no): ")
        if response.lower() == 'yes':
            deleted, failed = delete_blocks(doc, inactive_blocks)
            print(f"\nDeletion complete!")
            print(f"Successfully deleted: {deleted} blocks")
            if failed:
                print("\nFailed to delete the following blocks:")
                for block in failed:
                    print(f"- {block}")
    else:
        print("\nNo inactive blocks found in the drawing.") 