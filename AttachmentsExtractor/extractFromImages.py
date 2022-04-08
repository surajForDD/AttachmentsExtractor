import olefile 
import tempfile
import os
import shutil
import zipfile
import glob
from datetime import datetime
import json


error_files =[] 
def extract( path_zipped_xml, destination_folder):
    temp_dir = tempfile.mkdtemp()
    
    destination_folder = os.path.abspath(destination_folder)
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
   



    zip_file = zipfile.ZipFile( path_zipped_xml )
    zip_file.extractall( temp_dir )
    zip_file.close()

    subdir = {
            '.xlsx': 'xl',
            '.xlsm': 'xl',
            '.xltx': 'xl',
            '.xltm': 'xl',
            '.docx': 'word',
            '.dotx': 'word',
            '.docm': 'word',
            '.dotm': 'word',
            '.pptx': 'ppt',
            '.pptm': 'ppt',
            '.potx': 'ppt',
            '.potm': 'ppt',
        }[ os.path.splitext( path_zipped_xml )[ 1 ] ]


    if not os.path.isdir(os.path.join(temp_dir ,subdir ,'media')):
        return False
    # embeddings_dir = os.path.join(temp_dir ,subdir ,'embeddings','*.bin')

    result = {}
    # for bin_file in list( glob.glob( embeddings_dir ) ):
    #     result  = bin_embedding_to_dictionary( bin_file )
    #     if result == None:
    #         error_files.append({"fileName":path_zipped_xml,"obj":bin_file})
    #         continue

    #     with open(os.path.join(destination_folder,result[ 'original_filename' ]),'wb') as file:
    #         file.write(result[ 'contents' ])
    #         file.close()

  
    for entry in os.scandir(os.path.join(temp_dir ,subdir ,'media')):
        if not entry.name.endswith('.bin'):
            shutil.copy(entry.path,destination_folder)
    shutil.rmtree( temp_dir )

    return True
