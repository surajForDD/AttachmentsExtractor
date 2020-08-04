import olefile 
import tempfile
import os
import shutil
import zipfile
import glob
from datetime import datetime
import json


error_files =[] 
def read_zipped_xml_bin_embeddings( path_zipped_xml, destination_folder):
    temp_dir = tempfile.mkdtemp()

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


    if not os.path.isdir(os.path.join(temp_dir ,subdir ,'embeddings')):
        return False
    embeddings_dir = os.path.join(temp_dir ,subdir ,'embeddings','*.bin')

    result = {}
    for bin_file in list( glob.glob( embeddings_dir ) ):
        result  = bin_embedding_to_dictionary( bin_file )
        if result == None:
            error_files.append({"fileName":path_zipped_xml,"obj":bin_file})
            continue

        with open(os.path.join(destination_folder,result[ 'original_filename' ]),'wb') as file:
            file.write(result[ 'contents' ])
            file.close()

  
    for entry in os.scandir(os.path.join(temp_dir ,subdir ,'embeddings')):
        if not entry.name.endswith('.bin'):
            shutil.copy(entry.path,destination_folder)



    shutil.rmtree( temp_dir )

    return True


def bin_embedding_to_dictionary(path):
    with olefile.OleFileIO(path) as o:
        for entry in o.listdir():
            if str(entry[0]) =='\x01Ole10Native':
                fin = o.openstream(entry)
                fin.seek(6,0)
                result ={}
                result[ 'original_filename' ]=''
                while True:
                        ch = fin.read(1)
                        if ord(ch) == ord('\0'):
                            break
                        result[ 'original_filename' ] +=ch.decode()
                result[ 'original_filepath' ] = '' # original filepath in ANSI is next and is null terminated
                while True:
                    ch = fin.read( 1 )
                    if ord(ch) == ord('\0'):
                        break
                    result[ 'original_filepath' ] += ch.decode()
                fin.seek( 4, 1 ) # next 4 bytes is unused

                temporary_filepath_size = 0 # size of the temporary file path in ANSI in little endian
                temporary_filepath_size |= ord( fin.read( 1 ) ) << 0
                temporary_filepath_size |= ord( fin.read( 1 ) ) << 8
                temporary_filepath_size |= ord( fin.read( 1 ) ) << 16
                temporary_filepath_size |= ord( fin.read( 1 ) ) << 24
                result[ 'temporary_filepath' ] = fin.read( temporary_filepath_size ) # temporary file path in ANSI
                result[ 'size' ] = 0 # size of the contents in little endian
                result[ 'size' ] |= ord( fin.read( 1 ) ) << 0
                result[ 'size' ] |= ord( fin.read( 1 ) ) << 8
                result[ 'size' ] |= ord( fin.read( 1 ) ) << 16
                result[ 'size' ] |= ord( fin.read( 1 ) ) << 24
                result[ 'contents' ] = fin.read( result[ 'size' ] ) 
                return result 

