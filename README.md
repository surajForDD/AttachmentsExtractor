The Objects Embeded in files with extensions .xlsx, .ppt,.doc etc can be extracted using this library. 


list of file extensions supported 
1. .xlsx
2. .xlsm
3. .xltx
4. .xltm
5. .docx
6. .dotx
7. .docm
8. .dotm
9. .pptx
10. .pptm
11. .potx
12. .potm



Sample code to extract embeded files. 



from AttachmentsExtractor import extractor


abs_path_to_file='Please provide absolute path here '
path_to_destination_directory = 'Please provide path of the directory where the extracted attachments should be stored' 


extractor.extract(abs_path_to_file,path_to_destination_directory) # returns true if one or more attachments  are found else returns false.
