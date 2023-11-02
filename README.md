# EMLExtractor
EML attachment extractor

This is a small utility that solved 3 problems for me:

* Having an archive of eml files with a folder structure that is difficult to navigate
* Wanting to extract attachments from this archive
* Handling extensionless files in the resulting files

When choosing the option "Email", it will iterate the folder (and subfolders), read the eml file and extract the attachments into the destination folder.
As I've noticed that there are many duplicates, from conversation threads, it will not override existing files with the same name. Unless you alter the function StartEmail(true) to StartEmail(false).
(it would be more ideal to compare the bytesize, but for my purpose it's not been necessary)

When choosing the "Move" function, it will iterate the folder (and subfolders) and copy *ANY* file to the destination folder.
It will not skip duplicates, and just renames it.

When choosing the "Extensions" function, it will iterate every file without extension and save the file with the chosen extension into the destination path. Generally, it's safe to use the same folder (source and destination) as it will try to find a  unique name if the same filename is in the destination folder.



  
