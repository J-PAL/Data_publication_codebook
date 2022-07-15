# Data_publication_codebook
 Stata program to make a codebook for a directory that contains stata datasets. 
 
 ## Contents:
 This repo contains two main files: "jpal_codebook.do" and "jpal_codebook.ado." The outputs are identical, the difference is just in use. Adding the "ado" file to your personal ado library will make the command "jp_codebook" available to you to run. If you do not want to add the ado file, the do file can be run by itself.
 
 ## Use:
 
 - ado: run the command "jp_codebook" followed by a string containing the path of the directory you would like to make a codebook for.
 - do: insert a string containing the path of the directory you would like to make a codebook for in the space noted in the first line and then run the entire do-file.
 
 ## Output:
 
 - The program outputs an excel file containing a codebook to the current working directory. The excel file has two tabs:
  - "Variables" contains each distinct variable found in the .dta datasets in the specified folder, along with information about it, including but not limited to label, value label, # of distinct values, and mean, median, etc. for numeric variables
  - "Value labels" contains the value labels used in encoded variables in the dataset, and maps their #s to the underlying values.
