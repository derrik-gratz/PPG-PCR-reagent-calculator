# PPG-restriction-digest-reaction-reagent-calculator
A program to calculate restriction digest reaction reagent quantities for use at Paw Print Genetics

The assay dictionary is a reference file that can be accessed and changed by lab techs at PPG to update reagent information.

The excel template serves as a formated output file that the program adds information to.

For each 'run', an excel file is provided by our LIMS. After being pruned in excel, the input file can be passed into the program. The program references the assay dictionary to make the output file, which is saved in the directory of the input file. 
