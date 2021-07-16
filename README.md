# cmrt-consolidation

This is a program written by Lauren Cornell used to consolidate the CMRT files received from various suppliers. It takes a folder of responses and consolidates the Smelter List sheet into a new CMRT file, which is saved in the same location as the program. Should any smelter on the Smelter List not be included in the Smelter Look-Up sheet of the user provided CMRT template, this smelter is considered "unapproved" and saved in a separate file called unapproved_smelters. This file will also list which response file contained the unapproved smelter.

Each smelter is identified by an ID number. Only unique IDs are saved to the final file.

![image](https://user-images.githubusercontent.com/57767069/124303847-2ae4b480-db20-11eb-9820-41fc1e0c31cf.png)

Responses folder: The folder containing all CMRT responses, saved in xlsx format.

CMRT template: The empty CMRT file with the desired Smelter Look-Up sheet. This is the data that will be used to determine if a smelter is considered approved or not. The most up to date template can be downloaded directly from the [Responsible Minerals Initiative website](http://www.responsiblemineralsinitiative.org/reporting-templates/cmrt/?).

Save as: The name the newly generated consolidated list should be saved as. This file can be found in the same location as the program.
