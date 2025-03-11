This was a work project for a monthly data entry audit. 

Every month an audit would be conducted to flag missing/incorrect data entry as well as progress towards staff goals. 
This would come in the form of an Excel file with 4-8 cells generated in Tableau Prep Flow, however Tableau Prep cannot highlight or format cells in its output files.

This code does the following
  1. reads the Excel file created by Tableau Prep Flow
  2. highlights any missing/incorrect data (based on manager specifications)
  3. splits the data into Excel files for each site team (and folders if necessary)
  4. creates email text as a Word doc summarizing the finding for each site team
  5. creates an excel sheet summarizing the findings for admin

Permission was given to publish the code on the condition that all personal information is removed and replaced with dummy data.
The dummy data covers 190 clients at 3 work sites, though the original data was a much larger 1200+ clients at 15 work sites.
No included in the dummy data is seasonal data for the Tier2, SSP Goals, Attributes, and Contact Info sheets.
