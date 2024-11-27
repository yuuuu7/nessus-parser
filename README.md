# Nessus-Parser
How to use?
1. Load in your parent folder containing your Raw Nessus Excel/CSV outputs. (You can place your CSVs in sub-folders following their respective System names to make your life easier when debugging)
   ![image](https://github.com/user-attachments/assets/ff04305a-b86c-4e73-a66c-27c9a39c047a)

2. Run the script. (Make sure you have the required dependencies installed, just Pandas and Openpyxl)

3. Get your formatted output (The script will handle the generation of your output folder):
   ![image](https://github.com/user-attachments/assets/f588b6b0-b9cd-4009-825d-69651fcac346)

Notes:
1. `sorted.xlsx` is just `output_results_unique.xlsx` but sorted by name
2. A `.txt` format is given if you prefer to look-at/use it in that format
3. Nessus outputs are funky sometimes, there could be a 7-zip vulnerability but the port number will be written as 445. So do take note that the script only handles the parsing and formatting for the most part, it does not handle/fix minor discrepencies that stem from Nessus itself
