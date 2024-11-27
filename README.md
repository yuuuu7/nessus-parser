# Nessus-Parser

## Notes from me:
While I do hope that this parser makes reporting easier, I am also human and there may be logical errors within my code (I shat this out in 2 days). I do appreciate it, if you are able to give it a read and feedback to me if there are any issues.

## How to use?
1. Load in your parent folder containing your Raw Nessus Excel/CSV outputs. (You can place your CSVs in sub-folders following their respective System names to make your life easier when debugging)

![image](https://github.com/user-attachments/assets/ff04305a-b86c-4e73-a66c-27c9a39c047a)

2. Run the script. (Make sure you have the required dependencies installed, just Pandas and Openpyxl)

3. Get your formatted output (The script will handle the generation of your output folder):

![image](https://github.com/user-attachments/assets/f588b6b0-b9cd-4009-825d-69651fcac346)

## Other stuff to take note of:
1. `sorted.xlsx` is just `output_results_unique.xlsx` but sorted by name
2. A `.txt` format is given if you prefer to look-at/use it in that format
