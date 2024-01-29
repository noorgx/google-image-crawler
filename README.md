# Script README 🚀

## Introduction
This Python script performs the following tasks:
1. Takes an input Excel file containing names in the first three columns.
2. Performs Google Image searches for each name and extracts image URLs.
3. Filters and sorts the images based on color and resolution.
4. Writes the results to an output Excel file.

## Requirements 📦
Make sure you have the following libraries installed:
- `pandas`
- `deep_translator`
- `Pillow`
- `requests`
- `tqdm`

You can install them using the following:
```bash
pip install -r requirements.txt
```

Also, ensure that you have emojis support in your environment. 🌈

## Usage 🚀
Run the script with the following command in the terminal:
```bash
python script.py input_excel output_excel_name.xlsx
```

- `input_excel`: The path to the input Excel file.
- `output_excel_name.xlsx`: The desired name for the output Excel file.

## Script Workflow 🔄
1. The script reads the input Excel file and extracts names from the first three columns.
2. For each name, it performs Google Image searches to obtain image URLs.
3. The images are filtered and sorted based on color and resolution.
4. The script writes the results to the output Excel file.

## Important Notes ℹ️
- The script uses Google Image search and may be subject to rate limiting or changes in the Google search interface.
- Internet connectivity is required to fetch images.
- Some features in the script (commented out) can be uncommented based on specific requirements.

## Emojis Support 🌟
This script also supports emojis in your environment. Feel free to use emojis in the input Excel file or customize the script accordingly. 😎

## Example 🚀
```bash
python script.py input_data.xlsx output_results.xlsx
```

## Troubleshooting 🛠️
If you encounter any issues, make sure you have fulfilled the requirements and double-check the input parameters.

## Acknowledgments 🙌
- This script utilizes external libraries and services, including Google Image search.
- The script may need adjustments if there are changes in the Google search interface or other external factors.

**Note:** Always be mindful of the terms of service when using automated tools to interact with online services. 🚨
