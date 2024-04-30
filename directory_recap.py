import os
import csv

def files_to_csv(directory, output_csv):
    # Dictionary to store filenames without extensions as keys and a set of extensions as values
    files_dict = {}

    # Walk through the directory
    for filename in os.listdir(directory):
        if os.path.isfile(os.path.join(directory, filename)):
            filename=filename.lower()
            # Split the filename into name and extension
            file_base, file_ext = os.path.splitext(filename)
            file_ext = file_ext.lstrip('.')  # Remove the dot from the extension
            if file_ext=="txt" or file_ext=="doc" or file_ext=="docx" or file_ext=="rtf" or file_ext=="rol" :
                if file_base in files_dict:
                    # If the file_base is already in the dict, append the new extension
                    files_dict[file_base].add(file_ext)
                else:
                    # Otherwise, create a new set with the current extension
                    files_dict[file_base] = {file_ext}

    # Finding all unique extensions for headers in the CSV
    all_extensions = set()
    for extensions in files_dict.values():
        all_extensions.update(extensions)
    all_extensions = sorted(all_extensions)

    # Write to CSV
    with open(output_csv, 'w', newline='') as csvfile:
        fieldnames = ['filename'] + all_extensions
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

        writer.writeheader()
        for file_base, extensions in files_dict.items():
            # Create a row for each file
            row = {'filename': file_base}
            # Mark which extensions this file has
            for ext in extensions:
                row[ext] = 'Yes'
            writer.writerow(row)

# Example usage
files_to_csv('COMPTAGE', 'recap.csv')
