import os
import pandas as pd

def load_and_clean_csv(branch, folder_path):
    all_data = []
    
    # Loop through each semester in the folder
    for semester in ['sem1', 'sem3', 'sem5', 'sem7']:
        semester_path = os.path.join(folder_path, branch, semester)
        
        # Loop through each CSV file in the semester folder
        for csv_file in os.listdir(semester_path):
            if csv_file.endswith('.csv'):
                file_path = os.path.join(semester_path, csv_file)
                
                # Load and clean the CSV file
                df = pd.read_csv(file_path)
                
                # Perform necessary cleaning
                # Ensure faculty are split properly if multiple
                if 'Faculty' in df.columns:
                    df['Faculty'] = df['Faculty'].apply(lambda x: x.split('/'))
                
                # Append semester and branch information to each row
                df['Branch'] = branch
                df['Semester'] = semester.split('sem')[1]  # Extract semester number (1, 3, 5, or 7)
                
                all_data.append(df)
    
    # Combine all the cleaned data into a single DataFrame
    combined_df = pd.concat(all_data, ignore_index=True)
    
    # Reorganize columns for the final output
    combined_df = combined_df[['Branch', 'Semester', 'Course Code', 'Course Name', 'L-T-P-S-C', 'Faculty']]
    
    return combined_df

# Root directory where all the data is stored
root_dir = "/Users/arunodayahosakeri/Desktop/Data4timetable"

# List of all branches (folders)
branches = ['cse_a', 'cse_b', 'dsai', 'ece']

# Collect and clean data from each branch
all_cleaned_data = []
for branch in branches:
    branch_data = load_and_clean_csv(branch, root_dir)
    all_cleaned_data.append(branch_data)

# Combine all branch data into one DataFrame
final_df = pd.concat(all_cleaned_data, ignore_index=True)

# Save the cleaned data into a final CSV
final_df.to_csv("compiled_courses.csv", index=False)

print("CSV compiled successfully!")
