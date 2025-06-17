from pathlib import Path as p
import pandas as pd
import numpy as np

# Get the current script's directory, Import, and Export directories
SCRIPT_DIR = p(__file__).parent
IMPORT_DIR = SCRIPT_DIR / "Import"
EXPORT_DIR = SCRIPT_DIR / "Export"

# Create import and export directories if they do not exist
def ensure_dir():
    IMPORT_DIR.mkdir(exist_ok=True)
    EXPORT_DIR.mkdir(exist_ok=True)


# Allow user to select an Excel file from the Import directory
def select_excel_file():
    excel_files = [f for f in IMPORT_DIR.glob("*.xlsx") if not f.name.startswith("~$")]  # Exclude temporary files in list

    if not excel_files:
        raise FileNotFoundError("No Excel files found in the Import directory.")
    
    if len(excel_files) == 1:
        print(f"Only one Excel file found: {excel_files[0].name}")
        return excel_files[0]

    # Multiple Excel files found. Display available files
    print("\nAvailable Excel files:")
    for idx, file in enumerate(excel_files, 1):
        print(f"{idx}. {file.name}")

    # Prompt user to select a file
    while True:
        try:
            choice = int(input("\nSelect file number (1-{}): ".format(len(excel_files))))
            if 1 <= choice <= len(excel_files):
                return excel_files[choice - 1]
            print("Invalid choice. Please select a valid file number.")
        except ValueError:
            print("Please enter a valid number.")

# Function to make sure there is a unique export file name per run
def get_unique_export_path(base_path: p) -> p:
    if not base_path.exists():
        return base_path
    
    dir = base_path.parent
    name = base_path.stem
    ext = base_path.suffix
    counter = 1

    while True:
        new_path = dir / f"{name} ({counter}){ext}"
        if not new_path.exists():
            return new_path
        counter += 1


def calculate_baseball_stats(df):
    # Convert 'Outs' column to numeric
    df['Outs'] = pd.to_numeric(df['Outs'], errors='coerce')

    # Calculate probablity using the formula: IF(N2>0, 100/(N2+100), ABS(N2)/(ABS(N2)+100)), with N2 being Actual Odds (Home/Away)
    def calc_probailities(row):
        # Calculate home probability
        away_odds = row['Actual Odds (Away)']
        if away_odds > 0:
            away_prob = 100 / (away_odds + 100) # Formula = IF(N2>0, 100/(N2+100)
        else:
            away_prob = abs(away_odds) / (abs(away_odds) + 100) # The rest of the formula = ABS(N2)/(ABS(N2)+100)

        # Calculate home probabilitiy
        home_odds = row['Actual Odds (Home)']
        if home_odds > 0:
            home_prob = 100 / (home_odds + 100) # Formula = IF(N2>0, 100/(N2+100)
        else:
            home_prob = abs(home_odds) / (abs(home_odds) + 100) # The rest of the formula = ABS(N2)/(ABS(N2)+100)
            
        # Return probabilities
        return pd.Series({
            'Away Implied Probability': away_prob,
            'Home Implied Probability': home_prob                        
        })

    # Add probability columns
    prob_cols = df.apply(calc_probailities, axis=1)
    df = pd.concat([df, prob_cols], axis=1)

    # Calculate odds changes (actual)
    df['Odds Change Away (Actual)'] = df['Away Implied Probability'].diff().apply(lambda x: f'{x:.2f}%' if pd.notnull(x) else '')
    df['Odds Change Home (Actual)'] = df['Home Implied Probability'].diff().apply(lambda x: f'{x:.2f}%' if pd.notnull(x) else '')
    
    # Calculate percentage changes
    df['Odds Change Away (%)'] = (df['Away Implied Probability'].pct_change() * 100).round(2).apply(lambda x: f'{x:.2f}%' if pd.notnull(x) else '')
    df['Odds Change Home (%)'] = (df['Home Implied Probability'].pct_change() * 100).round(2).apply(lambda x: f'{x:.2f}%' if pd.notnull(x) else '')

    return df
    

def main():
    # Ensure directories exist
    ensure_dir()

    try:
        input_file = select_excel_file()
        base_output_file = EXPORT_DIR / f"Processed_{input_file.name}"
        output_file = get_unique_export_path(base_output_file)

        print(f"\nProcessing file: {input_file.name}")
        df = pd.read_excel(input_file)
        result = calculate_baseball_stats(df)
        result.to_excel(output_file, index=False)
        print(f"Successfully exported to: '{output_file.name}' in the Export directory.")

    except Exception as e:
        print(f"An error occurred: {str:e}")

   
if __name__ == "__main__":
    main()