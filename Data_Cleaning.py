import pandas as pd
import re
from datetime import datetime

class Data_Cleaning:
    def __init__(self, excel_file):
        # Load all sheets into a dictionary of DataFrames
        self.sheets = pd.read_excel(excel_file, sheet_name=None)

    def LocationBreakDown(self, df):
        def split_location(location):
            parts = str(location).split(', ')
            if len(parts) == 3:
                return pd.Series(parts, index=['City', 'State', 'Country'])
            elif len(parts) == 2:
                return pd.Series([parts[0], None, parts[1]], index=['City', 'State', 'Country'])
            elif len(parts) == 1:
                return pd.Series([None, None, parts[0]], index=['City', 'State', 'Country'])
            else:
                return pd.Series([None, None, None], index=['City', 'State', 'Country'])

        # Apply the split function to the 'Location' column
        if 'Location' in df.columns:
            location_breakdown = df['Location'].apply(split_location)
            df = pd.concat([df, location_breakdown], axis=1)
        return df

    def ExperienceBreakDown(self, df):
        current_year = datetime.now().year

        def parse_experience(experience_text):
            if pd.isna(experience_text):
                return {
                    "Current Position": None, 
                    "Current Company": None, 
                    "Former Positions": None, 
                    "Years of Experience": None
                }

            segments = [seg.strip() for seg in experience_text.split("\n\n") if seg.strip()]
            if not segments:
                return {
                    "Current Position": None, 
                    "Current Company": None, 
                    "Former Positions": None, 
                    "Years of Experience": None
                }

            current_position = segments[0].split("\n")[0].strip()
            if len(segments[0].split("\n")) > 1:
                raw_company = segments[0].split("\n")[1].strip()
                current_company = raw_company.split('·')[0].strip()
            else:
                current_company = None

            former_positions = [seg.split("\n")[0].strip() for seg in segments[1:]]
            four_digit_numbers = [int(num) for num in re.findall(r'\b\d{4}\b', experience_text)]

            if "Present" in experience_text and four_digit_numbers:
                total_years = current_year - min(four_digit_numbers)
            elif four_digit_numbers:
                total_years = max(four_digit_numbers) - min(four_digit_numbers)
            else:
                total_years = 0

            return {
                "Current Position": current_position,
                "Current Company": current_company,
                "Former Positions": ", ".join(former_positions) if former_positions else None,
                "Years of Experience": total_years
            }

        # Apply the parse function to the 'Experience' column
        if 'Experience' in df.columns:
            parsed_experience = df['Experience'].apply(parse_experience)
            experience_df = pd.DataFrame(parsed_experience.tolist())
            df = pd.concat([df, experience_df], axis=1)
        return df

    def IdentifyCorporateFinanceProfiles(self, df, roles_file):
        with open(roles_file, 'r') as file:
            finance_roles = [line.strip().lower() for line in file if line.strip()]

        def is_finance_related(row):
            combined_text = f"{row.get('Experience', '')} {row.get('About', '')} {row.get('Education', '')}".lower()
            keywords = ["finance", "financial", "fintech"]
            if any(keyword in combined_text for keyword in keywords):
                return "Yes"
            if any(role in combined_text for role in finance_roles):
                return "Yes"
            return "No"

        if {'Experience', 'About', 'Education'}.issubset(df.columns):
            df["In Corporate Finance Industry?"] = df.apply(is_finance_related, axis=1)
        return df

    def CombineAll(self, roles_file, output_xlsx):
        combined_df = pd.DataFrame()  # Create an empty DataFrame to store results

        for sheet_name, df in self.sheets.items():
            # Process the sheet through all functions
            df = self.LocationBreakDown(df)
            df = self.ExperienceBreakDown(df)
            df = self.IdentifyCorporateFinanceProfiles(df, roles_file)

            # Append processed DataFrame to the combined result
            combined_df = pd.concat([combined_df, df], ignore_index=True)

        # Write the combined DataFrame to a new Excel file
        combined_df.to_excel(output_xlsx, index=False, engine='openpyxl')
        print(f"Combined result saved to '{output_xlsx}'.")


# Example usage:
cleaner = Data_Cleaning('comments_pt1.xlsx')
cleaner.CombineAll(r'Input Files\Corporate Finance Jobs.txt', 'cleaned_Combined_attendees.xlsx')
