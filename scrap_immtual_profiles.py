import pandas as pd

# File paths
event_attendees_file = "Event Attendees.xlsx"
comments_file = "comments.xlsx"
output_file = "updated_comments.xlsx"

# Load Event Attendees file (all sheets combined into one DataFrame)
event_attendees = pd.concat(pd.read_excel(event_attendees_file, sheet_name=None), ignore_index=True)

# Load Comments file
comments = pd.read_excel(comments_file)

# Ensure column names are consistent (case-insensitive matching)
event_attendees.columns = event_attendees.columns.str.strip().str.lower()
comments.columns = comments.columns.str.strip().str.lower()

# Create a merged DataFrame to store results
updated_comments = comments.copy()

# Match and append data
for index, row in comments.iterrows():
    name = row['commentator name']
    profile_url = row['commentator profile url']

    # Find matching profile in Event Attendees
    match = event_attendees[(event_attendees['name'] == name) & (event_attendees['profile url'] == profile_url)]

    if not match.empty:
        # Add matched columns to the comments DataFrame
        for col in match.columns:
            if col not in updated_comments.columns:
                updated_comments[col] = None
            updated_comments.at[index, col] = match.iloc[0][col]

# Save the updated comments to a new Excel file
updated_comments.to_excel(output_file, index=False)

print(f"Updated comments saved to {output_file}")
