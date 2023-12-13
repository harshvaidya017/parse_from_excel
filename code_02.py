import pandas as pd

excel_file_path = "C:\\Users\\LENOVO\\Downloads\\Code 2\\Sheet1.xlsx"
excel_file_path1 = "C:\\Users\\LENOVO\\Downloads\\Code 2\\Sheet2.xlsx"

selected_columns = ['Candidate ID', 'Name', 'Email', 'Job Title', 'Company', 'Source', 'Lead owner', 'Added by', 'School 1', 'Degree 1', 'Education Start 1', 'Education End 1', 'Company 1', 'Title 1', 'Experience Start 1', 'Experience End 1','Phone']

df = pd.read_excel(excel_file_path, usecols=selected_columns)
df1 = pd.read_excel(excel_file_path1)

output_dicts = []

for index, row in df.iterrows():
    email_to_search = row['Email']
    phone_to_search = row['Phone'] 

    filtered_df = df1[(df1['Email'] == email_to_search) | (df1['Phone'] == phone_to_search)]

    if not filtered_df.empty:
        matching_jobs = filtered_df['Jobs'].tolist()

        edu_data = row[['School 1', 'Degree 1', 'Education Start 1', 'Education End 1']].rename(
            index={'School 1': 'school', 'Degree 1': 'degree', 'Education Start 1': 'from', 'Education End 1': 'to'}
        ).dropna().to_dict()

        exp_data = row[['Company 1', 'Title 1', 'Experience Start 1', 'Experience End 1']].rename(
            index={'Company 1': 'organization', 'Title 1': 'designation', 'Experience Start 1': 'from', 'Experience End 1': 'to'}
        ).dropna().to_dict()

        output_dict = {
            'id': row['Candidate ID'],
            'name': row['Name'],
            'email': row['Email'],
            'title': row['Job Title'],
            'org': row['Company'],
            'source': row['Source'],
            'lead_owner': row['Lead owner'],
            'added_by': row['Added by'],
            'jobs': matching_jobs,
            'experience': [exp_data],
            'education': [edu_data]
        }

        output_dicts.append(output_dict)

json_data = pd.io.json.dumps(output_dicts, indent=2)


json_file_path = 'output.json'
with open(json_file_path, 'w') as json_file:
    json_file.write(json_data)

print(f"Conversion successful. JSON file saved at: {json_file_path}")
