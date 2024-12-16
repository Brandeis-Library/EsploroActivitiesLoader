import pandas as pd
import numpy as np
import re
import warnings

# Suppress UserWarnings
warnings.filterwarnings("ignore", category=UserWarning)

def load_researcher_lookup(filename):
    lookup_df = pd.read_excel(filename)
    return dict(zip(lookup_df['Name'], lookup_df['researcherUserID']))

def get_researcher_ids(names, lookup_dict):
    if pd.isna(names):
        return ''
    if not isinstance(names, str):
        names = str(names)
    return ';'.join(lookup_dict.get(name.strip(), '') for name in names.split('\n') if name.strip())

def extract_genr_attributes(attr_string):
    if pd.isna(attr_string):
        return ''
    attributes = set()
    lines = str(attr_string).split('\n')
    for line in lines:
        if 'GENR-' in line:
            parts = line.split('GENR-')
            if len(parts) > 1:
                code = parts[1].split()[0].strip().lower()
                attributes.add(f"activity.{code}")
    return ';'.join(sorted(attributes))

def extract_course_subject(subject):
    if pd.isna(subject):
        return ''
    match = re.search(r'\(([^)]+)\)', str(subject))
    return match.group(1) if match else ''

def extract_course_section(section):
    if pd.isna(section):
        return ''
    parts = [part.strip() for part in str(section).split('-') if part.strip()]
    if len(parts) >= 2:
        return parts[1]
    return ''

def extract_first_word(period):
    if pd.isna(period):
        return ''
    return str(period).split()[0].lower() if period else ''

def transform_data(dataframe, researcher_lookup):
    # Filter out rows based on Section Status and Instructional Format
    dataframe = dataframe[
        (~dataframe['Section Status'].isin(['Canceled', 'Preliminary'])) &
        (~dataframe['Instructional Format'].isin(['Clinical', 'Independent Study', 'Internship']))
    ].copy()

    dataframe.loc[:, 'activityCategory'] = 'activity.teaching'
    dataframe.loc[:, 'activityType'] = 'activity.course'
    dataframe.loc[:, 'activityName'] = dataframe['Section']
    dataframe.loc[:, 'researcherUserID'] = dataframe['Instructors'].apply(lambda x: get_researcher_ids(x, researcher_lookup))
    dataframe.loc[:, 'activityDescription'] = ''
    dataframe.loc[:, 'activityKeywords'] = ''
    dataframe.loc[:, 'activityStartDate'] = pd.to_datetime(dataframe['Start Date'], errors='coerce').dt.strftime('%m/%d/%Y')
    dataframe.loc[:, 'activityEndDate'] = pd.to_datetime(dataframe['End Date'], errors='coerce').dt.strftime('%m/%d/%Y')
    dataframe.loc[:, 'activityAddress1'] = ''
    dataframe.loc[:, 'activityAddress2'] = ''
    dataframe.loc[:, 'activityCity'] = ''
    dataframe.loc[:, 'activityState'] = ''
    dataframe.loc[:, 'activityCountry'] = ''
    dataframe.loc[:, 'activityAttributes'] = dataframe['Course Tags'].apply(extract_genr_attributes)
    dataframe.loc[:, 'activityCourseID'] = dataframe['Course Subject'].apply(extract_course_subject) + ' ' + dataframe['Course Number'].astype(str)
    dataframe.loc[:, 'activityCourseName'] = dataframe['Title']
    dataframe.loc[:, 'activityCourseSection'] = dataframe['Section'].apply(extract_course_section)

    def format_course_type(format_str):
        if pd.isna(format_str):
            return ''
        if format_str.strip().lower() == 'independent study':
            return 'course.independentStudy'
        return 'course.' + format_str.strip().lower().replace(' ', '')

    dataframe.loc[:, 'activityCourseType'] = dataframe['Instructional Format'].apply(format_course_type)
    dataframe.loc[:, 'activityCourseEnrollment'] = dataframe['Enrollment Count']
    dataframe.loc[:, 'activityCourseHours'] = ''
    dataframe.loc[:, 'activityCourseLevel'] = dataframe['Academic Level']
    dataframe.loc[:, 'activityCourseTerm'] = 'term.' + dataframe['Academic Period'].apply(extract_first_word)
    dataframe.loc[:, 'activityLocalField1'] = 'Delivery Mode: ' + dataframe['Delivery Mode'].fillna('')

    for i in range(2, 16):
        dataframe.loc[:, f'activityLocalField{i}'] = ''

    columns = ['activityCategory', 'activityType', 'activityName', 'researcherUserID', 
               'activityDescription', 'activityKeywords', 'activityStartDate', 'activityEndDate',
               'activityAddress1', 'activityAddress2', 'activityCity', 'activityState', 'activityCountry',
               'activityAttributes', 'activityCourseID', 'activityCourseName', 'activityCourseSection',
               'activityCourseType', 'activityCourseEnrollment', 'activityCourseHours', 'activityCourseLevel',
               'activityCourseTerm', 'activityLocalField1']

    for i in range(2, 16):
        columns.append(f'activityLocalField{i}')

    columns += [col for col in dataframe.columns if col not in columns + ['Section', 'Instructors', 'Start Date', 'End Date', 
                                                                          'Course Subject', 'Course Number', 'Title', 'Instructional Format', 
                                                                          'Enrollment Count', 'Academic Level', 'Academic Period', 'Delivery Mode']]

    return dataframe[columns]

def finalize_spreadsheet(dataframe, output_file):
    dataframe.to_excel(output_file, index=False)

# Main execution
input_file = 'WorkdayCourses.xlsx'
output_file = 'esploro_course_loader.xlsx'
lookup_file = 'researcher_lookup.xlsx'

researcher_lookup = load_researcher_lookup(lookup_file)
df = pd.read_excel(input_file)
transformed_df = transform_data(df, researcher_lookup)
finalize_spreadsheet(transformed_df, output_file)

print(f"Transformation complete. Output saved to {output_file}")