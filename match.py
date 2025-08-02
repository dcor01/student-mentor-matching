import pandas as pd
from collections import defaultdict

# --- PREPROCESSING LOGIC ---

def preprocess_data(students_df, mentors_df):
    """
    Prepares student and mentor data for matching by standardizing columns,
    converting age to a number, mapping titles to gender, and mapping faculties to campuses.
    """
    # 1. Standardize all column names to lowercase and remove leading/trailing spaces
    students_df.columns = students_df.columns.str.lower().str.strip()
    mentors_df.columns = mentors_df.columns.str.lower().str.strip()
    
    # 2. Process Mentors: Map 'mr./ms.' to a new 'gender' column
    if 'mr./ms.' not in mentors_df.columns:
        raise KeyError("The required 'Mr./ms.' column was not found in the Mentors sheet.")
        
    def map_title_to_gender(title):
        if not isinstance(title, str): return 'Unknown'
        title_clean = title.lower().strip()
        if 'mr' in title_clean: return 'Male'
        # REVERTED as requested: The line below is from your original code.
        if 'ms' in title_clean or 'ms' in title_clean: return 'Female'
        return 'Unknown'
        
    mentors_df['gender'] = mentors_df['mr./ms.'].apply(map_title_to_gender)
    students_df['gender'] = students_df['mr./ms.'].apply(map_title_to_gender)
    print("Successfully mapped 'Mr./ms.' column to a 'gender' column for mentors and student.")
    
    # 3. Process Students: Clean up the gender_preference column
    if 'gender_preference' in students_df.columns:
        students_df['gender_preference'] = students_df['gender_preference'].str.strip()
        print("Successfully cleaned the 'gender_preference' column.")

    # 4. Process Both: Map 'faculty' to a new 'campus' column and clean 'age'
 
    campus_1_keys = ['Business','Medicine']
    campus_2_keys = ['Law']

    def assign_campus(faculty_name):
        if not isinstance(faculty_name, str): return 'Unknown'
        if any(key in faculty_name for key in campus_1_keys): return '1'
        if any(key in faculty_name for key in campus_2_keys): return '2'
        return 'Unknown'

    def shorten_school(faculty_name):
        """
        Checks if a faculty name string contains any of the keys.
        If it does, it returns the specific key that was found.
        """
        if not isinstance(faculty_name, str):
            return 'Unknown'

        # Check Campus 1 keys
        for key in campus_1_keys:
            if key in faculty_name:
                return key  # Return the specific key that matched

        # Check Campus 2 keys
        for key in campus_2_keys:
            if key in faculty_name:
                return key  # Return the specific key that matched

        # If no match was found in either list
        return 'Unknown'
    
    for df in [students_df, mentors_df]:
        if 'faculty' not in df.columns:
            raise KeyError("A 'faculty' column is required but was not found.")
        df['campus'] = df['faculty'].apply(assign_campus)
        df['campus_key'] = df['faculty'].apply(shorten_school)

        if 'age' not in df.columns:
            raise KeyError("An 'age' column is required but was not found.")
        df['age_str'] = df['age'].astype(str).str.extract('(\d+)')
        df['age'] = pd.to_numeric(df['age_str'], errors='coerce')
        df.drop(columns=['age_str'], inplace=True)
        
        if df['age'].isnull().any():
            print(f"Warning: Some 'age' values could not be converted to numbers and were set to NaN.")

    print("Successfully converted 'age' column to a numeric type.")
    print("Successfully mapped 'faculty' column to a 'campus' column for all participants.")
    return students_df, mentors_df


# --- MATCHING LOGIC ---

def match_by_age_and_rules(students_df, mentors_df, max_mentees_per_mentor):
    """
    Matches oldest students with oldest mentors, with a mandatory gender rule
    and a campus preference tie-breaker.
    """
    students_sorted = students_df.sort_values(by='age', ascending=False)
    mentors_sorted = mentors_df.sort_values(by='age', ascending=False)
    
    mentor_mentee_count = defaultdict(int)
    matches = []

    for student_index, student in students_sorted.iterrows():
        best_mentor = None
        highest_score = -1

        for mentor_index, mentor in mentors_sorted.iterrows():            
            # --- HARD CONSTRAINTS CHECK ---
            # Rule 1: Mandatory Gender Preference
            student_pref = student.get('gender_preference')
            mentor_gender = mentor.get('gender')
            
            is_male_preference_miss = (student_pref == 'Male' and mentor_gender != 'Male')
            is_female_preference_miss = (student_pref == 'Female' and mentor_gender != 'Female')

            if is_male_preference_miss or is_female_preference_miss:
                continue 

            # Rule 2: Mentor Capacity
            if mentor_mentee_count[mentor_index] >= max_mentees_per_mentor:
                continue 

            # --- Tie-Breaker Scoring ---
            score = 0
            if student.get('campus') == mentor.get('campus') and student.get('campus') != 'Unknown':
                score = 1 
            
            if score > highest_score:
                highest_score = score
                best_mentor = mentor

        if best_mentor is not None:
            matches.append({
                'Mentor Name': best_mentor.get('name'),
                'Student_Index': student_index,
                'Student_Age': student['age'],
                'Student Gender Preference': student['gender_preference'],
                'Mentor_Index': best_mentor.name,
                'Mentor_Age': best_mentor['age'],
                'Mentor_Title': best_mentor['mr./ms.'],
                'Match_Reason': 'Age Priority + Gender Rule + Campus Preference',
                'Same_Campus_Match': 'Yes' if highest_score > 0 else 'No',
                'Assigned_Campus': student['campus'],
                'Student Phone': student['student_phone'],
                'Student Email': student['student_email'],
                'Student Personal Email': student['student_personal_email'],
                'Disability': student['disability'],
                'Activity Preference': student['activity_1'],
            })
            mentor_mentee_count[best_mentor.name] += 1

    return pd.DataFrame(matches)

# --- Main Execution ---
if __name__ == "__main__":
    MENTOR_CAPACITY = 11
    EXCEL_FILE_NAME = 'mentors.xlsx'
    
    try:
        xls = pd.ExcelFile(EXCEL_FILE_NAME)
        students_df = pd.read_excel(xls, 'Students')
        mentors_df = pd.read_excel(xls, 'Mentors')
        
        # --- RUN PREPROCESSING ---
        print("Starting preprocessing...")
        students_processed, mentors_processed = preprocess_data(students_df, mentors_df)
        
        # --- RUN MATCHING ---
        print("\nPreprocessing complete. Running the matching algorithm...")
        matches_df = match_by_age_and_rules(
            students_processed, 
            mentors_processed, 
            MENTOR_CAPACITY
        )

        # --- SAVE RESULTS ---
        with pd.ExcelWriter('mentor_matches_final.xlsx') as writer:
            matches_df.to_excel(writer, sheet_name='Final Matches', index=False)

        print(f"\nMatching complete! Found and saved {len(matches_df)} matches.")
        print("The results have been saved to 'mentor_matches_final.xlsx'.")

        # --- IDENTIFY AND PRINT UNMATCHED STUDENTS ---
        if not matches_df.empty:
            matched_student_indices = matches_df['Student_Index'].tolist()
            unmatched_students_df = students_processed[~students_processed.index.isin(matched_student_indices)]
        else:
            unmatched_students_df = students_processed

        if not unmatched_students_df.empty:
            print("\n-------------------------------------------------")
            print(f"The following {len(unmatched_students_df)} students could not be matched:")
            print("-------------------------------------------------")
            print(unmatched_students_df[['age', 'gender_preference', 'faculty', 'campus','mr./ms.','name']].to_string())
        else:
            print("\nAll students were successfully matched!")


    except FileNotFoundError:
        print(f"Error: The file '{EXCEL_FILE_NAME}' was not found. Please ensure it is in the same directory.")
    except KeyError as e:
        print(f"Error: A required column could not be found or has a different name. Details: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")
