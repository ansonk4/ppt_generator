import pandas as pd
import random
import numpy as np

def generate_sample_data(num_rows=50):
    """Generate sample data based on the PPT Report Generator requirements"""
    
    # Set random seed for reproducibility
    random.seed(42)
    np.random.seed(42)
    
    data = {}
    
    # Demographics and Background
    data['性別'] = random.choices(['男', '女'], k=num_rows)
    data['Banding'] = random.choices(['Band 1', 'Band 2', 'Band 3'], 
                                   weights=[0.3, 0.4, 0.3], k=num_rows)
    data['學校編號'] = [f'SCH{str(i+1).zfill(3)}' for i in range(num_rows)]
    
    # Parental Education - random text values
    parent_education = ['小學', '中學', '大學', '碩士', '博士', '職業培訓']
    data['父母教育程度'] = random.choices(parent_education, k=num_rows)
    
    # High school subjects - single subject only
    subjects = ['物理', '化學', '生物', '經濟', '歷史', '地理', '資訊及通訊科技']
    data['高中選修學科'] = random.choices(subjects, k=num_rows)
    
    # Academic Scores
    score_options = ['< 25 分', '25-49 分', '50-75 分', '> 75 分']
    data['中文成績'] = random.choices(score_options, k=num_rows)
    data['英文成績'] = random.choices(score_options, k=num_rows)
    data['數學成績'] = random.choices(score_options, k=num_rows)
    
    # Post-DSE Plans (binary indicators)
    plan_columns = ['大學', '副學士', '文憑', '高級文憑', '工作', '工作假期', '其他']
    for col in plan_columns:
        data[col] = random.choices([0, 1], weights=[0.7, 0.3], k=num_rows)
    
    # Post-exam plans - text descriptions
    exam_plans = ['升讀大學', '尋找工作', '海外升學', '職業培訓', '休息一年', '創業']
    data['試後計劃'] = random.choices(exam_plans, k=num_rows)
    
    # Study Locations (binary indicators)
    location_columns = ['香港', '內地', '亞洲', '歐美澳']
    for col in location_columns:
        data[col] = random.choices([0, 1], weights=[0.6, 0.4], k=num_rows)
    
    # Preferred Universities (binary indicators)
    university_columns = ['浸會大學', '中文大學', '城市大學', '教育大學', '恒生大學', 
                         '香港大學', '嶺南大學', '都會大學', '理工大學', '聖方濟各大學', 
                         '樹仁大學', '科技大學', '自資學院']
    for col in university_columns:
        data[col] = random.choices([0, 1], weights=[0.8, 0.2], k=num_rows)
    
    # Academic Preferences - single subject only
    study_subjects = ['電腦工程', '電腦科學', '數學', '金融', '法律']
    preference_columns = ['希望修讀', '希望修讀_A', '希望修讀_B', 
                         '不希望修讀', '不希望修讀_A', '不希望修讀_B']
    for col in preference_columns:
        data[col] = random.choices(study_subjects, k=num_rows)
    
    # Career Guidance Activities
    activity_columns_a = ['大學入學講座_A', '升學展覽_A', '職業博覽_A', 
                         '生涯規劃_A', '團體師友_A', '工作影子_A']
    activity_columns_b = ['大學入學講座_B', '升學展覽_B', '職業博覽_B', 
                         '生涯規劃_B', '團體師友_B', '工作影子_B']
    
    for col in activity_columns_a + activity_columns_b:
        data[col] = random.choices(['有', '沒有'], weights=[0.6, 0.4], k=num_rows)
    
    # Decision Factors - random text values
    factors = ['非常重要', '重要', '一般', '不重要']
    factor_columns = ['學科知識', '院校因素', '大學學費', '助學金', '主要行業', 
                     '朋輩老師', '家庭因素', '預期收入', 'DSE成績', '高中選修科目']
    for col in factor_columns:
        data[col] = random.choices(factors, k=num_rows)
    
    # Greater Bay Area Knowledge
    data['大灣區了解'] = random.choices(['完全不了解', '不太了解', '了解', '非常了解'], 
                                   weights=[0.2, 0.4, 0.3, 0.1], k=num_rows)
    
    # GBA Information Sources
    gba_sources = ['公社科', '內地考察', '政府資訊', '新聞媒體', '網上資訊', 
                   '內地交流', '校內講座', '朋輩及老師']
    for col in gba_sources:
        data[col] = random.choices(['曾經 / 希望參與', '沒有 / 不會參與'], 
                                  weights=[0.4, 0.6], k=num_rows)
    
    # GBA Career Considerations - random text values
    gba_considerations = ['個人興趣及性格_gba', '個人能力_gba', '晉升機會_gba', 
                         '工作性質_gba', '行業前景_gba', '工作環境_gba', 
                         '工作量_gba', '薪水福利_gba', '生活成本_gba', '國家貢獻_gba']
    for col in gba_considerations:
        data[col] = random.choices(['非常重要', '重要', '一般', '不重要'], k=num_rows)
    
    # Job Preferences
    data['工作地方'] = random.choices(['香港', '內地', '國外 - 亞洲', '國外 - 歐美澳'], 
                                  weights=[0.6, 0.2, 0.1, 0.1], k=num_rows)
    
    # Job Importance Ratings
    importance_ratings = ['十分重要', '重要', '不太重要', '不重要']
    job_factor_columns = ['個人能力_B', '個人興趣性格_B', '成就感_B', '家庭因素_B', 
                         '人際關係_B', '工作性質_B', '工作模式_B', '工作量_B', 
                         '工作環境_B', '薪水及褔利_B', '晉升機會_B', '發展前景_B', 
                         '社會貢獻_B', '社會地位_B']
    for col in job_factor_columns:
        data[col] = random.choices(importance_ratings, k=num_rows)
    
    # Career Preferences - single job only
    career_options = ['資訊科技', '電腦工程', '銀行/金融', '創業']
    career_columns = ['希望從事', '希望從事_A', '希望從事_B', 
                     '不希望從事', '不希望從事_A', '不希望從事_B']
    for col in career_columns:
        data[col] = random.choices(career_options, k=num_rows)
    
    data['從事相關工作'] = random.choices(['絕對不會', '可能不會', '不確定', '可能會', '絕對會'], 
                                    weights=[0.1, 0.2, 0.4, 0.2, 0.1], k=num_rows)
    
    # STEM Education
    data['參加STEM'] = random.choices(['有', '沒有'], weights=[0.4, 0.6], k=num_rows)
    
    # STEM Influence (using random numeric values for demonstration)
    data['STEM影響職業選擇程度'] = [random.randint(1, 10) for _ in range(num_rows)]
    
    # STEM Skills - random ratings
    stem_skills = ['領導能力', '團隊合作', '創新思維', '科學知識', '解難能力']
    for col in stem_skills:
        data[col] = random.choices(['優秀', '良好', '一般', '需改善'], 
                                  weights=[0.2, 0.4, 0.3, 0.1], k=num_rows)
    
    # Add some 999 values (Not Applicable) randomly to demonstrate special values
    for col in random.sample(list(data.keys()), k=min(5, len(data.keys()))):
        if isinstance(data[col][0], str):  # Only for string columns
            indices = random.sample(range(num_rows), k=random.randint(1, 3))
            for idx in indices:
                data[col][idx] = '999'
    
    return pd.DataFrame(data)

def main():
    """Generate and save sample data to Excel file"""
    
    print("Generating sample data with 50 rows...")
    df = generate_sample_data(50)
    
    # Save to Excel file
    output_filename = "sample_dse_survey_data.xlsx"
    df.to_excel(output_filename, index=False)
    
    print(f"Sample data saved to '{output_filename}'")
    print(f"Generated {len(df)} rows with {len(df.columns)} columns")
    print("\nColumn names:")
    for i, col in enumerate(df.columns, 1):
        print(f"{i:2d}. {col}")
    
    print(f"\nFirst few rows:")
    print(df.head(3).to_string())
    
    # Display some statistics
    print(f"\nData summary:")
    print(f"- Gender distribution: {df['性別'].value_counts().to_dict()}")
    print(f"- School banding: {df['Banding'].value_counts().to_dict()}")
    print(f"- Work location preferences: {df['工作地方'].value_counts().to_dict()}")
    print(f"- STEM participation: {df['參加STEM'].value_counts().to_dict()}")

if __name__ == "__main__":
    main()