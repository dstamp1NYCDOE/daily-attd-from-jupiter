import pandas as pd


def main(data):
    school_year = data['school_year']
    attendance_filename = f'data/{school_year}/attendance.csv'
    attendance_df = pd.read_csv(attendance_filename)

    attd_pvt = pd.pivot_table(attendance_df, index=['StudentID','LastName','FirstName','Date'], columns=['Type'], values='Section', aggfunc='count').fillna(0)
    attd_pvt["in_class"] = attd_pvt["present"] + attd_pvt["tardy"]

    attd_pvt["present_in_school?"] = attd_pvt["in_class"] >=2
    attd_pvt["DailyAttd"] = attd_pvt["present_in_school?"].apply(lambda x: '' if x else 'A')
    attd_pvt = attd_pvt.reset_index()
    attd_pvt_cols = ["StudentID", "Date", "DailyAttd"]
    attd_pvt = attd_pvt[attd_pvt_cols]

    rosters_df = attendance_df.drop_duplicates(subset=['StudentID','Course'])
    rosters_cols = [
        "StudentID",
        "LastName",
        "FirstName",
        "Course",
        "Section",
        "Period",
        "Teacher",
    ]
    rosters_df = rosters_df[rosters_cols]

    period_3_only_flag = True
    if period_3_only_flag:
        rosters_df = rosters_df[rosters_df['Period'].isin(['3'])]

    for date, student_attd_df in attd_pvt.groupby('Date'):
        filename = f"data/{school_year}/{date}_Daily_Attd_Rosters.xlsx"
        writer = pd.ExcelWriter(filename)

        df = rosters_df.merge(student_attd_df, on=['StudentID'], how='left')
        for teacher, students_df in df.groupby('Teacher'):
            students_df = students_df.sort_values(by=['Period','Course','Section','LastName'])

            students_df.to_excel(writer, sheet_name=teacher,index=False)

        for sheet in writer.sheets:
            worksheet = writer.sheets[sheet]
            worksheet.freeze_panes(1, 3)
            worksheet.autofit()

        writer.close()
    return True

if __name__ == "__main__":
    data = {
        'school_year':'2023_2024'
    }
    main(data)
