import os

def list_files_in_sql_in_syntax(folder_path):
    filenames = os.listdir(folder_path)
    filenames_without_extension = [os.path.splitext(f)[0] for f in filenames]
    sql_in_clause = "IN ('{}')".format("','".join(filenames_without_extension))
    return sql_in_clause

if __name__ == "__main__":
    folder_path = r"C:\Users\EddieDavison\NHS\HealtheAnalytics Workstream - LTC LCS Workstream\QA\Case Finding Dashboard\September\Transformation SQL"
    result = list_files_in_sql_in_syntax(folder_path)
    print(result)
