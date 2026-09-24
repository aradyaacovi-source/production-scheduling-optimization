import time

from Tapi_Project_Utils import (
    lp_cache,
    solve_lp_subproblem,
    _read_input_from_workbook_pandas,
    _write_results_workbook,
)


def Run_Solver(Input_File_Name, Output_File_Name):
    """
  Stage 1: solves one fixed schedule (jobs in input order) and writes Excel results.

  Parameters:
      Input_File_Name: base name of the input .xlsx (without extension).
      Output_File_Name: base name for the results .xlsx to create.

  Returns:
      Nothing; creates the output workbook on disk.
  """
    run_start = time.perf_counter()

    input_path = f"{Input_File_Name}.xlsx"
    t, d, v, w, job_labels = _read_input_from_workbook_pandas(input_path)
    n = len(t)

    # סידור כרונולוגי 0,1,2,... — כמו סדר העמודות בקובץ הקלט, בלי GA
    original_sequence = list(range(n))
    obj_value, completion_times, earliness, tardiness = solve_lp_subproblem(
        original_sequence, t, d, v, w
    )

    start_times = [completion_times[j] - t[j] for j in range(n)]
    # קנס לכל עבודה = v*מוקדם + w*איחור (לשורת Pj בגיליון)
    penalties = [v[j] * earliness[j] + w[j] * tardiness[j] for j in range(n)]
    run_time = time.perf_counter() - run_start

    _write_results_workbook(
        f"{Output_File_Name}.xlsx",
        job_labels,
        t,
        d,
        v,
        w,
        original_sequence,
        completion_times,
        start_times,
        earliness,
        tardiness,
        penalties,
        obj_value,
        baseline_value=obj_value,
        run_time=run_time,
    )


if __name__ == "__main__":
    INPUT_NAME = "Input_Data_File_2_Example"
    OUTPUT_NAME = "My_Output_Results"

    print("מפעיל את הסולבר הלינארי ומייצר פלט...")
    try:
        Run_Solver(INPUT_NAME, OUTPUT_NAME)
        print(f"הצלחה! נוצר קובץ חדש בתיקייה בשם: {OUTPUT_NAME}.xlsx")
    except FileNotFoundError:
        print(f"שגיאה: לא מצאתי את קובץ הקלט '{INPUT_NAME}.xlsx'.")
        print("ודאי שהקובץ של המרצה יושב באותה תיקייה של קובץ הפייתון.")
    except PermissionError:
        print("שגיאה: קובץ הפלט פתוח כרגע באקסל! סגרי אותו והריצי שוב.")
