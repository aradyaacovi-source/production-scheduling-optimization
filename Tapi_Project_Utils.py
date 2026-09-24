import random

import pandas as pd
import pulp
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter

JOB_FILL = PatternFill(start_color="00B0B0", end_color="00B0B0", fill_type="solid")
IDLE_FILL = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

# מטמון משותף לכל הרצות ה-GA: מפתח = סידור עבודות (tuple), ערך = תוצאות LP מלאות
# Shared across GA runs: sequence tuple -> (obj, completion_times, earliness, tardiness).
lp_cache = {}


def solve_lp_subproblem(sequence, t, d, v, w):
    """
  Solves the linear program for one fixed job processing order.

  Parameters:
      sequence: list of job indices in the order they run on the machine.
      t: processing time for each job.
      d: due date for each job.
      v: earliness penalty weight for each job.
      w: tardiness penalty weight for each job.

  Returns:
      Total penalty (objective), completion times, earliness per job, tardiness per job.
  """
    n = len(t)

    prob = pulp.LpProblem("JIT_Scheduling", pulp.LpMinimize)

    # C_j = זמן סיום עבודה j | E_j = מוקדם | T_j = איחור (משתנים רציפים לא שליליים)
    C = [pulp.LpVariable(f"C_{i}", lowBound=0, cat="Continuous") for i in range(n)]
    E = [pulp.LpVariable(f"E_{i}", lowBound=0, cat="Continuous") for i in range(n)]
    T = [pulp.LpVariable(f"T_{i}", lowBound=0, cat="Continuous") for i in range(n)]

    # ממזערים סכום קנסות: v_j * E_j + w_j * T_j על כל העבודות
    prob += pulp.lpSum(v[i] * E[i] + w[i] * T[i] for i in range(n))

    if sequence:
        first_job = sequence[0]
        # העבודה הראשונה בסידור חייבת להסתיים לפחות אחרי זמן העיבוד שלה
        prob += C[first_job] >= t[first_job]

        for i in range(1, len(sequence)):
            curr_job = sequence[i]
            prev_job = sequence[i - 1]
            # סדר על המכונה: התחלת curr אחרי סיום prev (עם אפשרות להמתנה ביניהם)
            prob += C[curr_job] - t[curr_job] >= C[prev_job]

    for j in range(n):
        # E_j ו-T_j מחליפים max{0,...}: E>=d-C (מוקדם), T>=C-d (איחור) — בלי פונקציה לא לינארית
        prob += E[j] >= d[j] - C[j]
        prob += T[j] >= C[j] - d[j]

    # פותרים את התכנון הלינארי (מגבלת 5 שניות לכל סידור)
    prob.solve(pulp.PULP_CBC_CMD(timeLimit=5, msg=False))

    status = pulp.LpStatus[prob.status]
    if status != "Optimal":
        print(
            f"WARNING: LP solver finished with status '{status}' "
            f"(expected 'Optimal'). Check sequence or input data."
        )
#חילוץ התוצאות
    completion_times = [pulp.value(C[j]) for j in range(n)]
    earliness = [pulp.value(E[j]) for j in range(n)]
    tardiness = [pulp.value(T[j]) for j in range(n)]
    obj_value = pulp.value(prob.objective)

    return obj_value, completion_times, earliness, tardiness


def _fmt(value):
    """
  Rounds a number to one decimal place for display in Excel and the Gantt chart.

  Parameters:
      value: any numeric value.

  Returns:
      The value rounded to one decimal as a float.
  """
    return round(float(value), 1)

#קריאת קלט מאקסל
def _read_input_from_workbook_pandas(filepath):
    """
  Reads job times, due dates, penalties, and labels from the input Excel file.

  Parameters:
      filepath: full path to the .xlsx input workbook.

  Returns:
      Lists t, d, v, w and job_labels for all jobs.
  """
    df = pd.read_excel(filepath, sheet_name=0, header=None)
    # שורות 2-5 באקסל = t, d, v, w מתחילות מעמודה B
    t = [float(x) for x in df.iloc[1, 1:] if pd.notna(x)]
    d = [float(x) for x in df.iloc[2, 1:] if pd.notna(x)]
    v = [float(x) for x in df.iloc[3, 1:] if pd.notna(x)]
    w = [float(x) for x in df.iloc[4, 1:] if pd.notna(x)]
    n = len(t)
    raw_labels = df.iloc[0, 1 : n + 1]
    # אם חסר תווית לעבודה — משתמשים במספר סידורי 1..n
    job_labels = [int(x) if pd.notna(x) else i + 1 for i, x in enumerate(raw_labels)]
    return t, d, v, w, job_labels


def _build_schedule_columns(sequence, t, completion_times):
    """
  Builds the ordered list of Gantt columns (jobs and idle gaps) for Excel output.

  Parameters:
      sequence: job processing order on the machine.
      t: processing times.
      completion_times: LP completion time C_j for each job.

  Returns:
      List of dicts, each describing one Gantt column ('job' or 'idle').
  """
    columns = []

    if sequence:
        first_job = sequence[0]
        first_start = completion_times[first_job] - t[first_job]
        # מקטע IDLE לפני העבודה הראשונה — המכונה מחכה מתחילת הזמן
        if first_start > 1e-6:
            columns.append({"type": "idle", "start": 0.0, "end": first_start})

    for i, job_idx in enumerate(sequence):
        if i > 0:
            prev_job = sequence[i - 1]
            gap_start = completion_times[prev_job]
            gap_end = completion_times[job_idx] - t[job_idx]
            # אם יש רווח בין סיום prev להתחלת curr — מוסיפים עמודת המתנה (IDLE)
            if gap_end > gap_start + 1e-6:
                columns.append(
                    {"type": "idle", "start": gap_start, "end": gap_end}
                )
        columns.append({"type": "job", "idx": job_idx})

    return columns


def _apply_column_fill(ws, col, fill, rows=(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 17)):
    """
  Applies a background color fill to selected rows in one Excel column.

  Parameters:
      ws: openpyxl worksheet.
      col: column index.
      fill: PatternFill style to apply.
      rows: row numbers to color.
  """
    for row in rows:
        ws.cell(row=row, column=col).fill = fill


def _format_worksheet(ws):
    """
  Auto-sizes columns and center-aligns all cells on the results sheet.

  Parameters:
      ws: openpyxl worksheet to format.
  """
    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.alignment = CENTER_ALIGN
            if cell.value is not None:
                text = str(cell.value)
                # שורות מרובות בגאנט — לוקחים את השורה הארוכה ביותר לרוחב העמודה
                line_len = max(len(line) for line in text.split("\n"))
                max_len = max(max_len, line_len)
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)


def _write_results_workbook(
    output_path,
    job_labels,
    t,
    d,
    v,
    w,
    sequence,
    completion_times,
    start_times,
    earliness,
    tardiness,
    penalties,
    obj_value,
    baseline_value,
    run_time,
    total_generations="-",
):
    """
  Writes the full results workbook with schedule table and Gantt chart.

  Parameters:
      output_path: path to save the .xlsx file.
      job_labels, t, d, v, w: job input data.
      sequence: best processing order found.
      completion_times, start_times, earliness, tardiness, penalties: per-job results.
      obj_value: total penalty of the displayed solution.
      baseline_value: penalty of the chronological order (for comparison).
      run_time: wall-clock seconds.
      total_generations: number of GA generations, or "-" for Stage 1.

  Returns:
      Nothing; saves the file to disk.
  """
    schedule_columns = _build_schedule_columns(sequence, t, completion_times)

    wb_out = Workbook()
    ws = wb_out.active
    ws.title = "Results"

    row_labels = {
        1: "J",
        2: "tj",
        3: "dj",
        4: "vj",
        5: "wj",
        6: "Sj",
        7: "Cj",
        8: "Ej",
        9: "Tj",
        10: "Pj",
    }
    for row, label in row_labels.items():
        ws.cell(row=row, column=1, value=label)

    col = 2
    for slot in schedule_columns:
        if slot["type"] == "idle":
            idle_start = slot["start"]
            idle_end = slot["end"]
            fs, fe = _fmt(idle_start), _fmt(idle_end)

            ws.cell(row=1, column=col, value=f"IDLE {fs}-{fe}")
            for row in (2, 3, 4, 5, 8, 9, 10):
                ws.cell(row=row, column=col, value="-")
            ws.cell(row=6, column=col, value=idle_start)
            ws.cell(row=7, column=col, value=idle_end)
            ws.cell(row=17, column=col, value=f"IDLE\n({fs} - {fe})")
            _apply_column_fill(ws, col, IDLE_FILL)
        else:
            job_idx = slot["idx"]
            lbl = job_labels[job_idx]
            fs = _fmt(start_times[job_idx])
            fc = _fmt(completion_times[job_idx])

            ws.cell(row=1, column=col, value=f"j{lbl}")
            ws.cell(row=2, column=col, value=t[job_idx])
            ws.cell(row=3, column=col, value=d[job_idx])
            ws.cell(row=4, column=col, value=v[job_idx])
            ws.cell(row=5, column=col, value=w[job_idx])
            ws.cell(row=6, column=col, value=start_times[job_idx])
            ws.cell(row=7, column=col, value=completion_times[job_idx])
            ws.cell(row=8, column=col, value=earliness[job_idx])
            ws.cell(row=9, column=col, value=tardiness[job_idx])
            ws.cell(row=10, column=col, value=penalties[job_idx])
            ws.cell(row=17, column=col, value=f"j{lbl}\n({fs} - {fc})")
            _apply_column_fill(ws, col, JOB_FILL)

        col += 1

    # סיכום בתחתית הגיליון: קנס כולל, בסיס, דורות, זמן ריצה
    ws["A12"] = "Total Penalty"
    ws["B12"] = obj_value
    ws["A13"] = "Original Sol"
    ws["B13"] = baseline_value
    ws["A14"] = "Total Gen Created"
    ws["B14"] = total_generations
    ws["A15"] = "Run time"
    ws["B15"] = run_time
    ws["A17"] = "Gantt:"

    _format_worksheet(ws)
    wb_out.save(output_path)


def _read_ga_cell(ws, row):
    """
  Reads one GA setting from column A of the input sheet.

  Parameters:
      ws: openpyxl worksheet (input file).
      row: Excel row number (e.g. 7 for pop_size).

  Returns:
      The raw cell value from column A.
  """
    return ws.cell(row=row, column=1).value


def _parse_ga_int(value, param_name):
    """
  Converts an Excel cell value to a valid integer GA parameter.

  Parameters:
      value: raw cell contents.
      param_name: name used in error messages (e.g. 'pop_size').

  Returns:
      The parameter as an int.

  Raises:
      ValueError if the cell is empty or not a whole number.
  """
    if value is None or value == "":
        raise ValueError(f"{param_name} is missing (expected a value in the input sheet).")
    if isinstance(value, bool):
        raise ValueError(f"{param_name} must be an integer, got {value!r}.")
    try:
        numeric = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(
            f"{param_name} must be an integer, got {value!r}."
        ) from exc
    if numeric != int(numeric):
        raise ValueError(f"{param_name} must be an integer, got {value!r}.")
    return int(numeric)


def _parse_ga_float(value, param_name):
    """
  Converts an Excel cell value to a float GA parameter.

  Parameters:
      value: raw cell contents.
      param_name: name used in error messages.

  Returns:
      The parameter as a float.
  """
    if value is None or value == "":
        raise ValueError(f"{param_name} is missing (expected a value in the input sheet).")
    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{param_name} must be numeric, got {value!r}.") from exc

#בודק תקינות קלט 
def _validate_ga_parameters(pop_size, elite_size):
    """Validate and auto-correct GA parameters per project rules."""
    if not (20 <= pop_size <= 200):
        raise ValueError(
            f"pop_size must be between 20 and 200 (inclusive), got {pop_size}."
        )
    if pop_size % 2 != 0:
        pop_size += 1
        print(f"Warning: pop_size was odd, rounded up to {pop_size}.")

    if elite_size < 2:
        raise ValueError(f"elite_size must be at least 2, got {elite_size}.")

    max_elite = int(pop_size * 0.10)
    if elite_size > max_elite:
        raise ValueError(
            f"elite_size must not exceed 10% of pop_size ({max_elite}), "
            f"got {elite_size}."
        )
    if elite_size % 2 != 0:
        elite_size += 1
        print(f"Warning: elite_size was odd, rounded up to {elite_size}.")

    return pop_size, elite_size


def evaluate_population(population, t, d, v, w):
    """
  Computes the total penalty for every chromosome in the population.

  Parameters:
      population: list of chromosomes (each is a job permutation).
      t, d, v, w: job data passed to the LP solver.

  Returns:
      List of (chromosome, objective_value) pairs.
  """
    evaluated = []
    for sequence in population:
        # tuple כי רשימה לא ניתנת למפתח מילון — אותו סידור = אותו מפתח במטמון
        key = tuple(sequence)
        if key in lp_cache:
            # כבר פתרנו את הסידור הזה — חוסכים הרצת LP של ~5 שניות
            obj_value = lp_cache[key][0]
        else:
            result = solve_lp_subproblem(sequence, t, d, v, w)
            lp_cache[key] = result
            obj_value = result[0]
        evaluated.append((sequence, obj_value))
    return evaluated


def get_ranking_probabilities(evaluated_pop):
    """
  Assigns selection probabilities using linear ranking (better = higher chance).

  Parameters:
      evaluated_pop: list of (chromosome, penalty) from evaluate_population.

  Returns:
      sorted_pop: same individuals sorted by ascending penalty.
      probabilities: selection weight for each entry in sorted_pop.
  """
    # ממיינים מהטוב לגרוע — אינדקס 0 הוא הקטן ביותר (עונש נמוך)
    sorted_pop = sorted(evaluated_pop, key=lambda item: item[1])
    pop_count = len(sorted_pop)

    # דירוג לינארי: הטוב ביותר מקבל דירוג N, הגרוע ביותר דירוג 1
    ranks = [pop_count - i for i in range(pop_count)]
    rank_sum = sum(ranks)
    # הסתברות = דירוג / סכום דירוגים — טובים נבחרים יותר, אבל גרועים עדיין יכולים להיבחר
    probabilities = [rank / rank_sum for rank in ranks]

    return sorted_pop, probabilities


def select_parents(sorted_pop, probabilities, num_parents):
    """
  Randomly picks parent chromosomes according to ranking probabilities.

  Parameters:
      sorted_pop: evaluated population sorted by fitness.
      probabilities: weight for each individual in sorted_pop.
      num_parents: how many parents to draw (usually 2 per crossover).

  Returns:
      List of parent chromosomes (job index lists only).
  """
    selected = random.choices(sorted_pop, weights=probabilities, k=num_parents)
    # random.choices מחזיר זוגות (כרומוזום, עונש) — שומרים רק את הסידור
    return [chromosome for chromosome, _ in selected]


# OX (Order Crossover): שומר על תקינות התמורה — כל עבודה פעם אחת, בלי כפילויות
def crossover_ox(parent1, parent2):
    """
  Creates two child permutations from two parents using Order Crossover (OX).

  Parameters:
      parent1, parent2: valid job permutations (lists of indices).

  Returns:
      Two new child permutations (child1, child2).
  """
    n = len(parent1)
    # שני נקודות חיתוך אקראיות — הקטע ביניהן עובר לילד כמו שהוא
    point_a, point_b = sorted(random.sample(range(n), 2))

    def _ox_child(main_parent, fill_parent):
        child = [None] * n
        child[point_a : point_b + 1] = main_parent[point_a : point_b + 1]
        segment_jobs = set(child[point_a : point_b + 1])
        # שאר המקומות מתמלאים מסדר ההורה השני, מדלגים על עבודות שכבר בקטע
        fill_jobs = [job for job in fill_parent if job not in segment_jobs]
        fill_idx = 0
        for pos in range(n):
            if child[pos] is None:
                child[pos] = fill_jobs[fill_idx]
                fill_idx += 1
        return child

    return _ox_child(parent1, parent2), _ox_child(parent2, parent1)


def mutate_swap(chromosome, mutation_prob):
    """
  Optionally swaps two random positions in a chromosome (swap mutation).

  Parameters:
      chromosome: permutation to mutate (modified in place).
      mutation_prob: probability of applying one swap this generation.

  Returns:
      Nothing; may change chromosome in place.
  """
    if random.random() < mutation_prob:
        idx_a, idx_b = random.sample(range(len(chromosome)), 2)
        # החלפת שני מקומות — עדיין תמורה תקינה (אין הוספה/מחיקה של עבודות)
        chromosome[idx_a], chromosome[idx_b] = chromosome[idx_b], chromosome[idx_a]


def _get_cached_lp_solution(sequence, t, d, v, w):
    """
  Returns the full LP result for a sequence, reusing the cache when possible.

  Parameters:
      sequence: job processing order.
      t, d, v, w: job data for the LP.

  Returns:
      Tuple (obj_value, completion_times, earliness, tardiness).
  """
    key = tuple(sequence)
    if key in lp_cache:
        return lp_cache[key]
    result = solve_lp_subproblem(sequence, t, d, v, w)
    lp_cache[key] = result
    return result
