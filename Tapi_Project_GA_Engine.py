import random
import time

from openpyxl import load_workbook

from Tapi_Project_Utils import (
    lp_cache,
    solve_lp_subproblem,
    _read_input_from_workbook_pandas,
    _validate_ga_parameters,
    _parse_ga_int,
    _parse_ga_float,
    _read_ga_cell,
    evaluate_population,
    get_ranking_probabilities,
    select_parents,
    crossover_ox,
    mutate_swap,
    _get_cached_lp_solution,
    _write_results_workbook,
)


def Run_GA(Input_File_Name, Output_File_Name):
    """
  Runs the full genetic algorithm: evolve job sequences and export the best to Excel.

  Parameters:
      Input_File_Name: base name of the input .xlsx (without extension).
      Output_File_Name: base name for the results .xlsx to create.

  Returns:
      Tuple of (best job sequence found, its total penalty value).
  """
    run_start = time.perf_counter()
    # איפוס מטמון בתחילת ריצה — כל הרצת GA מתחילה נקייה
    lp_cache.clear()

    input_path = f"{Input_File_Name}.xlsx"
    # קוראים נתוני עבודות עם pandas
    t, d, v, w, job_labels = _read_input_from_workbook_pandas(input_path)

    # פרמטרי GA מעמודה A — openpyxl נוח לתאים בודדים
    wb_in = load_workbook(input_path, data_only=True)
    ws_in = wb_in.active
    n = len(t)
    pop_size = _parse_ga_int(_read_ga_cell(ws_in, 7), "pop_size")
    elite_size = _parse_ga_int(_read_ga_cell(ws_in, 9), "elite_size")
    mutation_prob = _parse_ga_float(_read_ga_cell(ws_in, 11), "mutation_prob")
    time_limit = _parse_ga_float(_read_ga_cell(ws_in, 13), "time_limit")
    wb_in.close()

    pop_size, elite_size = _validate_ga_parameters(pop_size, elite_size)

    # --- Step 0: פתרון בסיס כרונולוגי (0,1,2,...) — נקודת השוואה לפני ה-GA ---
    original_sequence = list(range(n))
    baseline_result = solve_lp_subproblem(original_sequence, t, d, v, w)
    baseline_value = baseline_result[0]
    lp_cache[tuple(original_sequence)] = baseline_result

    # --- Step 1: אוכלוסייה ראשונית — כל כרומוזום = סידור אקראי תקין של העבודות ---
    population = [random.sample(range(n), n) for _ in range(pop_size)]

    current_generation = 0
    best_overall_penalty = baseline_value
    best_overall_sequence = original_sequence[:]
    generation_of_last_update = 0
    updated_since_last_report = False

    # --- Step 8: לולאה עד שנגמר זמן הריצה מהאקסל ---
    while time.perf_counter() - run_start < time_limit:
        current_generation += 1

        # --- Step 2: מחשבים עונש (fitness) לכל פתרון באוכלוסייה הנוכחית ---
        evaluated = evaluate_population(population, t, d, v, w)
        sorted_eval = sorted(evaluated, key=lambda item: item[1])

        for chromosome, penalty in evaluated:
            if penalty < best_overall_penalty:
                best_overall_penalty = penalty
                best_overall_sequence = chromosome[:]
                generation_of_last_update = current_generation
                updated_since_last_report = True

        # --- Step 4: העברת אליטות — הטובים ביותר עוברים לדור הבא בלי שינוי ---
        next_population = [chromosome[:] for chromosome, _ in sorted_eval[:elite_size]]

        # --- Step 3: הסתברויות בחירה לפי דירוג (לשלב ההכלאה) ---
        sorted_pop, probabilities = get_ranking_probabilities(evaluated)

        # --- Step 5: בוחרים הורים, מכלאים OX, וממלאים עד גודל pop_size ---
        while len(next_population) < pop_size:
            parent1, parent2 = select_parents(sorted_pop, probabilities, 2)
            child1, child2 = crossover_ox(parent1, parent2)
            next_population.append(child1)
            if len(next_population) < pop_size:
                next_population.append(child2)

        # --- Steps 6+7: קודם מוטציה, אחר כך הערכה — כך הטוב ביותר הוא אחרי המוטציה ---
        for idx in range(elite_size, pop_size):
            mutate_swap(next_population[idx], mutation_prob)
            child = next_population[idx]
            penalty = _get_cached_lp_solution(child, t, d, v, w)[0]
            if penalty < best_overall_penalty:
                best_overall_penalty = penalty
                best_overall_sequence = child[:]
                generation_of_last_update = current_generation
                updated_since_last_report = True

        population = next_population

        if current_generation % 5 == 0:
            print(
                f"Generation {current_generation}: "
                f"Best penalty so far = {best_overall_penalty:.4f}"
            )
            if updated_since_last_report:
                print(
                    f"Best solution updated at generation "
                    f"{generation_of_last_update}."
                )
                updated_since_last_report = False

    # הדפסה אחרונה — גם אם הדור האחרון לא מתחלק ב-5
    print(
        f"Final generation {current_generation}: "
        f"Best penalty = {best_overall_penalty:.4f}"
    )
    if updated_since_last_report:
        print(
            f"Best solution last updated at generation "
            f"{generation_of_last_update}."
        )

    run_time = time.perf_counter() - run_start

    best_obj, completion_times, earliness, tardiness = _get_cached_lp_solution(
        best_overall_sequence, t, d, v, w
    )
    start_times = [completion_times[j] - t[j] for j in range(n)]
    penalties = [v[j] * earliness[j] + w[j] * tardiness[j] for j in range(n)]

    _write_results_workbook(
        f"{Output_File_Name}.xlsx",
        job_labels,
        t,
        d,
        v,
        w,
        best_overall_sequence,
        completion_times,
        start_times,
        earliness,
        tardiness,
        penalties,
        obj_value=best_obj,
        baseline_value=baseline_value,
        run_time=run_time,
        total_generations=current_generation,
    )

    print(
        f"\nGA finished: {current_generation} generations, "
        f"best penalty = {best_overall_penalty:.4f}, runtime = {run_time:.2f}s"
    )
    return best_overall_sequence, best_overall_penalty


if __name__ == "__main__":
    INPUT_NAME = "Input_Data_File_2_Example"
    OUTPUT_NAME = "My_GA_Output_Results"

    print("מפעיל את אלגוריתם הגנטי (GA) ומייצר פלט...")
    try:
        Run_GA(INPUT_NAME, OUTPUT_NAME)
        print(f"הצלחה! נוצר קובץ חדש בתיקייה בשם: {OUTPUT_NAME}.xlsx")
    except FileNotFoundError:
        print(f"שגיאה: לא מצאתי את קובץ הקלט '{INPUT_NAME}.xlsx'.")
        print("ודאי שהקובץ של המרצה יושב באותה תיקייה של קובץ הפייתון.")
    except PermissionError:
        print("שגיאה: קובץ הפלט פתוח כרגע באקסל! סגרי אותו והריצי שוב.")
    except ValueError as exc:
        print(f"שגיאה בפרמטרי GA: {exc}")
