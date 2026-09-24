# Production Scheduling Optimization (Genetic Algorithm + Linear Programming)

A university project in Industrial Engineering (B.Sc. Industrial Engineering & Management), done in a team of two. The program schedules jobs on a single machine. It uses a genetic algorithm to search for a good job order, and a linear program to time the jobs for each order it tries.

The code was written with AI tools (Cursor and Claude), working from our problem definition and requirements. The section [How this was built](#how-this-was-built-and-my-part) explains who did what.

## The problem

There is a single machine and a set of jobs. Each job `j` has:

| Symbol | Meaning |
|---|---|
| `t_j` | processing time |
| `d_j` | due date |
| `v_j` | penalty per time unit if the job finishes **early** (before `d_j`) |
| `w_j` | penalty per time unit if the job finishes **late** (after `d_j`) |

The machine processes one job at a time, and it may sit idle between jobs. The goal is to minimize the total weighted earliness and tardiness:

```
minimize  Σ_j ( v_j · E_j + w_j · T_j )
where     E_j = max(0, d_j − C_j)   (earliness)
          T_j = max(0, C_j − d_j)   (tardiness)
          C_j = completion time of job j
```

This is a "just-in-time" objective: finishing too early costs money (for example, storage), and so does finishing late.

## Approach

The problem splits into two decisions, and the code handles each one differently:

1. **Job order (genetic algorithm).** Each chromosome is a permutation of the jobs.
2. **Exact timing for a given order (linear program).** Once the order is fixed, `solve_lp_subproblem` in `Tapi_Project_Utils.py` builds an LP with PuLP (CBC solver). The LP finds the completion times that minimize the penalty. It is allowed to insert idle time, for example to avoid finishing a job much too early. The LP's optimal penalty is used as the fitness of that chromosome.

What the GA does each generation (`Tapi_Project_GA_Engine.py`):

- **Initial population:** random permutations
- **Fitness:** the LP penalty for each permutation. Results are cached, so an order that was already solved is not solved again.
- **Elitism:** the best `elite_size` solutions pass to the next generation unchanged
- **Selection:** linear ranking. The best solution gets the highest weight, and every solution has some chance to be picked.
- **Crossover:** Order Crossover (OX), which always produces a valid permutation
- **Mutation:** swap two jobs with probability `mutation_prob`
- **Stopping rule:** a run-time limit (in seconds) read from the input file

For comparison, the program also solves the LP for the original input order (1, 2, 3, …). This is called "Stage 1" and is saved as `Original Sol` in the output. The GA always reports the better of this baseline and the best order it found.

The input file sets the GA parameters, and the code checks them: population size must be 20–200 (rounded up to an even number), and elite size must be at least 2 and no more than 10% of the population.

## Input and output

**Input** is an Excel file (`examples/Input_Data_File_1_Example.xlsx`):

- Rows 1–5: job number, `t_j`, `d_j`, `v_j`, `w_j` (one column per job)
- Column A, rows 7 / 9 / 11 / 13: population size, number of elite solutions, mutation probability, time limit (seconds)

**Output** is an Excel file with one column per job or idle period, in schedule order. It lists start time `S_j`, completion time `C_j`, earliness `E_j`, tardiness `T_j` and penalty `P_j`, plus the total penalty, the baseline penalty, the number of generations, the run time, and a simple text Gantt row.

## Example: input → process → output

This comes from one of our test runs on the 10-job example file (population 26, 2 elite, mutation 0.2, 100-second limit). Both output files are in `examples/`.

**Baseline (Stage 1, jobs in input order 1→10):** total penalty **250**
(`examples/Output_Example_1_Stage1_input_order.xlsx`)

**GA result (8,270 generations in 100 s):** total penalty **207**
(`examples/Output_Example_1_GA.xlsx`)

| | Schedule |
|---|---|
| Order found | j1, j3, j4, j5, j2, j7, j6, j8, j9, j10 |
| Idle time inserted | 0–1 and 66–70 |

On this input, the GA found an order with a 17% lower penalty than the input order. This is one small example. The GA is a heuristic, so it does not guarantee an optimal schedule. Results also vary between runs and depend on the input: on some inputs we tested, it did not improve on the input order within the time limit.

## How to run

Requires Python 3 (the team used 3.13; also tested here on 3.11 and 3.12).

```bash
pip install -r requirements.txt
```

Run from Python. Pass the input file name without `.xlsx`:

```bash
# GA (runs for the time limit in the input file: 100 s for the example)
python -c "from Tapi_Project_GA_Engine import Run_GA; Run_GA('examples/Input_Data_File_1_Example', 'my_ga_result')"

# Stage 1 only (input order)
python -c "from Tapi_Project_Solver_Engine import Run_Solver; Run_Solver('examples/Input_Data_File_1_Example', 'my_solver_result')"
```

## Repository structure

```
Tapi_Project_GA_Engine.py        genetic algorithm main loop (Run_GA)
Tapi_Project_Solver_Engine.py    Stage 1: LP for the input order only (Run_Solver)
Tapi_Project_Utils.py            LP model (PuLP), Excel read/write, GA operators, parameter checks
examples/                        sample input and two output files from our runs
requirements.txt
```

Code comments are partly in Hebrew (the course language).

## How this was built, and my part

This was a two-person project. We used AI tools throughout and declared them to the course:

- **Cursor** was used to write and edit the code from the problem definition and requirements we gave it. It also fixed bugs, moved shared code into the `Utils` module, added pandas-based input reading, and added comments.
- **Claude** checked the code against the project requirements, helped find bugs (input validation, duplicated code, the order of mutation and evaluation), and helped us understand and explain the GA logic.

My part, as one of the two team members:

- Defining the problem and the requirements that the code was built from
- Working on the solution together with my teammate
- Testing the program and validating its outputs
- Processing and checking the input and output data (including with pandas)

I did not write the genetic algorithm from scratch. My work was on the problem definition, the requirements, and checking that the program did what we specified.

## Changes made for this repository

- Removed the small GUI window (`Project_Form.py`) and its background image. The window was based on a course-provided template, so it isn't included here. The program runs from Python, as shown above.
- Kept one example input and two of our own result files, and removed author metadata from them.
- No changes to the algorithm or the LP model.
