# I. Author's context:
In my first corporate finance job, one of my challenges was finding a reliable way of making 12 iterations of changes across 50 copies of a complex forecast model in Excel. The FP&A team only made it through this crisis because I introduced the idea of Excel macros.

However, if I faced another Excel automation challenge today, I would at least consider __different strategies outside of VBA and Excel macros__.

# II. Evaluation of competing tactics:
## 1. Why should we use VBA and Excel macros?:
### Security or cost concerns from IT
Maybe the company won't allow anything outside of Excel to be installed on their computers.
<br>
<br>
### End users have rigid constraints on process
The participants may be unwilling to adopt or learn new technologies
<br>
<br>

## 2. Why we should move beyond VBA and Excel Macros?
### Both de jur and de facto limits on data
While a sheet can theoretically store ~1M rows, my own work experience has shown noticeable performance issues (even full-blown crashes) under the weight of only 30,000 rows.
<br>
<br>
### Slowness
Python and Pandas simply move faster than Excel when it comes to complex transformations.
<br>
<br>
### Lack of isolation and controlled code execution
There's no step-by-step feedback system in VBA macros, making debugging very difficult. Jupyter notebooks allow for "literate programming" or "observable programming", meaning you can see the code and its output after running one code block at a time.
<br>
<br>
### Limited documentation
VBA allows for comments, but Jupyter allows the same markup language found in README files on GitHub repos, allowing for better readability and the use of images.
<br>
<br>
### No community support
VBA is essentially naked and unarmed, but Python can leverage specialized tools like:
* Pandas for spreadsheets and pivot tables
* NumPy for advanced numerical calculations
* Matplotlib for data visualization

<br>
<br>



# IV. What I built:
1. Each month, another department (the grocery stores division) hands me an Excel file with financial and operating data for grocery stores spread across multiple states. However, the data is in "crosstab" or "summary" format. That's like being sold marinara sauce, which isn't necessarily bad if everyone in your household wants to eat spaghetti.

2. The problem comes when someone wants tomato soup instead. My challenge is to decompose the marina sauce back down to its original components tomatoes, garlic, basil, and olive oil so I can make both marinara sauce and tomato soup. This is called "flattening the data".

3. Using a Jupyter notebook, I
* pulled data from an Excel file
* normalized/flattened it
* added my own transformations e.g.
  * summed up rental expense + utility expense to create new parent category "facilities expense"
  * created new calculations like corporate HQ expense and "net income"
* and then produced results as:
  * grids
  * pivot tables
  * charts
  * exported Excel workbooks with __some dynamic columns__

4. Tech used:
* Jupyter
* Pandas
* Matplotlib
* openpyxl

# IV. Screenshots (illustrative, but not comprehensive):
Findings are presented in screenshots below:


# V. Learnings:
- Intentionally left blank for now

# VI. Potential improvements:

>**Analysis**<br>
- We could run a loop where process and add multiple files at a time
- What if we had weekly reporting instead of monthly? Then we'd need to make adjustments such as having 13 "months" or absorbing the info by week and reading monthly reports with an understanding that each month is different based on day count or unalignment with a week's starting point