## VSD18 Setup

### (1) Download

- Download + unzip the VSD18 spreadsheet
- Enable macros when prompted

### (2) Configure

- Search for your event's team list at [frc-events.firstinspires.org/2018](https://frc-events.firstinspires.org/2018)
- Copy the list of team info and paste it to the right side of the table in `Team List`, but **do not overwrite the "INDEX" column (`D`)**.
- Copy the team numbers into the "Team #" column (`A`), rookie years into the "Rookie Year" column (`B`), and team names into the "Team Nickname" column (`C`)
- Select the lowest filled cell in the "INDEX" column (`D`) and drag the green square in the bottom-right corner down until there is an index value for each team in the list, then drag it down one extra cell and release
  - **Take note of the 'extra' cell's value** (you'll need it later)

### (3) Initialize

- Copy the list of team numbers from column `A` of `Team List` into column `A` of `Addl. Robot Data`, column `A` of `Average Comparison`, and column W of `Data Processor`
- Enter your team's team number into cell `C2` of `Picklist`
- Select columns `B` through `Q` of the bottom filled row in `Average Comparison` and drag the green square in the bottom-right corner down until there is a filled row for each team in column `A`
- Select the range `A2:T20` in `Data Processor` and drag the green square in the bottom-right corner down until you reach the row number that matches the value of the 'extra' cell from before
- One-by-one, from top to bottom, copy and paste each value from column `W` of `Data Processor` into the **next empty black cell in column `A`** of the same sheet.

## Using VSD18

### Addl. Robot Data

- General use sheet for tracking miscellaneous robot data, use however you wish

### RAW INPUT

- 1 row per data entry: provides space for up to 2000 data entries
- Fill in each column with the data type specified in square brackets `[]`
  - Data types in the form of `[X/Y/Z]` can be filled with either `X`, `Y`, or `Z`.
  - `[Auto-calculated]` columns will fill automatically as more data is inputted
- Only team number is required, everything else is optional

### Data Processor

- Process data by clicking on the "Process" button in cell `U1` of `Data Processor` or by using the shortcut `Ctrl+Shift+Q`

### Comparison sheets (Average Comparison, Single View, Tri-Team View)

- General match stats are shown, along with averages and linearly forecasted stats
  - `Average Comparison` shows every team's average stats
  - `Single View` shows a detailed breakdown of one team's stats
  - `Tri-Team View` shows a general breakdown of one team's stats
- `Single View` + `Tri-Team View`: Select a team to analyze by entering a team number into cell `B2` (+ `F2` and `J2` in `Tri-Team View`)
- `Single View`: View results for a team's *n*th match by entering said number *n* into cell `T3`

### Picklist

- Copy the list of teams from column `A` of `Team List` into column `F` of `Picklist` and remove any teams that did not attend your event
- Cut and paste the team numbers of your desired picks into column `C` of `Picklist` in the desired picklist position