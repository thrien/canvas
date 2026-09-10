# Canvas tools for GSIs in PHYS/BIOPHYS 251

This script currently contains six commands.

```
$ python canvas.py -h
usage: canvas.py [-h] [-v] command ...

Utilities for using Canvas as a GSI

positional arguments:
  command
    sheets              Draw a sign-in sheet showing what group/table students
                        are assigned to.
    introduction (intro, slides)
                        Create a template for introduction slides
    quiz_code (quiz)    Update the quiz code on the introduction slides
    new_quiz_code (new_code)
                        Generate a new quiz code on Canvas
    worksheet           Download the worksheet for the given lab from Canvas
    final_grades (grades)
                        FINAL LETTER GRADE CALCULATOR

options:
  -h, --help            show this help message and exit
  -v, --verbose         print status messages (default: 0)

This script is intended to be used by GSIs for PHYSICS 151/251 at the
University of Michigan.
```

Each command is described below. For questions, contact me at
[`thrien@umich.edu`](mailto:thrien@umich.edu) or open an issue here and I'll
be happy to help.

## Automatic sign-in sheets with groups

The `sheets` command draws a sign-in sheet showing what group/table students
are assigned to.

```
$ python canvas.py sheets -h
usage: canvas.py sheets [-h] [-f | --force | --no-force] [-e ext [ext ...]]
                        [-l numbers [numbers ...]] [-s section [section ...]]

Draw a sign-in sheet showing what group/table students are assigned to.

options:
  -h, --help            show this help message and exit
  -f, --force, --no-force
                        pull recent CSV from Canvas (default: False)
  -e, --extensions ext [ext ...]
                        output formats (default: ['pdf', 'png'])
  -l, --labs numbers [numbers ...]
                        integers or intervals like a..b (default: the next lab)
  -s, --sections section [section ...]
                        your section numbers (default: [15, 25])

Input and output files are organized in the current directory like this:
    .
    ├── canvas.py
    └── lab01
        ├── canvas.csv
        ├── groups015.png
        └── groups025.png
```

### Example

![Example group/table layout with fake names](example.png)

### Usage

There are two ways of using this command.

#### Manual download

For example, for Lab 1 navigate to **People > Groups > Lab 1** and
select **Download Group Category Roster CSV** under the three dots at the top.

Save the file as `./lab01/canvas.csv` and simply run
```
$ python canvas.py -v sheets
```

#### Automatic download

You can also use the Canvas API to automatically download the CSV when needed.
This only requires two simple steps described in the
[setup section](#canvas-api) below.

## Automatic templates for introduction slides

The `introduction` command creates a PowerPoint presentation from a template,
changes the title and inserts the sign-in sheet on the second slide. The third
slide contains the quiz code, which can be updated automatically later.

```
$ python canvas.py introduction -h
usage: canvas.py introduction [-h] [-u | --update | --no-update] [-l number]
                              [-s section [section ...]]

Create a template for introduction slides

options:
  -h, --help            show this help message and exit
  -u, --update, --no-update
                        update quiz code (default: False)
  -l, --lab number      the lab's number (default: latest existing lab)
  -s, --sections section [section ...]
                        your section numbers (default: [15, 25])

The template has three slides:
    - title page with lab and first section number
    - sign-in sheets stacked on top of each other
    - a placeholder for the quiz code

Since the quiz code changes quite frequently we put a placeholder in the
template. Use the quiz_code command to update it before class.
```

This part of the script has not been fully configured yet and contains some
hardcoded paths, chosen for my specific setup.

## Automatically update the quiz code

The `quiz_code` command pulls the current quiz access for the specified lab
from Canvas and updates it on the third slide of the introduction.

```
$ python canvas.py quiz_code -h
usage: canvas.py quiz_code [-h] [-l number]

Update the quiz code on the introduction slides

options:
  -h, --help        show this help message and exit
  -l, --lab number  the lab's number (default: latest existing lab)

This command pulls the latest quiz code from the Canvas API and updates
the corresponding slide in the introduction.
```

It assumes the slides have already been created.

## Generate a new quiz code on Canvas

The `new_quiz_code` command generates a new random access code for a given lab
and uploads it to Canvas. This is intended to be used in class after all
students have finished their quiz.

```
$ python canvas.py new_quiz_code -h
usage: canvas.py new_quiz_code [-h] [-l number]

Generate a new quiz code on Canvas

options:
  -h, --help        show this help message and exit
  -l, --lab number  the lab's number (default: latest existing lab)

Overwrites the quiz code on Canvas with a random number.
Intended to be used in class after students have finished the quiz.
```

## Download Worksheets

The `worksheet` command downloads the worksheet(s) for the specified lab(s) and
stores them in the same folders as the sign-in sheets.

```
$ python canvas.py worksheet -h
usage: canvas.py worksheet [-h] [-l number [number ...]] [-p path]

Download the worksheet for the given lab from Canvas

options:
  -h, --help            show this help message and exit
  -l, --labs number [number ...]
                        the lab's number (default: the next lab)
  -p, --path path       the path to the worksheet file on Canvas (default:
                        Physics 251 GSI Resources/Lab Worksheets)

The worksheets are stored in another Canvas course (at least for PHYS 251
WN26). Define another Canvas course ID with " GSI" appended to the name and
update the path to the top folder if necessary.
```

## Calculate final letter grades

The `final_grades` command converts the final scores (in %) from a gradebook
exported from Canvas into letter grades and prepares both a human-readable
gradebook to check the legitimacy of the results and a machine-readable CSV
file that can be uploaded to Wolverine Access.

```
$ python canvas.py final_grades -h
usage: canvas.py final_grades [-h] [-r path] [-u path] [-k name] gradebook

FINAL LETTER GRADE CALCULATOR

positional arguments:
  gradebook             path to gradebook exported from Canvas

options:
  -h, --help            show this help message and exit
  -r, --readable path   file to write human-readable grades (default:
                        grades/human-readable.csv)
  -u, --uploadable path
                        file to write machine-readable grades (default:
                        grades/wolverine_access.csv)
  -k, --grade-key name  the column name for final scores in the gradebook CSV
                        (default: Current Score)

PHYSICS 151, 251
WINTER 2025
Version 5.2
Liam Daly, Blake Bottesi, Michelle Thran
Last updated: 4/22/2025
```

This script was initially developed by Liam Daly in 2025 and was improved by
Blake Bottesi and Michelle Thran. It is included in this project with minimal
changes.

The embedded grade calculator metadata is from Winter 2025. Review its grading
thresholds and gradebook assumptions before reusing it for another term.

## Configuration

You might want to configure a few things before using this script.

### Default values

For example, you probably want to set some default values for convenience.
Specifically,
- your section numbers for the `sheets` command (see
  [`_sheets_parser`](canvas.py#L295) or
  [here](https://github.com/thrien/canvas/blob/main/canvas.py#L295:~:text=default%3D%5B15%2C%2025%5D))
- the Canvas course (if set up as described [below](#course-id))

### Table layout

The table layout for the sign-in sheets is defined like this:
```
table_layout = [[ 1 , 'I', '.'],
                [ 2 , '.',  8 ],
                [ 3 , '.',  7 ],
                [ 4 ,  5 ,  6 ]]
```

where each number labels a table with a particular group of students. `'I'` can
be used to define a cell for the instructor (you) if the variable `instructor`
holds a name, e.g., `"Thrien, Tobias"`.

The table layout is fed into
[plt.subplot_mosaic](https://matplotlib.org/stable/api/_as_gen/matplotlib.pyplot.subplot_mosaic.html#matplotlib.pyplot.subplot_mosaic).
Read the documentation for more details.

## Setup

### Dependencies

This script uses `python3`. Dependencies are loaded only when the corresponding
command is run:

- `sheets` requires `numpy` and `matplotlib`.
- `introduction` requires `python-pptx`, `numpy`, and `matplotlib`.
- `quiz_code` requires `python-pptx`.
- `final_grades` requires `pandas`.
- `new_quiz_code` and `worksheet` use only the Python standard library.

It has been tested on Linux and Windows, and is expected to run on macOS as
well.

### Canvas API

The `worksheet`, `quiz_code`, and `new_quiz_code` rely on the Canvas API. The
`sheets` (and indirectly `introduction`) command will try to use it if it has
been set up, but can be used without it (see
[Manual Download](#manual-download) above).

##### Canvas API token

First, you need an access token (i.e., a password) that you can generate under
**Account > Settings > Approved Integrations > New Access Token**.

For simplicity you can just copy the token into
```
TOKEN = "<your_access_token_here>"
```
but be warned: NEVER commit this token to GitHub!

Ideally you store this token in an environment variable on your machine and use
```
TOKEN = os.getenv("CANVAS_API_TOKEN")
```

On a Linux or macOS system, add
```
export CANVAS_API_TOKEN="<your_token>"
```
to `~/.profile` or `~/.zprofile`.

On Windows open a shell (`Ctrl-R "cmd"`) and type
```
setx CANVAS_API_TOKEN "<your_token>"
```

##### Course ID

Additionally you will have to find the ID of your course. Open it on Canvas and
read the URL. It should look something like
`https://umich.instructure.com/courses/734390`, where `734390` is the course
ID.

Paste it into your list of course IDs, e.g.,
```
COURSES = {"PHYS 151 WN25": 734390}
```

Once defined, a new option will show up as something like this
```
$ python canvas.py -h
usage: canvas.py [-h] [-v] [-c name] command ...

  -c, --course name     the Canvas course (default: PHYS 151 WN25) (choices:
                        PHYS 151 WN25, PHYS 251 WN26)
```
You can change the default value here:
```
    if COURSES:
        parser.add_argument("-c", "--course", choices=COURSE_NAMES,
                            default="PHYS 151 WN25", help="the Canvas course",
                            metavar="name")
```

## Documentation

The Canvas API is well documented. The function
[group_categories.export](https://developerdocs.instructure.com/services/canvas/resources/group_categories#method.group_categories.export)
exports a CSV file formatted like
[this](https://developerdocs.instructure.com/services/canvas/group-categories/file.group_category_csv)
for a group category with a given ID. We can query all existing group
categories with their names and IDs using
[group_categories](https://developerdocs.instructure.com/services/canvas/resources/group_categories#method.group_categories.index)
and find the one we need.

## Scheduling (optional)

If you already added an access token for the API, nothing stops you from
automating this even more, by scheduling this script to run once a week (e.g.
on Monday mornings) so that the new sign-in sheets will be ready when you need
them.

### Linux

#### systemd

To schedule this script to run once a week until it succeeds define a
[systemd timer](https://wiki.archlinux.org/title/Systemd/Timers) such as
`canvas_groups.timer` in `~/.config/systemd/user`. 

```
[Unit]
Description=Weekly trigger for PHYS/BIOPHYS 251 script

[Timer]
OnCalendar=Mon *-*-* 00:00:00
Persistent=true

[Install]
WantedBy=timers.target
```

It triggers a systemd service like `canvas_groups.service` in the same
directory. Change the `/path/to/your/directory/` to match your setup.

```
[Unit]
Description=Create new sign-in sheets for PHYS/BIOPHYS 251
Wants=network-online.target
After=network-online.target

# Allow retries for up to a week
StartLimitIntervalSec=1week
StartLimitBurst=28

[Service]
Type=oneshot
WorkingDirectory=/path/to/your/directory
ExecStart=/path/to/your/directory/canvas.py sheets

# Retry logic
Restart=on-failure
RestartSec=6h
```

Load the new definitions with
```
$ systemctl --user daemon-reload
```
and start and enable the timer with
```
$ systemctl --user enable --now myscript.timer
```

#### cron

Alternatively, use [cron](https://wiki.archlinux.org/title/Cron) and define a
simple `crontab` file that runs the script once a week on Mondays at 12PM.

```
0 12 * * 1 /path/to/canvas.py sheets
```

Load it with
```
$ crontab /path/to/crontab
```

This might not work if your machine is not running on Mondays at 12 PM and
won't repeat failed tasks.

### Windows

You can use the `Task Scheduler` program to create a new task that runs the
executable once a week at a specific time.

You can import the XML files in the `windows` folder and then modify them to
suit your needs. There are currently tasks for
- creating new sign-in sheets every Monday
- creating a template for introduction slides that includes the sign-in sheet
- printing the sign-in sheets on your default printer with `print.ps1`
before classes start
- updating the quiz code on the slides (previously created from the template)
before classes start
- creating a new quiz code on Canvas after students should have finished the
  quiz

### macOS

I don't have the money for that...

## TODOs

- document canvas functions in the source code
- include screenshots of the Windows Task Scheduler
- improve the `introduction` command
  - create sign-in sheets if necessary
  - allow configuration of file paths
  - add more options
- improve the `quiz_code` command
- revise the `final_grades` command
  - filter by section
    - create output files per section
  - download gradebook through API
  - check if Wolverine Access is accessible via an API as well