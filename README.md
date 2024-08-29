# Getting Started
1. Install **python** (tested on the v3.11.4 for Windows)

2. Download and unzip the script folder

3. Install the requirements:
	`cd <path to the script>`
	`py -m pip install -r requirements.txt`

4. Create a new template, if necessary:
	- Copy the folder: `assets/templates/template_metsys`
	- Paste the folder: `assets/templates/<template name>`
	- Update the folder:
		- Add/Remove fonts folders in the folder: `assets/templates/<template name>/fonts` (some fonts cannot be embedded in PDF files)
		- Update the files **header.docx** and **footer.docx**

5. Update the **config.txt** file, if necessary

6. Add the input files in the **input** folder:
	- PingCastle: `ad_hc_*.xml`
	- PurpleKnight: `Security_Assessment_Report_*.xlsx`

7. Execute the program:
	`cd <path to the script>`
	`py main.py`

8.	Get the final report in the **output** folder:
	`ActiveDirectoryAuditReport.pdf`

# Test mode
It is possible to generate a test report that includes all existing risks. To do this:

1. Uncomment the `TEST MODE` lines in the **mark_risks_found** function within the **main.py** file.
2. Generate the report.
3. Comment the `TEST MODE` lines again.