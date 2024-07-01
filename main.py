import argparse
import copy
import datetime
import json
import lib.config as config
import lib.docx_manager as docx_manager
import lib.pingcastle as pingcastle
import lib.purpleknight as purpleknight
import lib.logs as logging
import locale
import matplotlib
import numpy
import openpyxl
import os
import pandas
import pprint
import pytz
import seaborn
import shutil
import xml

############################################################################### CONSTANTS

# PATHS
PATH_FILE = os.path.abspath(__file__)
PATH_DIRECTORY = os.path.dirname(PATH_FILE)

# FILES
PATH_CONFIG = os.path.join(PATH_DIRECTORY, "config.txt")

# LOADED CONFIGURATION
config = config.Config()

############################################################################### FILE SYSTEM

#
# Remove the content of a folder
#
@logging.log_call
def delete_folder_contents(folder_path):
	# Iterate over all the contents of the folder
	for item in os.listdir(folder_path):
		item_path = os.path.join(folder_path, item)
		if "gitkeep" in item_path:
			continue
		# Check if the item is a file
		if os.path.isfile(item_path):
			# Remove the file
			os.remove(item_path)
		# Check if the item is a directory
		elif os.path.isdir(item_path):
			# Remove the directory and its contents recursively
			shutil.rmtree(item_path)

#
# Go through the subfolders of a folder and return the list of the files
#
@logging.log_call
def find_files(folder_path: str) -> list:
	#
	# List of files in the folder, with the expected extension
	#
	files_found = []
	#
	# Go through the subdirectory of the directory
	#
	for root, _, filenames in os.walk(folder_path):
		#
		# Go through the files of the current subdirectory
		#
		for filename in filenames:
			#
			# Add the file to the list
			#
			files_found.append(os.path.join(root, filename))
   	#
	# Return the list of matching files
	#
	return files_found

#
# Go through a list of files and filter them with a specific extension
#
@logging.log_call
def filter_files_by_extension(files: list, extension: str) -> list:
	#
	# List of files in the folder
	#
	files_found = []
	#
	# Go through the subdirectory of the directory
	#
	for path in files:
		#
		# If the current file has the right extension
		#
		if path.lower().endswith(extension):
			#
			# Add the file to the list
			#
			files_found.append(path)
   	#
	# Return the list of matching files
	#
	return files_found

#
# Get the default folder for the fonts files depending on the operating system
#
@logging.log_call
def get_default_font_folder_from_operating_system() -> str:
	#
	# Fonts folder of the user
	#
	default_font_folder = None
	#
	# If the operating system is Linux or Mac
	#
	if os.name == 'posix':
		#
		# Save the corresponding font folder
		#
		default_font_folder = os.path.expanduser("~/.local/share/fonts/")
	#
	# Else, if the operating system is Windows
	#
	elif os.name == 'nt':
		#
		# Save the corresponding font folder
		#
		default_font_folder = os.path.join(os.getenv('LOCALAPPDATA'), 'Microsoft', 'Windows', 'Fonts')
	#
	# Return the default font folder for the fonts files
	#
	return default_font_folder

#
# Go through a list of font files and install them in the scope of the user
#
@logging.log_call
def install_list_of_fonts_files(list_of_fonts_files:list) -> bool:
	#
	# Get the default folder for the fonts files depending on the operating system
	#
	default_font_folder = get_default_font_folder_from_operating_system()
	#
	# If the operating system is unknown
	#
	if default_font_folder is None:
		#
		# Write it in the console
		#
		print(f'Unable to install the fonts "{list_of_fonts_files}". Unsupported operating system.')
		#
		# Quit the function with an error code
		#
		return False
	#
	# Ensure the font folder exists
	#
	os.makedirs(default_font_folder, exist_ok=True)
	#
	# Go through all the files in the list
	#
	for file in list_of_fonts_files:
		#
		# Update the path from the template folder to the user's fonts folder
		#
		font_destination_path = os.path.join(default_font_folder, os.path.basename(file))
		#
		# If the file is not a font file
		#
		if not file.lower().endswith('.ttf'):
			#
			# Write it in the console
			#
			logging.log(f'File "{file}" is not a font file.', "info")
			#
			# Go to the next file
			#
			continue
		#
		# Else, if the font is already installed
		#
		elif os.path.exists(font_destination_path):
			#
			# Write it in the console
			#
			logging.log(f'Font "{font_destination_path}" already installed.', "info")
			#
			# Go to the next file
			#
			continue
		#
		# Else, if the file is a font file and is not already installed
		#
		else:
			#
			# Install the font
			#
			shutil.copy(file, default_font_folder)
			#
			# Write it in the console
			#
			logging.log(f'Font "{file}" installed.', "info")
	#
	# Quit the function with a success code
	#
	return True

#
# Convert formated data from XML to JSON
#
@logging.log_call
def convert_xml_to_json(xml_file_path):
	"""
	Convert XML file to JSON.

	Args:
	xml_file_path (str): Path to the XML file.

	Returns:
	dict: JSON representation of the XML data.
	"""
	# Parse XML
	tree = xml.etree.ElementTree.parse(xml_file_path)
	root = tree.getroot()

	# Convert XML to JSON
	json_data = {}

	# Convert XML tree to JSON recursively
	def parse_element(element):
		result = {}
		for child in element:
			if child:
				if child.tag in result:
					if isinstance(result[child.tag], list):
						result[child.tag].append(parse_element(child))
					else:
						result[child.tag] = [result[child.tag], parse_element(child)]
				else:
					result[child.tag] = parse_element(child)
			else:
				result[child.tag] = child.text
		return result

	json_data[root.tag] = parse_element(root)
	return json_data

#
# Install on the system the fonts from the template folder
#
@logging.log_call
def install_template_fonts() -> None:
	#
	# Get the files used in the template
	#
	template_files = find_files(config.get("PATH_TEMPLATE"))
	#
	# Get the fonts used in the tempalte
	# Example: ["./assets/templates/MyFirstTemplate/fonts/MyFirstFont/MyFirstFont.ttf"]
	#
	template_font_files = filter_files_by_extension(template_files, "ttf")
	#
	# If There is no font to install
	#
	if len(template_font_files) == 0 :
		#
		# Write it in the console
		#
		logging.log(f'No fonts found to install in the template folder "{config.get("PATH_TEMPLATE")}".', "info")
	#
	# Else, if there are fonts to install
	#
	else:
		#
		# Install them
		#
		install_list_of_fonts_files(template_font_files)

############################################################################### PROGRAM CONFIGURATION

#
# Load the configuration of the program into the session (log level, font size...)
#
def load_config() -> None:
	#
	# Load the content of the configuration file into a "config" variable
	#
	# Example of content of the configuration file: 
	#	
	#	# Level of informations logged during the execution of the program:
	#	#
	#	# 	debug 	-> 	All the available logs
	#	#	info 	->	Only the steps of the execution of the program
	#	#	warning ->	Only information that can affect the execution of the program
	#	#	error 	->	Only the failures of the execution of the program
	#	#	silent 	->	Nothing
	#	#
	#	# Default:
	#	#
	#	# 	LOG_LEVEL = info
	#	#
	#	LOG_LEVEL = info
	#	...
	#
	if not config.load(PATH_CONFIG):
		#
		# Log the error in the console
		#
		logging.log("Unable to load the configuration file. Exiting the program...", "error")
		#
		# Exit the program
		#
		exit()
	#
	# Update the level of logs from which the logs should be printed
	#
	update_log_level()
	#
	# Install on the system the fonts from the template folder (if any)
	#
	install_template_fonts()

#
# Get the data of risks from their PingCastle IDs
#
@logging.log_call
def get_json_database() -> list:
	#
	# File unifying the ID given by each tool for a same risk
	#
	json_database_file = config.get("PATH_MAPPED_RISKS")
	#
	# Open the file mapping the IDs of the referentials (PingCastle, ANSSI, ...)
	#
	with open(json_database_file, 'r', encoding=config.get("ENCODING_MAPPED_RISKS")) as json_database_fd:
		#
		# Convert the JSON text in a JSON object
		#
		return json.load(json_database_fd)

#
# Parse the arguments given to the program
#
@logging.log_call
def parse_arguments() -> None:
	#
	# Define a new argument parser
	#
	parser = argparse.ArgumentParser()
	#
	# Define the arguments that can be passed to the program
	# 
	parser.add_argument('-pc', '--pingcastle', type=str, default=f'{config.get("PATH_INPUT_PINGCASTLE")}/ad_hc_*.xml', help='Path to the PingCastle XML file.')
	#
	# Define the arguments that can be passed to the program
	# 
	parser.add_argument('-pk', '--purpleknight', type=str, default=f'{config.get("PATH_INPUT_PURPLEKNIGHT")}/Security_Assessment_Report_*.xlsx', help='Path to the PurpleKnight XLSX file.')
	#
	# Get the arguments passed to the program
	#
	args = parser.parse_args()
	#
	# If the path to the PingCastle file has been passed to the program
	#
	if hasattr(args, 'pingcastle') and args.pingcastle is not None:
		#
		# Save the path to the PingCastle file
		#
		config.set("pingcastle_file_path", args.pingcastle) 
	#
	# Else, if the path to the PingCastle file has not been passed to the program
	#
	else:
		#
		# Try to request it from the user
		#
		try:
			config.set("pingcastle_file_path", pingcastle.request_pingcastle_file_path())
		#
		# If it fails
		#
		except KeyboardInterrupt:
			#
			# Quit the program
			#
			return
	#
	# If the path to the PurpleKnight file has been passed to the program
	#
	if hasattr(args, 'purpleknight') and args.purpleknight is not None:
		#
		# Save the path to the PurpleKnight file
		#
		config.set("purpleknight_file_path", args.purpleknight) 
	#
	# Else, if the path to the PurpleKnight file has not been passed to the program
	#
	else:
		#
		# Try to request it from the user
		#
		try:
			config.set("purpleknight_file_path", purpleknight.request_pingcastle_file_path())
		#
		# If it fails
		#
		except KeyboardInterrupt:
			#
			# Quit the program
			#
			return

#
# Update the level of logs from which the logs should be printed
#
@logging.log_call
def update_log_level() -> None:
	#
	# Update the level of logs to write in the console.
	# Example: "info"
	#
	if not logging.update_log_level(config.get("LOG_LEVEL")):
		#
		# If it fails, Write it in the console
		#
		logging.log(f'Unable to update the log level to "{config.get("LOG_LEVEL")}". Default log level used.', "warning")

############################################################################### INPUT FILES

#
# Get the list of files in the input folder
#
@logging.log_call
def get_input_files() -> list:
	#
	# Return the list of files in the input folder
	#
	return find_files(config.get("PATH_INPUTS"))

#
# Filter the PingCastle files from a list of files
#
@logging.log_call
def filter_pingcastle_files(files: list) -> list:
	#
	# Filter out the input files without a XML extension
	#
	return filter_files_by_extension(files, "xml")

#
# Filter the PurpleKnight files from a list of files
#
@logging.log_call
def filter_purpleknight_files(files: list) -> list:
	#
	# Filter out the input files without a XLSX extension
	#
	return filter_files_by_extension(files, "xlsx")

#
# Get the datetime of the PingCastle report
#
@logging.log_call
def get_pingcastle_report_datetime(file_path:str) -> str:
	#
	# Get the content of the report
	#
	report = convert_xml_to_json(file_path)
	#
	# Return the datetime of the report
	#
	return report["HealthcheckData"]["GenerationDate"]

#
# Get the datetime of the PurpleKnight report
#
@logging.log_call
def get_purpleknight_report_datetime(file_path:str) -> str:
	# Open the workbook
	workbook = openpyxl.load_workbook(file_path)
	
	# Select the specified sheet
	sheet = workbook["Assessment summary"]
	
	# Extract data from the specified column
	dt = sheet["B3"].value

	# Define the default timezone (UTC+2)
	default_timezone = pytz.timezone("Europe/Berlin")

	# Localize the datetime object to the specified timezone
	localized_dt = default_timezone.localize(dt)

	# Convert to ISO 8601 format with the desired precision
	iso_format = localized_dt.strftime("%Y-%m-%dT%H:%M:%S.%f%z")

	# Format the timezone offset to include the colon
	iso_format = iso_format[:-2] + ":" + iso_format[-2:]

	# Return the formated datetime
	return iso_format

#
# Sort the PingCastle files by datetime
#
@logging.log_call
def sort_pingcastle_files_by_datetime(pingcastle_files:list) -> dict:
	#
	# Create the sorted PingCastle files
	#
	sorted_pingcastle_files = {}
	#
	# Go through all the PingCastle files
	#
	for path in pingcastle_files:
		#
		# Index the PingCastle file with its datetime
		#
		sorted_pingcastle_files[get_pingcastle_report_datetime(path)] = path
	#
	# Return the sorted PingCastle files
	#
	return sorted_pingcastle_files

#
# Sort the PurpleKnight files by datetime
#
@logging.log_call
def sort_purpleknight_files_by_datetime(purpleknight_files:list) -> dict:
	#
	# Create the sorted PingCastle files
	#
	sorted_purpleknight_files = {}
	#
	# Go through all the PingCastle files
	#
	for path in purpleknight_files:
		#
		# Index the PingCastle file with its datetime
		#
		sorted_purpleknight_files[get_purpleknight_report_datetime(path)] = path
	#
	# Return the sorted PingCastle files
	#
	return sorted_purpleknight_files

#
# Sort the list of input files by tool and datetime
#
@logging.log_call
def sort_input_files(input_files:list) -> dict:
	#
	# Create the sorted list
	#
	sorted_input_files = {
		"pingcastle": sort_pingcastle_files_by_datetime(filter_pingcastle_files(input_files)),
		"purpleknight": sort_purpleknight_files_by_datetime(filter_purpleknight_files(input_files))
	}
	#
	# Return the sorted input files
	#
	return sorted_input_files

#
# Process the tools input files
#
@logging.log_call
def process_input_files(json_database, sorted_input_files:dict) -> dict:
	#
	# Process the JSON database with the PingCastle data
	#
	json_database = process_pingcastle_files(json_database, sorted_input_files["pingcastle"])
	#
	# Process the JSON database with the PurpleKnight data
	#
	json_database = process_purpleknight_files(json_database, sorted_input_files["purpleknight"])
	#
	# Return the processed database
	#
	return json_database

############################################################################### PING CASTLE

#
# Mark the risks found in the PingCastle file
#
@logging.log_call
def mark_risks_found(json_database:list, datetime:str, ids_of_risks_to_mark:list, key_of_list_of_ids:str) -> list:
	#
	# Go through all the risks in the JSON database
	#
	for current_risk_index, current_risk in enumerate(json_database["risks"]):
		#
		# If the current risk does not have a PingCastle or PurpleKnight ID
		#
		if key_of_list_of_ids not in current_risk["frameworks"].keys():
			#
			# Mark the risk as found in the PingCastle file
			#
			current_risk["found"][datetime] = None
			#
			# Go to the next risk
			#
			continue
		#
		# Mark the risk as found in the PingCastle file
		#
		json_database["risks"][current_risk_index]["found"][datetime] = False
		#
		# Go through all the PingCastle or PurpleKnight IDs of the current risk
		#
		for current_pingcastle_id in current_risk["frameworks"][key_of_list_of_ids]:
			#
			# If the current PingCastle ID is not in the list of found IDs
			#
			if current_pingcastle_id not in ids_of_risks_to_mark:
				#
				# Go to the next PingCastle ID
				#
				continue
			#
			# Mark the risk as found in the PingCastle file
			#
			json_database["risks"][current_risk_index]["found"][datetime] = True
			#
			# Go to the next risk
			#
			break
	#
	# Return the list of unified risks
	#
	return json_database

#
# Process the JSON database with the PingCastle input files
#
@logging.log_call
def process_pingcastle_files(json_database:dict, sorted_pingcastle_files:dict) -> dict:
	#
	# Go through all the PingCastle sorted files
	#
	for datetime, file_path in sorted_pingcastle_files.items():
		#
		# Process the current PingCastle file
		#
		json_database = process_pingcastle_file(json_database, datetime, file_path)
	#
	# Return the processed database
	#
	return json_database

#
# Process the JSON database with the PingCastle data
#
@logging.log_call
def process_pingcastle_file(json_database:list, datetime:str, file_path:str) -> list:
	#
	# Get the ID of all the risks listed in the PingCastle file
	#
	risk_ids_from_pingcastle_file = pingcastle.get_risk_ids_from_pingcastle_file(file_path)
	#
	# Get the unified ID for all the PingCastle risks ID
	#
	json_database = mark_risks_found(json_database, datetime, risk_ids_from_pingcastle_file, "pingcastle")
	#
	# Return the filtered JSON database
	#
	return json_database

############################################################################### PURPLE KNIGHT

#
# Process the JSON database with the PurpleKnight input files
#
@logging.log_call
def process_purpleknight_files(json_database:dict, sorted_purpleknight_files:dict) -> dict:
	#
	# Go through all the PurpleKnight sorted files
	#
	for datetime, file_path in sorted_purpleknight_files.items():
		#
		# Process the current PurpleKnight file
		#
		json_database = process_purpleknight_file(json_database, datetime, file_path)
	#
	# Return the processed database
	#
	return json_database

#
# Process the JSON database with the PurpleKnight data
#
@logging.log_call
def process_purpleknight_file(json_database:list, datetime:str, file_path:str) -> list:
	#
	# Get the ID of all the risks listed in the PurpleKnight file
	#
	risk_ids_from_purpleknight_file = purpleknight.get_risk_ids_from_purpleknight_file(file_path)
	#
	# Get the unified ID for all the PurpleKnight risks ID
	#
	json_database = mark_risks_found(json_database, datetime, risk_ids_from_purpleknight_file, "purpleknight")
	#
	# Return the filtered JSON database
	#
	return json_database

############################################################################### CHARTS

def create_stacked_bar_chart(risks_data) -> None:

	datetimes = list(risks_data[0]["found"].keys())

	# Create a DataFrame from the input data
	data = {
		"Datetime": datetimes,
		"New positive": [0] * len(datetimes),
		"Positive, not retested": [0] * len(datetimes),
		"Still positive": [0] * len(datetimes),
		"Still negative": [0] * len(datetimes),
		"Negative, not retested": [0] * len(datetimes),
		"New negative": [0] * len(datetimes)
	}

	previous_datetime = None

	for risk_data in risks_data:
		for index, datetime in enumerate(risk_data["found"]):
			if index == 0:
				if (risk_data["found"][datetime] == True):
					data["New positive"][index] += 1
				elif (risk_data["found"][datetime] == False):
					data["New negative"][index] -= 1
			else:
				#
				if risk_data["found"][datetime] == True:
					#
					if len([value for value in list(risk_data["found"].values())[:index] if value != None]) == 0:
						#
						data["New positive"][index] += 1
					#
					elif risk_data["found"][previous_datetime] == True:
						#
						data["Still positive"][index] += 1
				#
				elif risk_data["found"][datetime] == False:
					#
					if len([value for value in list(risk_data["found"].values())[:index] if value != None]) == 0:
						#
						data["New negative"][index] -= 1
					#
					elif risk_data["found"][previous_datetime] == False:
						#
						data["Still negative"][index] += 1
				#
				elif risk_data["found"][datetime] == None:
					#
					if risk_data["found"][previous_datetime] == True:
						#
						data["Positive, not retested"][index] += 1
					#
					elif risk_data["found"][previous_datetime] == False:
						#
						data["Negative, not retested"][index] -= 1
			#
			previous_datetime = datetime

	data = pandas.DataFrame(data)

	# Melt the DataFrame to long format for seaborn
	data_melted = pandas.melt(data, id_vars='Datetime', var_name='Category', value_name='Value')
	
	print(data_melted)

	# Set up the matplotlib figure
	matplotlib.pyplot.figure(figsize=(12, 8))

	palette = {
		"New positive": "blue",
		"Positive, not retested": "red",
		"Still positive": "yellow",
		"Still negative": "green",
		"Negative, not retested": "purple",
		"New negative": "orange"
	}
	
	# Create the bar plot
	ax = seaborn.barplot(x='Datetime', y='Value', hue='Category', data=data_melted, palette=palette, dodge=False)
	
	# Set the title and labels
	ax.set_title(None)
	ax.set_xlabel(None)
	ax.set_ylabel(None)
	
	# Rotate x-tick labels for better readability
	matplotlib.pyplot.xticks(rotation=45)
	
	# Show the plot
	matplotlib.pyplot.tight_layout()
	#matplotlib.pyplot.show()
	#exit()

	#
	# Export the chart
	#
	#matplotlib.pyplot.savefig(chart_data["export"]["path"], format=chart_data["export"]["format"], transparent=chart_data["export"]["keep_transparency"])


def export_risk_history_graph(risk_data):

	#
	dates = [datetime.datetime.fromisoformat(str_datetime).strftime("%d %B %Y") for str_datetime in risk_data["found"].keys()]
	values = list(risk_data["found"].values())

	# Create a DataFrame from the input data
	data = pandas.DataFrame({'Datetime': dates, 'Still positive': values})

	# Create the plot
	matplotlib.pyplot.figure(figsize=(12, 2), facecolor="white")

	seaborn.set(font=config.get("FONT_NAME"))
	seaborn.despine()
	
	# Plot the line
	ax = seaborn.lineplot(x='Datetime', y=[1]*len(dates), data=data, marker='o', color='#E2E2E2')
	ax.grid(False)
	ax.set_xticks(dates)
	ax.set_xticklabels(dates)
	ax.tick_params(labelsize=config.get("FONT_SIZE"))
	ax.tick_params(labelsize=config.get("FONT_SIZE"))
	#ax.set_yticks([1])
	
	# Change marker colors
	for date, value in zip(dates, values):
		if value == None:
			continue
		elif value == True:
			ax.plot(date, 1, 'o', color='#D05252')
		elif value == False:
			ax.plot(date, 1, 'o', color='#60AD5E')

	#matplotlib.pyplot.xticks(rotation=90)
	matplotlib.pyplot.yticks([])
	matplotlib.pyplot.ylabel(None)
	matplotlib.pyplot.xlabel(None)
	matplotlib.pyplot.title(None)
	matplotlib.pyplot.tight_layout()

	export_path = os.path.join(config.get("CHARTS_FOLDER"), f'{risk_data["uid"]}.tiff')
	matplotlib.pyplot.savefig(export_path, format="tiff", transparent=True)
	matplotlib.pyplot.close()

	return export_path

#
# Create and export a bar chart
#
@logging.log_call
def create_bar_chart(chart_data) -> None:
	#
	# Example:
	#
	#	data_frame = {
	#		"Category": [A, B, C, D],
	#   	"Still positive": [17, 12, 15, 4],
	#   	"Not Still positive": [-53, -33, -35, -8]
	#	}
	#
	data_frame_values = {
		"Category": list(chart_data["categories"].keys())
	}
	#
	# Go through all the stacked bars
	#
	for stacked_bar in chart_data["stacked_bars"]:
		#
		# Add a new stacked bar with a value in each category
		#
		data_frame_values[stacked_bar["legend"]] = [stacked_bar["categories"][category]["value"] for category in data_frame_values["Category"]]
	#
	# Convert JSON data to DataFrame
	#
	data_frame = pandas.DataFrame(data_frame_values)
	#
	# Font
	#
	seaborn.set(font=chart_data["style"]["font"])
	#
	# Plot
	#
	matplotlib.pyplot.figure(figsize=(chart_data["style"]["width"], chart_data["style"]["height"]), facecolor=chart_data["style"]["background_color"])
	#
	# Go through all the stacked bars
	#
	for stacked_bar in chart_data["stacked_bars"]:
		#
		#
		#
		axis = seaborn.barplot(y=stacked_bar["legend"], x="Category", data=data_frame, color=stacked_bar["background_color"], width=stacked_bar["width_ratio"])
		# 
		# Go through all categories of the current stacked bar
		#
		for index, category in enumerate(stacked_bar['categories'].values()):
			#
			# Define the position of the values of the current category
			#
			middle_height = (category["value"] - (category["value"] / 2)) -1
			#
			# Define the text properties of the values of the current category
			#
			text = category["label"]["value"] if category["label"]["value"] else ""
			#
			#
			#
			axis.text(index, middle_height, text, ha=category["label"]["alignment"], color=category["label"]["font_color"])
		#
		# Adjust the position of the x-axis labels based on the bottom line of the light purple bars
		#
		for index, (category, bar) in enumerate(zip(chart_data["categories"].values(), axis.patches)):
			#
			# Define the position of the labels of the current category
			#
			top_position = bar.get_height() +2
			#
			# Define the text properties of the labels of the current category
			#
			matplotlib.pyplot.text(index, top_position, category["text"], ha=category["alignment"], color=category["font_color"], fontsize=category["font_size"])
	#
	# If the axis are hidden
	#
	if not chart_data["style"]["axis"]:
		#
		# Remove the title of the axis
		#
		matplotlib.pyplot.ylabel(None)
		matplotlib.pyplot.xlabel(None)
		#
		# Remove the graduation of the axis
		#
		matplotlib.pyplot.xticks([])
		matplotlib.pyplot.yticks([])
	#
	# If the legend is shown
	#
	if chart_data["style"]["legend"]["show"]:
		#
		#
		#
		patches_list = [matplotlib.patches.Patch(color=stacked_bar["background_color"], label=stacked_bar["legend"]) for stacked_bar in chart_data["stacked_bars"]]
		#
		# Place the legend
		#
		legend = matplotlib.pyplot.legend(handles=patches_list, fontsize=chart_data["style"]["legend"]["font_size"], loc='upper center', bbox_to_anchor=(chart_data["style"]["legend"]["x_position_ratio"], chart_data["style"]["legend"]["y_position_ratio"]), ncol=chart_data["style"]["legend"]["columns"])
		#
		# Go through all the legends
		#
		for text in legend.get_texts():
			#
			# Set the font color of the curent legend
			#
			text.set_color(chart_data["style"]["legend"]["font_color"])
		#
		# If the legend background transparency is true
		#
		if chart_data["style"]["legend"]["transparent"]:
			#
			# Keep the transparency
			#
			legend.get_frame().set_alpha(0)
	#
	# If the grid parameter is false
	#
	if not chart_data["style"]["grid"]:
		#
		# Remove the grid from the chart
		#
		seaborn.despine()
	#
	# Export the chart
	#
	matplotlib.pyplot.savefig(chart_data["export"]["path"], format=chart_data["export"]["format"], transparent=chart_data["export"]["keep_transparency"])
	matplotlib.pyplot.close()

#
# Create and export a line chart
#
@logging.log_call
def create_line_chart(chart_data) -> None:
	#
	# Create the chart with a custom size
	#
	matplotlib.pyplot.figure(figsize=(10, 6))
	#
	# Remove the title of the chart
	#
	matplotlib.pyplot.title(None)
	#
	# Set the font of the chart
	#
	seaborn.set(font=chart_data["style"]["font"])
	#
	# Add the parts of the lines to the chart
	#
	for line_key, line_data in chart_data["lines"].items():
		for severity, line_part_data in line_data["line_parts"].items():
			severity_color = chart_data["lines"]["average"]["line_parts"][severity]["legend"]["color"]
			matplotlib.pyplot.plot(range(0, len(line_part_data["y_values"])), line_part_data["y_values"], color=severity_color, marker=None, linestyle=line_data["line_style"])
	#
	# Add the filling between the lines
	#
	for severity, line_part_data in chart_data["lines"]["maximum"]["line_parts"].items():
		y1 = [] 
		y2 = [] 
		x1 = []
		x2 = []
		severity_color = chart_data["lines"]["average"]["line_parts"][severity]["legend"]["color"]

		for index, value in enumerate(line_part_data["y_values"]):
			if not numpy.isnan(value):
				x1.append(index)
				y1.append(value)

		for index, value in enumerate(chart_data["lines"]["minimum"]["line_parts"][severity]["y_values"]):
			if not numpy.isnan(value):
				x2.append(index)
				y2.append(value)

		last_y_value = numpy.nan

		x2.append(numpy.nan)
		y2.append(numpy.nan)
		
		for y_index, y_value in enumerate(y1):
			
			y2_value = y2[y_index]
			last_y2_value = y2[y_index-1]
			last_x2_value = x2[y_index -1]

			if (y_index > 0) and (y_value == last_y_value) and (y2_value != last_y2_value):
				x2.insert(y_index, last_x2_value)
				y2.insert(y_index, last_y2_value)

			last_y_value = y_value

		x2.pop()
		y2.pop()

		matplotlib.pyplot.fill_betweenx(y1, x1, x2, edgecolor=severity_color, label=severity, facecolor=severity_color, alpha=0.4)
		
	#
	# Add a legend to the chart
	#
	legend = matplotlib.pyplot.legend(frameon=False, fontsize=chart_data["style"]["legend"]["font_size"], loc='upper center', bbox_to_anchor=(-0.05, 0.6), ncol=1)
	#
	# Go through all the legends
	#
	for text in legend.get_texts():
		#
		# Set the font color of the curent legend
		#
		text.set_color(chart_data["style"]["legend"]["font_color"])
	#
	# Remove the frame of the chart
	#
	matplotlib.pyplot.gca().spines['left'].set_color('none')
	matplotlib.pyplot.gca().spines['bottom'].set_color('none')
	matplotlib.pyplot.gca().spines['right'].set_color('none')
	matplotlib.pyplot.gca().spines['top'].set_color('none')
	#
	# Remove the label of the axis
	#
	matplotlib.pyplot.xlabel(None)
	matplotlib.pyplot.ylabel(None)
	#
	# Customize the ticks of the X axis
	#
	step = 20
	margin = 0 if (len(chart_data["lines"]["maximum"]["line_parts"]["Niveau 1"]["y_values"]) % step) == 0 else step
	days_steps = range(0, len(chart_data["lines"]["maximum"]["line_parts"]["Niveau 1"]["y_values"]) +margin, step)
	matplotlib.pyplot.xticks(days_steps, [f"{days}j" for days in days_steps], size=chart_data["style"]["legend"]["font_size"], color=chart_data["style"]["legend"]["font_color"])
	#
	# Remove the ticks of the Y axis
	#
	matplotlib.pyplot.yticks([])
	#
	# Remove the grid of the chart
	#
	seaborn.despine()
	#
	# Export the chart
	#
	matplotlib.pyplot.savefig(chart_data["export"]["path"], format=chart_data["export"]["format"], transparent=chart_data["export"]["keep_transparency"])
	matplotlib.pyplot.close()

############################################################################### DOCX

#
# Order the risks from severity 1 (worst) to 5
#
@logging.log_call
def order_risks_by_severity(json_database):
	#
	# Create a copy of the json database
	#
	ordered_json_database = copy.deepcopy(json_database)
	#
	# Delete the list of risks from the copied database
	#
	ordered_json_database["risks"] = []
	#
	# Go through the severities from 1 to 5
	#
	for current_severity in range(1, 10):
		#
		# Go through all the risks in the json database
		#
		for current_risk in json_database["risks"]:
			#
			# If the current risk does not have a severity
			#
			if "severity" not in current_risk.keys():
				#
				# Log the error
				#
				logging.log(f'The risk \"{current_risk["uid"]}\" does not have a severity score and cannot be processed.', "error")
				#
				# Go to the next risk
				#
				continue
			#
			# If the severity of the risk match the one being listed
			#
			if current_risk["severity"] == current_severity:
				#
				# Add the current risk to the list
				#
				ordered_json_database["risks"].append(current_risk)
	#
	# Return the ordered list
	#
	return ordered_json_database

#
# Get the name of the concepts necessary to understand the documentation
#
@logging.log_call
def get_concepts_from_documentations(json_database, documentation_id: str, concepts: list) -> list:

	for new_concept in json_database["documentations"][documentation_id]["concepts"]:
		if new_concept not in json_database["documentations"].keys():
			continue
		concepts = get_concepts_from_documentations(json_database, new_concept, concepts)
		concepts.append(new_concept)
	return concepts

#
# Build the DOCX report page by page
#
@logging.log_call
def build_docx_document(sorted_input_files:dict, json_database:dict) -> None:
	#
	# Create the DOCX Manager object that has all the methods to create and export a DOCX document
	#
	my_docx_manager = docx_manager.DocxManager()
	#
	# DOCX header file for the report
	# Example: "./assets/templates/MyFirstTemplate/header.docx"
	#
	header_file = os.path.join(config.get("PATH_TEMPLATE"), "header.docx")
	#
	# Add the DOCX template to open in order to use the updated styles
	#
	my_docx_manager.header_file = header_file
	#
	# Define the path to the DOCX verson of the final report
	#
	my_docx_manager.path = config.get("PATH_OUTPUT_DOCX")
	#
	# Define the path to the PDF version of the final report
	#
	my_docx_manager.export_path = config.get("PATH_OUTPUT_PDF")
	#
	# Go to the next page of the DOCX report
	#
	my_docx_manager.break_page()
	#
	# Import the PingCastle data
	#
	pingcastle_data = convert_xml_to_json(list(sorted_input_files["pingcastle"].values())[-1])
	#
	# Create the base properties of the chart
	#
	chart_data = {
		"style": {
			"font": config.get("FONT_NAME"),
			"background_color": None,
			"width": 12,
			"height": 8,
			"legend": {
				"show": True,
				"columns": 2,
				"x_position_ratio": 0.5,
				"y_position_ratio": 0,
				"font_color" : config.get("CHART_LEGEND_COLOR"),
				"font_size": config.get("FONT_SIZE"),
				"transparent": True
			},
			"axis": False,
			"grid": False
		},
		"export": {
			"format": "tiff",
			"keep_transparency": True,
			"path": os.path.join(config.get("CHARTS_FOLDER"), f"risks_found.tiff")
		},
		"categories": {
			"Niveau 1": {
				"id": "Niveau 1",
				"text": "Niveau 1",
				"alignment": "center",
				"font_size": config.get("FONT_SIZE"),
				"font_color": config.get("CHART_LEGEND_COLOR")
			},
			"Niveau 2": {
				"id": "Niveau 2",
				"text": "Niveau 2",
				"alignment": "center",
				"font_size": config.get("FONT_SIZE"),
				"font_color": config.get("CHART_LEGEND_COLOR")
			},
			"Niveau 3": {
				"id": "Niveau 3",
				"text": "Niveau 3",
				"alignment": "center",
				"font_size": config.get("FONT_SIZE"),
				"font_color": config.get("CHART_LEGEND_COLOR")
			},
			"Niveau 4": {
				"id": "Niveau 4",
				"text": "Niveau 4",
				"alignment": "center",
				"font_size": config.get("FONT_SIZE"),
				"font_color": config.get("CHART_LEGEND_COLOR")
			},
			"Niveau 5": {
				"id": "Niveau 5",
				"text": "Niveau 5",
				"alignment": "center",
				"font_size": config.get("FONT_SIZE"),
				"font_color": config.get("CHART_LEGEND_COLOR")
			}
		},
		"stacked_bars": [
			{
				"legend": "Risques détectées",
				"width_ratio": 0.8,
				"background_color": config.get("CHART_PRIMARY_COLOR"),
				"categories": {
					"Niveau 1": {
						"id": "Niveau 1",
						"value": 0,
						"label": {
							"value": 0,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					},
					"Niveau 2": {
						"id": "Niveau 2",
						"value": 0,
						"label": {
							"value": 0,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					},
					"Niveau 3": {
						"id": "Niveau 3",
						"value": 0,
						"label": {
							"value": 0,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					},
					"Niveau 4": {
						"id": "Niveau 4",
						"value": 0,
						"label": {
							"value": 0,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					},
					"Niveau 5": {
						"id": "Niveau 5",
						"value": 0,
						"label": {
							"value": 0,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					}
				}
			},
			{
				"legend": "Risques non détectées",
				"width_ratio": 0.8,
				"background_color": config.get("CHART_SECONDARY_COLOR"),
				"categories": {
					"Niveau 1": {
						"id": "Niveau 1",
						"value": 0,
						"label": {
							"value": 0,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					},
					"Niveau 2": {
						"id": "Niveau 2",
						"value": 0,
						"label": {
							"value": 0,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					},
					"Niveau 3": {
						"id": "Niveau 3",
						"value": 0,
						"label": {
							"value": 0,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					},
					"Niveau 4": {
						"id": "Niveau 4",
						"value": 0,
						"label": {
							"value": 0,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					},
					"Niveau 5": {
						"id": "Niveau 5",
						"value": 0,
						"label": {
							"value": 0,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					}
				}
			}
		]
	}
	#
	# Go through all the risks of the PingCastle report
	#
	for current_risk in json_database["risks"]:
		#
		if list(current_risk["found"].values())[-1] == True:
			#
			# Increase the count of risks of the current category found
			#
			chart_data["stacked_bars"][0]["categories"][f'Niveau {current_risk["severity"]}']["value"] += 1
			chart_data["stacked_bars"][0]["categories"][f'Niveau {current_risk["severity"]}']["label"]["value"] += 1
		#
		elif list(current_risk["found"].values())[-1] == False:
			#
			# Decrease the count of risks of the current category not found
			#
			chart_data["stacked_bars"][1]["categories"][f'Niveau {current_risk["severity"]}']["value"] -= 1 # The value is already negative, so we add instead of substract
			chart_data["stacked_bars"][1]["categories"][f'Niveau {current_risk["severity"]}']["label"]["value"] -= 1 # The label is positive, to get rid of the minus sign
	#
	# Create a bar chart with the risks found compared to the total in each category
	#
	create_bar_chart(chart_data)

	# #
	# # Create the base properties of the chart
	# #
	# chart_data = {
	# 	"style": {
	# 		"font": config.get("FONT_NAME"),
	# 		"background_color": None,
	# 		"width": 12,
	# 		"height": 8,
	# 		"legend": {
	# 			"show": True,
	# 			"columns": 1,
	# 			"x_position_ratio": -0.2,
	# 			"y_position_ratio": 1,
	# 			"font_color" : config.get("CHART_LEGEND_COLOR"),
	# 			"font_size": config.get("FONT_SIZE"),
	# 			"transparent": True
	# 		},
	# 		"axis": False,
	# 		"grid": False
	# 	},
	# 	"export": {
	# 		"format": "tiff",
	# 		"keep_transparency": True,
	# 		"path": os.path.join(config.get("CHARTS_FOLDER"), f"risks_history.tiff")
	# 	},
	# 	"categories": {
	# 	},
	# 	"stacked_bars": [
	# 		{
	# 			"legend": "Niveau 5",
	# 			"width_ratio": 0.2,
	# 			"background_color": "#D5C1EB",
	# 			"categories": {}
	# 		},
	# 		{
	# 			"legend": "Niveau 4",
	# 			"width_ratio": 0.2,
	# 			"background_color": "#B592DD",
	# 			"categories": {}
	# 		},
	# 		{
	# 			"legend": "Niveau 3",
	# 			"width_ratio": 0.2,
	# 			"background_color": "#9564CE",
	# 			"categories": {}
	# 		},
	# 		{
	# 			"legend": "Niveau 2",
	# 			"width_ratio": 0.2,
	# 			"background_color": "#763BBA",
	# 			"categories": {}
	# 		},
	# 		{
	# 			"legend": "Niveau 1",
	# 			"width_ratio": 0.2,
	# 			"background_color": "#592C8C",
	# 			"categories": {}
	# 		},
	# 		{
	# 			"legend": "Niveau 5 X",
	# 			"width_ratio": 0.2,
	# 			"background_color": "#E5D6B3",
	# 			"categories": {}
	# 		},
	# 		{
	# 			"legend": "Niveau 4 X",
	# 			"width_ratio": 0.2,
	# 			"background_color": "#D5BD85",
	# 			"categories": {}
	# 		},
	# 		{
	# 			"legend": "Niveau 3 X",
	# 			"width_ratio": 0.2,
	# 			"background_color": "#C19D4C",
	# 			"categories": {}
	# 		},
	# 		{
	# 			"legend": "Niveau 2 X",
	# 			"width_ratio": 0.2,
	# 			"background_color": "#A78639",
	# 			"categories": {}
	# 		},
	# 		{
	# 			"legend": "Niveau 1 X",
	# 			"width_ratio": 0.2,
	# 			"background_color": "#7A622A",
	# 			"categories": {}
	# 		}
	# 	]
	# }
	# #
	# #
	# #
	# for current_risk in json_database["risks"]:
	# 	#
	# 	#
	# 	#
	# 	for current_datetime, current_state in current_risk["found"].items():
	# 		#
	# 		#
	# 		#
	# 		if current_datetime not in chart_data["categories"].keys():
	# 			#
	# 			#
	# 			#
	# 			chart_data["categories"][current_datetime] = {
	# 				"id": current_datetime,
	# 				"text": current_datetime[:10],
	# 				"alignment": "center",
	# 				"font_size": config.get("FONT_SIZE"),
	# 				"font_color": config.get("CHART_LEGEND_COLOR")
	# 			}
	# 		#
	# 		#
	# 		#
	# 		for current_severity in chart_data["stacked_bars"]:
	# 			#
	# 			#
	# 			#
	# 			if current_datetime not in current_severity["categories"].keys():
	# 				#
	# 				#
	# 				#
	# 				current_severity["categories"][current_datetime] = {
	# 					"id": current_datetime,
	# 					"value": 0,
	# 					"label": {
	# 						"value": 0,
	# 						"alignment": "center",
	# 						"font_size": config.get("FONT_SIZE"),
	# 						"font_color": "#ffffff"
	# 					}
	# 				}
	# 			#
	# 			#
	# 			#
	# 			if (current_state == True) and (current_severity["legend"] == f'Niveau {current_risk["severity"]}'):
	# 				#
	# 				#
	# 				#
	# 				current_severity["categories"][current_datetime]["value"] += 1
	# 				#
	# 				#
	# 				#
	# 			elif (current_state == False) and (current_severity["legend"] == f'Niveau {current_risk["severity"]} X'):
	# 				#
	# 				#
	# 				#
	# 				current_severity["categories"][current_datetime]["value"] -= 1
	# 				#
	# 				#
	# 				#
	# 			elif (current_state == None) and (current_severity["legend"] == f'Niveau {current_risk["severity"]} X'):
	# 				#
	# 				#
	# 				#
	# 				current_severity["categories"][current_datetime]["value"] += 1

	# pp = pprint.PrettyPrinter(indent=4)
	# #pp.pprint(json_database)
	# #pp.pprint(chart_data)
	# #
	# # Create a bar chart with the history of the risks found in each severity
	# #
	create_stacked_bar_chart(json_database["risks"])

	#
	# Add the title of the chart
	#
	my_docx_manager.title(text="Proportion de risques détectés", level=1, anchor=None)
	#
	# Calculate some stats for the description of the chart
	#
	total_risks_to_detect = sum([chart_data["stacked_bars"][0]["categories"][category]["value"] + (-chart_data["stacked_bars"][1]["categories"][category]["value"]) for category in chart_data["categories"].keys()])
	total_risks_detected = sum([chart_data["stacked_bars"][0]["categories"][category]["value"] for category in chart_data["categories"].keys()])
	detected_risks_ratio = int(round(total_risks_detected * 100 / total_risks_to_detect, 0))
	#
	# Add the description of the chart
	#
	my_docx_manager.add_text(text=f"Les risques détectables lors de cet audit sont catégorisés en fonction de leur sévérité. Celles-ci est notée de \"1\" à \"5\", avec \"1\" étant la plus grande sévérité. Au total, {total_risks_to_detect} risques étaient détectables lors de cet audit, dont {total_risks_detected} ont été détectés, soit {detected_risks_ratio}% de tests positifs (présentant un risque).", anchor=None)
	my_docx_manager.add_text(text=f"\nL'illustration ci-dessous représente la proportion de risques détectés et non-détectés lors de cet audit, pour chacune des sévérités:", anchor=None)
	#
	# Add the chart to the report
	#
	my_docx_manager.add_image(path=chart_data["export"]["path"], width=16, caption="Risques détectées", alignment="center", anchor=None)
	#
	# Go to the next page of the DOCX report
	#
	my_docx_manager.break_page(anchor=None)
	#
	# Create the base structure of the line chart
	#
	chart_data = {		
		"style": {
			"font": config.get("FONT_NAME"),
			"background_color": None,
			"width": 12,
			"height": 8,
			"legend": {
				"show": True,
				"columns": 1,
				"x_position_ratio": -0.1,
				"y_position_ratio": 0.4,
				"font_color" : config.get("CHART_LEGEND_COLOR"),
				"font_size": config.get("FONT_SIZE"),
				"transparent": True
			},
			"axis": False,
			"grid": False
		},
		"export": {
			"format": "tiff",
			"keep_transparency": True,
			"path": os.path.join(config.get("CHARTS_FOLDER"), f"days_to_fix.tiff")
		},
		"lines": {
			"minimum": {
				"line_parts": {
					"Niveau 1": {
						"y_values": []
					},
					"Niveau 2": {
						"y_values": []
					},
					"Niveau 3": {
						"y_values": []
					},
					"Niveau 4": {
						"y_values": []
					},
					"Niveau 5": {
						"y_values": []
					}
				},
				"line_style": "-"
			},
			"average": {
				"line_parts": {
					"Niveau 1": {
						"legend": {
							"value": "Niveau 1",
							"color": "#63329C"
						},
						"y_values": []
					},
					"Niveau 2": {
						"legend": {
							"value": "Niveau 2",
							"color": "#783CBD"
						},
						"y_values": []
					},
					"Niveau 3": {
						"legend": {
							"value": "Niveau 3",
							"color": "#A075D1"
						},
						"y_values": []
					},
					"Niveau 4": {
						"legend": {
							"value": "Niveau 4",
							"color": "#B491DB"
						},
						"y_values": []
					},
					"Niveau 5": {
						"legend": {
							"value": "Niveau 5",
							"color": "#C7ADE5"
						},
						"y_values": []
					}
				},
				"line_style": "--"
			},
			"maximum": {
				"line_parts": {
					"Niveau 1": {
						"y_values": []
					},
					"Niveau 2": {
						"y_values": []
					},
					"Niveau 3": {
						"y_values": []
					},
					"Niveau 4": {
						"y_values": []
					},
					"Niveau 5": {
						"y_values": []
					}
				},
				"line_style": "-"
			}
		}
	}
	#
	# 
	#
	ordered_json_database = order_risks_by_severity(json_database)
	#
	# For each line_key of the chart
	#
	for line_key, line_data in chart_data["lines"].items():
		#
		# Reset the index of the current risk
		#
		risk_index = 0
		#
		# Reset the count of day passed
		#
		day_passed = 0
		#
		#
		#
		risks_to_solve = [risk for risk in ordered_json_database["risks"] if list(risk["found"].values())[-1] == True]
		#
		# As long as there are risks to solve
		#
		while risk_index < len(risks_to_solve) -1:
			#
			#
			#
			if risks_to_solve[risk_index]["days_to_fix"][line_key] <= day_passed:
				#
				#
				#
				risk_index += 1
				#
				# Reset the count of day passed
				#
				day_passed = 0
			#
			# For each of the five severities
			#
			for current_severity in list(set([risk["severity"] for risk in risks_to_solve])):
				#
				# If the current risk is of the current severity
				#
				if current_severity == risks_to_solve[risk_index]["severity"]:
					#
					#
					#
					if (current_severity > 1) and len(line_data["line_parts"][f"Niveau {current_severity}"]["y_values"]) > 0 and numpy.isnan(line_data["line_parts"][f"Niveau {current_severity}"]["y_values"][-1]):
						#
						#
						#
						x = 1
						while True:
							if len(line_data["line_parts"][f"Niveau {current_severity -x}"]["y_values"]) == 0:
								x += 1
								continue
							break
						line_data["line_parts"][f"Niveau {current_severity}"]["y_values"][-1] = line_data["line_parts"][f"Niveau {current_severity -x}"]["y_values"][-2]
							
					#
					#
					#
					line_data["line_parts"][f"Niveau {current_severity}"]["y_values"].append(len(risks_to_solve) -1 - (risk_index +1))
				#
				# Else, if the current risk is not of the current severity
				#
				else:
					#
					# Add an empty value
					#
					line_data["line_parts"][f"Niveau {current_severity}"]["y_values"].append(numpy.nan)
			#
			# Increase the count of days passed
			#
			day_passed += 1
	#
	# Prepare some stats for the description of the chart
	#
	min_estimation = len(chart_data["lines"]["minimum"]["line_parts"][f"Niveau 5"]["y_values"])
	avg_estimation = len(chart_data["lines"]["average"]["line_parts"][f"Niveau 5"]["y_values"])
	max_estimation = len(chart_data["lines"]["maximum"]["line_parts"][f"Niveau 5"]["y_values"])
	avg_estimation_months = int(round(avg_estimation / (5 * 4), 0))
	#
	#
	#
	for line_key in ["minimum", "average"]:
		for severity in range(1, 6):
			# Calculate the difference in lengths
			length_diff = len(chart_data["lines"]["maximum"]["line_parts"][f"Niveau {severity}"]["y_values"]) - len(chart_data["lines"][line_key]["line_parts"][f"Niveau {severity}"]["y_values"])

			# Append nan to the shorter list until both lists have the same size
			if length_diff > 0:
				chart_data["lines"][line_key]["line_parts"][f"Niveau {severity}"]["y_values"] += [numpy.nan] * length_diff
	print(chart_data)
	#
	# Create a bar chart with the risks found compared to the total in each category
	#
	create_line_chart(chart_data)
	#
	# Add the title of the chart
	#
	my_docx_manager.title(text="Durée de correction", level=1, anchor=None)
	#
	# Add the description of the chart
	#
	my_docx_manager.add_text(text="La durée de correction d'une anomalie peut dépendre de beaucoup d'élements tels que la taille du système d'information, les protocoles de sécurité autour de celui-ci, ou encore la disponibilité des équipes compétentes. Il est donc impossible de prédire exactement la date à laquelle toutes les anomalies seraient corrigées. Cependant, il est possible d'utiliser le retour d'expérience de précédents audit Active Directory pour en dégager une tendance.", anchor=None)
	my_docx_manager.add_text(text="\nL'illustration ci-dessous représente les estimations de l'avancée de la correction des anomalies dans le temps. la ligne continue de gauche représente l'estimation optimiste, la ligne discontinue centrale représente l'estimation moyenne et la ligne continue de droite représente l'estimation péssimiste. Les trois estimations débutent en un point commun, en haut à gauche, qui représente le cumul, en hauteur, des anomalies à corriger. Celles-ci déscendent d'une hauteur pour chaque anomalie corrigée. Finalement, les estimations se terminent en bas à droite, lorsque toutes les anomalies sont corrigées. L'écart horizontal entre le point de départ et d'arrivée d'une estimation représente le temps nécessaire à la correction de toutes les anomalies.", anchor=None)
	my_docx_manager.add_text(text=f"\nDans le cas de cet audit, la durée de correction des {total_risks_detected} anomalies est estimée entre {min_estimation} et {max_estimation} jours, avec une moyenne de {avg_estimation} jours, soit environ {avg_estimation_months} mois.", anchor=None)
	#
	# Add the chart to the report
	#
	my_docx_manager.add_image(path=chart_data["export"]["path"], width=16, caption="Résolution des risques", alignment="center", anchor=None)
	#
	# Go to the next page of the DOCX report
	#
	my_docx_manager.break_page(anchor=None)
	#
	# 
	#
	my_docx_manager.title(text="Notions abordées", level=1, anchor=None)
	#
	# 
	#
	concepts = []
	#
	for risk in json_database["risks"]:
		if "concepts" not in risk.keys():
			continue

		for risk_concept in risk["concepts"]:
			if risk_concept in concepts:
				continue

			concepts = get_concepts_from_documentations(json_database, risk_concept, concepts)				
			concepts.append(risk_concept)
	#
	#
	#
	for documentation_id in concepts:
		#
		documentation_path = os.path.join(config.get("PATH_CONCEPTS_DOCUMENTATIONS"), json_database["documentations"][documentation_id]["file_name"])
		#
		title_level = 2
		#
		my_docx_manager.title(json_database["documentations"][documentation_id]["title"], title_level, anchor=None)
		#
		my_docx_manager.bookmark(documentation_id, anchor=None)
		#
		if len(json_database["documentations"][documentation_id]["concepts"]) > 0:
			#
			my_docx_manager.title("Notions", title_level +1, anchor=None)
			#
			for concept_id in json_database["documentations"][documentation_id]["concepts"]:
				#
				my_docx_manager.add_link(json_database["documentations"][concept_id]["title"], f"#{concept_id}", "List Bullet", anchor=None)
			#
			my_docx_manager.add_text("", anchor=None)
		#
		my_docx_manager.append(documentation_path, heading_offset=title_level, anchor=None)
		#
		# Go to the next page of the DOCX report
		#
		my_docx_manager.break_page(anchor=None)
	#
	#
	#
	my_docx_manager.title(text="Détails des risques détectés", level=1, anchor=None)
	#
	# Go through all the unified risks
	#
	for index, mapped_risk in enumerate(json_database["risks"]):
		#
		# If the "found" flag has not been added to the current risk
		#
		if list(mapped_risk["found"].values())[-1] != True:
			#
			# Write it in the console
			#
			logging.log(f'Risk "{mapped_risk["uid"]}" not found in the input files.')
			#
			# Go to the next risk
			#
			continue
		#
		# Get the path to the documentation of the current unified risk
		#
		file_path = os.path.join(config.get("PATH_RISKS_DOCUMENTATIONS"), mapped_risk["file_name"])
		#
		# If the file exists
		#
		if os.path.isfile(file_path):
			#
			title_level = 2
			#
			my_docx_manager.title(mapped_risk["title"], title_level)
			#
			locale.setlocale(locale.LC_TIME, 'fr_FR')
			#
			my_docx_manager.add_text(f'ID METSYS : {str(mapped_risk["uid"]).zfill(3)}', "Subtitle")
			#
			my_docx_manager.title("Historique", title_level +1)
			#
			export_path = export_risk_history_graph(mapped_risk)
			#
			my_docx_manager.add_image(path=export_path, width=16, caption=None, alignment="center", anchor=None)
			#
			my_docx_manager.title("Référentiels", title_level +1)
			#
			table = [["", ""]] * len(mapped_risk["frameworks"].keys())
			#
			table = my_docx_manager.table(table, border_color="#D9D9D9")
			#
			for line_index, framework_id in enumerate(mapped_risk["frameworks"].keys()):
				#
				table.rows[line_index].cells[0].paragraphs[0]._element = my_docx_manager.text(json_database["frameworks"][framework_id]["name"], "Subtitle")._element
				#
				for current_index, current_id in enumerate(mapped_risk["frameworks"][framework_id]):
					#
					current_link = ""
					#
					if framework_id == "pingcastle":
						#
						current_link = "https://pingcastle.com/PingCastleFiles/ad_hc_rules_list.html"
					#
					elif framework_id == "purpleknight":
						#
						current_link = "https://www.semperis.com/purple-knight/security-indicators/"
					#
					elif framework_id == "anssi":
						#
						current_id = current_id[:4] + current_id[5:]
						#
						current_link = f'https://www.cert.ssi.gouv.fr/uploads/ad_checklist.html#{current_id}'
					#
					elif framework_id == "mitre_att&ck":
						#
						if '.' in current_id:
							major_id, minor_id = current_id.split('.')
						#
						else:
							major_id = current_id
							minor_id = None
						#
						if current_id[0] == "M":
							#
							current_link = f'https://attack.mitre.org/mitigations/{major_id}'
						#
						elif current_id[0] == "T":
							#
							current_link = f'https://attack.mitre.org/techniques/{major_id}'
						#
						if minor_id:
							#
							current_link += f'/{minor_id}'
					
					#table[line_index][1].append(my_docx_manager.link(current_id, current_link, "List Bullet"))
			#
			if len(mapped_risk["concepts"]) > 0:
				#
				my_docx_manager.title("Notions", title_level +1)
				#
				for concept_id in mapped_risk["concepts"]:
					#
					my_docx_manager.add_link(json_database["documentations"][concept_id]["title"], f"#{concept_id}", "List Bullet")
			#
			# Add the content of the documentation to the DOCX report
			#
			my_docx_manager.append(file_path, heading_offset=title_level)
			#
			# If there is another documentation to add to the DOCX report
			#
			if (index +1) < len(json_database["risks"]):
				#
				# Go to the next page of the DOCX report
				#
				my_docx_manager.break_page()
		#
		# Else, if the documentation does not exist
		#
		else:
			#
			# Write it in the console
			#
			logging.log(f'Documentation not found at "{file_path}".', "error")
	#
	# Go to the next page of the DOCX report
	#
	my_docx_manager.break_page()
	#
	# DOCX footer file for the report
	# Example: "./assets/templates/MyFirstTemplate/footer.docx"
	#
	footer_file = os.path.join(config.get("PATH_TEMPLATE"), "footer.docx")
	#
	# Add the footer of the report
	#
	my_docx_manager.append(footer_file)
	#
	#
	#
	my_docx_manager.replace_text("[company_name]", "GERFLOR")
	my_docx_manager.replace_text("[company_address]", "26275 SAULCE-SUR-RHONE")
	#
	# Save the last modifications of the DOCX report
	#
	my_docx_manager.save_to_file()
	#
	# Update the table of content of the DOCX report
	#
	my_docx_manager.update_table_of_contents()
	#
	# Update the table of figures of the DOCX report
	#
	my_docx_manager.update_table_of_illustrations()
	#
	# Export the DOCX report to PDF
	#
	my_docx_manager.export()
	#
	# Open the exported PDF report
	#
	my_docx_manager.open_export()

############################################################################### MAIN

#
# Main program
#
@logging.log_call
def main() -> None:
	#
	# Remove the previous generated report
	#
	delete_folder_contents("./output")
	#
	# Load the content of the configuration (log level, fonts of the current template...)
	#
	load_config()
	#
	# Parse the arguments given to the script (PingCastle file)
	#
	#parse_arguments()
	#
	# Get the input files from PingCastle and PurpleKnight
	#
	input_files = get_input_files()
	#
	# Sort the list of input files by tool and datetime
	#
	sorted_input_files = sort_input_files(input_files)
	#
	# Get the JSON database of the risks
	#
	json_database = get_json_database()
	#
	# Process the database with the sorted input files
	#
	processed_database = process_input_files(json_database, sorted_input_files)
	#
	# Build the DOCX report page by page
	#
	build_docx_document(sorted_input_files, json_database)

#
# If the file is executed on its own, and not imported as part of a bigger program
#
if __name__ == '__main__':
	#
	# Execute the main function
	#
	main()