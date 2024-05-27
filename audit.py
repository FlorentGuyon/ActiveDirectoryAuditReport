from argparse import ArgumentParser
from copy import deepcopy
from json import load, loads
from lib.config import Config
from lib.docx_manager import DocxManager
from lib.get_risk_ids_from_pingcastle_file import request_file_path, get_risk_ids_from_pingcastle_file
from lib.logging import log, update_log_level, log_call
from matplotlib import patches, pyplot
from numpy import nan, isnan
from os import path, name, makedirs, getenv, walk, listdir, remove
from pandas import DataFrame, merge
import seaborn as sns
from shutil import copy, rmtree
from xml.etree import ElementTree

# PATHS
PATH_DIRECTORY = path.dirname(path.abspath(__file__))

# FILES
PATH_CONFIG = path.join(PATH_DIRECTORY, "config.txt")

# LOADED CONFIGURATION
config = Config()

#
# Remove the content of a folder
#
@log_call
def remove_contents(folder_path):
    # Iterate over all the contents of the folder
    for item in listdir(folder_path):
        item_path = path.join(folder_path, item)
        if "gitkeep" in item_path:
        	continue
        # Check if the item is a file
        if path.isfile(item_path):
            # Remove the file
            remove(item_path)
        # Check if the item is a directory
        elif path.isdir(item_path):
            # Remove the directory and its contents recursively
            rmtree(item_path)

#
# Get the the unified (or mapped) IDs of the PingCastle risks IDs
#
@log_call
def get_mapped_risks(pingcastle_ids:list) -> list:
	#
	# List of unified risks (risks with a unique ID, mapped with the IDs of each tool)
	#
	filtered_mapped_risks = {}
	#
	# List of PingCastle IDs already mapped
	#
	traited_pingcastle_ids = []
	#
	# File unifying the ID given by each tool for a same risk
	#
	mapped_risks_file = config.get("PATH_MAPPED_RISKS")
	#
	# Open the file mapping the IDs of the referentials (PingCastle, ANSSI, ...)
	#
	with open(mapped_risks_file, 'r', encoding=config.get("ENCODING_MAPPED_RISKS")) as mapped_risks_fd:
		#
		# Convert the JSON text in a JSON object
		#
		mapped_risks = load(mapped_risks_fd) 
	#
	#
	#
	filtered_mapped_risks = deepcopy(mapped_risks)
	filtered_mapped_risks["risks"] = []
	#
	# Go through all the PingCastle IDs
	#
	for pingcastle_id in pingcastle_ids:
		#
		# If the current PingCastle ID has already been listed
		#
		if pingcastle_id in traited_pingcastle_ids:
			#
			# Go to the next ID
			#
			continue
		#
		# Go through all the unified risks
		#
		for mapped_risk in mapped_risks["risks"]:
			#
			# If the current unified risk does not have an equivalent
			# PingCastle ID (the risk is not listed by PingCastle)
			#
			if "pingcastle_ids" not in mapped_risk.keys():
				#
				# Go to the next unified risk
				#
				continue
			#
			# If the current unified risk is equivalent to the risk of the
			# current PingCastle ID
			#
			if pingcastle_id in mapped_risk["pingcastle_ids"]:
				#
				# Add the current unified risk to the list of matches
				#
				filtered_mapped_risks["risks"].append(mapped_risk)
				#
				# Add the current PingCastle ID to the list of IDs already listed
				#
				traited_pingcastle_ids += mapped_risk["pingcastle_ids"]
	#
	# Return the list of unified risks
	#
	return mapped_risks #filtered_mapped_risks

#
# Go through a list of font files and install them in the scope of the user
#
@log_call
def install_fonts(font_file_list:list) -> bool:
	#
	# Fonts folder of the user
	#
	font_folder = None
	#
	# If the operating system is Linux or Mac
	#
	if name == 'posix':
		#
		# Save the corresponding font folder
		#
		font_folder = os.path.expanduser("~/.local/share/fonts/")
	#
	# Else, if the operating system is Windows
	#
	elif name == 'nt':
		#
		# Save the corresponding font folder
		#
		font_folder = path.join(getenv('LOCALAPPDATA'), 'Microsoft', 'Windows', 'Fonts')
	#
	# If the operating system is unknown
	#
	if font_folder is None:
		#
		# Write it in the console
		#
		print(f'Unable to install the fonts "{font_file_list}". Unsupported operating system.')
		#
		# Quit the function with an error code
		#
		return False
	#
	# Ensure the font folder exists
	#
	makedirs(font_folder, exist_ok=True)
	#
	# Go through all the files in the list
	#
	for file in font_file_list:
		#
		# Update the path from the template folder to the user's fonts folder
		#
		font_dest_path = path.join(font_folder, path.basename(file))
		#
		# If the file is not a font file
		#
		if not file.lower().endswith('.ttf'):
			#
			# Write it in the console
			#
			log(f'File "{file}" is not a font file.', "info")
			#
			# Go to the next file
			#
			continue
		#
		# Else, if the font is already installed
		#
		elif path.exists(font_dest_path):
			#
			# Write it in the console
			#
			log(f'Font "{font_dest_path}" already installed.', "info")
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
			copy(file, font_folder)
			#
			# Write it in the console
			#
			log(f'Font "{file}" installed.', "info")
	#
	# Quit the function with a success code
	#
	return True

#
# Go through the subfolders of a folder and return the list of the files with a
# specific extension
#
@log_call
def find_files_by_extension(folder_path: str, extension: str) -> list:
	#
	# List of files in the folder, with the expected extension
	#
	files = []
	#
	# Go through the subdirectory of the directory
	#
	for root, _, filenames in walk(folder_path):
		#
		# Go through the files of the current subdirectory
		#
		for filename in filenames:
			#
			# If the current file has the right extension
			#
			if filename.lower().endswith(extension):
				#
				# Add the file to the list
				#
				files.append(path.join(root, filename))
   	#
	# Return the list of matching files
	#
	return files

#
@log_call
def xml_to_json(xml_file_path):
    """
    Convert XML file to JSON.

    Args:
    xml_file_path (str): Path to the XML file.

    Returns:
    dict: JSON representation of the XML data.
    """
    # Parse XML
    tree = ElementTree.parse(xml_file_path)
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
# Order the risks from severity 1 to 5
#
@log_call
def order_by_severity(mapped_risks):
	#
	# Order the risks in a severity order
	#
	ordered_mapped_risks = deepcopy(mapped_risks)
	ordered_mapped_risks["risks"] = []
	default_severity = 5
	#
	# Go through the five possible severities
	#
	for listed_severity in range(1, 6):
		#
		# Go through all the mapped risks
		#
		for mapped_risk in mapped_risks["risks"]:
			#
			# Set a default severity for the current risk
			#
			max_anssi_severity = default_severity
			#
			# If the current risk is not referenced by ANSSI
			#
			if "anssi_ids" in mapped_risk.keys():
				#
				# Else, go through all the ANSSI IDs corresponding to the current risk
				#
				for anssi_id in mapped_risk["anssi_ids"]:
					#
					# Get the severity of the current ID
					#
					risk_severity = int(anssi_id[4:5])
					#
					# If the severity is more important than the current one
					#
					if risk_severity < max_anssi_severity:
						#
						# Set it as current severity
						#
						max_anssi_severity = risk_severity
			#
			# If the severity of the risk match the one being listed
			#
			if max_anssi_severity == listed_severity:
				#
				# Add the current risk to the list
				#
				ordered_mapped_risks["risks"].append(mapped_risk)
	#
	# Return the ordered list
	#
	return ordered_mapped_risks

#
@log_call
def create_bar_chart(chart_data) -> None:
	#
	# Example:
	#
	#	data_frame = {
	#		"Category": [A, B, C, D],
	#   	"Found": [17, 12, 15, 4],
	#   	"Not Found": [-53, -33, -35, -8]
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
		# Add a new stacked bar with a vlue in each category
		#
		data_frame_values[stacked_bar["legend"]] = [stacked_bar["categories"][category]["value"] for category in data_frame_values["Category"]]
	#
	# Convert JSON data to DataFrame
	#
	data_frame = DataFrame(data_frame_values)
	#
	# Font
	#
	sns.set(font=chart_data["style"]["font"])
	#
	# Plot
	#
	pyplot.figure(figsize=(chart_data["style"]["width"], chart_data["style"]["height"]), facecolor=chart_data["style"]["background_color"])
	#
	# Go through all the stacked bars
	#
	for stacked_bar in chart_data["stacked_bars"]:
		#
		#
		#
		axis = sns.barplot(y=stacked_bar["legend"], x="Category", data=data_frame, color=stacked_bar["background_color"], width=stacked_bar["width_ratio"])
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
			axis.text(index, middle_height, category["label"]["value"], ha=category["label"]["alignment"], color=category["label"]["font_color"])
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
			pyplot.text(index, top_position, category["text"], ha=category["alignment"], color=category["font_color"], fontsize=category["font_size"])
	#
	# If the axis are hidden
	#
	if not chart_data["style"]["axis"]:
		#
		# Remove the title of the axis
		#
		pyplot.ylabel(None)
		pyplot.xlabel(None)
		#
		# Remove the graduation of the axis
		#
		pyplot.xticks([])
		pyplot.yticks([])
	#
	# If the legend is shown
	#
	if chart_data["style"]["legend"]["show"]:
		#
		#
		#
		patches_list = [patches.Patch(color=stacked_bar["background_color"], label=stacked_bar["legend"]) for stacked_bar in chart_data["stacked_bars"]]
		#
		# Place the legend
		#
		legend = pyplot.legend(handles=patches_list, fontsize=chart_data["style"]["legend"]["font_size"], loc='upper center', bbox_to_anchor=(chart_data["style"]["legend"]["x_position_ratio"], chart_data["style"]["legend"]["y_position_ratio"]), ncol=chart_data["style"]["legend"]["columns"])
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
		sns.despine()
	#
	# Export the chart
	#
	pyplot.savefig(chart_data["export"]["path"], format=chart_data["export"]["format"], transparent=chart_data["export"]["keep_transparency"])

#
@log_call
def create_line_chart(chart_data) -> None:
	#
	# Create the chart with a custom size
	#
	pyplot.figure(figsize=(10, 6))
	#
	# Remove the title of the chart
	#
	pyplot.title(None)
	#
	# Set the font of the chart
	#
	sns.set(font=chart_data["style"]["font"])
	#
	# Add the parts of the lines to the chart
	#
	for line_key, line_data in chart_data["lines"].items():
		for severity, line_part_data in line_data["line_parts"].items():
			severity_color = chart_data["lines"]["average"]["line_parts"][severity]["legend"]["color"]
			pyplot.plot(range(0, len(line_part_data["y_values"])), line_part_data["y_values"], color=severity_color, marker=None, linestyle=line_data["line_style"])
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
			if not isnan(value):
				x1.append(index)
				y1.append(value)

		for index, value in enumerate(chart_data["lines"]["minimum"]["line_parts"][severity]["y_values"]):
			if not isnan(value):
				x2.append(index)
				y2.append(value)

		last_y_value = nan

		x2.append(nan)
		y2.append(nan)
		
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

		pyplot.fill_betweenx(y1, x1, x2, edgecolor=severity_color, label=severity, facecolor=severity_color, alpha=0.4)
		
	#
	# Add a legend to the chart
	#
	legend = pyplot.legend(frameon=False, fontsize=chart_data["style"]["legend"]["font_size"], loc='upper center', bbox_to_anchor=(-0.05, 0.6), ncol=1)
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
	pyplot.gca().spines['left'].set_color('none')
	pyplot.gca().spines['bottom'].set_color('none')
	pyplot.gca().spines['right'].set_color('none')
	pyplot.gca().spines['top'].set_color('none')
	#
	# Remove the label of the axis
	#
	pyplot.xlabel(None)
	pyplot.ylabel(None)
	#
	# Customize the ticks of the X axis
	#
	step = 20
	margin = 0 if (len(chart_data["lines"]["maximum"]["line_parts"]["Niveau 1"]["y_values"]) % step) == 0 else step
	days_steps = range(0, len(chart_data["lines"]["maximum"]["line_parts"]["Niveau 1"]["y_values"]) +margin, step)
	pyplot.xticks(days_steps, [f"{days}j" for days in days_steps], size=chart_data["style"]["legend"]["font_size"], color=chart_data["style"]["legend"]["font_color"])
	#
	# Remove the ticks of the Y axis
	#
	pyplot.yticks([])
	#
	# Remove the grid of the chart
	#
	sns.despine()
	#
	# Export the chart
	#
	pyplot.savefig(chart_data["export"]["path"], format=chart_data["export"]["format"], transparent=chart_data["export"]["keep_transparency"])



def get_concepts_from_documentations(mapped_risks, documentation_id: str, concepts: list) -> list:

	for new_concept in mapped_risks["documentations"][documentation_id]["concepts"]:
		if new_concept not in mapped_risks["documentations"].keys():
			continue
		concepts = get_concepts_from_documentations(mapped_risks, new_concept, concepts)
		concepts.append(new_concept)
	return concepts


#
# Main program
#
@log_call
def main() -> None:
	#
	# Remove the previous report
	#
	remove_contents("./output")
	#
	# Load the configuration file
	# Example: {"LOG_LEVEL": "info", ...}
	#
	if not config.load(PATH_CONFIG):
		log("Unable to load the configuration file. Exiting the program...", "error")
		return
	#
	# Update the level of logs to write in the console. If it fails
	# Example: "info"
	#
	if not update_log_level(config.get("LOG_LEVEL")):
		#
		# Write it in the console
		#
		log(f'Unable to update the log level to "{config.get("LOG_LEVEL")}". Default log level used.', "warning")
	#
	# DOCX header file for the report
	# Example: "./assets/templates/MyFirstTemplate/header.docx"
	#
	header_file = path.join(config.get("PATH_TEMPLATE"), "header.docx")
	#
	# DOCX footer file for the report
	# Example: "./assets/templates/MyFirstTemplate/footer.docx"
	#
	footer_file = path.join(config.get("PATH_TEMPLATE"), "footer.docx")
	#
	# Get the fonts to install to use the tempalte
	# Example: ["./assets/templates/MyFirstTemplate/fonts/MyFirstFont/MyFirstFont.ttf"]
	#
	template_fonts = find_files_by_extension(config.get("PATH_TEMPLATE"), "ttf")
	#
	# If There is no font to install
	#
	if len(template_fonts) == 0 :
		#
		# Write it in the console
		#
		log(f'No fonts found to install in the template folder "{config.get("PATH_TEMPLATE")}".', "info")
	#
	# Else, if there are fonts to install
	#
	else:
		#
		# Install them
		#
		install_fonts(template_fonts)
	#
	# Define the available arguments to pass to the program
	# 
	parser = ArgumentParser(description='Parse a PingCastle XML report and extract the list of the risks ID')
	parser.add_argument('-f', '--file', type=str, default="input/ad_hc_*.xml", help='Path to the PingCastle XML file.')
	#
	# Parse the arguments passed to the program
	#
	args = parser.parse_args()
	#
	# If the path to the PingCastle file has been passed to the program
	#
	if hasattr(args, 'file') and args.file is not None:
		#
		# Save the path to the PingCastle file
		#
		file_path = args.file 
	#
	# Else, if the path to the PingCastle file has not been passed to the program
	#
	else:
		#
		# Try to request it from the user
		#
		try:
			file_path = request_file_path()
		#
		# If it fails
		#
		except KeyboardInterrupt:
			#
			# Quit the program
			#
			return
	#
	# Get the ID of all the risks listed in the PingCastle report
	#
	pingcastle_ids = get_risk_ids_from_pingcastle_file(file_path)
	#
	# Get the unified ID for all the PingCastle risks ID
	#
	mapped_risks = get_mapped_risks(pingcastle_ids)
	#
	# Write the list of unified risks ID in the console
	#
	log(f'Mapped risk ids : {", ".join([str(mapped_risk["uid"]) for mapped_risk in mapped_risks["risks"]])}')
	#
	# Create the DOCX Manager object that has all the methods to create and export a DOCX document
	#
	docx_manager = DocxManager()
	#
	# Add the DOCX template to open in order to use the updated styles
	#
	docx_manager.header_file = header_file
	#
	# Define the path to the DOCX verson of the final report
	#
	docx_manager.path = config.get("PATH_OUTPUT_DOCX")
	#
	# Define the path to the PDF version of the final report
	#
	docx_manager.export_path = config.get("PATH_OUTPUT_PDF")
	#
	# Go to the next page of the DOCX report
	#
	docx_manager.break_page()
	#
	# Add the RACI table to the DOCX document
	#
	#docx_manager.append(config.get("PATH_RACI"))
	#
	# Go to the next page of the DOCX report
	#
	#docx_manager.break_page()
	#
	# Write it in the console
	#
	#log(f'RACI table added.')
	#
	# Import the PingCastle data
	#
	pingcastle_data = xml_to_json(find_files_by_extension("input", "xml")[0])
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
			"path": path.join(config.get("CHARTS_FOLDER"), f"risks_found.tiff")
		},
		"categories": {
			"Anomalies": {
				"id": "Anomalies",
				"text": "Anomalies",
				"alignment": "center",
				"font_size": config.get("FONT_SIZE"),
				"font_color": config.get("CHART_LEGEND_COLOR")
			},
			"PrivilegedAccounts": {
				"id": "PrivilegedAccounts",
				"text": "Comptes à privilèges",
				"alignment": "center",
				"font_size": config.get("FONT_SIZE"),
				"font_color": config.get("CHART_LEGEND_COLOR")
			},
			"StaleObjects": {
				"id": "StaleObjects",
				"text": "Objets périmés",
				"alignment": "center",
				"font_size": config.get("FONT_SIZE"),
				"font_color": config.get("CHART_LEGEND_COLOR")
			},
			"Trusts": {
				"id": "Trusts",
				"text": "Relations de confiance",
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
					"Anomalies": {
						"id": "Anomalies",
						"value": 0,
						"label": {
							"value": 0,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					},
					"PrivilegedAccounts": {
						"id": "PrivilegedAccounts",
						"value": 0,
						"label": {
							"value": 0,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					},
					"StaleObjects": {
						"id": "StaleObjects",
						"value": 0,
						"label": {
							"value": 0,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					},
					"Trusts": {
						"id": "Trusts",
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
					"Anomalies": {
						"id": "Anomalies",
						"value": -71,
						"label": {
							"value": 71,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					},
					"PrivilegedAccounts": {
						"id": "PrivilegedAccounts",
						"value": -46,
						"label": {
							"value": 46,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					},
					"StaleObjects": {
						"id": "StaleObjects",
						"value": -51,
						"label": {
							"value": 51,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					},
					"Trusts": {
						"id": "Trusts",
						"value": -12,
						"label": {
							"value": 12,
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
	for risk_rule in pingcastle_data["HealthcheckData"]["RiskRules"]["HealthcheckRiskRule"]:
		#
		# Get the category of the current risk
		#
		category = risk_rule["Category"]
		#
		# Increase the count of risks of the current category found
		#
		chart_data["stacked_bars"][0]["categories"][category]["value"] += 1
		chart_data["stacked_bars"][0]["categories"][category]["label"]["value"] += 1
		#
		# Decrease the count of risks of the current category not found
		#
		chart_data["stacked_bars"][1]["categories"][category]["value"] += 1 # The value is already negative, so we add instead of substract
		chart_data["stacked_bars"][1]["categories"][category]["label"]["value"] -= 1 # The label is positive, to get rid of the minus sign
	#
	# Create a bar chart with the risks found compared to the total in each category
	#
	create_bar_chart(chart_data)
	#
	# Add the title of the chart
	#
	docx_manager.title(text="Proportion de risques détectés", level=1)
	#
	# Calculate some stats for the description of the chart
	#
	total_risks_to_detect = sum([chart_data["stacked_bars"][0]["categories"][category]["value"] + (-chart_data["stacked_bars"][1]["categories"][category]["value"]) for category in chart_data["categories"].keys()])
	total_risks_detected = sum([chart_data["stacked_bars"][0]["categories"][category]["value"] for category in chart_data["categories"].keys()])
	detected_risks_ratio = int(round(total_risks_detected * 100 / total_risks_to_detect, 0))
	#
	# Add the description of the chart
	#
	docx_manager.text(text=f"Les risques détectables lors de cet audit sont catégorisés en fonction du type de donnée qui est à l'origine du risque. Les catégories sont \"Anomalies\", \"Comptes à privilèges\", \"Objets périmés\" et \"Relations de confiance\". Au total, {total_risks_to_detect} risques étaient détectables lors de cet audit, dont {total_risks_detected} ont été détectés, soit {detected_risks_ratio}% de tests positifs (présentant une anomalie).")
	docx_manager.text(text=f"\nL'illustration ci-dessous représente la proportion de risques détectés et non-détectés lors de cet audit, pour chacune des quatre catégories:")
	#
	# Add the chart to the report
	#
	docx_manager.add_image(path=chart_data["export"]["path"], width=16, caption="Risques détectées", alignment="center")
	#
	# Go to the next page of the DOCX report
	#
	docx_manager.break_page()
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
			"path": path.join(config.get("CHARTS_FOLDER"), f"days_to_fix.tiff")
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
	ordered_mapped_risks = order_by_severity(mapped_risks)
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
		# As long as there are risks to solve
		#
		while risk_index < len(ordered_mapped_risks["risks"]) -1:

			if ordered_mapped_risks["risks"][risk_index]["days_to_fix"][line_key] <= day_passed:
				risk_index += 1
				#
				# Reset the count of day passed
				#
				day_passed = 0
			#
			# For each of the five severities
			#
			for severity in range(1, 6):
				#
				# Set a default severity for the current risk
				#
				anssi_severity = 5
				#
				# If the current risk is not referenced by ANSSI
				#
				if "anssi_ids" in ordered_mapped_risks["risks"][risk_index].keys():
					#
					# Else, go through all the ANSSI IDs corresponding to the current risk
					#
					for anssi_id in ordered_mapped_risks["risks"][risk_index]["anssi_ids"]:
						#
						# Get the severity of the current ID
						#
						risk_severity = int(anssi_id[4:5])
						#
						# If the severity is more important than the current one
						#
						if risk_severity < anssi_severity:
							#
							# Set it as current severity
							#
							anssi_severity = risk_severity
				#
				# If the current risk is of the current severity
				#
				if severity == anssi_severity:
					if (severity > 1) and len(line_data["line_parts"][f"Niveau {severity}"]["y_values"]) > 0 and isnan(line_data["line_parts"][f"Niveau {severity}"]["y_values"][-1]):
						line_data["line_parts"][f"Niveau {severity}"]["y_values"][-1] = line_data["line_parts"][f"Niveau {severity -1}"]["y_values"][-2]
					line_data["line_parts"][f"Niveau {severity}"]["y_values"].append(len(mapped_risks["risks"]) -(risk_index +1))
				#
				# Else, if the current risk is not of the current severity
				#
				else:
					#
					# Add an empty value
					#
					line_data["line_parts"][f"Niveau {severity}"]["y_values"].append(nan)
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
			    chart_data["lines"][line_key]["line_parts"][f"Niveau {severity}"]["y_values"] += [nan] * length_diff

	#
	# Create a bar chart with the risks found compared to the total in each category
	#
	create_line_chart(chart_data)
	#
	# Add the title of the chart
	#
	docx_manager.title(text="Durée de correction des risques détectés", level=1)
	#
	# Add the description of the chart
	#
	docx_manager.text(text="La durée de correction d'une anomalie peut dépendre de beaucoup d'élements tels que la taille du système d'information, les protocoles de sécurité autour de celui-ci, ou encore la disponibilité des équipes compétentes. Il est donc impossible de prédire exactement la date à laquelle toutes les anomalies seraient corrigées. Cependant, il est possible d'utiliser le retour d'expérience de précédents audit Active Directory pour en dégager une tendance.")
	docx_manager.text(text="\nL'illustration ci-dessous représente les estimations de l'avancée de la correction des anomalies dans le temps. la ligne continue de gauche représente l'estimation optimiste, la ligne discontinue centrale représente l'estimation moyenne et la ligne continue de droite représente l'estimation péssimiste. Les trois estimations débutent en un point commun, en haut à gauche, qui représente le cumul, en hauteur, des anomalies à corriger. Celles-ci déscendent d'une hauteur pour chaque anomalie corrigée. Finalement, les estimations se terminent en bas à droite, lorsque toutes les anomalies sont corrigées. L'écart horizontal entre le point de départ et d'arrivée d'une estimation représente le temps nécessaire à la correction de toutes les anomalies.")
	docx_manager.text(text=f"\nDans le cas de cet audit, la durée de correction des {total_risks_detected} anomalies est estimée entre {min_estimation} et {max_estimation} jours, avec une moyenne de {avg_estimation} jours, soit environ {avg_estimation_months} mois.")
	#
	# Add the chart to the report
	#
	docx_manager.add_image(path=chart_data["export"]["path"], width=16, caption="Résolution des risques", alignment="center")
	#
	# Go to the next page of the DOCX report
	#
	docx_manager.break_page()
	#
	# 
	#
	docx_manager.title(text="Notions abordées", level=1)
	#
	# 
	#
	concepts = []
	#
	for risk in mapped_risks["risks"]:
		if "concepts" not in risk.keys():
			continue

		for risk_concept in risk["concepts"]:
			if risk_concept in concepts:
				continue

			concepts = get_concepts_from_documentations(mapped_risks, risk_concept, concepts)				
			concepts.append(risk_concept)
	#
	#
	#
	for documentation_id in concepts:
		#
		documentation_path = path.join(config.get("PATH_DOCUMENTATIONS"), mapped_risks["documentations"][documentation_id]["file_name"])
		#
		title_level = 2
		#
		docx_manager.title(mapped_risks["documentations"][documentation_id]["title"], title_level)
		#
		docx_manager.bookmark(documentation_id)
		#
		if len(mapped_risks["documentations"][documentation_id]["concepts"]) > 0:
			#
			docx_manager.title("Notions", title_level +1)
			#
			for concept_id in mapped_risks["documentations"][documentation_id]["concepts"]:
				#
				docx_manager.link(mapped_risks["documentations"][concept_id]["title"], f"#{concept_id}", "List Bullet")
			#
			docx_manager.text("")
		#
		docx_manager.append(documentation_path, heading_offset=title_level)
		#
		# Go to the next page of the DOCX report
		#
		docx_manager.break_page()
	#
	#
	#
	docx_manager.title(text="Détails des risques détectés", level=1)
	#
	# Go through all the unified risks
	#
	for index, mapped_risk in enumerate(mapped_risks["risks"]):
		#
		# Get the path to the documentation of the current unified risk
		#
		file_path = path.join(config.get("PATH_DOCUMENTATIONS"), mapped_risk["file_name"])
		#
		# If the file exists
		#
		if path.isfile(file_path):
			#
			title_level = 2
			#
			docx_manager.title(mapped_risk["title"], title_level)
			#
			if len(mapped_risk["concepts"]) > 0:
				#
				docx_manager.title("Notions", title_level +1)
				#
				for concept_id in mapped_risk["concepts"]:
					#
					docx_manager.link(mapped_risks["documentations"][concept_id]["title"], f"#{concept_id}", "List Bullet")
			#
			# Add the content of the documentation to the DOCX report
			#
			docx_manager.append(file_path, heading_offset=title_level)
			#
			# If there is another documentation to add to the DOCX report
			#
			if (index +1) < len(mapped_risks["risks"]):
				#
				# Go to the next page of the DOCX report
				#
				docx_manager.break_page()
		#
		# Else, if the documentation does not exist
		#
		else:
			#
			# Write it in the console
			#
			log(f'Documentation not found at "{file_path}".', "error")
	#
	# Go to the next page of the DOCX report
	#
	docx_manager.break_page()
	#
	# Add the footer of the report
	#
	docx_manager.append(config.get("PATH_MAPPED_REFERENCES"))
	#
	# Go to the next page of the DOCX report
	#
	docx_manager.break_page()
	#
	# Add the footer of the report
	#
	docx_manager.append(footer_file)
	#
	# Save the last modifications of the DOCX report
	#
	docx_manager.save_to_file()
	#
	# Update the table of content of the DOCX report
	#
	docx_manager.update_table_of_contents()
	#
	# Update the table of figures of the DOCX report
	#
	docx_manager.update_table_of_illustrations()
	#
	# Export the DOCX report to PDF
	#
	docx_manager.export()
	#
	# Open the exported PDF report
	#
	docx_manager.open_export()


if __name__ == '__main__':
	main()