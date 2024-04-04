from argparse import ArgumentParser
from json import load, loads
from lib.config import Config
from lib.docx_manager import DocxManager
from lib.get_risk_ids_from_pingcastle_file import request_file_path, get_risk_ids_from_pingcastle_file
from lib.logging import log, update_log_level, log_call
from matplotlib import patches, pyplot
from os import path, name, makedirs, getenv, walk
from pandas import DataFrame, merge
import seaborn as sns
from shutil import copy
from xml.etree import ElementTree

# PATHS
PATH_DIRECTORY = path.dirname(path.abspath(__file__))

# FILES
PATH_CONFIG = path.join(PATH_DIRECTORY, "config.txt")

# LOADED CONFIGURATION
config = Config()

#
# Get the the unified (or mapped) IDs of the PingCastle risks IDs
#
@log_call
def get_mapped_risks(pingcastle_ids:list) -> list:
	#
	# List of unified risks (risks with a unique ID, mapped with the IDs of each tool)
	#
	filtered_mapped_risks = []
	#
	# List of PingCastle IDs already mapped
	#
	traited_pingcastle_ids = []
	#
	# File unifying the ID given by each tool for a same risk
	#
	mapped_risks_file = path.join(config.get("PATH_MAPPED_RISKS"))
	#
	# Open the file mapping the IDs of the referentials (PingCastle, ANSSI, ...)
	#
	with open(mapped_risks_file, 'r', encoding=config.get("ENCODING_MAPPED_RISKS")) as mapped_risks_fd:
		#
		# Convert the JSON text in a JSON object
		#
		mapped_risks = load(mapped_risks_fd)
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
		for mapped_risk in mapped_risks:
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
				filtered_mapped_risks.append(mapped_risk)
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
# Main program
#
@log_call
def main() -> None:
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
	log(f'Mapped risk ids : {", ".join([str(mapped_risk["uid"]) for mapped_risk in mapped_risks])}')
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
	# Add the RACI table to the DOCX document
	#
	docx_manager.append(config.get("PATH_RACI"))
	#
	# Go to the next page of the DOCX report
	#
	docx_manager.break_page()
	#
	# Write it in the console
	#
	log(f'RACI table added.')
	#
	# Add the Risks page
	#
	docx_manager.title("Risques", 1)
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
						"value": -70,
						"label": {
							"value": 70,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					},
					"PrivilegedAccounts": {
						"id": "PrivilegedAccounts",
						"value": -45,
						"label": {
							"value": 45,
							"alignment": "center",
							"font_size": config.get("FONT_SIZE"),
							"font_color": "#ffffff"
						}
					},
					"StaleObjects": {
						"id": "StaleObjects",
						"value": -50,
						"label": {
							"value": 50,
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
	# Add the chart to the report
	#
	docx_manager.add_image(path=chart_data["export"]["path"], width=18.5, caption="Risques détectées", alignment="center")
	#
	# Go to the next page of the DOCX report
	#
	docx_manager.break_page()
	#
	# Go through all the unified risks
	#
	for index, mapped_risk in enumerate(mapped_risks):
		#
		# Get the current unified risk
		#
		uid = mapped_risk["uid"]
		#
		# Get the path to the documentation of the current unified risk
		#
		documentation_file = path.join(config.get("PATH_DOCUMENTATIONS"), f'{uid}.docx')
		#
		# If the file exists
		#
		if path.isfile(documentation_file):
			#
			# Add the content of the documentation to the DOCX report
			#
			docx_manager.append(documentation_file, heading_offset=1)
			#
			# If there is another documentation to add to the DOCX report
			#
			if (index +1) < len(mapped_risks):
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
			log(f'Documentation not found at "{documentation_file}".', "error")
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