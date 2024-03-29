from argparse import ArgumentParser
from os import path, name, makedirs, getenv, walk
from shutil import copy
from lib.config import Config
from lib.docx_manager import DocxManager
from lib.get_risk_ids_from_pingcastle_file import request_file_path, get_risk_ids_from_pingcastle_file
from json import load, loads
from lib.logging import log, update_log_level, log_call

# PATHS
PATH_FILE_DIRECTORY = path.dirname(path.abspath(__file__))
PATH_ASSETS = path.join(PATH_FILE_DIRECTORY, "assets")
PATH_TEMPLATES = path.join(PATH_ASSETS, "templates")
PATH_OUTPUT = path.join(PATH_FILE_DIRECTORY, "output")
PATH_DOCUMENTATIONS = path.join(PATH_ASSETS, "documentations")

# FILES
FILE_CONFIG = path.join(PATH_FILE_DIRECTORY, "config.txt")
FILE_RACI = path.join(PATH_ASSETS, "raci", "raci.docx")
FILE_MAPPED_RISKS = path.join(PATH_ASSETS, "mapped risks", "mapped_risks.json")
FILE_STYLE_TEMPLATE = path.join(PATH_TEMPLATES, "template_metsys.docx")
FILE_OUTPUT_DOCX = path.join(PATH_OUTPUT, "ActiveDirectoryAuditReport.docx")
FILE_OUTPUT_PDF = path.join(PATH_OUTPUT, "ActiveDirectoryAuditReport.pdf")

# LOADED CONFIGURATION
config = Config()

#
# Get the the unified (or mapped) IDs of the PingCastle risks IDs
#
@log_call
def get_mapped_risks(pingcastle_ids):
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
	mapped_risks_file = path.join(FILE_MAPPED_RISKS)
	#
	# Open the file mapping the IDs of the referentials (PingCastle, ANSSI, ...)
	#
	with open(mapped_risks_file, 'r', encoding=config.get("ENCODING_FILE_MAPPED_RISKS")) as mapped_risks_fd:
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
		print(f"Unable to install the fonts \"{font_file_list}\". Unsupported operating system.")
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
			log(f"File \"{file}\" is not a font file.", "info")
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
			log(f"Font \"{font_dest_path}\" already installed.", "info")
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
			log(f"Font \"{file}\" installed.", "info")
	#
	# Quit the function with a success code
	#
	return True

#
# Go through the subfolders of a folder and return the list of the files with a
# specific extension
#
@log_call
def find_files(folder_path: str, extension: str) -> list:
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
def main():
	#
	# Load the configuration file
	# Example: {"LOG_LEVEL": "info", ...}
	#
	if not config.load(FILE_CONFIG):
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
		log(f"Unable to update the log level to \"{config.get('LOG_LEVEL')}\". Default log level used.", "warning")
	#
	# Path to the template folder
	# Example: "MyFirstTemplate"
	#
	template_path = path.join(PATH_TEMPLATES, config.get("TEMPLATE_NAME"))
	#
	# DOCX template file for the report
	# Example: "./assets/templates/MyFirstTemplate/MyFirstTemplate.docx"
	#
	template_files = find_files(template_path, "docx")
	#
	# If there is no DOCX template file
	#
	if len(template_files) == 0 :
		#
		# Write it in the console
		#
		log(f"Unable to find a DOCX template file in the template folder \"{template_path}\". Exiting the program...", "error")
		#
		# Quit the program
		#
		return
	#
	# Get the fonts to install to use the tempalte
	# Example: ["./assets/templates/MyFirstTemplate/fonts/MyFirstFont/MyFirstFont.ttf"]
	#
	template_fonts = find_files(template_path, "ttf")
	#
	# If There is no font to install
	#
	if len(template_fonts) == 0 :
		#
		# Write it in the console
		#
		log(f"No fonts found to install in the template folder \"{template_path}\".", "info")
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
	parser = ArgumentParser(description='Parse a PingCastle HTML report and extract the list of the risks ID')
	parser.add_argument('-f', '--file', type=str, default="input/ad_hc_*.*", help='Path to the PingCastle HTML or XML file.')
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
	docx_manager.style_template = template_files[0]
	#
	# Define the path to the DOCX verson of the final report
	#
	docx_manager.path = FILE_OUTPUT_DOCX
	#
	# Define the path to the PDF version of the final report
	#
	docx_manager.export_path = FILE_OUTPUT_PDF
	#
	# Add the RACI table to the DOCX document
	#
	docx_manager.append(FILE_RACI)
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
		documentation_file = path.join(PATH_DOCUMENTATIONS, f'{uid}.docx')
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
	# Save the last modifications of the DOCX report
	#
	docx_manager.save_to_file()
	#
	# Update the table of content of the DOCX report
	#
	docx_manager.update_table_of_contents()
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