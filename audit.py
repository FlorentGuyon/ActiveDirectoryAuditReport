from argparse import ArgumentParser
from os import path
from lib.docx_manager import DocxManager
from lib.get_risk_ids_from_pingcastle_file import request_file_path, get_risk_ids_from_pingcastle_file
from json import load, loads
from lib.logging import log, update_log_level, log_call

# PATHS
PATH_FILE_DIRECTORY = path.dirname(path.abspath(__file__))
PATH_ASSETS = path.join(PATH_FILE_DIRECTORY, "assets")
PATH_OUTPUT = path.join(PATH_FILE_DIRECTORY, "output")
PATH_DOCUMENTATIONS = path.join(PATH_ASSETS, "documentations")

# FILES
FILE_RACI = path.join(PATH_ASSETS, "raci", "raci.docx")
FILE_MAPPED_RISKS = path.join(PATH_ASSETS, "mapped risks", "mapped_risks.json")
FILE_STYLE_TEMPLATE = path.join(PATH_ASSETS, "templates", "template_metsys.docx")
FILE_OUTPUT_DOCX = path.join(PATH_OUTPUT, "ActiveDirectoryAuditReport.docx")
FILE_OUTPUT_PDF = path.join(PATH_OUTPUT, "ActiveDirectoryAuditReport.pdf")

# ENCODING
ENCODING_FILE_MAPPED_RISKS = "utf-8"

# LOGS
LOG_LEVEL = "info"

@log_call
def get_mapped_risks(pingcastle_ids):
	filtered_mapped_risks = []
	traited_pingcastle_ids = []
	mapped_risks_file = path.join(FILE_MAPPED_RISKS)
	
	with open(mapped_risks_file, 'r', encoding=ENCODING_FILE_MAPPED_RISKS) as mapped_risks_fd:
		mapped_risks = load(mapped_risks_fd)
	
	for pingcastle_id in pingcastle_ids:
		if pingcastle_id in traited_pingcastle_ids:
			continue
		for mapped_risk in mapped_risks:
			if "pingcastle_ids" not in mapped_risk.keys():
				continue
			if pingcastle_id in mapped_risk["pingcastle_ids"]:
				filtered_mapped_risks.append(mapped_risk)
				traited_pingcastle_ids += mapped_risk["pingcastle_ids"]

	return mapped_risks #filtered_mapped_risks

@log_call
def main():
	# LOGS
	update_log_level(LOG_LEVEL)

	# ARGUMENTS
	parser = ArgumentParser(description='Parse a PingCastle HTML report and extract the list of the risks ID')
	parser.add_argument('-f', '--file', type=str, default="input/ad_hc_*.*", help='Path to the PingCastle HTML or XML file.')
	args = parser.parse_args()

	if hasattr(args, 'file') and args.file is not None:
		file_path = args.file 
	else:
		try:
			file_path = request_file_path()
		except KeyboardInterrupt:
			return
	# IDS
	pingcastle_ids = get_risk_ids_from_pingcastle_file(file_path)
	mapped_risks = get_mapped_risks(pingcastle_ids)

	log(f'Mapped risk ids : {", ".join([str(mapped_risk["uid"]) for mapped_risk in mapped_risks])}')

	# DOCX MANAGER
	docx_manager = DocxManager()
	docx_manager.style_template = FILE_STYLE_TEMPLATE
	docx_manager.path = FILE_OUTPUT_DOCX
	docx_manager.export_path = FILE_OUTPUT_PDF
	
	# RACI
	docx_manager.append(FILE_RACI)
	docx_manager.break_page()
	log(f'RACI table added.')

	docx_manager.title("Risques", 1)

	for index, mapped_risk in enumerate(mapped_risks):
		uid = mapped_risk["uid"]
		documentation_file = path.join(PATH_DOCUMENTATIONS, f'{uid}.docx')
		if path.isfile(documentation_file):
			docx_manager.append(documentation_file, heading_offset=1)
			if (index +1) < len(mapped_risks):
				docx_manager.break_page()
		else:
			log(f'Documentation not found at "{documentation_file}".', "error")

	docx_manager.save_to_file()
	docx_manager.update_table_of_contents()
	docx_manager.export()
	docx_manager.open_export()

if __name__ == '__main__':
	main()