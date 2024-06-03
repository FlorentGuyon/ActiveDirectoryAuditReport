import datetime
import os
import inspect
import colorama

COLORS = {
	"reset": colorama.Style.RESET_ALL,
	"debug": colorama.Fore.BLUE,
	"info": colorama.Style.RESET_ALL,
	"warning": colorama.Fore.YELLOW,
	"error": colorama.Fore.RED
}
LOG_LEVELS = {
	"debug": 0,
	"info": 1,
	"warning": 2,
	"error": 3,
	"silent": 4
}
LOG_LEVEL = "info"
DATE_FORMAT = "%Y-%m-%d %H:%M:%S" # 2023-12-31 23:59:59

# Initialize colorama
colorama.init()

def print_available_log_levels() -> None:
	print(f'Available log levels : {", ".join(LOG_LEVELS.keys())}')

def log(text:str, level:str="info", caller:str=None) -> bool:
	try:
		level = level.lower()
		minimum_log_level_to_print = LOG_LEVELS[LOG_LEVEL.lower()]
		log_level_of_the_log = LOG_LEVELS[level]
		if log_level_of_the_log < minimum_log_level_to_print:
			return True
		caller = caller if caller else os.path.basename(inspect.getmodule(inspect.stack()[1][0]).__file__)
		formated_datetime = datetime.datetime.now().strftime(DATE_FORMAT)
		formated_log = f'{formated_datetime}\t{caller}\t{level.capitalize()}\t{text.capitalize()}' # 2023-12-31 23:59:59	myfile.py	Info	My log message
		colored_log = f'{COLORS[level]}{formated_log}{COLORS["reset"]}'
		print(colored_log)
	except Exception as e:
		print(f'{COLORS["error"]}Error while logging the message "{text}"" with the log level "{level}" and the date format "{DATE_FORMAT}" : {e}{COLORS["reset"]}')
		print_available_log_levels()

def log_call(method):
	def wrapper(*args, **kwargs):
		text = f'{method.__name__}({str(args[1:])}, {str(kwargs)})'
		caller = method.__module__.split('.')[-1] + ".py"
		log(text, level="debug", caller=caller)
		return method(*args, **kwargs)
	return wrapper

@log_call
def update_log_level(log_level:str) -> bool:
	log_level = log_level.lower()
	if not log_level in LOG_LEVELS.keys():
		log(f'Impossible to update the log level to "{log_level}". Available log levels: {", ".join(LOG_LEVELS.keys())}', "error")
		return False
	global LOG_LEVEL
	LOG_LEVEL = log_level
	log(f'Log level updated to "{log_level}".')
	return True