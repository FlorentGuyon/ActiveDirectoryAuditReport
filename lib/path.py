import os

class Path():

	################################################################# SURCHARGE

	def __init__(self, path:str=None) -> None:
		self._abs = None		# C:\Users\MyUser\MyScriptDirectory\MyLogDirectory\Myfile.log
		self._abs_dir = None	# C:\Users\MyUser\MyScriptDirectory\MyLogDirectory
		self._base = None		# C:\Users\MyUser\MyScriptDirectory
		self._extension = None	# log
		self._full_name = None	# Myfile.log
		self._name = None		# Myfile
		self._rel = None		# .\MyLogDirectory\Myfile.log
		self._rel_dir = None	# .\MyLogDirectory

		if path:
			self.update(path)

	def __str__(self) -> str:
		substrings = []
		for attribute, value in vars(self).items():
			substrings.append(f"{attribute}: {str(value)}")
		return "\n".join(substrings)

	################################################################### GETTERS

	@property
	def abs(self) -> str:
		return self._abs

	@property
	def abs_dir(self) -> str:
		return self._abs_dir

	@property
	def base(self) -> str:
		return os.path.dirname(os.path.abspath(__file__))

	@property
	def extension(self) -> str:
		return self._extension

	@property
	def full_name(self) -> str:
		return self._full_name

	@property
	def name(self) -> str:
		return self._name

	@property
	def rel(self) -> str:
		return self._rel

	@property
	def rel_dir(self) -> str:
		return self._rel_dir
	
	################################################################### SETTERS

	@abs.setter
	def abs(self, abs:str) -> None:
		self._abs = abs

	@base.setter
	def base(self, base:str) -> None:
		self._base = base

	@abs_dir.setter
	def abs_dir(self, abs_dir:str) -> None:
		self._abs_dir = abs_dir

	@extension.setter
	def extension(self, extension:str) -> None:
		self._extension = extension

	@full_name.setter
	def full_name(self, full_name:str) -> None:
		self._full_name = full_name

	@name.setter
	def name(self, name:str) -> None:
		self._name = name

	@rel.setter
	def rel(self, rel:str) -> None:
		self._rel = rel

	@rel_dir.setter
	def rel_dir(self, rel_dir:str) -> None:
		self._rel_dir = rel_dir
	
	################################################################### METHODS

	def update(self, path:str) -> None:
		if os.path.isabs(path):
			if os.path.isdir(path):
				self.abs_dir = path
				self.rel_dir = os.path.relpath(path, self.base)
			else:
				self.abs = path
				self.rel = os.path.relpath(path, self.base)
				self.abs_dir = os.path.dirname(path)
				self.rel_dir = os.path.relpath(self.abs_dir)
				self.full_name = os.path.basename(path)
				self.name, self.extension = self.full_name.split('.')
		else:
			if os.path.isdir(path):
				self.rel_dir = path
				self.abs_dir = os.path.abspath(path)
			else:
				self.rel = path
				self.abs = os.path.abspath(path)
				self.rel_dir = os.path.dirname(path)
				self.abs_dir = os.path.abspath(self.rel_dir)
				self.full_name = os.path.basename(path)
				self.name, self.extension = self.full_name.split('.')
		self.base = self.base