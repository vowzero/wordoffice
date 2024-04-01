import abc
import zipfile


class Driver(metaclass=abc.ABCMeta):
	# noinspection PyUnusedLocal
	@abc.abstractmethod
	def __init__(self, docx_path):
		...
	
	@abc.abstractmethod
	def write(self, item_name, blob):
		...
	
	@abc.abstractmethod
	def write_file(self, item_name, filename):
		...
	
	@abc.abstractmethod
	def close(self):
		...
	
	@abc.abstractmethod
	def __enter__(self):
		...
	
	@abc.abstractmethod
	def __exit__(self, *exc_info):
		...


class FileDriver(Driver):
	def __init__(self, docx_path):
		self._zip = zipfile.ZipFile(docx_path, 'w', compression=zipfile.ZIP_DEFLATED)
	
	def write(self, item_name, blob):
		self._zip.writestr(item_name, blob)
	
	def write_file(self, item_name, filename):
		self._zip.write(filename, item_name)
	
	def close(self):
		self._zip.close()
	
	def __enter__(self):
		return self
	
	def __exit__(self, *exc_info):
		self.close()
