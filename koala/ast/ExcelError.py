
ErrorCodes = ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A")

class ExcelError(Exception):
	def __init__(self, value, info = None):
		self.value = value
		self.info = info

class EmptyCellError(ExcelError):
    pass
