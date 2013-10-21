# coding: utf-8
header_tmpl = '''// 
// Virtual table file.
// Version 4.0.01
// 

// Type.
ConnectType = "CSV"

// Encoding.
Encoding = "UTF_8"

// DataSource.
DataSourceName = "%s"
UserID = ""
PassWord = ""

// Select statement.
SQL_SELECT = ""
SQL_FROM = ""
SQL_WHERE = ""
SQL_OTHER = ""

// Date-Time format.
DateType = "0"
DateDelimit = "/"
TimeType = "0"
TimeDelimit = ":"

// CSV options.
Begin CSVOption
	Delimit = ","
	ColNameHeader = "-1"
	CharacterEnclose = """
End CSVOption

// FetchSize.
FetchSize = "0"
// AutoCommit.
AutoCommit = "0"
'''

field_tmpl = '''
Begin VirtualField
	VirtualFieldName = "%s"
	RealFieldName = ""
	Type = "VARCHAR"
	Precision = "0"
	Scale = "0"
	Position = "%d"
	Expression = ""
End VirtualField
'''
