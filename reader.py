# reader.py
# functions for a reader

import json, re

sku_idx = 0
handle_idx = 1 # obsolete
collection_idx = 1
title_idx = 2
intro_idx = 3
color_idx = 4
mat_idx = 5
finish_idx = 6
width_idx = 7
depth_idx = 8
height_idx = 9
weight_idx = 10
features_idx = 11
cost_idx = 12
img_src_idx = 13
barcode_idx = 14
gen_vendor_idx = 15
gen_handle_idx = 16

class Measurement:
	def __init__(self, meas_value, meas_type):
		self.meas_value = meas_value
		self.meas_type = meas_type
		
# remove extra whitespaces from all elements
def remove_extra_whitespaces(all_data):
	all_scores_scrubbed = []
	for scores in all_data:
		scrubbed_scores = []
		for value in scores:
			scrubbed_val = re.sub(" ","",value)
			scrubbed_scores.append(scrubbed_val)
		all_scores_scrubbed.append(scrubbed_scores)
	
	return all_scores_scrubbed
	
def strip_data(all_data):
	all_scores_scrubbed = []
	for scores in all_data:
		scrubbed_scores = []
		for value in scores:
			scrubbed_val = value.strip()
			scrubbed_scores.append(scrubbed_val)
		all_scores_scrubbed.append(scrubbed_scores)
	
	return all_scores_scrubbed
	
def strip_list(list):
	scrubbed_scores = []
	for value in list:
		scrubbed_val = value.strip()
		scrubbed_scores.append(scrubbed_val)
	
	return scrubbed_scores
	
def remove_blank_lines(all_data):

	no_blank_data = []
	
	for item in all_data:
		
		if len(item) > 0:
			no_blank_data.append(item)

	return no_blank_data

# get vendor data from a file and format into a list
def extract_vendor_data(vendor, input, extension):
	vendor = re.sub("\s","-",vendor)
	catalog_filename = ''
	if input == "name":
		catalog_filename = "../Data/product-names - " + vendor + "." + extension
	elif input == "handle":
		catalog_filename = "../Data/" + vendor + "-product-import - " + input.capitalize() + "s." + extension
	elif input == "raw data":
		catalog_filename = "../Data/" + vendor + "-catalog - " + input.title() + "." + extension
	elif input == "inventory import":
		input = re.sub("\s","-",input)
		catalog_filename = "../Data/" + vendor + "-" + input + " - New." + extension
	else:
		vendor = re.sub(' ','-',vendor)
		catalog_filename = catalog_filename = "../Data/" + vendor + "-catalog - " + input.capitalize() + "." + extension
	#print("catalog_filename: " + catalog_filename)

	lines = []
	data = []
	all_data = []

	with open(catalog_filename, encoding="UTF8") as catalog_file:

		current_line = ""
		for catalog_info in catalog_file:
			current_line = catalog_info.strip()
			lines.append(current_line)

		catalog_file.close()

	# skip header line
	for line in lines[1:]:
		#print("line: '" + line + "'")
		data = []
		if len(line) > 0:
			if extension == "csv":
				data = line.split(",")
			else:
				data = line.split("\t")
		all_data.append(data)
		
	all_data = remove_blank_lines(all_data)

	all_data = strip_data(all_data)

	return all_data
	
# get warehouse data from a file and format into a list
def extract_warehouse_data(vendor, warehouse, cmd, extension):
	cmd = cmd.title()
	catalog_filename = "../Data/" + vendor + "-" + warehouse + "-stock-status - " + cmd + "." + extension
	
	lines = []
	data = []
	all_data = []

	with open(catalog_filename, encoding="UTF8") as catalog_file:

		current_line = ""
		for catalog_info in catalog_file:
			current_line = catalog_info.strip()
			lines.append(current_line)

		catalog_file.close()

	# skip header line
	for line in lines[1:]:
		if len(line) > 0:
			if extension == "csv":
				data = line.split(",")
			else:
				data = line.split("\t")
		all_data.append(data)

	return all_data
	
# get warehouse data from a file and format into a list
def extract_inventory_data(warehouse, item_type, source_type, cmd, extension):
	cmd = cmd.title()
	catalog_filename = "../Data/" + warehouse + "-" + item_type + "-" + source_type + " - " + cmd + "." + extension
	
	lines = []
	data = []
	all_data = []

	with open(catalog_filename, encoding="UTF8") as catalog_file:

		current_line = ""
		for catalog_info in catalog_file:
			current_line = catalog_info.strip()
			lines.append(current_line)

		catalog_file.close()

	# skip header line
	for line in lines[1:]:
		if len(line) > 0:
			if extension == "csv":
				data = line.split(",")
			else:
				data = line.split("\t")
		all_data.append(data)

	return all_data
	
# get warehouse data from a file and format into a list
def extract_stock_status_data(warehouse, date, table_name, extension):
	cmd = cmd.title()
	catalog_filename = "../Data/" + warehouse + " Stock Status " + date + "." + extension
	
	lines = []
	data = []
	all_data = []

	with open(catalog_filename, encoding="UTF8") as catalog_file:

		current_line = ""
		for catalog_info in catalog_file:
			current_line = catalog_info.strip()
			lines.append(current_line)

		catalog_file.close()

	# skip header line
	for line in lines[1:]:
		if len(line) > 0:
			if extension == "csv":
				data = line.split(",")
			else:
				data = line.split("\t")
		all_data.append(data)

	return all_data

def extract_data(data_type, input='', extension='csv', field_title=False):
	#input = re.sub(' ','-',input)
	#catalog_filename = "../Data/" + input + "-init-data." + extension
	catalog_filename = "../Data/" + input + "-details." + extension
	#catalog_filename = "../Data/" + vendor + "-" + input + "-details." + extension

	

	data_type = re.sub("\s","-",data_type)
	if data_type == 'comparisons':
		catalog_filename = "../Data/" + data_type + " - " + input + "." + extension
	elif data_type == 'products':
		catalog_filename = "../Data/" + data_type + " - " + input.title() + "." + extension
	elif data_type == 'gb-principles': # principles as in numbered statements in gb
		catalog_filename = "../Data/" + data_type + " - " + input.title() + "." + extension
	elif data_type == 'aw-origin' or data_type == 'gb-source' or data_type == 'gb-appendix': # list of aw origin paragraphs or reference/source data or appendix data referenced in book main content or other appendix table
		catalog_filename = "../Data/GB-appendix-data" + " - " + input.title() + "." + extension
	elif data_type == 'aw-tables' or data_type == 'gb-tables': # list of aw or gb tables
		catalog_filename = "../Data/GB-appendix-data" + " - " + input.title() + "." + extension
	elif data_type == 'reference-data': # for references section
		catalog_filename = "../Data/GB-" + data_type + " - " + input + "." + extension
	elif data_type == 'acknowledgements-data': # for acknowledgements section
		catalog_filename = "../Data/GB-" + data_type + " - " + input + "." + extension
	elif data_type == 'gb-intro' or data_type == 'gb-conc': # intro and conc together bc in same spreadsheet GB-intro-conc-data
		catalog_filename = "../Data/GB-intro-conc-data" + " - " + input.title() + "." + extension
	else:
		catalog_filename = "../Data/" + data_type + " - " + input.title() + "." + extension
		
	#print('\ncatalog_filename: \"' + catalog_filename + "\"\n")

	lines = []
	data = []
	all_data = []

	with open(catalog_filename, encoding="UTF8") as catalog_file:

		current_line = ""
		for catalog_info in catalog_file:
			current_line = catalog_info.strip()
			lines.append(current_line)

		catalog_file.close()

	# skip header line if none given
	if field_title == False:
		final_lines = lines[1:]
	else:
		final_lines = lines
		
	for line in final_lines:
		
		if len(line) > 0:
			if extension == "csv":
				data = line.split(",")
			else:
				data = line.split("\t")
		
			all_data.append(data)

	return all_data
	
def extract_appendix_tables(table_nums=[], extension='tsv'):

	#table_nums = ['1','2','3','4','5','6','7','8','9','10','11','12','13']
	
	all_tables = []
	
	if len(table_nums) > 0:
	
		for table_num in table_nums:
		
			#table_num = re.sub("\.","_",table_num)

			catalog_filename = "../Data/appendix-tables/GB-appendix-data - A" + table_num + "." + extension
		
			#print('\ncatalog_filename: \"' + catalog_filename + "\"\n")

			lines = []
			data = []
			all_data = []

			with open(catalog_filename, encoding="UTF8") as catalog_file:

				current_line = ""
				for catalog_info in catalog_file:
					current_line = catalog_info.strip()
					lines.append(current_line)

				catalog_file.close()

			# skip header line
			for line in lines:
		
				if len(line) > 0:
					if extension == "csv":
						data = line.split(",")
					else:
						data = line.split("\t")
		
					all_data.append(data)
				
			all_tables.append(all_data)
			
	else:
		print("Warning: no table nums to extract!")
		
	#print("all_tables: " + str(all_tables))

	return all_tables

def write_data(arranged_data, input):
	input = re.sub(' ','-',input)
	catalog_filename = "../Data/" + input + "-final-data.csv"
	catalog_file = open(catalog_filename, "w", encoding="utf8") # overwrite existing content

	for row_idx in range(len(arranged_data)):
		catalog_file.write(arranged_data[row_idx])
		catalog_file.write("\n")
		#print(catalog[row_idx])

	catalog_file.close()

def read_json(data_type, value_type=''):
	data_type = re.sub(' ','-',data_type)
	value_type = re.sub(' ','-',value_type)
	json_filename = "../Data/" + data_type + ".json"
	if value_type != '':
		json_filename = "../Data/" + value_type + "s/" + data_type + "-" + value_type + "s.json"

	lines = [] # capture each line in the document

	try:
		with open(json_filename, encoding="UTF8") as json_file:
			line = ''
			for json_info in json_file:
				line = json_info.strip()
				lines.append(line)

			json_file.close()
	except:
		print("Warning: No json file!")

	# combine into 1 line
	condensed_json = ''
	for line in lines:
		condensed_json += line

	#print("Condensed JSON: " + condensed_json)

	# parse condensed_json
	dict = json.loads(condensed_json)

	return dict

# valid for json files
def read_keywords(key_type):
	key_type = re.sub(' ','-',key_type)
	keys_filename = "../Data/keywords/" + key_type + "-keywords.json"
	#print("keys_filename: " + keys_filename)

	lines = [] # capture each line in the document

	try:
		with open(keys_filename, encoding="UTF8") as keys_file:
			line = ''
			for key_info in keys_file:
				line = key_info.strip()
				lines.append(line)

			keys_file.close()
	except:
		print("Warning: No keywords file!")

	# combine into 1 line
	condensed_json = ''
	for line in lines:
		condensed_json += line

	#print("Condensed JSON: " + condensed_json)

	# parse condensed_json
	keys = json.loads(condensed_json)

	return keys
	
# valid for json files
def read_glossary(subject):
	subject = re.sub(' ','-',subject)
	keys_filename = "../Data/" + subject + "-glossary.json"
	#print("keys_filename: " + keys_filename)

	lines = [] # capture each line in the document

	try:
		with open(keys_filename, encoding="UTF8") as keys_file:
			line = ''
			for key_info in keys_file:
				line = key_info.strip()
				lines.append(line)

			keys_file.close()
	except:
		print("Warning: No keywords file!")

	# combine into 1 line
	condensed_json = ''
	for line in lines:
		condensed_json += line

	#print("Condensed JSON: " + condensed_json)

	# parse condensed_json
	keys = json.loads(condensed_json)

	return keys

# valid for json files
def read_standards(standard_type):
	standard_type = re.sub(' ','-',standard_type)
	keys_filename = "../Data/standards/" + standard_type + ".json"

	lines = [] # capture each line in the document

	try:
		with open(keys_filename, encoding="UTF8") as keys_file:
			line = ''
			for key_info in keys_file:
				line = key_info.strip()
				lines.append(line)

			keys_file.close()
	except:
		print("Warning: No keywords file!")

	# combine into 1 line
	condensed_json = ''
	for line in lines:
		condensed_json += line

	#print("Condensed JSON: " + condensed_json)

	# parse condensed_json
	keys = json.loads(condensed_json)

	return keys
	


def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

def determine_measurement_type(measurement, handle):

	#print("=== Determine Measurement Type=== ")

	meas_type = 'rectangular' # default to rectangular WxD b/c most common. alt option: round

	measurement = measurement.lower() # could contain words such as "Round" or "n/a"
	#print("Measurement: " + measurement)

	blank_meas = True
	if measurement != 'n/a' and measurement != '':
		blank_meas = False

	if not blank_meas:
		# eg "9\' x 2\'"
		if re.search('\'|\"|”',measurement): # given in units of feet and/or inches

			#print("===Measurement Given in Custom Format===")

			meas_ft_value = meas_in_value = 0.0

			#if re.search("round",measurement):
				#print("Measurement is of Round Object, so Diam or Rad?")

			meas_vars = []
			meas_value = measurement # from here on use value

			if re.search('\'\d+(\"|”)\s*\d+(\'|\"|”)',meas_value):
				meas_type = 'combined_rect' #already set by default
				print("Warning for " + handle + ": Width and Depth given in same field, while determining measurement type: \"" + meas_value + "\"!")
			elif re.search('(\"|”)\s*.+(\"|”)',meas_value) or re.search('(\')\s*.+(\')',meas_value):
				meas_type = 'invalid'
				print("Warning for " + handle + ": 2 values with the same unit given while determining measurement type: \"" + meas_value + "\"!")
			# assuming meas value followed by meas type
			elif re.search('\s+',meas_value):
				#print("Measurement contains a space character.")
				meas_vars = re.split('\s+',meas_value)
				#print("Meas vars: " + str(meas_vars))
				meas_value = meas_vars[0]
				meas_type = meas_vars[1]

	#print("Measurement value: " + meas_value)
	#print("Measurement type: \"" + meas_type + "\"\n")

	return meas_type

# input could already be in standard format, which is plain number with default unit of inches
# if input in format a'b" need to convert to standard format and keep original format for use later or make a function to convert it back
def format_dimension(measurement, handle):
	#print("\n=== Format Dimension ===\n")
	#print("raw measurement for " + handle + ": '" + measurement + "'")

	# define local variables
	total_meas = '1' # return output, result of this function. default=1 for zoho inventory

	measurement = measurement.lower() # could contain words such as "Round", "Square" or "n/a"

	blank_meas = True
	if measurement != 'n/a' and measurement != '':
		blank_meas = False

	if not blank_meas:
	
		# remove extraneous chars
		if re.search("[a-z]\s*(\"|”)",measurement):
			#print("found incorrect format 0.44h")
			measurement = re.sub("[a-z](\"|”)","",measurement) # incorrect format 0.44h"
		if re.search("(\"|”)\s*[a-z]",measurement):
			#print("found incorrect format 0.44"h)
			measurement = re.sub("(\"|”)[a-z]","",measurement) # incorrect format 0.44"h. correct to 0.44
		if re.search("[a-z]\d",measurement):
			measurement = re.sub("[a-z](?=\d)","",measurement) # incorrect format l52. correct to 52
		
		#print("measurement: " + measurement)
	
		# if two nums like 55(79), take larger one to ensure fits in customer room
		# isolate dims given in parentheses
		if re.search("\(",measurement):
			measurement = re.sub("([a-z]|\s+)","",measurement) # if given (4.5 max)
			#print("parentheses in measurement: " + measurement)
			min_max_dims = re.split("\(",measurement) # ['55','79)']
			min_dim = min_max_dims[0]
			#print("min_dim: '" + min_dim + "'")
			max_dim = min_max_dims[1].rstrip(")") #re.sub("\)","",min_max_dims[1]) # remove rear parenthesis
			#print("max_dim: '" + max_dim + "'")
			
			total_meas = max_dim
			
		# isolate dims with whole number and fraction, like 30 1/2
		elif re.search("\d\s+\d/\d",measurement):
			dim_parts = re.split("\s+",measurement)
			dim = dim_parts[0] # for simplicity take the whole number part as the whole measurement, until time to integrate measurement converter so it is decimal
			total_meas = dim
		
		# isolate dims separated by a slash, like 30/39
		elif re.search("/",measurement):
			#print("slash")
			min_max_dims = re.split("/",measurement) # ['30','39']
			min_dim = min_max_dims[0]
			#print("min_dim: '" + min_dim + "'")
			max_dim = min_max_dims[1]
			#print("max_dim: '" + max_dim + "'")
			
			total_meas = max_dim

		# eg "9\' x 2\'"
		elif re.search('\'|\"|”',measurement): # given in units of feet and/or inches

			#print("===Measurement Given in Custom Format===")

			meas_ft_value = meas_in_value = 0.0

			#if re.search("round",measurement):
				#print("Measurement is of Round Object, so Diam or Rad?")

			meas_vars = []
			meas_value = measurement # from here on use value
			meas_type = 'rectangular' # default to rectangular WxD b/c most common. option: round

			# assuming meas value followed by meas type
			if re.search('\s+',meas_value):
				#print("Measurement contains a space character.")
				meas_vars = re.split('\s+',meas_value,1)
				#print("Meas vars: " + str(meas_vars))
				meas_value = meas_vars[0]
				meas_type = meas_vars[1]

			#print("Measurement value: " + meas_value)
			#print("Measurement type: \"" + meas_type + "\"\n")

			# if we find a foot symbol in the measurment
			if re.search("\'",measurement):
				#print("Measurement before removing non-digits: " + measurement)
				# TODO: remove all non-digits

				meas_ft_and_in = meas_value.split("\'")
				#print("Measurement Feet and Inches: " + str(meas_ft_and_in))
				meas_ft = meas_ft_and_in[0].strip()
				if meas_ft != '' and meas_ft.lower() != 'n/a':
					#print("Meas Ft: \"" + meas_ft + "\"")
					meas_ft_value = float(meas_ft)

# 				if meas_ft == '':
# 					print("Meas Ft is blank!")
				#print("Measurement Feet: " + meas_ft)
				#print("Measurement Feet Value: " + str(meas_ft_value))

				meas_in = meas_ft_and_in[1].rstrip("\"").rstrip("”").strip()
				if meas_in != '' and meas_in.lower() != 'n/a':
					#print("Meas In: \"" + meas_in + "\"")
					meas_in_value = float(meas_in)

# 				if meas_in == '':
# 					print("Meas In is blank!")
				#print("Measurement Inches: " + meas_in)
				#print("Measurement Inches Value: " + str(meas_in_value))
			# if measured in inches, not feet
			elif re.search("(\"|”)",meas_value):
				meas_value_data = re.split('(\"|”)',meas_value,1)
				meas_in = meas_value_data[0].rstrip("\"").rstrip("”").strip()

				if meas_in != '' and meas_in.lower() != 'n/a':
					#print("Meas In: \"" + meas_in + "\"")
					meas_in_value = float(meas_in)

# 				if meas_in == '':
# 					print("Meas In is blank!")
				#print("Measurement Inches: " + meas_in)
				#print("Measurement Inches Value: " + str(meas_in_value))

			total_meas = str(int(round(meas_ft_value * 12.0 + meas_in_value)))

		else:
			total_meas = measurement

	else:
		print("Warning for " + handle + ": Invalid measurement found while formatting a dimension: \"" + measurement + "\"!")

	total_meas = re.sub("”","",total_meas)
	#print("Total Measurement (in): " + total_meas)
	return total_meas
	
def format_all_volume_dims(all_dims):

	#print("\n=== Format All Volume Dims ===\n")
	
	#print("all_dims: " + str(all_dims))

	all_valid_dims = []

	# volume_dims in format 94.5 x 35.5 x 30H
	for volume_dim_idx in range(len(all_dims)):
		volume_dims = all_dims[volume_dim_idx].lower()
		
		# if 2 sets of dims given then set to n/a 
		if re.search("h/\s*\d",volume_dims):
			volume_dims = 'n/a'
		
		# remove fb prefix before dims
		if re.search("[a-z]\s*\:",volume_dims):
			#print("found prefix in volume dims at idx " + str(volume_dim_idx))
			volume_dims = re.sub("[a-z]+\s*\:\s*","",volume_dims)
			#print("volume_dims: " + volume_dims)
			
		# remove max bc x used as separator
		if re.search("max",volume_dims):
			#print("found max in volume dims at idx " + str(volume_dim_idx))
			volume_dims = re.sub("\s*max","",volume_dims)
			#print("volume_dims: " + volume_dims)
	
		valid_volume_dims = ''
	
		# if dim has x then we want to split by that delimiter
		if re.search("x", volume_dims.lower()):
		
			dim_data = re.split('\s*x\s*', volume_dims.lower())
			
			# separate dims with \s bt bc maybe typo only included 1 x
			final_dim_data = []
			for dim_datum in dim_data:
				#if re.search("\s+",dim_datum):
					#print("found dim separated by x and space - typo")
					
				dim_data_part = re.split("\s+",dim_datum)
				for part in dim_data_part:
					final_dim_data.append(part)
			
			dim_data = final_dim_data
			
			valid_dim_data = []
		
			for dim in dim_data:
			
				valid_dim = format_dimension(dim, "item")
				#print("valid_dim: " + valid_dim)
				
				valid_dim_data.append(valid_dim)
				
			for dim_idx in range(len(valid_dim_data)):
			
				dim = valid_dim_data[dim_idx]
				
				if dim_idx == 0:
					valid_volume_dims = dim
				else:
					valid_volume_dims += " x " + dim
			
		else:
		
			#print("Warning no x found in volume dims! " + volume_dims.lower())
			
			# if dims has slash we must be sure not fraction
			# if regex /\\s*[lwdh]
			# remove inch markings for single dim later in format dim fcn
			if re.search('/\\s*[lwdh]', volume_dims.lower()):
			
				#print("volume dims separated by slashes")
				
				dim_data = re.split('\s*/\s*', volume_dims.lower())
				#print("dim_data: " + str(dim_data))
			
				valid_dim_data = []
		
				for dim in dim_data:
			
					valid_dim = format_dimension(dim, "item")
					#print("valid_dim: " + valid_dim)
				
					valid_dim_data.append(valid_dim)
				
				for dim_idx in range(len(valid_dim_data)):
			
					dim = valid_dim_data[dim_idx]
				
					if dim_idx == 0:
						valid_volume_dims = dim
					else:
						valid_volume_dims += " x " + dim
						
			elif re.search('\\s*[lwdh]', volume_dims.lower()):
			
				#print("volume dims separated by spaces")
				
				dim_data = re.split('\s+', volume_dims.lower())
				#print("dim_data: " + str(dim_data))
			
				valid_dim_data = []
		
				for dim in dim_data:
			
					valid_dim = format_dimension(dim, "item")
					#print("valid_dim: " + valid_dim)
				
					valid_dim_data.append(valid_dim)
				
				for dim_idx in range(len(valid_dim_data)):
			
					dim = valid_dim_data[dim_idx]
				
					if dim_idx == 0:
						valid_volume_dims = dim
					else:
						valid_volume_dims += " x " + dim
				
			
		#print("valid_volume_dims: " + valid_volume_dims)
		all_valid_dims.append(valid_volume_dims)
		
	return all_valid_dims
	
def read_list_from_file(vendor, source, input, extension):

	#print("\n=== Read List From File ===\n")

	source = re.sub("\s","-", source)
	
	catalog_filename = "../Data/" + vendor + "-" + source + " - " + input.title() + "s." + extension
	if vendor == "":
			catalog_filename = "../Data/" + source + " - " + input.title() + "s." + extension
	#print("catalog_filename: " + catalog_filename)

	lines = []
	data = []
	all_data = []

	with open(catalog_filename, encoding="UTF8") as catalog_file:

		current_line = ""
		for catalog_info in catalog_file:
			current_line = catalog_info.strip()
			lines.append(current_line)

		catalog_file.close()

	# skip header line
	for line in lines[1:]:
		#print("line: '" + line + "'")
		data = []
		if len(line) > 0:
			if extension == "csv":
				data = line.split(",")
			else:
				data = line.split("\t")
				
		item = data[0]
		all_data.append(item)
		
	#all_data = remove_blank_lines(all_data)

	#all_data = strip_data(all_data)

	return all_data
	
def read_product_types():

	#print("\n=== Read List From File ===\n")
	extension = "csv"
	
	catalog_filename = "../Data/product-types.csv"
	#print("catalog_filename: " + catalog_filename)

	lines = []
	data = []
	all_data = []

	with open(catalog_filename, encoding="UTF8") as catalog_file:

		current_line = ""
		for catalog_info in catalog_file:
			current_line = catalog_info.strip()
			lines.append(current_line)

		catalog_file.close()

	# skip header line
	for line in lines[1:]:
		#print("line: '" + line + "'")
		data = []
		if len(line) > 0:
			if extension == "csv":
				data = line.split(",")
			else:
				data = line.split("\t")
				
		item = data[0]
		all_data.append(item)
		
	#all_data = remove_blank_lines(all_data)

	#all_data = strip_data(all_data)

	return all_data
	
# if the width field has w, d, and h then separate vars
# take all_details like [sku,...,width-depth-height,...,barcode]
# make all_details like [sku,...,width,depth,height,...,barcode]
def interpret_catalog_dimensions(all_catalog_details):

	#print("\n=== Interpret Catalog Dimensions ===\n")

	all_details_with_dims = []

	for item_idx in range(len(all_catalog_details)):
		item_catalog_details = all_catalog_details[item_idx]
		
		#print("item_catalog_details: " + str(item_catalog_details))
	
		width_or_dims = item_catalog_details[width_idx]
		
		if re.search("x",width_or_dims):
			# width field contains all dims
			dim_data = re.split("\sx\s",width_or_dims) # separate dims
			#print("dim_data: " + str(dim_data))
			width = dim_data[0]
			depth = dim_data[1]
			height = dim_data[2]
			
			item_catalog_details[width_idx] = width
			item_catalog_details[depth_idx] = depth
			item_catalog_details[height_idx] = height
			
		all_details_with_dims.append(item_catalog_details)
		
	#print("all_details_with_dims: " + str(all_details_with_dims))
			
	#print("\n=== Interpreted Catalog Dimensions ===\n")
			
	return all_details_with_dims
	
def add_rear_dot(init_princ_num):

	final_princ_num = init_princ_num

	# determine if dot already at end of string
	if not re.search("\.$",init_princ_num):
		final_princ_num = init_princ_num + "."
	
	return final_princ_num
	
def sub_special_characters(init_string):

	final_string = re.sub("’","\'",init_string)
	final_string = re.sub("“|”","\"",final_string)
	
	return final_string
	
def read_section_titles(book_title, ch_num):

	#print("\n=== Read Section Titles ===\n")
	
	ch_section_titles = []
	
	if re.search("gb",book_title):
		book_title = "gb"
	
	if ch_num != '': # we cannot accept blank ch num bc we are searching for a specific chapter
		all_section_titles_dicts = read_json("chapter section titles")
		#print("all_section_titles_dicts: " + str(all_section_titles_dicts))
		section_titles_dict = all_section_titles_dicts[book_title]
		ch_section_titles = section_titles_dict[ch_num]
	else:
		print("Warning: ch_num is blank so we cannot get ch_section_titles!")
	
	#print("ch_section_titles: " + str(ch_section_titles))
	return ch_section_titles
	
def read_comparison_section_titles(book_title, ch_num):

	all_ch_section_titles = []

	if re.search(",",book_title):
		# separate by comma and get all section titles
		all_book_titles = book_title.split(",")
		
		for title in all_book_titles:

			ch_section_titles = read_section_titles(title, ch_num)
			
			all_ch_section_titles.append(ch_section_titles)
		
	else:
		print("Warning: comparison title must have comma!")
		
	return all_ch_section_titles