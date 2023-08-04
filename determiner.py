# determiner.py
# determine true/false or variables

import reader, isolator
import re

# unique variant

# opt_name like 'Size'
def determine_opt_idx(opt_name, item_data):
	opt_idx = 0
	for idx in range(len(item_data)):
		datum = item_data[idx]
		if datum == opt_name:
			opt_name_idx = idx
			opt_idx = opt_name_idx + 1
			#print("found opt name " + opt_name + " at idx " + str(opt_name_idx))
			break
			
	return opt_idx

def determine_field_idx(vendor, keyword, data_type):
	final_field_idx = 0
	
	value_type = "field"
	raw_data_fields = reader.read_json(data_type, value_type)
	
	vendor_fields = raw_data_fields[vendor]
	
	for field_idx in range(len(vendor_fields)):
		field = vendor_fields[field_idx]
		#print("field: " + field)
		#print("keyword: " + keyword.title())
		if field.lower() == keyword.lower():
			final_field_idx = field_idx
			
	#print(final_field_idx)
	return final_field_idx
	
def determine_pkg_row(coll, row_idx):
	pkg_row = False
	prod_row_len = 20
	
	if len(coll) > 1  and row_idx+1 != len(coll):
		row = coll[row_idx]
		next_row = coll[row_idx+1]
		
		if len(row) == prod_row_len and len(next_row) != prod_row_len:
			pkg_row = True
	
	return pkg_row
	
def determine_coll_name(vendor, prod_name_data, item_sku):

	#print("\n=== Determine Collection Name: " + item_sku + " ===\n")

	# deconstruct item sku to get coll sku
	
	coll_name = ''
	
	table_title = 'name'
	field_title = 'sku'
	skus = isolator.isolate_data_field(prod_name_data, table_title, field_title)
	field_title = 'coll name'
	coll_names = isolator.isolate_data_field(prod_name_data, table_title, field_title)
	field_title = 'handle'
	handles = isolator.isolate_data_field(prod_name_data, table_title, field_title)
	
	item_coll_sku = isolator.isolate_coll_sku(vendor, item_sku).strip()
	
	# first search for exact match sku to get coll name, 
	# and then check coll sku if no exact match
	for item_idx in range(len(prod_name_data)):
		current_item_sku = skus[item_idx]
		if item_sku == current_item_sku:
			#print("found matching sku in product name data")
			coll_name = coll_names[item_idx]
			if coll_name == 'n/a':
				# get coll name from handle
				handle = handles[item_idx]
				handle_data = handle.split('-')
				coll_name = handle_data[0]
	
	if coll_name == '': # still blank after searching for exact match
		for item_idx in range(len(prod_name_data)):
			current_item_sku = skus[item_idx]
			current_coll_sku = isolator.isolate_coll_sku(vendor, current_item_sku).strip()
		
			#print("item_coll_sku: " + item_coll_sku)
			#print("current_coll_sku: " + current_coll_sku)
			# remove parenthese bc affect pattern match
			item_coll_sku = re.sub("\(|\)","",item_coll_sku)
			current_coll_sku = re.sub("\(|\)","",current_coll_sku)
		
			if re.search(item_coll_sku, current_coll_sku):
				coll_name = coll_names[item_idx]
				if coll_name == 'n/a':
					# get coll name from handle
					handle = handles[item_idx]
					handle_data = handle.split('-')
					coll_name = handle_data[0]
	
	return coll_name
	
def determine_all_coll_names(vendor, all_item_skus):

	#print("\n=== Determine All Collection Names ===\n")

	coll_names = []

	# use vendor to get list of coll names
	input = 'name'
	extension = 'tsv'
	prod_name_data = reader.extract_vendor_data(vendor, input, extension)
	
	for item_sku in all_item_skus:
		coll_name = ''
		coll_name = determine_coll_name(vendor, prod_name_data, item_sku)
		
		coll_names.append(coll_name)

	return coll_names
	
def determine_vague_item_descrip(item_descrip):

	vague = True
	
	#vague_descrips = ["leg","base","top"]
	key_type = "description"
	descrip_keywords = reader.read_keywords(key_type)
	descrip_type = "defined"
	defined_descrips = descrip_keywords[descrip_type] #["chest","dresser","mirror","nightstand","entertainment unit"]
	#print("defined_descrips: " + str(defined_descrips))
	
	# nested keywords
	key_type = "title"
	title_keywords = reader.read_keywords(key_type)
	defined_keywords = []
	for defined_descrip in defined_descrips:
		item_keywords = title_keywords[defined_descrip]
		for item_key in item_keywords:
			defined_keywords.append(item_key)
	#print("defined_keywords: " + str(defined_keywords))
	
	# is it more accurate to check if item descrip contained in vague descrips or defined descrips?
	# if found item descrip in defined descrips we know it does not need product descrip added
	# if item descrip = vague descrip then we know it is vague but it must be exact match
	# default is vague so include more info unless we know it is defined already
	for defined_descrip_key in defined_keywords:
	
		if re.search(defined_descrip_key,item_descrip.lower()):
			vague = False
			break
	
	return vague
	
def determine_product_part(type):

	part = True
	
	key_type = "product"
	product_keywords = reader.read_keywords(key_type)
	keyword_type = "whole product types"
	whole_product_types = product_keywords[keyword_type] #["chest","dresser","mirror","nightstand","entertainment unit"]
	#print("whole_product_types: " + str(whole_product_types))
	
	# given that there are more parts than whole products, use shorter list of products
	for whole_product_type in whole_product_types:
		if type == whole_product_type:
			part = False
			break
	
	return part
	
def determine_composite_item(sku,catalog_title=''):

	composite = False
	
	if re.search("\+",sku):
		composite = True
	
	#print("catalog_title: " + catalog_title)
	if not catalog_title == '':
		if re.search("[2-9].*cn",catalog_title.lower()): # 2-9 bc finite num pieces in set (single digit)
			#print("found pack")
			composite = True
	
	return composite
	
# def determine_composite_item(type):
# 
# 	composite = False
# 	
# 	key_type = "product"
# 	product_keywords = reader.read_keywords(key_type)
# 	keyword_type = "composite product types"
# 	composite_product_types = product_keywords[keyword_type] #["sectionals","sets"]
# 	#print("composite_product_types: " + str(composite_product_types))
# 	
# 	# given that there are more parts than whole products, use shorter list of products
# 	for composite_product_type in composite_product_types:
# 		if re.search(composite_product_type,type):
# 			composite = True
# 			break
# 	
# 	return composite
	
def determine_valid_item(item, valid_ref_nums):
	valid = False
	
	for ref_num in valid_ref_nums:
	
		item_ref_num = item[0]
		
		if item_ref_num == ref_num:
			valid = True
	
	return valid 
	
def determine_published_status(img_src,product_type):

	published = 'FALSE'
	
	if img_src != '' and not determine_product_part(product_type):
		published = 'TRUE'
	
	return published
	
def determine_remove_opt_name(product_handle):

	#print("\n=== Determine Remove Option Name for " + product_handle + " ===\n")

	remove_opt_name = ''
	
	#all_opts = ['size','color','material','version']
	output = "option"
	all_keywords = reader.read_keywords(output)
	
	dashless_handle = re.sub('-',' ',product_handle)
	#print("dashless_handle: \"" + dashless_handle + "\"")
	
	# search all opt vals to see if in handle. then take opt name for opt val found
	# loop for each type of option, b/c need to fill in value for each possible option (eg loop for size and then loop for color in case item has both size and color options)
	for option_name, option_dict in all_keywords.items():
		#print("======Check for Option Name: " + option_name)
		#print("Option Dict: " + str(option_dict))

		final_opt_value = final_opt_name = ''

		for option_value, option_keywords in option_dict.items():
			#print("Option Value: " + option_value)
			#print("Option Keywords: " + str(option_keywords))

			#print("Plain SKU: " + dashless_sku)
			compare_option_value_regex = "\s" + option_value.lower() + "(\s|$)" # ensure exact match by requiring match to have space before it so "tan" is not found in "nightstand"
			#print("compare_option_value_regex: \"" + compare_option_value_regex + "\"")
			if re.search(compare_option_value_regex,dashless_handle):
				#print("found " + str(compare_option_value_regex) + " in " + dashless_handle)
				#print("found option with name " + option_name + " and value " + option_value + " in handle")
				remove_opt_name = option_name

				break
			#else:
				#print("did not find " + compare_option_value_regex + " in " + dashless_handle)
				
		if remove_opt_name != '':
			#print("remove_opt_name: " + remove_opt_name)
			#print("Final Option Value: " + final_opt_value)
			break
	# for opt in all_opts:
# 		if re.search(opt,product_handle):
# 			remove_opt_name = opt
# 			break
			
	#print("remove_opt_name: " + remove_opt_name)
	
	return remove_opt_name
	
# ====== Principle Generator ======	

def determine_principle_class(principle_num):

	prin_cl = ''

	num_dots = principle_num.count('.')
	#print('num_dots: ' + str(num_dots))
	
	if num_dots == 1:
		prin_cl = 'gob-preview-principle'
	elif num_dots == 2:
		prin_cl = 'gob-sub-principle'
	elif num_dots == 3:
		prin_cl = 'gob-sub-sub-principle'

	return prin_cl
	
def determine_demo_principle(demo_principles, principle_id):

	print("\n=== Determine Demo Principle: " + principle_id + " ===\n")

	demo = False

	# principle id like gob-p1-1-1
	# demo princ like 1-1.1.
	
	princ_num_id = re.sub('gob-p','',principle_id)
	print("princ_num_id: " + princ_num_id + "\n")
	
	for demo_princ in demo_principles:
	
		#print("demo_princ: " + demo_princ)
	
		demo_princ_id = re.sub('\.','-',demo_princ).rstrip('-') # 1-1-1
		print("demo_princ_id: " + demo_princ_id)
		
		if demo_princ_id == princ_num_id:
			print("Principle " + principle_id + " is a demo principle!")
			demo = True
	
	return demo
	
# for a given line of data copied from aw or gb, determine if it is a principle by its format
def determine_principle(raw_data_line):

	line_principle = False

	# if starting with \d+\. and not ending with \w then it is principle
	if re.search('^\d+\.',raw_data_line) and not re.search('\w$',raw_data_line): 
		line_principle = True
	# if starting with \d+\. and ending with 'and','or','nor' (used in lists for second to last list item) then it is principle
	elif re.search('^\d+\.',raw_data_line) and re.search('(and|n?or)$',raw_data_line): 
		line_principle = True
		
	return line_principle
	
def determine_sub_sub_principle(princ):

	#print("\n=== Determine Sub-Sub-Principle: " + princ + " ===\n")

	sub_sub_princ = False
	
	# if it has 3 dots it is a sub-sub-principle so it should be kept with previous line unless sub-princ extends >1 page
	if princ.count('.') == 3:
		sub_sub_princ = True
	
	return sub_sub_princ
	
def determine_sub_principle(princ):

	#print("\n=== Determine Sub-Principle: " + princ + " ===\n")

	sub_princ = False
	
	# if it has 3 dots it is a sub-sub-principle so it should be kept with previous line unless sub-princ extends >1 page
	if princ.count('.') == 2:
		sub_princ = True
	
	return sub_princ
	
# if first principle in set contains the phrase "the following", 
# then keep the following sub-principles in the same sub-princ-set (even though not sub-sub-principle)
def determine_strict_princ_set(princ, ch_princ_data):
	
	#print("\n=== Determine Strict Principle Set: " + princ + " ===\n")
	
	ch_princ_nums = []
	for princ_data in ch_princ_data:
		ch_princ_nums.append(princ_data[0])
	#print("ch_princ_nums: " + str(ch_princ_nums))
	ch_princ_contents = []
	for princ_data in ch_princ_data:
		ch_princ_contents.append(princ_data[1])
	#print("ch_princ_contents: " + str(ch_princ_contents))
	
	strict_set = False
	
	# if sub-princ, then check if related princ has words "the following"
	if determine_sub_principle(princ):
		#print("sub-princ found: " + princ)
		related_princ_num = princ.split('.')[0]
		related_princ_id = related_princ_num + "."
		#print("related_princ_id: " + related_princ_id)
		related_princ_idx = ch_princ_nums.index(related_princ_id)
		#print("related_princ_idx: " + str(related_princ_idx))
		related_princ_content = ch_princ_contents[related_princ_idx]
		#print("related_princ_content: " + str(related_princ_content))
		
		key_string = 'the following'
		if re.search(key_string, related_princ_content.lower()):
			#print("found key phrase \'the following\' in related principle content: \'" + related_princ_content + "\'")
			strict_set = True
	
	return strict_set

def determine_has_keywords(content,keywords):

	found_key = False
	
	for key in keywords:
		if re.search(key,content):
			found_key = True
	
	return found_key
	
def determine_keyword(content,keywords):

	keyword = False
	
	for key in keywords:
		if key == content:
			keyword = True
	
	return keyword