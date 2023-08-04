# converter.py
# for every time you need to convert something

def convert_variant_strings_to_data(variant_strings):
	all_variant_data = []
	for variant_string in variant_strings:
		variant_data = variant_string.split(';')
		all_variant_data.append(variant_data)
	
	return all_variant_data
	
def convert_variant_data_to_strings(all_variant_data):
	variant_strings = []
	
	for variant_data in all_variant_data:
		variant_string = ''
		for datum_idx in range(len(variant_data)):
			variant_datum = variant_data[datum_idx]
			if datum_idx == 0:
				variant_string += variant_datum 
			else:
				variant_string += ';' + variant_datum 
		
		variant_strings.append(variant_string)
	
	return variant_strings
	
def convert_prod_opt_string_to_data(opt_string):

	#print("\n=== Convert Product Option String to Data ===\n")
	#print("opt_string: " + opt_string)

	# convert prod opts string to data
	opt_data = opt_string.split(',')
	
	final_opt_names = []
	final_opt_values = []
	
	for opt_datum_idx in range(len(opt_data)):
		opt_datum = opt_data[opt_datum_idx]
		if opt_datum_idx % 2 != 0: # odd number
			final_opt_values.append(opt_datum)
		else: # even number
			final_opt_names.append(opt_datum)
	
	#print("final_opt_names: " + str(final_opt_names))
	#print("final_opt_values: " + str(final_opt_values))
	
	final_opt_data = [final_opt_names,final_opt_values]
	#print("final_opt_data: " + str(final_opt_data))
	
	#print("\n=== Converted Product Option String to Data ===\n")
	
	return final_opt_data