# sorter.py
# for every time you need to sort something

import isolator, generator, converter, reader, determiner, writer
import numpy as np

# reverse order of json dict
def reverse_json():
	in_json = "\"King\":\"\",\"Queen\":\"\""
	
variants = [['gabe-bed','Size','Queen'],['gabe-bed','Size','King'],['gabe-bed','Size','Full']]

# choose to sort variants by size, by cross-referencing option values with standards we set in data>standards folder
def sort_variants_by_size(variants):
	sorted_variants = []

	opt_name = 'Size'
	sizes = reader.read_standards("sizes")[opt_name]
	#print(sizes)
	# get size opt val for variant 
	# what idx is size opt name? look for size keyword so it could be at any idx
	
	for variant in variants:
		size_opt_idx = determiner.determine_opt_idx(opt_name,variant)
		#print("size_opt_idx: " + str(size_opt_idx))
	
		variant_size = variant[size_opt_idx]
		#print("variant_size: " + variant_size)
		
		sizes[variant_size] = variant
		
	#print(sizes)
	for variant in sizes.values():
		if variant != '':
			#print(variant)
			sorted_variants.append(variant)
	
	return sorted_variants
	
#sorted_variants = sort_variants_by_size(variants)

def sort_variants(variants):
	#if input is string 
	#variants = converter.convert_variant_strings_to_data(variant_strings)

	# first sort by size, then we can sort by other options
	sorted_variants = sort_variants_by_size(variants)
	#print("sorted_variants: " + str(sorted_variants))
	
	sorted_variant_strings = converter.convert_variant_data_to_strings(sorted_variants)
	#print("sorted_variant_strings: " + str(sorted_variant_strings))
	
	return sorted_variant_strings
	
#sorted_variant_strings = sort_variants(variants)
#writer.display_list(sorted_variant_strings)

def sort_all_variants(all_item_info):
	all_sorted_variants = []
	# isolate products in unsorted import table
	import_type = "shopify"
	isolated_product_imports = generator.isolate_product_strings(all_item_info, import_type)
	#print("Unsorted Product Imports: " + str(isolated_product_imports) + "\n")
	
	for product in isolated_product_imports:
		variants = converter.convert_variant_strings_to_data(product)
		sorted_variants = sort_variants(variants)
		for sorted_variant in sorted_variants:
			all_sorted_variants.append(sorted_variant)
		
	return all_sorted_variants
	
#invalid format all_item_info = [['gabe-bed','Size','Queen'],['gabe-bed','Size','King'],['gabe-bed','Size','Full'],['gino-bed','Size','Queen'],['gino-bed','Size','King'],['gino-bed','Size','Full']]

#all_sorted_variant_strings = sort_all_variants(all_item_info)
#writer.display_list(all_sorted_variant_strings)

def sort_items_by_handle(unsorted_handles, all_details):
	all_sorted_items = all_details
	
	handles_array = np.array(unsorted_handles)
	sorted_indices = np.argsort(handles_array)
	
	for item_idx in range(len(all_details)):
		sorted_idx = sorted_indices[item_idx]
		sorted_item = all_details[sorted_idx]
		all_sorted_items.append(sorted_item)
	
	return all_sorted_items

# sort from small to large so price user sees in grid view is first price shown on product page
def sort_items_by_size(all_item_info, import_type, all_details):
	#print("\n=== Sort Items by Size ===\n")
	all_sorted_items = all_item_info # init list of strings in shopify import format

	# isolate products in unsorted import table
	isolated_product_imports = generator.isolate_product_strings(all_item_info, import_type)
	#print("Unsorted Product Imports: " + str(isolated_product_imports) + "\n")
	num_product_imports = len(isolated_product_imports)
	#print("Num Unsorted Product Imports: " + str(num_product_imports) + "\n")

	isolated_product_details = generator.isolate_products(all_details)
	#print("Unsorted Product Details: " + str(isolated_product_details) + "\n")
	num_product_details = len(isolated_product_details)
	#print("Num Unsorted Product Details: " + str(num_product_details) + "\n")

	if num_product_imports == num_product_details:
		all_sorted_products = []
		all_sorted_items = []

		num_isolated_products = len(isolated_product_details)
		#print("Num Isolated Products: " + str(num_isolated_products))

		for product_idx in range(num_isolated_products):
			product_details = isolated_product_details[product_idx] # list of data
			product_imports = isolated_product_imports[product_idx] # string of data

			sorted_indices = generator.get_sorted_indices(product_details) # reqs valid dims
			num_sorted_indices = sorted_indices.size
			#print("Num Sorted Indices = Num Widths: " + str(num_sorted_indices))

			num_variants = len(product_details)
			#print("Num Variants: " + str(num_variants))

			sorted_variants = product_details # init as unsorted variants
			# only sort variants if we have valid values for sorting
			if num_variants == num_sorted_indices:
				sorted_variants = [] # How can we be sure there are valid dimensions given?
				for idx in range(num_variants):
					#print("Index: " + str(idx))
					sorted_idx = sorted_indices[idx]
					#print("Sorted Index: " + str(sorted_idx))
					sorted_variant = product_imports[sorted_idx]
					sorted_variants.append(sorted_variant)
			else:
				print("Warning: Num Variants != Num Sorted Indices while sorting items!")

			all_sorted_products.append(sorted_variants)

		all_sorted_items = isolator.isolate_unique_variants(all_sorted_products, import_type)

	else:
		print("Warning: Num Product Imports != Num Product Details (" + str(num_product_imports) + " != " + str(num_product_details) + ") while sorting items!\n")

	#print("\n=== Sorted Items by Size ===\n")

	return all_sorted_items
	

