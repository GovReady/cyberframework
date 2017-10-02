# Parse the NIST Cybersecurity Framework Excel file
# and convert it to YAML.
#
# Install the following modules first:
#
# pip3 install openpyxl rtyaml
#
# The sheet is organized by row, with cells merged vertically
# to indicate that its label applies as a family to the cells
# to its right. In openpyxl, merged cells have their value in
# the top row and None's in subsequent rows.

import re
import tempfile
import urllib
from collections import OrderedDict

import openpyxl
import rtyaml

# Download the file. Since openpyxl requires a filename
# and one that ends in .xlsx, save it to a temporary file.
with tempfile.NamedTemporaryFile(suffix=".xlsx") as f:
	#urllib.request.urlretrieve("https://www.nist.gov/document-3764", f.name)
	fn = "/tmp/framework-for-improving-critical-infrastructure-cybersecurity-core.xlsx"
	xlsx = openpyxl.load_workbook(fn) # f.name

# Read the rows.
root = []
stack = [root]
for i, row in enumerate(xlsx.worksheets[0].rows):
	if i == 0: continue # skip header row
	for col in range(0, 3):
		# If the cell's value is not None, then we have a new
		# Function, Category, or Sub-Category. Otherwise we
		# have a new row in Informative References that falls
		# under the last seen Function/Category/Sub-Category.
		if row[col].value != None:
			# Pop the stack to the right level.
			while len(stack) > col+1: stack.pop(-1)

			# Parse this entry.
			m = re.match(r"(?P<name>[^():]+)(?: \((?P<id>[^)]+?)\))?(?:: (?P<descr>.*))?$", row[col].value)

			# Construct a new dict for it.
			val = OrderedDict()
			if m.group("id"):
				val["id"] = m.group("id")
				val["name"] = m.group("name")
			else:
				val["id"] = m.group("name")
			val["type"] = ("function", "category", "subcategory")[col]
			if m.group("descr"):
				val["description"] = m.group("descr")

			# Append it into its parent.
			stack[-1].append(val)

			# Insert a new entry into the stack.
			sublist = []
			subattr = ("categories", "subcategories", "references")[col]
			val[subattr] = sublist
			stack.append(sublist)

	# Add the Informative References in this row.
	val = row[3].value
	val = val.replace("NIST SP 800-53 Rev.4", "NIST SP 800-53 Rev. 4") # data error
	m = re.match(r"·\s+(CCS CSC|COBIT 5|ISA 62443-2-1:2009|ISA 62443-3-3:2013|ISA 62443-2-1|ISO/IEC 27001:2013|NIST SP 800-53 Rev. 4),? (.*)", val)
	standard, controls = m.groups()
	controls = [c.strip() for c in controls.split(", ")]
	stack[-1].append(OrderedDict([("standard", standard), ("controls", controls)]))

print(rtyaml.dump(root))
