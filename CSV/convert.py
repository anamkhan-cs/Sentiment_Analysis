import pyexcel
import re

for file in re.findall('*.xls',file):
	 pyexcel.transcode("file", "filex")