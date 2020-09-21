import requests
from bs4 import BeautifulSoup
import os

totalTaskCount = 0
tasksDoneCount = 0
def taskDone():
    global tasksDoneCount
    tasksDoneCount = tasksDoneCount + 1
    print("progress: " + str(
        round(tasksDoneCount / totalTaskCount * 100)
        ).rjust(3) + "%", end = "\r")

# returns strings
# could return None
# could return "Multiple"
# 是上市公司 iff QCC报告有A股代码。利用该A股代码查中财网。
def getA股代码(qccSearchParam):
	IPO_URL = "https://ipo.qichacha.com/search?key="

	soup = BeautifulSoup(requests.get(IPO_URL + qccSearchParam).content, "html.parser")
	container_found = soup.find("tbody")

	taskDone()

	if (container_found is None):
		return

	# only iterates through first page of search results, that is, at most 10 companies
	# not an issue as we only want 1 anyway
	companyFound = False
	for tr in container_found.children:
		if tr.name == "tr":
			if companyFound:
				return "Multiple"
			companyFound = True
			A股代码 = tr.contents[3].a.span.string
	return A股代码

# splits a file containing a company name per line into 3:
# Firstly, a file containing the A股代码 of all companies that have them
# Secondly, a file containing companies whose name, queried on QCC, gave multiple results
# Thirdly, a file containing companies that cannot be found via QCC
def split_name_list(full_names_file, output_dir):
	with (open(full_names_file, "r")) as f:
		full_names = f.read().splitlines()
	
	full_names = [x for x in full_names if x != ""]

	global totalTaskCount
	totalTaskCount = len(full_names)

	with open(os.path.join(
	output_dir, "A股上市公司.txt"), "w") as single_file, open(os.path.join(
	output_dir, "未找到.txt"), "w") as none_file, open(os.path.join(
	output_dir, "找到多个qcc结果.txt"), "w") as multiple_file:
		for full_name in full_names:
			query_result = getA股代码(full_name)
			if (query_result is None):
				none_file.write(full_name + "\n")
			elif (query_result == "Multiple"):
				multiple_file.write(full_name + "\n")
			else:
				single_file.write(full_name + "\\" + query_result + "\n")

if __name__ == "__main__":
	from tkinter import Tk
	from tkinter.filedialog import askopenfilename, askdirectory
	Tk().withdraw()

	print("请选择要自动分成三类（企查查上市公司，企查查找到多个结果，企查查找不到）的公司列表")
	full_names_file = askopenfilename()
	print("请选择保存位置")
	outputDir = askdirectory()
	split_name_list(full_names_file, outputDir)
