import openpyxl
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory
import os, shutil

totalTaskCount = 0
tasksDoneCount = 0
def taskDone():
    global tasksDoneCount
    tasksDoneCount = tasksDoneCount + 1
    print("progress: " + str(
        round(tasksDoneCount / totalTaskCount * 100)
        ).rjust(3) + "%", end = "\r")

if __name__ == "__main__":
    Tk().withdraw()

    source_dir_path = askdirectory(title="Select parent folder containing 1月/ 2月/ ...")
    workbook = openpyxl.load_workbook(
        askopenfilename(title="Select sheet containing row IDs to delete"),
        read_only=True
    )
    dest_dir_path = askdirectory(title="Select an EMPTY folder to download to")

    source_dir = {}
    for entry in os.listdir(source_dir_path):
        if (entry.endswith("月")):
            source_dir[entry[:-1]] = {}
            for xlsm in os.listdir(os.path.join(source_dir_path, entry, "全网络汇总")):
                if xlsm.startswith("SA") and xlsm.endswith(".xlsm"):
                    source_dir[entry[:-1]][xlsm.split("_")[0]] = xlsm
    
    workbook_data = workbook.active.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True)

    to_download = {}

    for row in workbook_data:
        month_string = "{dt.month}".format(dt=row[1])
        if month_string in source_dir and row[0] in source_dir[month_string]:
            if month_string not in to_download:
                to_download[month_string] = set()
            to_download[month_string].add(source_dir[month_string][row[0]])
    
    totalTaskCount = sum(len(files) for files in to_download.values())

    for month_string, files in to_download.items():
        os.mkdir(os.path.join(dest_dir_path, month_string))
        for file_name in files:
            shutil.copy(
                os.path.join(source_dir_path, month_string + "月", "全网络汇总", file_name),
                os.path.join(dest_dir_path, month_string)
                )
            taskDone()

    workbook.close()

    

