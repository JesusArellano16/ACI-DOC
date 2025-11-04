import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) 
SRC_DIR = os.path.join(BASE_DIR, "src")
RESULTS_DIR = os.path.join(BASE_DIR, "results")
TEMPLATE_PATH = os.path.join(SRC_DIR, "Template.xlsx")



def rename_txt():
    for filename in os.listdir(SRC_DIR):
        if filename.lower().endswith(".txt"):
            old_path = os.path.join(SRC_DIR, filename)

            base_name = os.path.splitext(filename)[0][:17]
            new_name = f"{base_name}.txt"
            new_path = os.path.join(SRC_DIR, new_name)

            if old_path == new_path:
                continue

            counter = 1
            while os.path.exists(new_path):
                new_name = f"{base_name}_{counter}.txt"
                new_path = os.path.join(SRC_DIR, new_name)
                counter += 1

            os.rename(old_path, new_path)