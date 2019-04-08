import glob

folders = glob.glob("files/*")

for folder in folders:
    print(folder)