import os
import sys

try:
    file_paths = open(sys.argv[1], 'r').read().split("\n")
    for file_path in file_paths:
        if(os.path.exists(file_path)):
            os.remove(file_path)
    print("success")

except Exception as e:
    print(e)

