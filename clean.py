import os

def clean():
    for file in os.listdir():
        if ".txt" in file or ".dat" in file:
            os.remove(file)

clean()