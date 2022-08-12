import os

def clean():
    try:
        for file in os.listdir():
            if ".txt" in file or ".dat" in file:
                os.remove(file)
        return True
    except:
        print("中间数据清扫失败，程序中断")
        return False

if __name__ == "__main__":
    clean()