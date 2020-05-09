import os
import os.path

fileType = ('xlsx', 'xls')
ls = []

def getfile_fix(filename) -> str:
    return filename[filename.rfind('.')+1:]


def getfile(filepath) -> str:
    return filepath[filepath.rfind('\\')+1:]


def getAppointFile(path, ls, traverse=True) -> None:
    fileList = os.listdir(path)
    try:
        for tmp in fileList:
            pathTmp = os.path.join(path, tmp)
            if traverse and os.path.isdir(pathTmp) == True:
                getAppointFile(pathTmp, ls)
            elif getfile_fix(pathTmp).lower() in fileType:
                ls.append(pathTmp)
    except PermissionError:
        pass


def getAppointFileList(path: str, _fileType=('xlsx', 'xls'), traverse=True) -> list:
    fileType = _fileType
    getAppointFile(path, ls, traverse)
    return ls


def main():

    while True:
        path = input('input a pathname: ').strip()
        # Return true if the pathname refers to an existing directory
        if os.path.isdir(path) == True:
            break
        print("The pathname doesn't refer to an existing directory")

    getAppointFile(path, ls)
    for fp in ls:
        print(getfile(fp))
    print(len(ls))
    print(getAppointFileList(path))


if __name__ == '__main__':
    main()