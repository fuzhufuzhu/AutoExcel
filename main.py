
import xlwings as xw
import os
import argparse


def get_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument('-r',type=str,required=True,help='Enter the report path to process')
    parser.add_argument('-output',type=str,help='Enter the path to output the results')
    return parser

def help():
    print("")

#用于获得文件夹下的所有文件名
def get_file_list(dir):
    print("进行导入的文件列表如下")
    fileList=[]
    i=0
    for home, dirs, files in os.walk(dir):
        for filename in files:
            fullname = os.path.join(home, filename)
            fileList.insert(i,fullname)
            i+=1
            print(fullname)
    print("导入文件数量："+str(len(fileList)))
    return fileList

def creatExcel(excelPath):
    app = xw.App(visible=False, add_book=False)
    wb = app.books.add()
    wb.sheets.add('周报')
    wb.save(excelPath)
    print("创建excel表格成功....................")
    path = wb.fullname
    wb.app.quit()
    print("创建新excel表格路径为"+path)
    return path

def test(path):
    app = xw.App(visible=False, add_book=True)
    wb = app.books.open(path)
    excel = wb.sheets["周报"]
    rows = excel.used_range.last_cell.row
    print(excel.range('A2').expand().value)



def readExcel(path,List):
    # 创建总周报app对象
    sum_app = xw.App(visible=False, add_book=True)
    sum_wb = sum_app.books.open(path)
    sum_excel = sum_wb.sheets["周报"]
    num = 0

    #创建成员周报app对象
    for i in range(len(List)):
        app = xw.App(visible=False, add_book=True)
        print("开始写入"+List[i]+"-------------------------")
        wb = app.books.open(List[i])
        excel = wb.sheets["周报"]
        rows = excel.used_range.last_cell.row

        for x in range(rows):
            if(i==0):
                print("[+]当前写入" + str(excel.range('A'+str(x+1)).expand().value))
                sum_excel.range('A'+str(num+1+x)).value = excel.range('A'+str(x+1)).expand().value
                continue
            else:
                if(x==0):
                    continue
                print("[+]当前写入" + str(excel.range('A' + str(x + 1)).expand().value))
                sum_excel.range('A' + str(num + x)).value = excel.range('A' + str(x + 1)).expand().value
        num = rows-1+num
        wb.save()
        wb.app.quit()
    sum_wb.save()
    sum_wb.app.quit()

if __name__ == '__main__':
    parser = get_parser()
    args=parser.parse_args()
    name = args.r
    excelPath = args.output
    print(excelPath)
    List =get_file_list(name)
    path =creatExcel(excelPath)
    readExcel(str(path),List)




