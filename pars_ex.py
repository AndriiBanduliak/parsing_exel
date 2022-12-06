import xlrd3 as xlrd


FILE = "Test.xls"


def parsing(file):
    book = xlrd.open_workbook(file)
    sh = book.sheet_by_index(0)
    print(sh)
    for row_number in range(sh.nrows):
        for col_number in range(sh.ncols):
            print(sh.cell_value(rowx=row_number, colx=col_number))


def do(file):
    art = []
    exp = []
    book = xlrd.open_workbook(file)
    sh = book.sheet_by_index(0)
    for row_number in range(sh.nrows):
        row = sh.row_values(row_number)
        if row[1]:
            if row[0]!="Итого:" and isinstance(row[1], float):
                art.append(row[0])
                exp.append(row[1])
    print_result(art,exp)

def print_result(art, exp):
    index_min = get_extr_key(exp, min)
    index_max = get_extr_key(exp, max)
    print('минимальный расход:')
    print("Статья:", art[index_min])
    print("Сумма:", exp[index_min])
    print('---------------------------')
    print('максимальный расход')
    print("Статья:", art[index_max])
    print("Сумма:", exp[index_max])
    
    

def get_extr_key(arry, compare):
    extr_index = 0
    extreme = arry[extr_index]
    i = 1
    while i < len(arry):
        if compare(arry[i], extreme) == arry[i]:
            extreme = arry[i]
            extr_index = i
        i +=1
    return extr_index
    
   
                

if __name__ == "__main__":
    # parsing(FILE)
    do(FILE)
