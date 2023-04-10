import re

def order_table(table,xy_order_str=''):
    if len(xy_order_str)==0:
        print(table)
    else:
        list1 = []
        try:
            for i in list(table['yx']):
                list1.append(re.search(i, xy_order_str).span(0)[0])
        except:
            print('error!')
        table['order'] = list1
        table.sort_values(by='order',inplace=True)
        table.drop('order',axis=1,inplace=True)
    return table

