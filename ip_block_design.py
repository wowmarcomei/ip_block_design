from IPy import IP
import xlsxwriter

max_mask = 18
min_mask = 30
original_ip_block = IP('10.75.96.0')

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('IP_blocks.xlsx')
worksheet = workbook.add_worksheet()
# Create a format to use in the merged range.
cell_format = workbook.add_format({
    'font_name': 'Times New Roman',
    'font_size': 11,
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'gray'
    })
head_format = workbook.add_format({
    'font_name': 'Times New Roman',
    'font_size': 14,
    'bold': 2,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#538dd5',
    'color': '#ffffff'
    })

def initial_ipblocks(max_mask=max_mask,min_mask=min_mask,original_ip_block = original_ip_block,column = 1):
    ip_numbers = 2**(32-max_mask)
    unit = (2**(32-min_mask+column-1))
    column_numbers = min_mask - max_mask +1
    print('总共IP数：',ip_numbers,'\n每次递增数:',unit)
    # 将原始IP地址段转换为十进制，便于计算，在后续的函数中计算完成后，可以再次转换为IP对象
    convert_ip_to_dec = original_ip_block.strDec()

    # 根据最大掩码数与最小掩码数计算出写入哪一列，IP地址将会被写在excel中的对应列
    # 创建一个包含元祖对象的列表tuple1
    tuple1 = []
    for i in range(min_mask-max_mask+1):
        # append()一个元祖过来，元祖对象为括号包含，所以格式为append((x,y))
        tuple1.append(('{}'.format(min_mask-max_mask-i+1),chr(i+ord('A'))))
    # 将元祖转换为字典，后续在字典中查找Key为1，查找出字典中对应的value为：dict['1'],如果该值等于'K',则IP地址写入Excel的K列，即倒数第1列
    excel_dict = {k: v for k, v in tuple1}
    
    # 返回列表：
    # 1.ip数
    # 2.每次间隔
    # 3.转换后的十进制IP
    # 4.需要写入的列数
    # 5.Excel字典，如excel_dict={'7': 'A', '6': 'B', '5': 'C', '4': 'D', '3': 'E', '2': 'F', '1': 'G'}，excel_dict['1']则会在第G列写入
    return [ip_numbers,unit,convert_ip_to_dec,column_numbers,excel_dict]

def write_to_first_row(min_mask=min_mask,max_mask=max_mask):
    # 根据最大掩码数与最小掩码数计算出写入哪一列，IP地址将会被写在excel中的对应列
    # 创建一个包含元祖对象的列表tuple1
    tuple1 = []
    for i in range(min_mask-max_mask+1):
        # append()一个元祖过来，元祖对象为括号包含，所以格式为append((x,y)), chr(i+ord('A'))表示的是字母表A,B,C,D,E,F...
        tuple1.append(  (chr(i+ord('A')),i+max_mask)  )
    # 将包含元祖的列表转换为字典
    excel_dict = {k: v for k, v in tuple1}

    for key, value in excel_dict.items():
        # print(excel_dict)
        worksheet.write('{}1'.format(key),'/{}'.format(value),head_format)


def write_to_column(column):
    #软参column表示第几列
    caculate_unit = initial_ipblocks(max_mask=max_mask,min_mask=min_mask,original_ip_block = IP('10.75.96.0'),column = column)
    total_value = caculate_unit[0]
    add_value = caculate_unit[1]
    dec_value = caculate_unit[2]
    unit = 2**(column-1)

    # 考虑到A1:A1自己不能合并，所以最后一列单独调用函数write,而不是merge
    if column == 1:
        for time in range(int(total_value/add_value)):
            # 十进制 --> 字符串 --> IP对象 --> 字符串
            ips = IP(str(int(dec_value)+time*add_value)).strNormal(0)
            # print('{}'.format(2+unit*time))
            worksheet.write('{}{}'.format(caculate_unit[4][str(column)],2+unit*time),ips,cell_format)
            # 设置第column列的宽度为16
            worksheet.set_column('{}:{}'.format(caculate_unit[4][str(column)],caculate_unit[4][str(column)]),12)
    else:
        for time in range(int(total_value/add_value)):
            # 十进制 --> 字符串 --> IP对象 --> 字符串
            ips = IP(str(int(dec_value)+time*add_value)).strNormal(0)
            print('第{}次:{}{}:{}{}****IP-{}'.format(time+1,caculate_unit[4][str(column)],1+unit*time,caculate_unit[4][str(column)],unit*(time+1),ips))
            # 写进Excel文件，以合并单元格的方式写数据，如merge_range('A1:A2','data',format)
            # worksheet.merge_range('{}{}:{}{}'.format(caculate_unit[4][str(column)],1+unit*time,caculate_unit[4][str(column)],unit*(time+1)),ips,merge_format) #从第一行写起
            worksheet.merge_range('{}{}:{}{}'.format(caculate_unit[4][str(column)],2+unit*time,caculate_unit[4][str(column)],1+unit*(time+1)),ips,cell_format) #从第二行写起
            # 设置第column列的宽度为16
            worksheet.set_column('{}:{}'.format(caculate_unit[4][str(column)],caculate_unit[4][str(column)]),12)

# def save_ip_address(content):
#     with open('ip.txt', 'a', encoding='utf-8') as f:
#         f.write( content + '\n')

def main():
    # 获取到需要写入几列
    column_numbers = initial_ipblocks()[3]
    write_to_first_row()
    for i in range(1,column_numbers+1):
        write_to_column(i)

if __name__ == '__main__':
    main()
    workbook.close()

