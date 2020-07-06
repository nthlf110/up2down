import os
import sys
import xlwt
import xlrd
import optparse
import datetime


def read_from_xlsx(xls_file_name, sheet_name_in_xls, header="T"):
    data_info = xlrd.open_workbook(xls_file_name)
    try:
        data_sh = data_info.sheet_by_name(sheet_name_in_xls)
    except:
        print("no sheet in %s named data" % xls_file_name)
        os._exit()

    data_nrows = data_sh.nrows
    data_ncols = data_sh.ncols

    result_data = []
    if header == "T":
        for i in range(1, data_nrows):
            result_data.append(dict(zip(data_sh.row_values(0), data_sh.row_values(i))))
    elif header == "F":
        for i in range(data_nrows):
            result_data.append(data_sh.row_values(i))
    else:
        print("header = ", header)
        print("The parameter 'header' is undefined, please check!")
        os._exit()

    return result_data


def read_from_xlsx_rich(xls_file_name, sheet_name_in_xls, header="T"):
    data_info = xlrd.open_workbook(xls_file_name, formatting_info=True)
    try:
        data_sh = data_info.sheet_by_name(sheet_name_in_xls)
    except:
        print("no sheet in %s named data" % xls_file_name)
        os._exit()

    data_nrows = data_sh.nrows
    data_ncols = data_sh.ncols

    result_data = []
    if header == "T":
        for i in range(1, data_nrows):
            temp = data_sh.row_values(i)
            for j in range(0, data_ncols):
                temp[j] = cell_real_value(data_sh, i, j)
            result_data.append(dict(zip(data_sh.row_values(0), temp)))
    elif header == "F":
        for i in range(data_nrows):
            temp = data_sh.row_values(i)
            for j in range(0, data_nrows):
                temp = cell_real_value(data_sh, i, j)
            result_data.append(temp)
    else:
        print("header = ", header)
        print("The parameter 'header' is undefined, please check!")
        os._exit()

    return result_data


def cell_real_value(sh, row, col):
    for merged in sh.merged_cells:
        if (merged[0] <= row < merged[1]
                and merged[2] <= col < merged[3]):
            return sh.cell_value(merged[0], merged[2])
    return sh.cell_value(row, col)


def info_to_xlsx(list_head_names, list_info, output_file_name, sheet_name):
    if list_head_names == '':
        print("Please set the list of head names of the output .xls file.")
        print("If you don't set the names, default value is empty.")
    if list_info == []:
        print("The content to be filled is empty, ")
        print("please check the input parameters of function info_to_xlsx")
        return
        # os._exit()
    if output_file_name == '':
        print("Please assign the names of the output xls file's name.")
        os._exit()

    wb = xlwt.Workbook(encoding='utf-8')
    # wb = xlrd.open_workbook(output_file_name)
    ws = wb.add_sheet(sheet_name)
    if isinstance(list_info[0], list):
        if list_head_names != '':
            for i in range(len(list_head_names)):
                ws.write(0, i, list_head_names[i])
            for i in range(len(list_info)):
                for j in range(len(list_info[i])):
                    ws.write((i + 1), j, (list_info[i])[j])
        else:
            for i in range(len(list_info)):
                for j in range(len(list_info[i])):
                    ws.write(i, j, (list_info[i])[j])
    elif isinstance(list_info[0], str):
        if list_head_names != '':
            for i in range(len(list_head_names)):
                ws.write(0, i, list_head_names[i])
            for i in range(len(list_info)):
                ws.write(1, i, list_info[i])
        else:
            for i in range(len(list_info)):
                ws.write(0, i, list_info[i])
    # output_name = output_file_name + ".xls"
    wb.save(output_file_name)


def query_ts_id(ts_id, summary):
    for i in iter(summary):
        if i['检测编号'].__contains__(ts_id):
            return i['检测编号']
    return ts_id + 'Not Found'


def query(ts_id, summary):
    for i in iter(summary):
        if i['检测编号'] == ts_id:
            return i['收样日期'], i['销售'], i['癌种'], i['其他']
    return '', '', '', ''


def query_sample_type(ts_id, summary):
    tissue = []
    for i in iter(summary):
        if i['检测编号'] == ts_id:
            tissue.append(i['组织类型'])
    if_blood = sum([i.__contains__('血') for i in tissue])
    if_tissue = sum([i.__contains__('组织') for i in tissue]) + \
                sum([i.__contains__('蜡') for i in tissue]) + \
                sum([i.__contains__('切片') for i in tissue])
    if if_blood:
        if if_tissue:
            return 'T+B'
        else:
            return 'B'
    else:
        if if_tissue:
            return 'T'
        else:
            return 'UoN'


def if_pdl1(ts_id, summary):
    panel = []
    for i in iter(summary):
        if i['检测编号'] == ts_id:
            panel.append(i['检测项目'])
    if sum([i.__contains__('PD-L1') for i in panel]) != 0:
        return True
    else:
        return False


def query_report(sales_name, dictionary):
    for i in iter(dictionary):
        if i['姓名'] == sales_name:
            return i['负责人']
    return 'N/A'


def get_date(date):
    if date == '':
        return ''
    else:
        delta = datetime.timedelta(days=date)
        output_date = datetime.datetime.strptime('1899-12-30', '%Y-%m-%d') + delta
        return datetime.datetime.strftime(output_date, '%Y/%m/%d')


if __name__ == '__main__':
    parser = optparse.OptionParser()
    parser.add_option('-i', '--input', dest='input')
    parser.add_option('-o', '--output', dest='output',
                      default=datetime.datetime.strftime(datetime.date.today(), '%Y%m%d') + 'I--明码_raw.xlsx')
    parser.add_option('-s', '--summary', dest='summary', default='../外送统计表/桐树NGS外送样本统计表.xlsx')
    parser.add_option('-r', '--report', dest='report', default='../外送统计表/报告分配.xlsx')
    (options, args) = parser.parse_args()
    input_xlsx = read_from_xlsx_rich(options.input, 'Sheet1', header='T')
    summary_xlsx = read_from_xlsx(options.summary, '北东区', header='T')
    summary_xlsx.extend(read_from_xlsx(options.summary, '中区', header='T'))
    summary_xlsx.extend(read_from_xlsx(options.summary, '南区', header='T'))
    summary_xlsx.extend(read_from_xlsx(options.summary, '方华', header='T'))
    report_dict = read_from_xlsx(options.report, 'Sheet1', header='T')

    output_xlsx = []
    treatment = []
    control = []
    index = 0
    for i in iter(input_xlsx):
        if i['TS编号'] == '':
            control.append(i)
        else:
            treatment.append(i)
    no_treatment = control

    for i in iter(treatment):
        index += 1
        sample_id = i['TS编号']
        if isinstance(sample_id, (int, float)):
            sample_id = query_ts_id(str(int(sample_id)), summary_xlsx)
        patient_name = i['患者姓名'].split('-')[0]
        sample_name = i['样本名'].split('-')[0]
        test_panel = i['样本组成'] + i['文库名*'] + ',' + str(int(i['要求测序数据量（G）']))
        note_1 = i['备注']
        if not sample_id.__contains__('Not Found'):
            receive_date, sales, cancer, note_2 = query(sample_id, summary_xlsx)
            report = query_report(sales, report_dict)

            sample_type = 'UoN'
            if test_panel.__contains__('小') or test_panel.__contains__('中'):
                sample_type = 'B'
            elif test_panel.__contains__('肾癌'):
                sample_type = query_sample_type(sample_id, summary_xlsx)
            elif test_panel.__contains__('癌组织'):
                for j in iter(control):
                    if j['患者姓名'].__contains__(patient_name):
                        if j['患者姓名'].__contains__('白细胞'):
                            sample_type = 'T+B'
                        else:
                            sample_type = '癌T+正T'
                        test_panel += '.' + j['样本组成'] + j['文库名*'] + ',' + str(int(j['要求测序数据量（G）']))
                        no_treatment.remove(j)
                        break
                    else:
                        sample_type = '癌T'
            elif test_panel.__contains__('全外'):
                sample_type = query_sample_type(sample_id, summary_xlsx)
                for j in iter(control):
                    if j['患者姓名'].__contains__(patient_name):
                        no_treatment.remove(j)
                        test_panel += '.' + j['样本组成'] + j['文库名*'] + ',' + str(int(j['要求测序数据量（G）']))
                        break
                    else:
                        sample_type = 'T'
            elif test_panel.__contains__('cfDNA'):
                sample_type = 'B'
                for j in iter(control):
                    if j['患者姓名'].__contains__(patient_name):
                        no_treatment.remove(j)
                        test_panel += '.' + j['样本组成'] + j['文库名*'] + ',' + str(int(j['要求测序数据量（G）']))
                        break

            if not if_pdl1(sample_id, summary_xlsx):
                if note_1 == '':
                    note_1 = '不要pdl1'
                else:
                    note_1 += '\n不要pdl1'

            if receive_date != '':
                deadline = receive_date + 9
            else:
                deadline = ''

            output_xlsx.append(
                dict(
                    序号=index, 收样日期=get_date(receive_date), 应出报告时间=get_date(deadline),
                    加急时间='', 样本名=sample_name, 检测编号=sample_id, 患者姓名=patient_name, 销售=sales,
                    数据量G=test_panel, 样本类型=sample_type,
                    癌种=cancer, 电子报告='', 盖章='', 快递='',
                    备注1=note_1, 医学部出具报告情况=note_2, 负责人=report))
        else:
            output_xlsx.append(
                dict(
                    序号=index, 收样日期='', 应出报告时间='',
                    加急时间='', 样本名=sample_name, 检测编号='', 患者姓名=patient_name, 销售='',
                    数据量G=test_panel, 样本类型='',
                    癌种='', 电子报告='', 盖章='', 快递='',
                    备注1=note_1, 医学部出具报告情况='', 负责人=''))

    for i in iter(no_treatment):
        index += 1
        test_panel = i['样本组成'] + i['文库名*'] + ',' + str(int(i['要求测序数据量（G）']))
        patient_name = i['患者姓名'].split('-')[0]
        sample_name = i['样本名'].split('-')[0]
        note_1 = i['备注']
        output_xlsx.append(
            dict(
                序号=index, 收样日期='', 应出报告时间='',
                加急时间='', 样本名=sample_name, 检测编号='', 患者姓名=patient_name, 销售='',
                数据量G=test_panel, 样本类型='',
                癌种='', 电子报告='', 盖章='', 快递='',
                备注1=note_1, 医学部出具报告情况='', 负责人=''))

    info_to_xlsx(list(output_xlsx[0].keys()), [list(i.values()) for i in output_xlsx],
                 options.output, 'Sheet1')
