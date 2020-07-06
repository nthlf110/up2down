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

    data_nrows = data_sh.nrows
    data_ncols = data_sh.ncols

    result_data = []
    if header == "T":
    	# 默认1行表头
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


def query(ts_id, summary):
    for i in iter(summary):
        if i['检测编号'] == ts_id:
            return i['送样日期'], i['收样日期'], i['销售'], i['癌种'], i['其他']
    return '', '', '', '', ''


def if_pdl1(ts_id, summary):
    panel = []
    for i in iter(summary):
        if i['检测编号'] == ts_id:
            panel.append(i['检测项目'])
    if sum([i.__contains__('PD-L1') for i in panel]) != 0:
        return True
    else:
        return False


def if_msi(ts_id, summary):
    panel = []
    for i in iter(summary):
        if i['检测编号'] == ts_id:
            panel.append(i['检测项目'])
    if sum([i.__contains__('MSI') for i in panel]) != 0:
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


def formatting(panel):
    if isinstance(panel, str):
        if panel.__contains__('NCCN'):
            return '结直肠癌NCCN指南必选检测'
        elif panel.__contains__('肠') and panel.__contains__('14'):
            return '结直肠癌14驱动基因'
        elif panel.__contains__('肠') and panel.__contains__('12'):
            return '结直肠癌12'
        elif sum([panel.upper().__contains__(i) for i in ['C','I','K','T']]) == 4:
            return 'ckit'
        else:
            return panel
    else:
        return panel


if __name__ == '__main__':
    parser = optparse.OptionParser()
    parser.add_option('-i', '--input', dest='input')
    parser.add_option('-o', '--output', dest='output',
                      default=datetime.datetime.strftime(datetime.date.today(), '%Y%m%d') + '-ffpe_raw.xlsx')
    parser.add_option('-s', '--summary', dest='summary', default='../外送统计表/桐树NGS外送样本统计表.xlsx')
    parser.add_option('-r', '--report', dest='report', default='../外送统计表/报告分配.xlsx')
    (options, args) = parser.parse_args()
    input_xlsx = read_from_xlsx(options.input, 'Sheet1', header='T')
    summary_xlsx = read_from_xlsx(options.summary, '北东区', header='T')
    summary_xlsx.extend(read_from_xlsx(options.summary, '中区', header='T'))
    summary_xlsx.extend(read_from_xlsx(options.summary, '南区', header='T'))
    summary_xlsx.extend(read_from_xlsx(options.summary, '方华', header='T'))
    report_dict = read_from_xlsx(options.report, 'Sheet1', header='T')

    output_xlsx = []
    for i in iter(input_xlsx):
        index = int(i['序号'])
        sample_id = i['检测编号']
        if isinstance(sample_id, (int, float)):
            sample_id = str(int(sample_id))
        sample_id = sample_id.upper()
        sample_name = i['样本姓名']
        sample_type = i['样本类型']
        test_panel = formatting(i['检测项目'])
        DNA = i['DNA标签']
        if isinstance(DNA, float): DNA = str(int(DNA))
        RNA = i['RNA标签']
        if isinstance(RNA, float): RNA = str(int(RNA))
        note_1 = i['备注']
        if sample_id != '':
            delivery_date, receive_date, sales, cancer, note_3 = query(sample_id, summary_xlsx)
            report = query_report(sales, report_dict)

            if if_pdl1(sample_id, summary_xlsx):
                note_2 = ''
            else:
                note_2 = '不要pdl1'
            if if_msi(sample_id, summary_xlsx):
                note_2 += ''
            else:
                note_2 += '不要msi'

            if receive_date:
                deadline = receive_date + 6
            else:
                deadline = ''

            output_xlsx.append(
                dict(
                    序号=index, 病例号='', 检测编号=sample_id, 样本姓名=sample_name, 销售=sales, 检测项目=test_panel,
                    DNA标签=DNA, RNA标签=RNA, 癌种=cancer, 样本类型=sample_type, 送检日期=get_date(delivery_date),
                    收样日期=get_date(receive_date), 最后出报告日期=get_date(deadline), 电子报告='',
                    备注1=note_1, 备注2=note_2, 备注3=note_3, 负责人=report))
        else:
            output_xlsx.append(
                dict(
                    序号=index, 病例号='', 检测编号=sample_id, 样本姓名=sample_name, 销售='', 检测项目=test_panel,
                    DNA标签=DNA, RNA标签=RNA, 癌种='', 样本类型=sample_type, 送检日期='',
                    收样日期='', 最后出报告日期='', 电子报告='',
                    备注1=note_1, 备注2='', 备注3='', 负责人=''))
    info_to_xlsx(list(output_xlsx[0].keys()), [list(i.values()) for i in output_xlsx],
                 options.output, 'Sheet1')
