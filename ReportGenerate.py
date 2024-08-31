s = ('\n'
     '（航运周报）\n'
     '2）8月12日上海出口集装箱运价指数报{a1}，环比{c1}{b1}。最新一期上海出口集装箱运价指数（SCFI）报{a2}，环比上周{c2}{b2}，对应{c3}{b3}。其中SCFI欧洲航线报${'
     'a3}/TEU，较上周{c4}${a4}/TEU，对应{c5}{b4}。SCFI美西较上周{c6}{b5}'
     '，SCFI美东较上周{c7}{b6}。\n'
     '3）本周成品油大西洋运价{c8}：8月16日BCTI TC7运价为{a5}美元/天，环比上周{a6}美元/天；太平洋一篮子运价为{a7}美元/天，环比上周{a8}'
     '美元/天；大西洋一篮子运价为{a9}美元/天，环比上周{a10}美元/天。\n'
     '4）本周原油TD3C、TD15{c9}，TD22{c10}：8月16日BDTI TD3C运价为{a11}美元/天，环比上周{c11}{a12}'
     '美元/天；8月16日苏伊士船型运价为{a13}美元/天，环比上周{a14}美元/天；8月16日阿芙拉船型运价为{a15}美元/天，环比上周{a16}美元/天。\n'
     '5）本周散运{c12}BDI ：8月16日BDI指数环比上周{b7}，至{a17}点；BCI指数环比上周{b8}，至{a18}点。\n'
     '(船舶周报）6）本周气体船、集装箱船价上涨： 8月16日新造船价指数环比上周{c13}{b9}点，为{a19}'
     '点；散货船/油轮/气体船船价指数环比{c14}/{c15}/{c16}{b10} ；当月集装箱船价指数环比{c17}{b11}。\n')
import openpyxl
import warnings

warnings.simplefilter("ignore")

MARITIME_REPORT = 'ShippingResult.xlsx'
VESSEL_REPORT = 'VesselResult.xlsx'


# Integer Number: a
# Signed Percent: b
# Up or Down: c

def format_value(sheet, row, col):
    cell_value = abs(sheet[f"{col}{row}"].value)
    cell_format = sheet[f"{col}{row}"].number_format

    if '0%' in cell_format or '0.00%' in cell_format:
        formatted_value = f"{cell_value * 100:.1f}%"
    else:
        formatted_value = cell_value

    return formatted_value


def set_precision(value, precision=0):
    if not precision:
        return round(value)
    return round(value, precision)


def add_sign(value):
    if isinstance(value, int) or isinstance(value, float):
        return f"+{value}" if value > 0 else f"{value}"
    elif isinstance(value, str):
        return f"+{value}" if float(value.rstrip('%')) > 0 else f"{value}"


def main():
    maritime = openpyxl.load_workbook(MARITIME_REPORT, data_only=True)
    vessel = openpyxl.load_workbook(VESSEL_REPORT, data_only=True)

    maritime_sheet = maritime['航运周报']
    vessel_sheet = vessel['船舶周报']

    format_replacement = {
        'a1': set_precision(maritime_sheet['L34'].value),
        'a2': set_precision(maritime_sheet['L34'].value),
        'a3': set_precision(maritime_sheet['L35'].value),
        'a4': set_precision(abs(maritime_sheet['L35'].value - maritime_sheet['M35'].value)),
        'a5': set_precision(maritime_sheet['L31'].value),
        'a6': add_sign(set_precision(maritime_sheet['L31'].value - maritime_sheet['M31'].value)),
        'a7': set_precision(maritime_sheet['L30'].value),
        'a8': add_sign(set_precision(maritime_sheet['L30'].value - maritime_sheet['M30'].value)),
        'a9': set_precision(maritime_sheet['L32'].value),
        'a10': add_sign(set_precision(maritime_sheet['L32'].value - maritime_sheet['M32'].value)),
        'a11': set_precision(maritime_sheet['L22'].value),
        'a12': set_precision(abs(maritime_sheet['L22'].value - maritime_sheet['M22'].value)),
        'a13': set_precision(maritime_sheet['L25'].value),
        'a14': add_sign(set_precision(maritime_sheet['L25'].value - maritime_sheet['M25'].value)),
        'a15': set_precision(maritime_sheet['L26'].value),
        'a16': add_sign(set_precision(maritime_sheet['L26'].value - maritime_sheet['M26'].value)),
        'a17': set_precision(maritime_sheet['L41'].value),
        'a18': set_precision(maritime_sheet['L42'].value),
        'a19': set_precision(vessel_sheet['L20'].value, 1),

        'b1': format_value(maritime_sheet, 34, 'N').lstrip('-'),
        'b2': abs(set_precision(maritime_sheet['L34'].value - maritime_sheet['M34'].value)),
        'b3': format_value(maritime_sheet, 34, 'N').lstrip('-'),
        'b4': format_value(maritime_sheet, 35, 'N'),
        'b5': format_value(maritime_sheet, 37, 'N'),
        'b6': format_value(maritime_sheet, 38, 'N'),
        'b7': add_sign(format_value(maritime_sheet, 41, 'N')),
        'b8': add_sign(format_value(maritime_sheet, 42, 'N')),
        'b9': set_precision(abs(vessel_sheet['L20'].value - vessel_sheet['M20'].value), 1),
        'b10': format_value(vessel_sheet, 35, 'N'),
        'b11': format_value(vessel_sheet, 30, 'N'),

        'c1': '上涨' if maritime_sheet['N34'].value > 0 else '下跌',
        'c2': '上涨' if maritime_sheet['L34'].value - maritime_sheet['M34'].value > 0 else '下跌',
        'c3': '增加' if maritime_sheet['N34'].value > 0 else '下降',
        'c4': '上涨' if maritime_sheet['L35'].value - maritime_sheet['M35'].value > 0 else '下跌',
        'c5': '上涨' if maritime_sheet['N35'].value > 0 else '下跌',
        'c6': '上涨' if maritime_sheet['N37'].value > 0 else '下跌',
        'c7': '上涨' if maritime_sheet['N38'].value > 0 else '下跌',
        'c8': '上涨' if maritime_sheet['N32'].value > 0 else '下跌',
        'c9': '(TO BE MODIFIED)',
        'c10': '(TO BE MODIFIED)',
        'c11': '上涨' if maritime_sheet['L22'].value - maritime_sheet['M22'].value > 0 else '下跌',
        'c12': '大船带涨' if maritime_sheet['N42'].value > maritime_sheet['N41'].value else '小船带涨',
        'c13': '上涨' if vessel_sheet['L20'].value > vessel_sheet['M20'].value else '下跌',
        'c14': '上升' if vessel_sheet['N26'].value > 0 else '持平' if vessel_sheet['N26'].value == 0 else '下跌',
        'c15': '上升' if vessel_sheet['N21'].value > 0 else '持平' if vessel_sheet['N26'].value == 0 else '下跌',
        'c16': '上升' if vessel_sheet['N35'].value > 0 else '持平' if vessel_sheet['N26'].value == 0 else '下跌',
        'c17': '上升' if vessel_sheet['N30'].value > 0 else '持平' if vessel_sheet['N30'].value == 0 else '下跌'
    }

    print(s.format(**format_replacement))


if __name__ == '__main__':
    main()
