# coding=utf-8
# 2018.7.3 从文件中提取所需数据
# python2

import os
import re
import xlwings as xw
import copy
import pandas as pd
from xlwings.constants import LineStyle

class Constants:
    xlNextToAxis = 4  # from enum Constants
    xlNoDocuments = 3  # from enum Constants
    xlNone = -4142  # from enum Constants
    xlNotes = -4144  # from enum Constants
    xlOff = -4146  # from enum Constants
    xl3DEffects1 = 13  # from enum Constants
    xl3DBar = -4099  # from enum Constants
    xl3DEffects2 = 14  # from enum Constants
    xl3DSurface = -4103  # from enum Constants
    xlAbove = 0  # from enum Constants
    xlAccounting1 = 4  # from enum Constants
    xlAccounting2 = 5  # from enum Constants
    xlAccounting3 = 6  # from enum Constants
    xlAccounting4 = 17  # from enum Constants
    xlAdd = 2  # from enum Constants
    xlAll = -4104  # from enum Constants
    xlAllExceptBorders = 7  # from enum Constants
    xlAutomatic = -4105  # from enum Constants
    xlBar = 2  # from enum Constants
    xlBelow = 1  # from enum Constants
    xlBidi = -5000  # from enum Constants
    xlBidiCalendar = 3  # from enum Constants
    xlBoth = 1  # from enum Constants
    xlBottom = -4107  # from enum Constants
    xlCascade = 7  # from enum Constants
    xlCenter = -4108  # from enum Constants
    xlCenterAcrossSelection = 7  # from enum Constants
    xlChart4 = 2  # from enum Constants
    xlChartSeries = 17  # from enum Constants
    xlChartShort = 6  # from enum Constants
    xlChartTitles = 18  # from enum Constants
    xlChecker = 9  # from enum Constants
    xlCircle = 8  # from enum Constants
    xlClassic1 = 1  # from enum Constants
    xlClassic2 = 2  # from enum Constants
    xlClassic3 = 3  # from enum Constants
    xlClosed = 3  # from enum Constants
    xlColor1 = 7  # from enum Constants
    xlColor2 = 8  # from enum Constants
    xlColor3 = 9  # from enum Constants
    xlColumn = 3  # from enum Constants
    xlCombination = -4111  # from enum Constants
    xlComplete = 4  # from enum Constants
    xlConstants = 2  # from enum Constants
    xlContents = 2  # from enum Constants
    xlContext = -5002  # from enum Constants
    xlCorner = 2  # from enum Constants
    xlCrissCross = 16  # from enum Constants
    xlCross = 4  # from enum Constants
    xlCustom = -4114  # from enum Constants
    xlDebugCodePane = 13  # from enum Constants
    xlDefaultAutoFormat = -1  # from enum Constants
    xlDesktop = 9  # from enum Constants
    xlDiamond = 2  # from enum Constants
    xlDirect = 1  # from enum Constants
    xlDistributed = -4117  # from enum Constants
    xlDivide = 5  # from enum Constants
    xlDoubleAccounting = 5  # from enum Constants
    xlDoubleClosed = 5  # from enum Constants
    xlDoubleOpen = 4  # from enum Constants
    xlDoubleQuote = 1  # from enum Constants
    xlDrawingObject = 14  # from enum Constants
    xlEntireChart = 20  # from enum Constants
    xlExcelMenus = 1  # from enum Constants
    xlExtended = 3  # from enum Constants
    xlFill = 5  # from enum Constants
    xlFirst = 0  # from enum Constants
    xlFixedValue = 1  # from enum Constants
    xlFloating = 5  # from enum Constants
    xlFormats = -4122  # from enum Constants
    xlFormula = 5  # from enum Constants
    xlFullScript = 1  # from enum Constants
    xlGeneral = 1  # from enum Constants
    xlGray16 = 17  # from enum Constants
    xlGray25 = -4124  # from enum Constants
    xlGray50 = -4125  # from enum Constants
    xlGray75 = -4126  # from enum Constants
    xlGray8 = 18  # from enum Constants
    xlGregorian = 2  # from enum Constants
    xlGrid = 15  # from enum Constants
    xlGridline = 22  # from enum Constants
    xlHigh = -4127  # from enum Constants
    xlHindiNumerals = 3  # from enum Constants
    xlIcons = 1  # from enum Constants
    xlImmediatePane = 12  # from enum Constants
    xlInside = 2  # from enum Constants
    xlInteger = 2  # from enum Constants
    xlJustify = -4130  # from enum Constants
    xlLTR = -5003  # from enum Constants
    xlLast = 1  # from enum Constants
    xlLastCell = 11  # from enum Constants
    xlLatin = -5001  # from enum Constants
    xlLeft = -4131  # from enum Constants
    xlLeftToRight = 2  # from enum Constants
    xlLightDown = 13  # from enum Constants
    xlLightHorizontal = 11  # from enum Constants
    xlLightUp = 14  # from enum Constants
    xlLightVertical = 12  # from enum Constants
    xlList1 = 10  # from enum Constants
    xlList2 = 11  # from enum Constants
    xlList3 = 12  # from enum Constants
    xlLocalFormat1 = 15  # from enum Constants
    xlLocalFormat2 = 16  # from enum Constants
    xlLogicalCursor = 1  # from enum Constants
    xlLong = 3  # from enum Constants
    xlLotusHelp = 2  # from enum Constants
    xlLow = -4134  # from enum Constants
    xlMacrosheetCell = 7  # from enum Constants
    xlManual = -4135  # from enum Constants
    xlMaximum = 2  # from enum Constants
    xlMinimum = 4  # from enum Constants
    xlMinusValues = 3  # from enum Constants
    xlMixed = 2  # from enum Constants
    xlMixedAuthorizedScript = 4  # from enum Constants
    xlMixedScript = 3  # from enum Constants
    xlModule = -4141  # from enum Constants
    xlMultiply = 4  # from enum Constants
    xlNarrow = 1  # from enum Constants
    xlOn = 1  # from enum Constants
    xlOpaque = 3  # from enum Constants
    xlOpen = 2  # from enum Constants
    xlOutside = 3  # from enum Constants
    xlPartial = 3  # from enum Constants
    xlPartialScript = 2  # from enum Constants
    xlPercent = 2  # from enum Constants
    xlPlus = 9  # from enum Constants
    xlPlusValues = 2  # from enum Constants
    xlRTL = -5004  # from enum Constants
    xlReference = 4  # from enum Constants
    xlRight = -4152  # from enum Constants
    xlScale = 3  # from enum Constants
    xlSemiGray75 = 10  # from enum Constants
    xlSemiautomatic = 2  # from enum Constants
    xlShort = 1  # from enum Constants
    xlShowLabel = 4  # from enum Constants
    xlShowLabelAndPercent = 5  # from enum Constants
    xlShowPercent = 3  # from enum Constants
    xlShowValue = 2  # from enum Constants
    xlSimple = -4154  # from enum Constants
    xlSingle = 2  # from enum Constants
    xlSingleAccounting = 4  # from enum Constants
    xlSingleQuote = 2  # from enum Constants
    xlSolid = 1  # from enum Constants
    xlSquare = 1  # from enum Constants
    xlStError = 4  # from enum Constants
    xlStar = 5  # from enum Constants
    xlStrict = 2  # from enum Constants
    xlSubtract = 3  # from enum Constants
    xlSystem = 1  # from enum Constants
    xlTextBox = 16  # from enum Constants
    xlTiled = 1  # from enum Constants
    xlTitleBar = 8  # from enum Constants
    xlToolbar = 1  # from enum Constants
    xlToolbarButton = 2  # from enum Constants
    xlTop = -4160  # from enum Constants
    xlTopToBottom = 1  # from enum Constants
    xlTransparent = 2  # from enum Constants
    xlTriangle = 3  # from enum Constants
    xlVeryHidden = 2  # from enum Constants
    xlVisible = 12  # from enum Constants
    xlVisualCursor = 2  # from enum Constants
    xlWatchPane = 11  # from enum Constants
    xlWide = 3  # from enum Constants
    xlWorkbookTab = 6  # from enum Constants
    xlWorksheet4 = 1  # from enum Constants
    xlWorksheetCell = 3  # from enum Constants
    xlWorksheetShort = 5  # from enum Constants
class LineStyle:
    xlContinuous = 1  # from enum XlLineStyle
    xlDash = -4115  # from enum XlLineStyle
    xlDashDot = 4  # from enum XlLineStyle
    xlDashDotDot = 5  # from enum XlLineStyle
    xlDot = -4118  # from enum XlLineStyle
    xlDouble = -4119  # from enum XlLineStyle
    xlLineStyleNone = -4142  # from enum XlLineStyle
    xlSlantDashDot = 13  # from enum XlLineStyle

class BorderWeight:
    xlHairline = 1  # from enum XlBorderWeight
    xlMedium = -4138  # from enum XlBorderWeight
    xlThick = 4  # from enum XlBorderWeight
    xlThin = 2  # from enum XlBorderWeight
class BordersIndex:
    xlDiagonalDown = 5  # from enum XlBordersIndex
    xlDiagonalUp = 6  # from enum XlBordersIndex
    xlEdgeBottom = 9  # from enum XlBordersIndex
    xlEdgeLeft = 7  # from enum XlBordersIndex
    xlEdgeRight = 10  # from enum XlBordersIndex
    xlEdgeTop = 8  # from enum XlBordersIndex
    xlInsideHorizontal = 12  # from enum XlBordersIndex
    xlInsideVertical = 11  # from enum XlBordersIndex
class comp_device:
    def __init__(self, name, model_info, pin_, etch_, pin_net_, pinx_, piny_, coNNsch_):
        self.name = name
        self.model_info = model_info
        self.pin_ = tuple(sorted(list(pin_)))
        self.etch_ = etch_
        self.pin_net_ = pin_net_
        self.pinx_ = pinx_
        self.piny_ = piny_
        self.coNNsch_ = coNNsch_
    def GetName(self):
        return self.name
    def GetModel(self):
        return self.model_info
    def GetPinList(self):
        return self.pin_
    def GetNet(self, pin_n):
        return self.pin_net_.get(pin_n)
    def GetNetList(self):
        return self.etch_
    def GetXPoint(self, pin_n):
        return self.pinx_[pin_n]
    def GetYPoint(self, pin_n):
        return self.piny_[pin_n]
    def GetXY(self, pin_n):
        return (self.pinx_[pin_n], self.piny_[pin_n])
def _flatten(a):
    if not isinstance(a, (list, )):
        return [a]
    else:
        b = []
        for item in a:
            b += _flatten(item)
    return b
def get_exclude_netlist(netlist):  # netlist = All_Net_List
    # Get pwr and gnd net list

    PWR_Net_KeyWord_List = ['^\+.*', '^-.*',
                            'VREF|PWR|VPP|VSS|VREG|VCORE|VCC|VT|VDD|VLED|PWM|VDIMM|VGT|VIN|[^S](VID)|VR',
                            'VOUT|VGG|VGPS|VNN|VOL|VSD|VSYS|VCM|VSA', '.*V[0-9]A.*', '.*V[0-9]\.[0-9]A.*',
                            '.*V[0-9]_[0-9]A.*', '.*V[0-9]S.*', '^V[0-9].*', '.*_V[0-9]', '.*_V[0-9][0-9]',
                            '.*V[0-9]P.*', '.*V[0-9]V.*', '.*[0-9]V[0-9].*', '^[0-9]V.*', '^[0-9][0-9]V.*',
                            '.*[0-9]\.[0-9]V.*', '.*[0-9]_[0-9]V.*', '.*_[0-9]V.*', '.*_[0-9][0-9]V.*',
                            '.*_[0-9]\.[0-9]V.*', '.*[0-9]P[0-9]V.*', '.[0-9]*P[0-9][0-9]V.*', '.*V_[0-9]P[0-9].*',
                            '.*\+[0-9]V.*', '.*\+[0-9][0-9]V.*']
    PWR_Net_List = [net for net in netlist for keyword in PWR_Net_KeyWord_List if re.findall(keyword, net) != []]
    # myprint(PWR_Net_List)
    PWR_Net_List = sorted(list(set(PWR_Net_List)))

    GND_Net_List = [net for net in netlist if net.find('GND') > -1]
    GND_Net_List = sorted(list(set(GND_Net_List)))

    # 被排除的线：地线和电源线
    Exclude_Net_List = sorted(list(set(PWR_Net_List + GND_Net_List)))

    return Exclude_Net_List, PWR_Net_List, GND_Net_List

# 设置单元格内容的字型字体大小和字体位置
def SetCellFont_current_region(sheet, start_cell_ind, Font_Name, Font_Size, horizon_alignment):
    sheet.range(start_cell_ind).current_region.api.Font.Name = Font_Name
    sheet.range(start_cell_ind).current_region.api.Font.Size = Font_Size
    if horizon_alignment == 'c':
        sheet.range(start_cell_ind).current_region.api.HorizontalAlignment = Constants.xlCenter
    elif horizon_alignment == 'r':
        sheet.range(start_cell_ind).current_region.api.HorizontalAlignment = Constants.xlRight
    elif horizon_alignment == 'l':
        sheet.range(start_cell_ind).current_region.api.HorizontalAlignment = Constants.xlLeft
# 设置表格边框
def SetCellBorder_current_region(sheet, start_cell_ind):
    sheet.range(start_cell_ind).current_region.api.Borders.LineStyle = LineStyle.xlContinuous

# def managet_diff_data():
#     '''
#     将intel 给的表格中Pin List 中的pin name按照顺序排列，并且保持差分对在一起的形式
#     '''
#     # root_path = os.getcwd()
#     root_path =sys.path[0]
#     root_path = os.path.join('\\'.join(root_path.split('\\')))
#     print("",root_path)
#     result_path = os.path.join('\\'.join(root_path.split('\\')[:-2]))
#     print(result_path)
#     location_info_path=os.path.join(result_path,'info')
#     print(location_info_path)
#     for item in os.listdir(location_info_path):
#         if item.find('.xlsx') > -1 and item.find('~$') == -1:
#             GPIO_table = os.path.join(location_info_path, item)
#     app = xw.App(visible=False, add_book=False)
#     app.display_alerts = False
#     app.screen_updating = False
#     # print(GPIO_table)
#     wb = app.books.open(GPIO_table)
#     # wb = xw.Book(GPIO_table)
#     sht = wb.sheets["Pin List"]
#     col_idx = len(sht.range('A2').options(expand='table').value) + 1
#     list_length= len(sht.range('A2').options(expand='table').value)
#     # print(list_length)
#
#     Pin_name_list = [str(x) for x in sht.range('A2:A{}'.format(col_idx)).value]  # Pin name 数据
#     # Pin_name_list.sort(key=lambda i:len(i),reverse= True)
#     # print('pin',Pin_name_list)
#     Pin_Location_list = [str(x) for x in sht.range('B2:B{}'.format(col_idx)).value]  # Pin list 数据
#     name_location_list=[]
#     for i in range(len(Pin_name_list)):
#         name_location=[Pin_name_list[i],Pin_Location_list[i]]
#         name_location_list.append(name_location)
#     # print("DICT",name_location_list)
#
#     wb.close()
#     app.quit()

def get_net_info():
    """
    获取所有的net name
    :return:
    """
    root_path = os.getcwd()

    RUL_table = None
    for item in os.listdir(root_path):
        if item.find('.xlsx') > -1 and item.find('~$') == -1:
            RUL_table = os.path.join(root_path, item)

    wb = xw.Book(RUL_table)
    active_sheet = wb.sheets.active  # Get the active sheet object
    selection_range = wb.app.selection
    start_ind = (selection_range.row, selection_range.column)
    all_net_list = []
    net_node_dict = {}
    net_node_dict1 = {}
    net_location_list = []
    net_node_list = []
    # 输出每条信号线相接的元器件及其pin脚
    with open(os.path.join(root_path, 'pstxnet.dat'), 'r') as file1:

        # print("pstxnet.dat",dat_path)
        content1 = file1.read().split('NET_NAME')
        # print(content1)
        for ind in range(len(content1)):
            content1[ind] = content1[ind].split('\n')
            # print(content1[ind])
        for x in content1:
            node_list = []
            node_location_list = []
            all_net_list.append(x[1][1:-1])
            # print("all_net_list",all_net_list)
            for y_idx in range(len(x)):
                if x[y_idx].find('NODE_NAME') > -1:
                    node_location_list.append([x[y_idx].split('NODE_NAME\t')[-1].split(' ')[0],x[y_idx].split('NODE_NAME\t')[-1].split(' ')[1], x[y_idx + 2].split("'")[1]])
                    node_list.append([x[y_idx].split('NODE_NAME\t')[-1].split(' ')[0], x[y_idx + 2].split("'")[1]]) # 不包含Location信息
            node_flatten_list = list(_flatten([[x[1][1:-1]] + node_location_list]))
            net_node_dict[x[1][1:-1]] = node_flatten_list
            # print("net_node_dict", net_node_dict)  # net线对应经过的芯片，pin loaction ,pin name
            net_location_list.append(node_flatten_list)
            node_flatten_list1 = list(_flatten([[x[1][1:-1]] + node_list]))
            net_node_dict1[x[1][1:-1]] = node_flatten_list1
            net_node_list.append(node_flatten_list1)
            # print(325,net_node_list)#['M_DA7', 'XMM3', 'DQ7', 'XMM4', 'DQ7', 'XU1', 'DDR0_DQ_7/DDR0_DQ_7']
        all_net_list = all_net_list[1:]
    # print(all_net_list)
    # print(322,net_node_dict)
    # print(323,net_location_list)  # net_node_dict中的value值
    # print(324,all_net_list)  # net_node_dict中所有的key值
  ####################################################
    ic_pin_name = []
    for i in range(1,len(net_location_list)):
        # 数据形式 net_name ( IC   pin_number  pin_location  )
        # 'PME_OUT#_R', 'U5', '2', 'GPIO25/SIOPME#/SP_C_CFG0/PWM5/S', 'TPVIA4', '1', 'TP'
        net_loc = net_location_list[i]
        print(net_loc[0])
        ic_pin = [net_loc[1::3],net_loc[2::3],net_loc[3::3]]
        for i in range(len(ic_pin[0])):
            ic_pin_name.append([ic_pin[0][i],ic_pin[1][i],net_loc[0],ic_pin[2][i]])
    # print(ic_pin_name)# 将net线经过的ic  和pin 都与其实线名关联
    def find_last(string, str):
        last_position = -1
        while True:
            position = string.find(str, last_position + 1)
            if position == -1:
                return last_position
            last_position = position
    # print(sorted(all_net_name))
    temp1 = all_net_list
    temp2, temp3 = [], []
    diff_pair_list = []
    for i in range(len(temp1)):
        if temp1[i].count('N')==1:
            temp2.append(temp1[i])
        elif temp1[i].count("N") > 1:
            temp3.append(temp1[i])
    for i in range(len(temp2)):
        trans2_data = temp2[i].replace("N", "P")
        if trans2_data in temp1:
            diff_pair_list.append([temp2[i],trans2_data])
            # diff_pair_list.append(temp2[i])
            # diff_pair_list.append(trans2_data)
            temp1.remove(temp2[i])
            temp1.remove(trans2_data)
    # print(len(temp1))
    for i in range(len(temp3)):
        last_ind = find_last(temp3[i], 'N')
        list3 = list(temp3[i])
        list3[last_ind] = 'P'
        trans3_data = ''.join(list3)
        if trans3_data in temp1:
            diff_pair_list.append([temp3[i], trans3_data])
            # diff_pair_list.append(temp3[i])
            # diff_pair_list.append(trans3_data)
            temp1.remove(temp3[i])
            temp1.remove(trans3_data)
    # print(diff_pair_list)# 所有差分
    # print(temp1)# 除去差分的所有net线
    temp1 = sorted(temp1)
    diff_pair_list = sorted(diff_pair_list)
    # print(diff_pair_list)

    # 将net 写入excel中，添加对应的数据（线宽，线距，space）


    output_path = os.path.join(root_path, 'output')
    if os.path.exists(output_path):
        pass
    else:
        os.mkdir(output_path)
    result_excel_path = os.path.join(output_path, 'result.xlsx')
    app = xw.App(visible=False)
    try:
        wb = app.books.open(result_excel_path)
    except IOError:
        wb = app.books.add()
        wb.save(result_excel_path)
        wb = app.books.open(result_excel_path)

    sh1 = wb.sheets[0]
    sh1.clear()
    # sh1.range('A1').value = 'Type'
    sh1.range('A1').value = 'Name'
    diff_data=[]
    for x in range(len(diff_pair_list)):
        for z in range(len(diff_pair_list[x])):
            diff_data.append(diff_pair_list[x][z])
    # print(diff_data)

    for i in range(len(diff_data)):
        sh1.range(i+ 2, 1).value =diff_data[i]
    # print(temp1)
    for y in range(len(temp1)):
        sh1.range(len(diff_pair_list)*2+2+y,1).value = temp1[y]

    sh1.autofit('c')
    wb.save()
    wb.close()
    app.quit()
    ####################################################
 #  获取表格中写入的chip 以及 pin location
 #
 #    chip_table = active_sheet.range(start_ind[0], start_ind[1]).expand('table').value
 #    net_name_list = []
 #    all_pin_name = []
 #    all_ic_name = []
 #    if chip_table is not None:
 #        for idx in range(1,len(chip_table)):
 #            if active_sheet.range(start_ind[0], start_ind[1]).value.upper() == 'CHIP(SINGLE)':
 #                chip_table = active_sheet.range(start_ind[0] + idx, start_ind[1]).value.upper()
 #                pin_table = active_sheet.range(start_ind[0] + idx, start_ind[1] + 1).value
 #                if type(pin_table) == float:
 #                    pin_table = str(int(pin_table)).upper()
 #                else:
 #                    pin_table = pin_table.upper()
 #                for i in range(len(ic_pin_name)):
 #                    if ic_pin_name[i][0] == chip_table and ic_pin_name[i][1] == pin_table:
 #                        active_sheet.range((start_ind[0] + idx, start_ind[1] + 2)).value = ic_pin_name[i][2]
 #                active_sheet.range((start_ind[0], start_ind[1])).api.Interior.ColorIndex = 48
 #                active_sheet.range((start_ind[0] + idx, start_ind[1])).api.Interior.ColorIndex = 43
 #                active_sheet.range((start_ind[0], start_ind[1] + 1)).api.Interior.ColorIndex = 48
 #                active_sheet.range((start_ind[0] +  idx, start_ind[1] + 1)).api.Interior.ColorIndex = 43
 #                active_sheet.range((start_ind[0], start_ind[1] + 2)).api.Interior.ColorIndex = 48
 #                active_sheet.range((start_ind[0], start_ind[1] + 1)).value = 'Pin location'
 #                active_sheet.range((start_ind[0], start_ind[1] + 2)).value = 'Net name'
 #                active_sheet.range((start_ind[0] +  idx, start_ind[1] + 2)).api.Interior.ColorIndex = 43
 #                # chip_single_list.append(chip_table)
 #                # location_single_list.append(pin_table)
 #            elif active_sheet.range(start_ind[0], start_ind[1]).value.upper() == 'CHIP(DIFF)':
 #                chip_diff_table1 = active_sheet.range(start_ind[0]+idx, start_ind[1]).value.upper()
 #                pin_diff_table1 = active_sheet.range(start_ind[0]+idx, start_ind[1] + 1).value
 #                if type(pin_diff_table1) == float:
 #                    pin_diff_table1 = str(int(pin_diff_table1)).upper()
 #                else:
 #                    pin_diff_table1 = pin_diff_table1.upper()
 #                for i in range(len(ic_pin_name)):
 #                    if ic_pin_name[i][0] == chip_diff_table1 and ic_pin_name[i][1] == pin_diff_table1:
 #                        active_sheet.range((start_ind[0] + idx, start_ind[1] + 2)).value = ic_pin_name[i][2]
 #                        net_name_list.append(ic_pin_name[i][2])
 #                        all_ic_name.append(ic_pin_name[i][0])
 #                        all_pin_name.append(ic_pin_name[i][3])
 #                # chip_diff_list.append(chip_diff_table1)
 #                # location_diff_list.append(pin_diff_table1)
 #    # print('name_list',net_name_list)
 #    # print(111,all_ic_name)
 #    # print(222,all_pin_name)
 #    # # 根据pin location 和Ic 获取的net name 并初步写入相应的表格中
 #    # # net_name_dict = {}
 #    #
 #    # # all_pin_name = []
 #    # # all_ic_name = []
 #    #
 #    # # net_org_list = []
 #    # # net_number_list = []
 #    #
 #    # #
 #    # # if active_sheet.range(start_ind[0], start_ind[1]).value.upper() == 'CHIP(SINGLE)':
 #    # #     for i in range(len(net_location_list)):
 #    # #         data_list = net_location_list[i]
 #    # #         for x in range(len(data_list)):
 #    # #             # 获取输入的IC 以及 pin location
 #    # #             for m in range(len(location_single_list)):
 #    # #                 if type(location_single_list[m]) == float:
 #    # #                     location_single_list[m]=str(int(location_single_list[m]))
 #    # #                 if data_list[x] == chip_single_list[m].upper() and data_list[x+1] == location_single_list[m].upper():
 #    # #                     active_sheet.range((start_ind[0] + m + 1, start_ind[1] + 2)).value = data_list[0]
 #    # #                     net_name_dict[m] = data_list[0]
 #    # #                     active_sheet.range((start_ind[0] , start_ind[1])).api.Interior.ColorIndex = 48
 #    # #                     active_sheet.range((start_ind[0]+ 1 + m, start_ind[1])).api.Interior.ColorIndex = 43
 #    # #                     active_sheet.range((start_ind[0], start_ind[1] + 1)).api.Interior.ColorIndex = 48
 #    # #                     active_sheet.range((start_ind[0] + 1 + m, start_ind[1] +1)).api.Interior.ColorIndex = 43
 #    # #                     active_sheet.range((start_ind[0], start_ind[1] + 2)).api.Interior.ColorIndex = 48
 #    # #                     active_sheet.range((start_ind[0], start_ind[1] + 1)).value = 'Pin location'
 #    # #                     active_sheet.range((start_ind[0], start_ind[1] + 2)).value = 'Net name'
 #    # #                     active_sheet.range((start_ind[0] + 1 + m, start_ind[1]+2)).api.Interior.ColorIndex = 43
 #    #
 #    # # if active_sheet.range(start_ind[0], start_ind[1]).value.upper() == 'CHIP(DIFF)':
 #    # #     # print(374,net_location_list)
 #    # #     for i in range(len(net_location_list)):
 #    # #         data_list = net_location_list[i]
 #    # #         for x in range(len(data_list)):
 #    # #             for n in range(len(location_diff_list)):
 #    # #                 if type(location_diff_list[n]) == float:
 #    # #                     location_diff_list[n]=str(int(location_diff_list[n]))
 #    # #                 if data_list[x] == chip_diff_list[n].upper() and data_list[x+1] == location_diff_list[n].upper():
 #    # #                     active_sheet.range(start_ind[0] + n+1, start_ind[1] + 2).value = data_list[0]
 #    # #                     net_org_list.append(data_list[0])
 #    # #                     net_number_list.append(n)
 #    # #                     net_name_dict[n] = data_list[0]
 #    # #                     all_pin_name.append(data_list[x+2])
 #    # #                     all_ic_name.append(data_list[x])
 #    # #
 #    # # print('all_pin_name',all_pin_name)
 #    # # print('all_ic_name',all_ic_name)
 #    #
 #    # # net_name_list = []
 #    # # for n_idx in range(len(net_name_dict)):
 #    # #     net_name_list.append(net_name_dict[n_idx])
 #    # # print(409,net_name_list)
 #    # # print("net_name",net_name_list)
 #    # # 输出每个元器件的pin脚数目以确定是否是芯片
 #    with open(os.path.join(root_path, 'pstchip.dat'), 'r') as file:
 #                content = file.read().split('end_primitive')
 #                # myprint(content)
 #                pattern = re.compile(r".*?primitive '(.*?)'")
 #                ext_icname_pin_num_dict = {}
 #                for c_item in content:
 #                    key_item = pattern.findall(c_item)
 #                    # print(1, c_item)
 #                    # print(2, key_item)
 #                    if key_item:
 #                        ext_icname_pin_num_dict[key_item[0]] = c_item.count('PIN_NUMBER')    # PIN_NUMBER 即为Pin Location
 #    # print(416,ext_icname_pin_num_dict)  # key值对应几个pin脚
 #
 #    with open(os.path.join(root_path, 'pstxprt.dat'), 'r') as file2:
 #        content2 = file2.read().split('PART_NAME')
 #        # myprint(content2)
 #        all_node_list = []
 #        all_res_dict = {}
 #        # all_res_list = []
 #        node_page_dict = {}
 #        primitive_list = []
 #        ic_ext_icname_dict = {}
 #        for ind in range(len(content2)):
 #            content2[ind] = content2[ind].split('\n')
 #        # myprint(content2)
 #        for x in content2:
 #            # print(x)
 #            # print('\n')
 #            node = x[1].split(' ')[1]
 #            # print(node)
 #            all_node_list.append(node)
 #            if x[0] == '':
 #                pattern2 = re.compile(r".*?'(.*?)'.*?")
 #                # pattern3 = re.compile(r".*?@.*?@(.*?)\..*?")
 #                node1 = x[1].split(' ')[1]
 #                ic_ext_icname_dict[node1] = pattern2.findall(x[1])[0]
 #                # print(430,pattern3.findall(x[5]))
 #                # if pattern3.findall(x[5])[0].upper() == 'RESISTOR':
 #                #     all_res_list.append(node1)
 #            pattern = re.compile(r".*?_.*?_(.*?)_.*?")
 #            res_val = pattern.findall(x[1].split(' ')[2])
 #            # print('x', x[1].split(' ')[2])
 #            # print('res', res_val)
 #            if res_val:
 #                # if res_val[0][-1] == 'K':
 #                #     all_res_dict[node] = res_val[0][:-1]
 #                # else:
 #                all_res_dict[node] = res_val[0]
 #
 #            if x[0] == '':
 #                primitive_list.append(x[1].split("\'")[1])
 #            if 'page' in x[6]:
 #                page_now = x[6].split(':')[-1].split('_')[0]
 #            else:
 #                page_now = x[7].split(':')[-1].split('_')[0]
 #            if node_page_dict.get(page_now):
 #                node_page_dict[page_now] += [all_node_list[-1]]
 #            else:
 #                node_page_dict[page_now] = [all_node_list[-1]]
 #    # print(572,node_page_dict)
 #    # print(467,ic_ext_icname_dict)  # 给出IC以及IC对应的关键字，通过关键字连接到pstchip.dat文件中，可获得对应的pin脚数目
 #
 #    # 获得所有电源线
 #            # 获得电源线及地线信息
 #    Exclude_Net_List, PWR_Net_List, GND_Net_List = get_exclude_netlist(all_net_list)
 #    IC_pin_num_dict = {}
 #    for ic_item in ic_ext_icname_dict.keys():
 #        IC_pin_num_dict[ic_item] = ext_icname_pin_num_dict[ic_ext_icname_dict.get(ic_item)] # 输入Pin脚数目判断是否进入下一个IC，#通过ic_ext_icname_dict.get(ic_item)即key获得value，再将value作为key获取value，其结果获得对应IC的pin脚数目
 #    # print(IC_pin_num_dict) # 储存结果 IC:pin 脚数量'@U4': 1, 'AC1': 2
 #    pin_net_node_list = []
 #    pin_net_node_dict = {}
 #    net_node_copy_list = copy.deepcopy(net_node_list) # ['M_DA7', 'XMM3', 'DQ7', 'XMM4', 'DQ7', 'XU1', 'DDR0_DQ_7/DDR0_DQ_7']
 #    error_all_list = []
 #    # 找出详细的连接信息
 #    # print("type",all_pin_name)
 #
 #    for pin_idx in range(len(all_pin_name)):
 #        # print(all_pin_name[pin_idx])
 #        # 遍历没有拼写错误的所有pin脚
 #        if all_pin_name[pin_idx]:
 #                node_item_flag = False
 #                for node_item in net_node_list:
 #                    net_item = node_item[0]
 #                    # print("node_item", node_item)
 #                    # 找到pin脚连接的信号线信息，找到出pin的连接信号线
 #                    if all_pin_name[pin_idx] in node_item:
 #                        if all_ic_name[pin_idx] in node_item:
 #                            node_item3 = []
 #                            pin_net_node_list1 = [net_item]  # 初始的net线名
 #                            # print(1,pin_net_node_list1)
 #                            node_item1 = copy.deepcopy(node_item)
 #                            node_item1.pop(node_item1.index(all_pin_name[pin_idx], 1) - 1)
 #                            node_item1.pop(node_item1.index(all_pin_name[pin_idx], 1))
 #                            node_item1 = node_item1[1::2]
 #                            # print('node_item1',node_item1)
 #                            # print("Ic_name",node_item1)
 #                            flagfour = True
 #                            split_flag = True
 #                            split_node_list = []
 #                            layer_num = 0
 #                            layer_add_num_dict = {}
 #                            # 如果pin未连接信号线
 #                            if pin_net_node_list1[0] == 'NC':
 #                                pin_net_node_list1 = []
 #                                pin_net_node_dict[net_name_list[pin_idx]] = pin_net_node_list1
 #                                node_item_flag = True
 #                            else:
 #                                while flagfour:
 #                                    node_item3 = []
 #                                    break_flag = False
 #                                    split_out_flag = False
 #                                    if split_flag:
 #                                        split_node_list.append(copy.deepcopy(node_item1))
 #                                        # print(517,node_item1)
 #                                    for x_idx in range(len(node_item1)):
 #                                        # IC_flag = False
 #                                        next_flag = False
 #                                        all_break = False
 #                                        add_num = 0
 #                                        # myprint(len(node_item1), x_idx, node_item1)
 #                                        item1 = node_item1[x_idx]  # 储存net线经过的每经过的一个芯片名称
 #                                        split_node_list[-1].pop(split_node_list[-1].index(node_item1[x_idx]))
 #
 #                                        #利用key值item1，获取字典IC_pin_num_dict的value，即对应IC的pin脚数目，然后进行判断
 #                                        # 大于3说明到另外一个芯片了，停止
 #                                        if IC_pin_num_dict[item1] > 3:
 #                                            # print('pin_num', item1, IC_pin_num_dict[item1])
 #                                            split_flag = False
 #                                            add_num += 1
 #                                            split_node_list.append([])
 #                                        else:
 #                                            # 判断是否为终止端元器件（中间的元器件会出现两次）
 #                                            if node_item1.count(item1) > 1:
 #                                                # print('pin_net_node_list1',pin_net_node_list1)
 #                                                add_num += 1
 #                                            if node_item1.count(item1) == 1:
 #                                                # print(item1)
 #                                                count = 0
 #                                                # print(561, pin_net_node_list1)
 #                                                for node_item2 in net_node_copy_list:
 #                                                    # print(550,net_node_copy_list[-1],node_item2,node_item,node_item3)
 #                                                    count += 1
 #                                                    add_sch_flag = False
 #                                                    # 找到元器件所连接的另一根线
 #                                                    # 如果这次经过的线与上次或第一次相同，则退出
 #                                                    if item1 in node_item2 and node_item2 != node_item \
 #                                                            and node_item2 != node_item3:
 #
 #                                                        # 如果中间没有经过过这个元器件则进入循环
 #                                                        if node_item2[0] not in pin_net_node_list1:
 #                                                            add_sch_flag = True
 #                                                            # 此处判断，如果电源线或者GND则不做添加
 #                                                            if node_item2[0] not in Exclude_Net_List:
 #                                                                pin_net_node_list1.append(node_item2[0])
 #                                                            # print(552,pin_net_node_list1)
 #                                                            add_num += 2
 #                                                            if node_item2[0] != node_item[0] and node_item2 != node_item3:
 #                                                                #
 #                                                                # if node_item2[0] in Exclude_Net_List:  # node_item2[0]判断其是否为电源线或者接地GND线
 #                                                                #     # myprint('layer_num', layer_num)
 #                                                                #     split_node_list.append([])
 #                                                                #     # layer_add_num_dict[layer_num] = add_num
 #                                                                #     split_flag = False
 #                                                                #     break
 #                                                                node_item1 = copy.deepcopy(node_item2)
 #                                                                node_item3 = copy.deepcopy(node_item2)
 #                                                                # myprint(4, node_item1)
 #                                                                node_item1.pop(node_item1.index(item1) - 1)
 #                                                                node_item1.pop(node_item1.index(item1))
 #                                                                # myprint(4, node_item1)
 #                                                                node_item1 = node_item1[1::2]
 #                                                                next_flag = True
 #                                                                split_flag = True
 #                                                                break_flag = True
 #                                                                break
 #                                                                # break 不要break，是因为可能元器件有超过两个pin，要所有都遍历到
 #                                                                # 虽然速度会变慢
 #                                                            else:
 #                                                                all_break = True
 #
 #                                                    if net_node_copy_list[-1] == node_item2 and add_sch_flag is False:  # 如果未进入上一个if判断，add_sch_flag flag为 False
 #                                                        split_flag = False
 #                                                        add_num += 1
 #                                                        split_node_list.append([])
 #                                                        pin_net_node_list1.append(item1)
 #                                                        # print(599,pin_net_node_list1)
 #                                                        if pin_net_node_list1 not in pin_net_node_list :
 #                                                            pin_net_node_list.append(pin_net_node_list1)
 #
 #                                                        if pin_net_node_dict.get(net_name_list[pin_idx]):
 #                                                            pin_net_dict_list = pin_net_node_dict[net_name_list[pin_idx]]
 #                                                            pin_net_dict_list.append(pin_net_node_list1)
 #                                                            pin_net_node_dict[net_name_list[pin_idx]] = \
 #                                                                pin_net_dict_list
 #                                                        else:
 #                                                            pin_net_node_dict[net_name_list[pin_idx]] = \
 #                                                                [pin_net_node_list1]
 #
 #                                        if all_break:
 #                                            pass
 #                                        else:
 #                                            layer_add_num_dict[layer_num] = add_num
 #                                            split_node_flag = True
 #                                            before_layer_num = 0
 #                                            if next_flag is False:
 #                                                if node_item1[-1] == item1 or split_flag is False:
 #                                                    split_flag = True
 #                                                    split_node_list.pop(-1)
 #                                                    before_layer_num = copy.deepcopy(layer_num)
 #                                                    layer_num = len(split_node_list)
 #                                                    if split_node_list:
 #                                                        try:
 #                                                            while not split_node_list[-1]:
 #                                                                split_node_list.pop(-1)
 #                                                                layer_num -= 1
 #                                                        except IndexError:
 #                                                            pass
 #
 #                                                    if split_node_list:
 #                                                        node_item1 = split_node_list[-1]
 #                                                        # if len(node_item1) == 1:
 #                                                        split_node_list.pop(-1)
 #                                                        split_node_flag = True
 #                                                        break_flag = True
 #                                                    else:
 #                                                        split_node_flag = False
 #                                                        flagfour = False
 #                                                        split_out_flag = True
 #
 #                                                    if pin_net_node_list1 not in pin_net_node_list:
 #                                                        pin_net_node_list.append(pin_net_node_list1)
 #
 #                                                    if pin_net_node_dict.get(net_name_list[pin_idx]):
 #                                                        pin_net_dict_list = pin_net_node_dict[net_name_list[pin_idx]]
 #                                                        pin_net_dict_list.append(pin_net_dict_list)
 #                                                        pin_net_node_dict[net_name_list[pin_idx]] = \
 #                                                                pin_net_dict_list
 #                                                        # print(pin_net_node_dict)
 #                                                    else:
 #                                                        pin_net_node_dict[net_name_list[pin_idx]] = \
 #                                                            [pin_net_node_list1]
 #
 #                                            if split_node_flag:
 #                                                if before_layer_num:
 #                                                    for layer_idx in range(layer_num, before_layer_num + 1):
 #                                                        if layer_add_num_dict[layer_idx] != 0:
 #                                                            pin_net_node_list1 = pin_net_node_list1[:-layer_add_num_dict[layer_idx]]
 #                                                        layer_add_num_dict.pop(layer_idx)
 #                                                    layer_num -= 1
 #                                            if break_flag:
 #                                                break
 #
 #                                            if split_out_flag:
 #                                                # myprint('split_break')
 #                                                break
 #                            if node_item_flag:
 #                                break
 #        else:
 #            error_all_list.append(pin_idx)
 #            pin_net_node_dict[net_name_list[pin_idx]] = [['misspelled']]
 #
 #
 #    # print('all_info',pin_net_node_list)
 #    net_last_list = []
 #    # all_info_list = []
 #    # 筛选出起始net名为一开始出pin的net名，若一开始出pin的net不一致，则不作考虑
 #    for x in range(len(pin_net_node_list)):
 #        if pin_net_node_list[x][0] in net_name_list and pin_net_node_list[x-1][0] !=pin_net_node_list[x][0]:
 #            net_last_list.append(pin_net_node_list[x])
 #    # print('net_last_list',net_last_list)
 #    net_idx=(start_ind[0]+ 1, start_ind[1] + 2)
 #    for m in range(len(net_last_list)):
 #        # print(net_last_list[m])
 #
 #        last_net = net_last_list[m]
 #        for i in range(len(last_net)):
 #            active_sheet.range(net_idx[0] + m, net_idx[1] + i).value = last_net[i]
 #            active_sheet.range((start_ind[0], start_ind[1])).api.Interior.ColorIndex = 48
 #            active_sheet.range((start_ind[0] + 1 + m , start_ind[1])).api.Interior.ColorIndex = 43
 #            active_sheet.range((start_ind[0], start_ind[1] + 1)).api.Interior.ColorIndex = 48
 #            active_sheet.range((start_ind[0], start_ind[1] + 1)).value = 'Pin location'
 #            active_sheet.range((start_ind[0], start_ind[1] + 2)).value = 'Net name'
 #            active_sheet.range((start_ind[0] + 1 + m, start_ind[1] + 1)).api.Interior.ColorIndex = 43
 #            active_sheet.range((start_ind[0], start_ind[1] + 2 + i)).api.Interior.ColorIndex = 48
 #            active_sheet.range((start_ind[0]+ 1 + m, start_ind[1] + 2 + i )).api.Interior.ColorIndex = 43
 #
 #    SetCellFont_current_region(active_sheet, start_ind, '等线', 12, 'l')
 #    SetCellBorder_current_region(active_sheet, start_ind)
 #    wb.save()
 #    # wb.close()s



def get_Topology_Info():
    """
    获取Topology的所有值
    :return:
    """

    root_path=os.getcwd()
    RUL_table = None
    for item in os.listdir(root_path):
        if item.find('.xlsx') > -1 and item.find('~$') == -1:
            RUL_table = os.path.join(root_path, item)
    print(RUL_table)
    wb = xw.Book(RUL_table)
    sheet_names = []
    for i in range(wb.sheets.count):
        if i>3:
            sheet_names.append(wb.sheets[i].name)
    print(sheet_names)
    excel_data=[]
    for i in sheet_names:
        # print(i)#每个sheet表名称
        df = pd.read_excel(RUL_table,sheet_name=i )
        # print(df.index.values)# 索引
        # print(df.columns.values)# 列号  nan值均为float类型
        total_Padstack = df.values# 所有值
        excel_data.append(total_Padstack)
    for x in range(len(excel_data)):
        data = excel_data[x]
        for y in range(len(data)):
            if type(data[y][1]) == str:
                print(y, data[y])



    # # 单根：net name,line width,max length(差分：space)
    # 数据的取舍

    # active_sheet = wb.sheets.active  # Get the active sheet object
    # Topology_data_list=[]
    # for cell in active_sheet.api.UsedRange.Cells:
    #     if cell.Value== '                               Segments\nRules':
    #         cell_idx = (cell.Row-1,cell.Column)
    #         print('cell_idx',cell_idx)
    #
    #         Table_value=active_sheet.range(cell_idx).current_region.value
    #
    #         for i in range(len(Table_value)):
    #             for j in range(len(Table_value[i])):
    #                 if Table_value[i][j] is not None:
    #                     Topology_data_list.append(Table_value[i][j])
    # p
                    # # if j == len(Table_value[0])-1:   # 表示是完整的一行数据
                    #     Topology_data_list.append('\n')
                    # if Table_value[i][j] == 'W/S/W (mils)': # 通过判断输出topology中net的线宽
                    #     flag = True
    # print(Topology_data_list)

    # wb.close()



# managet_diff_data()
if __name__ == '__main__':
    # getpinnumio()
    get_Topology_Info()
    # get_net_info()
