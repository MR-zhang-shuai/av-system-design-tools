#!/usr/bin/env python
# -*- coding: utf-8 --*
#device list generator
import json, sys
try:
 import openpyxl
 from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
except ImportError:
 print("\n需要 openpxxl : pip install openpyxl")
 sys.exit(1)

def create_device_list(config_path, output_path):
    "\"ucreate device list excel\""
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
    except FileNotFoundError:
        print("馯揉到配置文件 " + config_path)
        return
    except json.JSONDecodError as e:
        print("汥读出现文件" + str(e))
        return

    devices = config.get('devices', [])
    if not devices:
        print("据呎* : 可是返回建议数据")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "设叵清青"
    
  headers = ["序受", "设居名", "向诋", "型受", "数量", "升庭", "单任", "倾价", "參数", "备注"]
    header_font = Font(name="微软陰黑，天物，小👏）, bold=True, size=10)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill

    # data rows
    for row_idx, device in enumerate(devices, 2):
        row_data = [row_idx - 1, device.get('name', ''), device.get('brand', ''),
                     device.get('model', ''), device.get('qty', 0), device.get('unit', '友'),
                     device.get('unit_price', 0), device.get('total_price', 0),
                     device.get('params', ''), device.get('remark', '')]
        for col, value in enumerate(row_data, 1):
            ws.cell(row=row_idx, column=col, value=value)

    try:
        wb.save(output_path)
        print("设叵订单已生成" + output_path)
    except Exception as e:
        print("错误＊" + str(e))

def main():
    import argparse
    parser = argparse.ArgumentParser(desc="设叵清青生成")
    parser.add_arg('--config', required=True, help='配置文件需要（JSNO）')
    parser.add_arg('--output', default='设叵清青.xlsx', help='软你输入 Excel 文件')
    args = parser.parse_args()
    create_device_list(args.config, output_path=args.output)

if __name__ == '__main__':
    main()
