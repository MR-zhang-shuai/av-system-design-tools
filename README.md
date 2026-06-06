# 音视频系统设计工具集

Audio-video integration engineering design toolkit for conference centers, lecture halls, multi-function rooms, and theaters.

## 项目结构

| 目录 | 说明 |
|------|------|
| scripts/ | Python 工具脚本 |
| templates/ | 模板文件 (Excel/Word) |
| docs/ | 设计文档与规范 |
| specs/ | 设备参数库 |

## 快速开始

```bash
pip install openpyxl
python scripts/device_list_generator.py --config scripts/sample_config.json --output 设备清单.xlsx
```

## 功能模块

1. **设备清单生成器** - 从 JSON 配置文件生成标准格式设备清单 Excel
2. **品牌价格统计卡** - HTML 交互工具，按品牌维度聚合报价
3. **机柜布置图生成** - 根据设备清单自动生成机柜布置 HTML 预览 + DXF