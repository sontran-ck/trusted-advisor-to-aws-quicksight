from file_processing import ExcelProcessing, JsonProcessing
from openpyxl import load_workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.fill import SolidColorFillProperties
from openpyxl.drawing.colors import ColorChoice
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image
import os
if __name__ == "__main__":
    try:
        EXCEL_PROCESSING = ExcelProcessing()
    except Exception as e:
        print(f"Error: {e}")
        raise e
    
    extracted_json_data = EXCEL_PROCESSING.excel_to_json("all.xlsx")
    
    EXCEL_PROCESSING.save_json(extracted_json_data, "all.json")
    try:
        JSON_PROCESSING = JsonProcessing()
    except Exception as e:
        print(f"Error: {e}")
        raise e
    parse_data = JSON_PROCESSING.get_parse_data(extracted_json_data)

    insert_data_detail_alert_criteria = JSON_PROCESSING.insert_data_detail_alert_criteria(extracted_json_data, null_value=None)

    summary_alert_criteria = JSON_PROCESSING.summary_alert_criteria(parse_data)

    
    JSON_PROCESSING.save_excel(insert_data_detail_alert_criteria, "new_all.xlsx")

    JSON_PROCESSING.create_chart_alert_criteria(summary_alert_criteria, "new_all.xlsx", "new_all_2.xlsx")

    EXCEL_PROCESSING.save_json(insert_data_detail_alert_criteria, "new_all_2.json")
    

    
