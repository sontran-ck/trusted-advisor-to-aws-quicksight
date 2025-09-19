from file_processing import ExcelProcessing, JsonProcessing
                
if __name__ == "__main__":
    try:
        file_processing = ExcelProcessing("all.xlsx")
    except Exception as e:
        print(f"Error: {e}")
        raise e
    output = file_processing.excel_to_json()
    file_processing.save_json(output, "all.json")
    json_processing = JsonProcessing("all.json")
    json_processing.json_to_excel("new_all.xlsx")
