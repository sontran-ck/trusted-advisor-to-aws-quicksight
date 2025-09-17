from file_processing import ExcelProcessing
                
if __name__ == "__main__":
    try:
        file_processing = ExcelProcessing("all.xlsx")
    except Exception as e:
        print(f"Error: {e}")
        raise e
    output = file_processing.excel_to_json()
    flattened_output = file_processing.flatten_json_data(output)
    file_processing.save_json(flattened_output, "all.json")
