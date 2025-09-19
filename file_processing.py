import pandas as pd
import re
import json
class ExcelProcessing:
    def __init__(self, file_path):
        self.file_path = file_path
        try: 
            df = pd.read_excel(file_path, sheet_name=None, header=None)
        except Exception as e:
            print(f"Error reading excel file: {e}")
            raise e
        self.df = df

    def read_file(self):
        with open(self.file_path, 'r') as file:
            return file.read()

    def write_file(self, data):
        with open(self.file_path, 'w') as file:
            file.write(data)

    def handle_rows(self,sheet_name,df_sheet):
        
        title = df_sheet.iat[0,0]
        account_id = None
        real_description = None
        status = None
        total_number_of_resources_processed = None
        number_of_resources_flagged = None
        number_of_suppressed_resources = None
        source = None
        alert_criteria = []
        recommended_action = None
        additional_resources = None
        
        if "AWS Account ID:" in df_sheet.iat[1,0]:
            account_id = df_sheet.iat[1,0].partition(':')[2].strip()
        if "Description:" in df_sheet.iat[2,0]:
            description_text = df_sheet.iat[2, 0]

            # Tách phần mô tả chính
            real_description_match = re.split(r"\n\n(Source|Alert Criteria|Recommended Action|Additional Resources)", description_text, maxsplit=1)
            real_description = real_description_match[0].partition(":")[2].strip()

            # Chuẩn bị biến kết quả
            source = recommended_action = additional_resources = None
            alert_criteria = []

            # Regex pattern tìm từng section
            pattern = re.compile(
                r"(Source|Alert Criteria|Recommended Action|Additional Resources)\s*(.*?)\s*(?=(Source|Alert Criteria|Recommended Action|Additional Resources|$))",
                re.DOTALL,
            )

            for match in pattern.finditer(description_text):
                section, content = match.group(1), match.group(2).strip()
                if section == "Source":
                    source = content.strip()
                elif section == "Alert Criteria":
                    alert_criteria = self.parse_alert_criteria(content.strip())
                elif section == "Recommended Action":
                    recommended_action = content.strip()
                elif section == "Additional Resources":
                    additional_resources = content.strip()

        if "Status:" in df_sheet.iat[3,0]:
            status = df_sheet.iat[3,0].partition(':')[2].strip()
        if "Total number of resources processed:" in df_sheet.iat[5,1]:
            total_number_of_resources_processed = df_sheet.iat[5,1].partition(':')[2].strip()
        if "Number of resources flagged:" in df_sheet.iat[6,1]:
            number_of_resources_flagged = df_sheet.iat[6,1].partition(':')[2].strip()
        if "Number of suppressed resources:" in df_sheet.iat[7,1]:
            number_of_suppressed_resources = df_sheet.iat[7,1].partition(':')[2].strip()

        if status == "not_available":
            return {
                "check_title": title,
                "account_id": account_id,
                "description": None,
                "status": "not_available",
                "total_number_of_resources_processed": None,
                "number_of_resources_flagged": None,
                "number_of_suppressed_resources": None,
                "source": None,
                "alert_criteria": None,
                "recommended_action": None,
                "additional_resources": None
            }
        return {
            "check_title": title,
            "account_id": account_id,
            "description": real_description,
            "status": status,
            "total_number_of_resources_processed": total_number_of_resources_processed,
            "number_of_resources_flagged": number_of_resources_flagged,
            "number_of_suppressed_resources": number_of_suppressed_resources,
            "source": source,
            "alert_criteria": alert_criteria,
            "recommended_action": recommended_action,
            "additional_resources": additional_resources
        }

    def parse_alert_criteria(self, alert_criteria_text):
        """
        Parse alert criteria text into structured JSON format
        Input: "Red: description\nYellow: description\nGreen: description"
        Output: [{"level": "Red", "description": "description"}, ...]
        """
        if not alert_criteria_text or alert_criteria_text.strip() == "":
            return []
        
        criteria_list = []
        lines = alert_criteria_text.strip().split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # Tìm pattern "Color:" ở đầu dòng
            color_match = re.match(r'^(Red|Yellow|Green):\s*(.+)$', line, re.IGNORECASE)
            if color_match:
                level = color_match.group(1).capitalize()
                description = color_match.group(2).strip()
                criteria_list.append({
                    "level": level,
                    "description": description
                })
        
        return criteria_list

    def excel_to_json(self):
        output = []
        for sheet_name, df_sheet in self.df.items():
            output.append(self.handle_rows(sheet_name,df_sheet))  
        return output

    def flatten_json_data(self, data):
        """
        Flatten JSON data - tách alert_criteria thành các records riêng biệt
        
        Args:
            data: List of dictionaries từ excel_to_json()
        
        Returns:
            List of flattened dictionaries
        """
        flattened_data = []
        
        for record in data:
            if record.get('alert_criteria'):
                # Tách alert_criteria thành các record riêng biệt
                base_record = record.copy()
                alert_criteria_list = base_record.pop('alert_criteria', [])
                
                if alert_criteria_list:
                    # Tạo record cho mỗi alert criterion
                    for criterion in alert_criteria_list:
                        flattened_record = base_record.copy()
                        flattened_record['alert_level'] = criterion.get('level')
                        flattened_record['alert_description'] = criterion.get('description')
                        flattened_data.append(flattened_record)
                else:
                    # Nếu không có alert_criteria thì giữ nguyên record
                    base_record['alert_level'] = None
                    base_record['alert_description'] = None
                    flattened_data.append(base_record)
            else:
                # Nếu không có alert_criteria thì thêm fields null và giữ nguyên
                record_copy = record.copy()
                record_copy['alert_level'] = None
                record_copy['alert_description'] = None
                flattened_data.append(record_copy)
        
        return flattened_data

    def convert_to_ndjson(self, data, output_file=None, flatten_alert_criteria=False):
        """
        Convert data to NDJSON format (Newline Delimited JSON)
        
        Args:
            data: List of dictionaries từ excel_to_json()
            output_file: Tên file để save, nếu None thì return string
            flatten_alert_criteria: Nếu True, tách alert_criteria thành các dòng riêng biệt
        
        Returns:
            String NDJSON hoặc save to file
        """
        if flatten_alert_criteria:
            # Sử dụng function flatten_json_data
            data_to_process = self.flatten_json_data(data)
        else:
            # Giữ nguyên format gốc
            data_to_process = data
        
        ndjson_lines = []
        for record in data_to_process:
            ndjson_lines.append(json.dumps(record, ensure_ascii=False))
        
        ndjson_content = '\n'.join(ndjson_lines)
        
        if output_file:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(ndjson_content)
            print(f"NDJSON data saved to {output_file}")
            return output_file
        else:
            return ndjson_content

    def save_json(self, data, file_path):
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)



class JsonProcessing:
    def __init__(self, file_path):
        self.file_path = file_path
        self.data = self.read_json()

    def read_json(self):
        with open(self.file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def json_to_excel(self, output_file):
        df = pd.read_json(output_file, lines=True)
        df.to_excel(output_file, index=False)


