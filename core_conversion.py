import pandas as pd
import json
import ast

def excel_to_json_format(excel_file):
    """
    Convert Excel sheet with 'rules' and 'connections' sheets to corresponding JSON format examples (used by INA-tool)
    
    Args:
        excel_file: Path to Excel file
        
    Returns:
        rules_json: Rules in JSON format (used by INA-tool)
        connections_json: Connections in JSON format (used by INA-tool)
    """
    # Read Excel sheets
    rules_df = pd.read_excel(excel_file, sheet_name='rules')
    connections_df = pd.read_excel(excel_file, sheet_name='connections')
    
    # Convert rules to expected format
    rules_list = []
    for _, row in rules_df.iterrows():
        # Convert row to dictionary and handle missing values
        row_dict = row.to_dict()
        
        rule_dict = {
            "Id": str(row_dict.get('id', '')),
            "Statement": row_dict.get('Statement', '...'),
            "Statement Type": row_dict.get('Statement Type', ''),
            "Attribute": row_dict.get('Attribute', ''),
            "Deontic": row_dict.get('Deontic', ''),
            "Aim": row_dict.get('Aim', ''),
            "Direct Object": row_dict.get('Direct Object', ''),
            "Type of Direct Object": row_dict.get('Type of Direct Object', ''),
            "Indirect Object": row_dict.get('Indirect Object', ''),
            "Type of Indirect Object": row_dict.get('Type of Indirect Object', ''),
            "Activation Condition": row_dict.get('Activation Condition', ''),
            "Execution Constraint": row_dict.get('Execution Constraint', ''),
            "Or Else": row_dict.get('Or Else', ''),
        }
        
        # Remove None values and replace with empty strings
        for k, v in rule_dict.items():
            if pd.isna(v):
                rule_dict[k] = ''
        rules_list.append(rule_dict)
    
    # Convert connections to expected format
    connections_list = []
    for _, row in connections_df.iterrows():
        conn_dict = {
            "driven_by": row['driven_by'],
            "source_component": row['source_component'],
            "source_statement": str(row['source_statement']),
            "target_component": row['target_component'],
            "target_statement": str(row['target_statement'])
        }
        connections_list.append(conn_dict)
    
    # Convert to JSON strings
    rules_json = json.dumps(rules_list, indent=2)
    connections_json = json.dumps(connections_list, indent=2)
    
    return rules_json, connections_json

def parse_text_file(file_content):
    """
    Parse a text file that might be in JSON format or JavaScript object format (used by INA-tool for rules or connections)
    
    Args:
        file_content: Content of the text file
        
    Returns:
        Parsed data from the input file as a list of dictionaries
    """
    # Clean up the content to make it more compatible with JSON parsing
    # Replace JavaScript property names without quotes with quoted property names
    import re
    
    # Add quotes to unquoted keys
    cleaned_content = re.sub(r'(\s*)(\w+)(\s*):(\s*)', r'\1"\2"\3:\4', file_content)
    
    # Try to parse as JSON
    try:
        data = json.loads(cleaned_content)
        return data
    except json.JSONDecodeError:
        # If that fails, try with ast.literal_eval
        try:
            data = ast.literal_eval(cleaned_content)
            return data
        except (SyntaxError, ValueError):
            # If all parsing fails, raise an error
            raise ValueError("Could not parse the text file format")

def json_to_excel_format(rules_file, connections_file, output_excel):
    """
    Convert JSON-like format (used by INA-tool) files to Excel (used by LA and students)
    
    Args:
        rules_file: Path to rules file
        connections_file: Path to connections file
        output_excel: Path to save the Excel output
    """
    # Read the text files
    with open(rules_file, 'r') as f:
        rules_text = f.read()
    
    with open(connections_file, 'r') as f:
        connections_text = f.read()
    
    # Parse the text files
    rules_data = parse_text_file(rules_text)
    connections_data = parse_text_file(connections_text)
    
    # Convert to DataFrames
    rules_df = pd.DataFrame(rules_data)
    connections_df = pd.DataFrame(connections_data)
    
    # For connections, ensure the statement columns are properly typed
    # Convert string IDs to integers for the Excel format
    if 'source_statement' in connections_df.columns:
        connections_df['source_statement'] = connections_df['source_statement'].astype(int)
    if 'target_statement' in connections_df.columns:
        connections_df['target_statement'] = connections_df['target_statement'].astype(int)
    
    # For rules, convert "Id" to "id" to match Excel format
    if 'Id' in rules_df.columns:
        rules_df = rules_df.rename(columns={'Id': 'id'})
    
    # Save to Excel
    with pd.ExcelWriter(output_excel) as writer:
        rules_df.to_excel(writer, sheet_name='rules', index=False) # rules sheet
        connections_df.to_excel(writer, sheet_name='connections', index=False) # connections sheet

# Additional helper functions
def validate_data(data, format_type):
    """
    Validate that the data has the expected structure
    
    Args:
        data: Data to validate
        format_type: 'rules' or 'connections'
        
    Returns:
        List of validation errors
    """
    errors = []
    
    if not isinstance(data, list):
        errors.append(f"{format_type} data is not a list")
        return errors
    
    if len(data) == 0:
        errors.append(f"{format_type} data is empty")
        return errors
    
    # Check required fields based on type
    if format_type == 'rules':
        required_fields = ['Id', 'Statement Type', 'Attribute', 'Deontic', 'Aim']
        for i, item in enumerate(data):
            for field in required_fields:
                if field not in item:
                    errors.append(f"Rule at index {i} is missing required field '{field}'")
    
    elif format_type == 'connections':
        required_fields = ['driven_by', 'source_component', 'source_statement', 
                           'target_component', 'target_statement']
        for i, item in enumerate(data):
            for field in required_fields:
                if field not in item:
                    errors.append(f"Connection at index {i} is missing required field '{field}'")
    
    return errors

# Example usage
if __name__ == "__main__":
    try:
        # Convert from Excel to JSON
        print("Converting Excel to JSON format...")
        rules_json, connections_json = excel_to_json_format('hychain_wp3_connections_ground_truth.xlsx')
        
        # Save to text files
        with open('rules_output.txt', 'w') as f:
            f.write(rules_json)
        
        with open('connections_output.txt', 'w') as f:
            f.write(connections_json)
        
        print("Excel to JSON conversion complete. Files saved as rules_output.txt and connections_output.txt")
        
        # Convert from text to Excel
        print("\nConverting JSON-like data to Excel format...")
        json_to_excel_format('example_rules.txt', 'example_connections.txt', 'converted_output.xlsx')
        print("JSON to Excel conversion complete. File saved as converted_output.xlsx")
        
    except Exception as e:
        print(f"Error during conversion: {str(e)}")