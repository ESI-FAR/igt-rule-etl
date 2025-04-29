#!/usr/bin/env python3
import argparse
import os
from core_conversion import excel_to_json_format, json_to_excel_format

def main():
    parser = argparse.ArgumentParser(description='Convert between Excel and JSON formats for rules and connections')
    subparsers = parser.add_subparsers(dest='command', help='Command to execute')
    
    # Excel to JSON command
    excel_to_json = subparsers.add_parser('excel_to_json', help='Convert Excel to JSON format')
    excel_to_json.add_argument('excel_file', help='Path to the Excel file')
    excel_to_json.add_argument('--rules_output', default='rules_output.txt', help='Path to save rules JSON output')
    excel_to_json.add_argument('--connections_output', default='connections_output.txt', help='Path to save connections JSON output')
    
    # JSON to Excel command
    json_to_excel = subparsers.add_parser('json_to_excel', help='Convert JSON files to Excel format')
    json_to_excel.add_argument('rules_file', help='Path to the rules JSON file')
    json_to_excel.add_argument('connections_file', help='Path to the connections JSON file')
    json_to_excel.add_argument('--excel_output', default='converted_output.xlsx', help='Path to save Excel output')
    
    args = parser.parse_args()
    
    if args.command == 'excel_to_json':
        try:
            if not os.path.exists(args.excel_file):
                print(f"Error: Excel file '{args.excel_file}' not found")
                return 1
                
            print(f"Converting '{args.excel_file}' to JSON format...")
            rules_json, connections_json = excel_to_json_format(args.excel_file)
            
            with open(args.rules_output, 'w') as f:
                f.write(rules_json)
            
            with open(args.connections_output, 'w') as f:
                f.write(connections_json)
                
            print(f"Conversion complete. Rules saved to '{args.rules_output}' and connections saved to '{args.connections_output}'")
            return 0
            
        except Exception as e:
            print(f"Error during Excel to JSON conversion: {str(e)}")
            return 1
            
    elif args.command == 'json_to_excel':
        try:
            if not os.path.exists(args.rules_file):
                print(f"Error: Rules file '{args.rules_file}' not found")
                return 1
                
            if not os.path.exists(args.connections_file):
                print(f"Error: Connections file '{args.connections_file}' not found")
                return 1
                
            print(f"Converting '{args.rules_file}' and '{args.connections_file}' to Excel format...")
            json_to_excel_format(args.rules_file, args.connections_file, args.excel_output)
            
            print(f"Conversion complete. Excel file saved as '{args.excel_output}'")
            return 0
            
        except Exception as e:
            print(f"Error during JSON to Excel conversion: {str(e)}")
            return 1
    else:
        parser.print_help()
        return 1

if __name__ == "__main__":
    exit(main())