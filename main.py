from load_files import load_data
from generate_reports_deliverable1 import create_excel_report
from generate_reports_deliverable1 import create_excel_report_graph
from generate_reports_deliverable2 import create_deliverable_2

def main():
    print("Starting the script...")
    data_path = '/Users/andreacontreras/Downloads/FP&ADeveloperAssessment/Transactions'
    print(f"Loading data from {data_path}")
    final_df = load_data(data_path)
    print("Data loaded successfully, now creating Excel report...")
    create_excel_report(final_df)
    print("Excel report created successfully!")
    create_deliverable_2(final_df)
    print("Deliverable 2 completed.")

if __name__ == "__main__":
    main()