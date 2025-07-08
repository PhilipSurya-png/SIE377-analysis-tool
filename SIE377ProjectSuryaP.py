import pandas as pd
from tkinter import Tk, filedialog

def process_file(file_path):
    # Read the CSV file into a DataFrame
    data = pd.read_csv(file_path)

    # Create a new Excel file
    output_file = "processed_hospital_staffing_data.xlsx"
    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        # Write the "Original Data" worksheet without the index
        data.to_excel(writer, sheet_name="Original Data", index=False)

        # Calculate descriptive statistics
        stats = {
            "Statistic": ["Mean", "Min", "Max"],
            "Productive Hours": [
                data["Productive Hours"].mean(),
                data["Productive Hours"].min(),
                data["Productive Hours"].max(),
            ],
            "Productive Hours per Adjusted Patient Day": [
                data["Productive Hours per Adjusted Patient Day"].mean(),
                data["Productive Hours per Adjusted Patient Day"].min(),
                data["Productive Hours per Adjusted Patient Day"].max(),
            ],
        }
        stats_df = pd.DataFrame(stats)

        # Write the "Statistics" worksheet
        stats_df.to_excel(writer, sheet_name="Statistics", index=False, startrow=1)

        # Format the headers
        worksheet = writer.sheets["Statistics"]
        bold_format = writer.book.add_format({"bold": True})

        worksheet.write(0, 0, "Statistic", bold_format)
        worksheet.write(0, 1, "Productive Hours", bold_format)
        worksheet.write(0, 2, "Productive Hours per Adjusted Patient Day", bold_format)

    print(f"File processed and saved as {output_file}")

def main():
    # Display file dialog for user to select the CSV file
    root = Tk()
    root.withdraw()  # Hide the main tkinter window
    file_path = filedialog.askopenfilename(
        title="Please select the hospital staffing data file:",
        filetypes=[("CSV files", "*.csv")]
    )

    if not file_path:
        print("No file selected. Exiting...")
        return

    try:
        process_file(file_path)
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()

