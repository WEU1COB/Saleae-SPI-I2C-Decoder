import grpc
import pandas as pd
import os
from saleae.grpc import saleae_pb2, saleae_pb2_grpc

# Default I2C analyzer settings
# Maps SDA and SCL lines to specific Saleae channels
I2C_SETTINGS = {
    "SDA": saleae_pb2.AnalyzerSettingValue(int64_value=2),
    "SCL": saleae_pb2.AnalyzerSettingValue(int64_value=1),
}

def decode_i2c(
    sal_file_path,
    export_csv_path="C:/Temp/i2c_decoded.csv",
    output_txt_path="C:/Temp/i2c_formatted_output.txt",
    output_xlsx_path="C:/Temp/i2c_formatted_output.xlsx",
    settings_override=None,
    progress_callback=None,
    export_excel=False
):
    """
    Decode I2C frames from a Saleae capture file and export them to CSV, TXT, and optionally Excel.

    Parameters:
        sal_file_path (str): Path to the .sal capture file
        export_csv_path (str): Path to save raw CSV export from Saleae
        output_txt_path (str): Path to save formatted TXT output
        output_xlsx_path (str): Path to save formatted Excel output
        settings_override (dict): Optional analyzer settings override
        progress_callback (function): Optional function to report progress (0-100)
        export_excel (bool): Whether to create Excel output
    Returns:
        Tuple of TXT and Excel file paths (Excel path may be None)
    """
    
    # Use provided settings override if available
    settings = settings_override or I2C_SETTINGS

    # Ensure the output directories exist
    os.makedirs(os.path.dirname(export_csv_path), exist_ok=True)
    os.makedirs(os.path.dirname(output_txt_path), exist_ok=True)
    if export_excel:
        os.makedirs(os.path.dirname(output_xlsx_path), exist_ok=True)

    # Connect to Saleae gRPC server
    channel = grpc.insecure_channel("localhost:10430")
    stub = saleae_pb2_grpc.ManagerStub(channel)

    # Initial progress
    if progress_callback:
        progress_callback(5)

    # Load the capture file
    capture_reply = stub.LoadCapture(saleae_pb2.LoadCaptureRequest(filepath=sal_file_path))
    capture_id = capture_reply.capture_info.capture_id

    if progress_callback:
        progress_callback(15)

    # Add an I2C analyzer to the capture
    add_analyzer_reply = stub.AddAnalyzer(saleae_pb2.AddAnalyzerRequest(
        capture_id=capture_id,
        analyzer_name="I2C",
        analyzer_label="MyI2C",
        settings=settings
    ))
    analyzer_id = add_analyzer_reply.analyzer_id

    if progress_callback:
        progress_callback(25)

    # Export the analyzer data to CSV
    export_request = saleae_pb2.ExportDataTableCsvRequest(
        capture_id=capture_id,
        filepath=export_csv_path,
        analyzers=[
            saleae_pb2.DataTableAnalyzerConfiguration(
                analyzer_id=analyzer_id,
                radix_type=saleae_pb2.RadixType.RADIX_TYPE_HEXADECIMAL
            )
        ],
        iso8601_timestamp=False  # Use float timestamps instead of ISO strings
    )
    stub.ExportDataTableCsv(export_request)

    if progress_callback:
        progress_callback(50)

    # Read exported CSV into a DataFrame
    df = pd.read_csv(export_csv_path)

    # Ensure all expected columns exist
    for col in ["data", "type", "address", "read"]:
        if col not in df.columns:
            df[col] = ""

    # Convert columns to string type to avoid issues
    df["data"] = df["data"].fillna("").astype(str)
    df["type"] = df["type"].fillna("").astype(str).str.lower()
    df["address"] = df["address"].fillna("").astype(str)
    df["read"] = df["read"].fillna("").astype(str)

    if progress_callback:
        progress_callback(60)

    # Prepare containers for formatted output
    output_lines = []   # Lines for TXT output
    excel_rows = []     # Rows for Excel output

    # Variables for building each I2C frame
    frame_count = 0
    frame_start_time = None
    addr_info = ""
    direction = ""
    data_bytes = []

    # Iterate over each row in the CSV
    for _, row in df.iterrows():
        t = row["type"]         # Event type: start, address, data, stop
        ts = row["start_time"]  # Timestamp
        value = row["data"]     # Data byte

        if t == "start":
            # New I2C frame started
            frame_start_time = ts
            addr_info = ""
            data_bytes = []
            direction = ""

        elif t == "address":
            # Capture address and read/write direction
            read_val = str(row["read"]).strip().lower()
            if read_val == "true":
                direction = "Read"
            elif read_val == "false":
                direction = "Write"
            else:
                # Fallback in case Saleae encoding is unusual
                direction = "Read" if "read" in value.lower() else "Write"
            addr_info = row["address"]

        elif t == "data":
            # Append data bytes for this frame
            data_bytes.append(value)

        elif t == "stop" and frame_start_time is not None:
            # End of the frame, create formatted outputs
            frame_count += 1
            end_time = ts

            # TXT line
            frame_line = (
                f"[FRAME {frame_count}] {frame_start_time:.9f}s START "
                f"ADDR: {addr_info} ({direction}), DATA: {' '.join(data_bytes)} STOP"
            )
            output_lines.append(frame_line)

            # Excel row
            excel_rows.append([
                frame_count,
                f"{frame_start_time:.9f}s - {end_time:.9f}s",
                addr_info,
                direction,
                " ".join(data_bytes)
            ])

            # Reset for next frame
            frame_start_time = None

        # Update progress dynamically
        if progress_callback:
            progress_callback(60 + int((frame_count / max(len(df), 1)) * 35))

    # Write formatted TXT output
    with open(output_txt_path, "w") as f:
        f.write("\n".join(output_lines))

    # Write Excel output if requested
    if export_excel:
        df_out = pd.DataFrame(excel_rows, columns=["Frame", "Timestamp Range", "Address", "Direction", "Data"])
        df_out.to_excel(output_xlsx_path, index=False)

    # Final progress
    if progress_callback:
        progress_callback(100)

    # Return output paths
    return (output_txt_path, output_xlsx_path if export_excel else None)
