import grpc
import pandas as pd
import os
from saleae.grpc import saleae_pb2, saleae_pb2_grpc

# Default SPI analyzer settings
# Configure channels, bits per transfer, clock polarity/phase, and enable line
SPI_SETTINGS = {
    "Clock": saleae_pb2.AnalyzerSettingValue(int64_value=2),
    "MISO": saleae_pb2.AnalyzerSettingValue(int64_value=0),
    "MOSI": saleae_pb2.AnalyzerSettingValue(int64_value=1),
    "Enable": saleae_pb2.AnalyzerSettingValue(int64_value=3),
    "Bits per Transfer": saleae_pb2.AnalyzerSettingValue(int64_value=8),
    "Significant Bit": saleae_pb2.AnalyzerSettingValue(
        string_value="Most Significant Bit First (Standard)"
    ),
    "Clock State": saleae_pb2.AnalyzerSettingValue(
        string_value="Clock is Low when inactive (CPOL = 0)"
    ),
    "Clock Phase": saleae_pb2.AnalyzerSettingValue(
        string_value="Data is Valid on Clock Leading Edge (CPHA = 0)"
    ),
    "Enable Line": saleae_pb2.AnalyzerSettingValue(
        string_value="Enable line is Active Low (Standard)"
    ),
}


def decode_spi(
    sal_file_path,
    export_csv_path="C:/Temp/spi_decoded.csv",
    output_txt_path="C:/Temp/spi_formatted_output.txt",
    output_xlsx_path="C:/Temp/spi_formatted_output.xlsx",
    settings_override=None,
    progress_callback=None,
    export_excel=False,
):
    """
    Decode SPI frames from a Saleae capture file and export formatted results.

    Parameters:
        sal_file_path (str): Path to the Saleae .sal capture file
        export_csv_path (str): Path for intermediate CSV export
        output_txt_path (str): Path for formatted TXT output
        output_xlsx_path (str): Path for formatted Excel output
        settings_override (dict): Optional analyzer settings override
        progress_callback (function): Optional function to report progress (0-100)
        export_excel (bool): Whether to create Excel output
    Returns:
        Tuple of TXT and Excel file paths (Excel may be None)
    """

    # Use provided settings override or default SPI_SETTINGS
    settings = settings_override or SPI_SETTINGS

    # Ensure the output directory exists
    os.makedirs(os.path.dirname(export_csv_path), exist_ok=True)

    # Connect to Saleae gRPC server
    channel = grpc.insecure_channel("localhost:10430")
    stub = saleae_pb2_grpc.ManagerStub(channel)

    # Initial progress
    if progress_callback: 
        progress_callback(5)

    # Load the capture file
    capture_reply = stub.LoadCapture(
        saleae_pb2.LoadCaptureRequest(filepath=sal_file_path)
    )
    capture_id = capture_reply.capture_info.capture_id

    if progress_callback: 
        progress_callback(15)

    # Add an SPI analyzer to the capture
    add_analyzer_reply = stub.AddAnalyzer(
        saleae_pb2.AddAnalyzerRequest(
            capture_id=capture_id,
            analyzer_name="SPI",
            analyzer_label="MySPI",
            settings=settings,
        )
    )
    analyzer_id = add_analyzer_reply.analyzer_id

    if progress_callback: 
        progress_callback(25)

    # Export the analyzer results to CSV
    export_request = saleae_pb2.ExportDataTableCsvRequest(
        capture_id=capture_id,
        filepath=export_csv_path,
        analyzers=[
            saleae_pb2.DataTableAnalyzerConfiguration(
                analyzer_id=analyzer_id,
                radix_type=saleae_pb2.RadixType.RADIX_TYPE_HEXADECIMAL,
            )
        ],
        iso8601_timestamp=False,
    )
    stub.ExportDataTableCsv(export_request)

    if progress_callback: 
        progress_callback(50)

    # Read CSV into a DataFrame
    df = pd.read_csv(export_csv_path)

    # Clean up MOSI/MISO columns: remove "0x" prefix, fill missing, pad with zeros
    df["mosi"] = (
        df["mosi"].fillna("").astype(str).str.replace("0x", "", case=False).str.zfill(2)
    )
    df["miso"] = (
        df["miso"].fillna("").astype(str).str.replace("0x", "", case=False).str.zfill(2)
    )
    df["type"] = df["type"].fillna("").astype(str).str.lower()

    if progress_callback: 
        progress_callback(60)

    # Identify indices where Enable events occur
    enable_indices = df[df["type"] == "enable"].index.tolist()
    enable_indices.append(len(df))  # Add end for final frame processing

    # Helper function to format bytes in 8+8 style per line
    def format_bytes(byte_list):
        lines = []
        for i in range(0, len(byte_list), 16):
            chunk = byte_list[i : i + 16]
            left = " ".join(chunk[:8])
            right = " ".join(chunk[8:16])
            lines.append(f"{left:<23} {right}".rstrip())
        return "\n".join(lines)

    output_lines = []  # Lines for TXT output
    excel_rows = []    # Rows for Excel output
    total_frames = len(enable_indices) - 1

    # Process each SPI frame based on Enable events
    for idx in range(total_frames):
        start = enable_indices[idx] + 1
        end = enable_indices[idx + 1]
        frame = df.iloc[start:end]
        frame = frame[frame["type"] == "result"]  # Only keep actual data results

        if frame.empty:
            continue

        # Extract MOSI and MISO bytes
        mosi_bytes = frame["mosi"].tolist()
        miso_bytes = frame["miso"].tolist()

        # Frame timestamps
        start_time = frame["start_time"].iloc[0]
        end_time = frame["start_time"].iloc[-1] + frame["duration"].iloc[-1]

        # Append to TXT output
        output_lines.append(
            f"Frame {idx + 1} | Timestamp Range: {start_time:.9f}s - {end_time:.9f}s"
        )
        output_lines.append("MOSI:")
        output_lines.append(format_bytes(mosi_bytes))
        output_lines.append("MISO:")
        output_lines.append(format_bytes(miso_bytes))
        output_lines.append("=" * 40 + "\n")

        # Prepare Excel row
        excel_rows.append(
            {
                "Frame": idx + 1,
                "Start Time (s)": f"{start_time:.9f}",
                "End Time (s)": f"{end_time:.9f}",
                "MOSI Bytes": " ".join(mosi_bytes),
                "MISO Bytes": " ".join(miso_bytes),
            }
        )

        # Update progress
        if progress_callback:
            progress_callback(60 + int((idx + 1) / total_frames * 35))

    # Write TXT output
    with open(output_txt_path, "w") as f:
        f.write("\n".join(output_lines))

    # Write Excel output if requested
    if export_excel and excel_rows:
        pd.DataFrame(excel_rows).to_excel(output_xlsx_path, index=False)

    # Final progress
    if progress_callback:
        progress_callback(100)

    return output_txt_path, (output_xlsx_path if export_excel else None)
