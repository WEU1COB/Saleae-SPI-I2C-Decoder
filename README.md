# Saleae-SPI-I2C-Decoder
The Saleae SPI/I2C Decoder Tool is a companion utility for Saleae Logic 2. It automates the decoding and formatting of captured SPI and I2C transactions into human-readable text and optional Excel reports. This tool is built using the Saleae Automation API (gRPC) and works alongside Saleae Logic 2 software.

A lightweight GUI tool to decode `.sal` captures from Saleae Logic 2 using gRPC Automation API.

---

##  Features

- Decode SPI and I2C captures
- Modify analyzer settings before decoding
- Logs analyzer config and decoding status
- CLI-free, no programming needed
- Built with Python, packaged as standalone `.exe`

---

##  How to Use

1. **Start Logic 2 with Automation API enabled.**
2. **Open the `SaleaeDecoderTool.exe`**
3. **Select protocol (SPI/I2C).**
4. **Browse and select your `.sal` file.**
5. **Edit analyzer settings if needed.**
6. **Click `Start Decode`.**

 Output will be saved to:
- `C:/Temp/spi_formatted_output.txt` or
- `C:/Temp/i2c_formatted_output.txt`

---

##  Requirements (For Source Use)

- Python 3.8+
- Saleae Logic 2 with Automation API
- `saleae` gRPC API
- `protobuf`, `pandas`, `grpcio`, `tkinter`

```bash
pip install -r requirements.txt
