SAPGUI Auto VBA Method

SAPGUI Auto VBA Method is a tool designed for automating interactions with SAPGUI using VBA (Visual Basic for Applications). This method was developed by Parmenas Santos.

Features:
Automatically connects to SAPGUI.
Allows opening or connecting to a specific session.
Supports various SAP transactions, such as ECC, EWM, etc.
Manages SAP sessions and performs automated actions.

Installation:
Clone this repository to your local machine using the following command:
git clone https://github.com/parmenassantos/SAPGUI-Auto-VBA-Method.git

Usage:
Manually open SAPGUI or let the tool open it automatically.
Copy and paste the provided VBA code into your Excel macro or other VBA environment.
Customize the uSession function call as needed for your specific transactions.
Execute the macro to interact automatically with SAPGUI.

Example Usage:
    Dim Session As Object
    Set Session = uSession("LOGON_SAP", "TRANSACTION")

    If Session Is Nothing Then
        MsgBox "Session not found.", vbInformation, "SCRIPT: Error Onto Session Script"
        Exit Sub
    End If

References - VBA Projects Used:
Visual Basic For Applications
Microsoft Excel 16.0 Object Library
OLE Automation

Contribution:
Contributions are welcome! Feel free to open issues or submit pull requests.

License:
This project is licensed under the MIT License - see the LICENSE file for details.

Contact:
For more information, contact Parmenas Santos: parmenassantos@gmail.com

Enjoy automating your SAPGUI interactions!
