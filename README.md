<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body>

<h1>SAPGUI Auto VBA Method</h1>

<p>SAPGUI Auto VBA Method is a tool designed for automating interactions with 
SAPGUI using VBA (Visual Basic for Applications). This method was developed by Parmenas Santos.</p>

<h2>Features:</h2>
<ul>
    <li>Automatically connects to SAPGUI.</li>
    <li>Allows opening or connecting to a specific session.</li>
    <li>Supports various SAP transactions, such as ECC, EWM, etc.</li>
    <li>Manages SAP sessions and performs automated actions.</li>
</ul>

<h2>Installation:</h2>
<p>Clone this repository to your local machine using the following command:</p>

<pre><code>git clone https://github.com/parmenassantos/SAPGUI-Auto-VBA-Method.git</code></pre>

<h2>Usage:</h2>
<ol>
    <li>Manually open SAPGUI or let the tool open it automatically.</li>
    <li>Copy and paste the provided VBA code into your Excel macro or other VBA environment.</li>
    <li>Customize the uSession function call as needed for your specific transactions.</li>
    <li>Execute the macro to interact automatically with SAPGUI.</li>
</ol>

<h2>Example Usage:</h2>
<pre><code>
Dim Session As Object
Set Session = uSession("LOGON_SAP", "TRANSACTION")

If Session Is Nothing Then
    MsgBox "Session not found.", vbInformation, "SCRIPT: Error Onto Session Script"
    Exit Sub
End If
</code></pre>

<h2>References - VBA Projects Used:</h2>
<ul>
    <li>Visual Basic For Applications</li>
    <li>Microsoft Excel 16.0 Object Library</li>
    <li>OLE Automation</li>
</ul>

<h2>Contribution:</h2>
<p>Contributions are welcome! Feel free to open issues or submit pull requests.</p>

<h2>License:</h2>
<p>This project is licensed under the MIT License - see the LICENSE file for details.</p>

<h2>Contact:</h2>
<p>For more information, contact Parmenas Santos: parmenassantos@gmail.com</p>

<a href="https://github.com/parmenassantos/SAPGUI-Auto-VBA-Method.git">GitHub Repository</a>

<p>Enjoy automating your SAPGUI interactions!</p>

</body>
</html>
