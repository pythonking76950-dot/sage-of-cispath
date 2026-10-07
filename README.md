# app.py
from flask import Flask, render_template_string, request, send_file
import csv
import io
import openpyxl

app = Flask(__name__)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ta">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sage of Cispath - Data Entry & Conversion Tool</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { 
            font-family: Arial, sans-serif; 
            background-color: #f8f9fa; 
            display: flex; 
            flex-direction: column; 
            height: 100vh; 
            color: #333;
        }
        
        /* Ad Banner Section */
        .ad-banner { 
            background: #2c3e50; 
            color: #ecf0f1; 
            text-align: center; 
            padding: 10px; 
            font-size: 14px; 
            font-weight: bold;
            border-bottom: 2px solid #1a252f;
            flex-shrink: 0;
        }

        /* Header Info */
        .app-header {
            background: #ffffff;
            padding: 10px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid #ddd;
            flex-shrink: 0;
        }
        .app-title { font-size: 18px; font-weight: bold; color: #2c3e50; }
        .app-subtitle { font-size: 12px; color: #7f8c8d; }

        /* Table Container */
        .table-container { 
            flex: 1; 
            overflow: auto; 
            background: white; 
            position: relative;
        }
        table { 
            border-collapse: collapse; 
            width: 100%; 
            table-layout: fixed;
        }
        th, td { 
            border: 1px solid #dcdcdc; 
            padding: 4px 8px; 
            text-align: center; 
            min-width: 120px; 
            height: 35px;
        }
        th { 
            background-color: #34495e; 
            color: white; 
            position: sticky; 
            top: 0; 
            z-index: 10;
            font-size: 14px;
        }
        th.row-num-header, td.row-num {
            width: 50px;
            min-width: 50px;
            background-color: #eaeded;
            color: #555;
            font-weight: bold;
            position: sticky;
            left: 0;
            z-index: 5;
        }
        td input { 
            width: 100%; 
            height: 100%;
            border: none; 
            outline: none; 
            font-size: 14px; 
            text-align: left; 
            background: transparent;
            padding: 0 4px;
        }
        td input:focus {
            background: #e8f4fd;
        }

        /* Control Panel */
        .control-panel { 
            background: #ffffff; 
            padding: 12px 20px; 
            display: flex; 
            justify-content: space-between; 
            align-items: center; 
            box-shadow: 0 -2px 10px rgba(0,0,0,0.05);
            flex-shrink: 0;
            border-top: 1px solid #ddd;
            flex-wrap: wrap;
            gap: 10px;
        }
        .btn-group { display: flex; gap: 10px; align-items: center; }
        button { 
            background: #27ae60; 
            color: white; 
            border: none; 
            padding: 10px 18px; 
            font-size: 14px; 
            font-weight: bold;
            cursor: pointer; 
            border-radius: 4px; 
            transition: background 0.2s;
        }
        button:hover { background: #219653; }
        .btn-secondary { background: #2980b9; }
        .btn-secondary:hover { background: #2471a3; }
        .format-select {
            padding: 9px 12px;
            font-size: 14px;
            border-radius: 4px;
            border: 1px solid #ccc;
            background: white;
            font-weight: bold;
        }
    </style>
</head>
<body>

    <!-- விளம்பரப் பகுதி (Ad Banner Space) -->
    <div class="ad-banner">
        📢 Sage of Cispath - Free Data Conversion & Entry Tool (Advertisement Space)
    </div>

    <!-- தலைப்பு பகுதி -->
    <div class="app-header">
        <div>
            <div class="app-title">Sage of Cispath</div>
            <div class="app-subtitle">A-Z Dynamic Columns & Infinite Rows Spreadsheet</div>
        </div>
        <div id="statsInfo" style="font-size: 13px; color: #555;">Rows: 50 | Columns: 5 (A-E)</div>
    </div>

    <!-- அட்டவணைப் பகுதி (Dynamic Table) -->
    <div class="table-container">
        <form id="dataForm" action="/convert" method="POST">
            <input type="hidden" name="active_cols" id="activeColsInput" value="5">
            <table id="dynamicTable">
                <thead>
                    <tr id="headerRow">
                        <th class="row-num-header">#</th>
                    </tr>
                </thead>
                <tbody id="tableBody">
                    <!-- Rows injected via JavaScript -->
                </tbody>
            </table>
        </form>
    </div>

    <!-- கட்டுப்பாட்டுப் பகுதி (Control Panel) -->
    <div class="control-panel">
        <div class="btn-group">
            <button type="button" class="btn-secondary" onclick="addRow()">+ புதிய வரிசை (Add Row)</button>
            <button type="button" class="btn-secondary" onclick="addColumn()">+ காலம் கூட்டு (Add Column)</button>
        </div>
        <div class="btn-group">
            <label for="formatSelect" style="font-weight: bold; font-size: 14px;">Format:</label>
            <select name="export_format" id="formatSelect" form="dataForm" class="format-select">
                <option value="csv">CSV (.csv)</option>
                <option value="xlsx">Excel (.xlsx)</option>
            </select>
            <button type="submit" form="dataForm">📥 கன்வெர்ட் & டவுன்லோட் (Convert)</button>
        </div>
    </div>

    <script>
        let colCount = 5; // Start with 5 columns (A to E)
        let rowCount = 50; // Start with 50 rows for immediate entry
        const maxCols = 26; // A to Z
        const columns = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];

        function initTable() {
            const headerRow = document.getElementById("headerRow");
            headerRow.innerHTML = '<th class="row-num-header">#</th>';
            
            for (let i = 0; i < colCount; i++) {
                let th = document.createElement("th");
                th.innerText = columns[i];
                headerRow.appendChild(th);
            }

            const tableBody = document.getElementById("tableBody");
            tableBody.innerHTML = "";
            for (let r = 1; r <= rowCount; r++) {
                appendRow(r);
            }
            updateStats();
        }

        function appendRow(rowNum) {
            const tableBody = document.getElementById("tableBody");
            let tr = document.createElement("tr");
            
            let tdIndex = document.createElement("td");
            tdIndex.className = "row-num";
            tdIndex.innerText = rowNum;
            tr.appendChild(tdIndex);

            for (let c = 0; c < colCount; c++) {
                let td = document.createElement("td");
                let input = document.createElement("input");
                input.type = "text";
                input.name = `cell_${rowNum}_${columns[c]}`;
                
                // Enter அழுத்தும்போது அடுத்த வரிசைக்குச் செல்ல / புதிய வரிசை உருவாக்க
                input.addEventListener("keydown", function(e) {
                    if (e.key === "Enter") {
                        e.preventDefault();
                        let nextRow = rowNum + 1;
                        let nextInput = document.querySelector(`input[name="cell_${nextRow}_${columns[c]}"]`);
                        if (nextInput) {
                            nextInput.focus();
                        } else {
                            addRow();
                            setTimeout(() => {
                                let newlyAdded = document.querySelector(`input[name="cell_${rowCount}_${columns[c]}"]`);
                                if (newlyAdded) newlyAdded.focus();
                            }, 50);
                        }
                    }
                });
                
                td.appendChild(input);
                tr.appendChild(td);
            }
            tableBody.appendChild(tr);
        }

        function addRow() {
            rowCount++;
            appendRow(rowCount);
            updateStats();
            const container = document.querySelector(".table-container");
            container.scrollTop = container.scrollHeight;
        }

        function addColumn() {
            if (colCount < maxCols) {
                colCount++;
                document.getElementById("activeColsInput").value = colCount;
                initTable();
            } else {
                alert("Z வரை காலம்கள் மட்டுமே உள்ளது!");
            }
        }

        function updateStats() {
            document.getElementById("statsInfo").innerText = `Rows: ${rowCount} | Columns: ${colCount} (${columns[0]}-${columns[colCount-1]})`;
        }

        window.onload = initTable;
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/convert', methods=['POST'])
def convert():
    form_data = request.form.to_dict()
    active_cols_count = int(form_data.get('active_cols', 5))
    export_format = form_data.get('export_format', 'csv')
    
    columns = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"][:active_cols_count]
    
    # Organize form data into structured rows
    rows = {}
    for key, value in form_data.items():
        if key.startswith('cell_'):
            parts = key.split('_')
            r_num = int(parts[1])
            col_name = parts[2]
            if r_num not in rows:
                rows[r_num] = {}
            rows[r_num][col_name] = value

    sorted_row_nums = sorted(rows.keys())
    if not sorted_row_nums:
        max_r = 1
    else:
        max_r = max(sorted_row_nums)

    if export_format == 'xlsx':
        # Create Excel file using openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sage Data"
        
        # Header row
        header = ["#"] + columns
        ws.append(header)
        
        # Data rows
        for r in range(1, max_r + 1):
            row_data = [r]
            for col in columns:
                row_data.append(rows.get(r, {}).get(col, ""))
            ws.append(row_data)
            
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='sage_of_cispath_data.xlsx'
        )
    else:
        # Create CSV file
        output = io.StringIO()
        writer = csv.writer(output)
        
        writer.writerow(["Row"] + columns)
        for r in range(1, max_r + 1):
            row_data = [r]
            for col in columns:
                row_data.append(rows.get(r, {}).get(col, ""))
            writer.writerow(row_data)
            
        output.seek(0)
        return send_file(
            io.BytesIO(output.getvalue().encode('utf-8')),
            mimetype='text/csv',
            as_attachment=True,
            download_name='sage_of_cispath_data.csv'
        )

if __name__ == '__main__':
    app.run(debug=True)

