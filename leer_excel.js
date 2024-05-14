class Excel {
    constructor(content) {
        this.content = content;
    }
    
    header() {
        return this.content[0];
    }
    
    rows() {
        return new RowCollection(this.content.slice(1, this.content.length));
    } 
}

class RowCollection {
    constructor(rows) {
        this.rows = rows;
    }
 
    first() {
        return this.rows[0];
    }
    
    get() {
        return this.rows;
    }
    
    count() {
        return this.rows.length;
    }
}

class Row {
    constructor(row) {
        this.row = row;
    }

    NumeroEmpleado() {
        return this.row[0];
    }
    Colaborador() {
        return this.row[1];
    }   
    Puesto() {
        return this.row[2];
    }       
    Correo() {
        return this.row[3];
    }
}

class ExcelPrinter {
    static print(tableId, excel) {
        const table = document.getElementById(tableId);

        excel.header().forEach(title => {
            table.querySelector("thead>tr").innerHTML += `<td>${title}</td>`;
        });

        const rows = excel.rows().get();
        rows.forEach(rowData => {
            const rowInstance = new Row(rowData);
            table.querySelector("tbody").innerHTML += `
                <tr>
                    <td>${rowInstance.NumeroEmpleado()}</td>
                    <td>${rowInstance.Colaborador()}</td>
                    <td>${rowInstance.Puesto()}</td>
                    <td>${rowInstance.Correo()}</td>
                </tr>`;
        });
    }
}

const excelInput = document.getElementById('excel-file');

excelInput.addEventListener('change', async function() {
    const content = await readXlsxFile(excelInput.files[0]);
    const excel = new Excel(content);
    ExcelPrinter.print('tabla-excel', excel);
});
