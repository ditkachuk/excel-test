(function() {
    window.onload = function() {
        window.app = new application();
    };

    function application() {
        var self = this;
        self.table = document.getElementById('table');

		this.loadExcel = function() {
            data = getDataFromTable(self.table);
            widths = getWidthFromTable(self.table);
            saveExcel(data, widths);
        };

        function saveExcel(data, wscols) {
            var ws_name = "Table";

            var wc_style_border = {
                left: {style: 'thin', color: {auto: 1}},
                right: {style: 'thin', color: {auto: 1}},
                top: {style: 'thin', color: {auto: 1}},
                bottom: {style: 'thin', color: {auto: 1}}
            };

            var wc_style_font_header = {
                bold: true
            };

            /* dummy workbook constructor */
            function Workbook() {
                if(!(this instanceof Workbook)) return new Workbook();
                this.SheetNames = [];
                this.Sheets = {};
            }
            var wb = new Workbook();

            function datenum(v, date1904) {
                if(date1904) v+=1462;
                var epoch = Date.parse(v);
                return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
            }

            /* convert an array of arrays in JS to a CSF spreadsheet */
            function sheet_from_array_of_arrays(data, opts) {
                var ws = {};
                var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
                for(var R = 0; R != data.length; ++R) {
                    for(var C = 0; C != data[R].length; ++C) {
                        if(range.s.r > R) range.s.r = R;
                        if(range.s.c > C) range.s.c = C;
                        if(range.e.r < R) range.e.r = R;
                        if(range.e.c < C) range.e.c = C;
                        var cell = {v: data[R][C] };
                        if(cell.v == null) continue;
                        var cell_ref = XLSX.utils.encode_cell({c:C,r:R});

                        /* TEST: proper cell types and value handling */
                        if(typeof cell.v === 'number') cell.t = 'n';
                        else if(typeof cell.v === 'boolean') cell.t = 'b';
                        else if(cell.v instanceof Date) {
                            cell.t = 'n'; cell.z = XLSX.SSF._table[14];
                            cell.v = datenum(cell.v);
                        }
                        else cell.t = 's';
                        cell.s = {border: wc_style_border};
                        if (R == 0) cell.s.font = wc_style_font_header;
                        ws[cell_ref] = cell;
                    }
                }

                /* TEST: proper range */
                if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
                return ws;
            }
            var ws = sheet_from_array_of_arrays(data);

            /* TEST: add worksheet to workbook */
            wb.SheetNames.push(ws_name);
            wb.Sheets[ws_name] = ws;

            /* TEST: column widths */
            ws['!cols'] = wscols;

            var wopts = { bookType:'xlsx', bookSST:false, type:'binary' };
            console.log(wb);
            var wbout = XLSX.write(wb, wopts);

            function s2ab(s) {
              var buf = new ArrayBuffer(s.length);
              var view = new Uint8Array(buf);
              for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
              return buf;
            }

            /* the saveAs call downloads a file on the local machine */
            saveAs(new Blob([s2ab(wbout)],{type:""}), "test.xlsx");
        }

        function getDataFromTable(element) {
            var myTableArray = [];

            $(element).find("tr").each(function() {
                var arrayOfThisRow = [];

                var tableHeader = $(this).find('th');
                var tableData = $(this).find('td');
                tableData = tableHeader.length > 0 ? tableHeader : tableData;
                if (tableData.length > 0) {
                    tableData.each(function() { arrayOfThisRow.push($(this).text()); });
                    myTableArray.push(arrayOfThisRow);
                }
            });

            return myTableArray;
        }

        function getWidthFromTable(element) {
            var widths = [];

            $(element).find("th").each(function() {
                widths.push({wch: $(this).width() / 8});
            });

            return widths;
        }
    }
})();
