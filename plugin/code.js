(function(window, undefined){

    window.Asc.plugin.init = function(){

        console.log('plugin loaded');
        console.log('window.Asc:', window.Asc);

        document.getElementById('btn-copy').addEventListener('click', FCopy);    
    };

    function FCopy(){
        console.log('copy button clicked');

        window.Asc.plugin.callCommand(function() {
            
            //geting active table borders
            console.log('borders processing');
            let oWorksheet = Api.GetActiveSheet();
            console.log('sheet gotten', oWorksheet);
            
            const usedRange = oWorksheet.UsedRange;
            console.log('usedrange gotten');

            const row1 = usedRange.Row;
            const col1 = usedRange.Col;
            const rowsN = usedRange.Rows.Count;
            const colsN = usedRange.Cols.Count;
            console.log('borders gotten');

            //getting active table data
            const tableData = [];
            for (let i = 0; i < rowsN; i++) {
                const oRow = [];
                for (let j = 0; j < colsN; j++) {
                    const oValue = oWorksheet.GetRangeByNumber(row1 - 1 + i, col1 - 1 + j).Text ?? ' ';
                    oRow.push(oValue);
                }
                tableData.push(oRow);
            }
            console.log('table', tableData);

            //formating markdown string
            const mdHeader = '|' + tableData[0].join('|') + '|';
            const mdSeparator = '|' + tableData[0].map(() => '-:').join('|') + '|';
            const mdBody = '|' + tableData.slice(1).map(row => row.join('|')).join('|\n|');
            const mdTable = `${mdHeader}\n${mdSeparator}\n${mdBody}`;
            console.log('markdown table', mdTable);
            alert(mdTable);

            // copying .md table to clipboard
            navigator.clipboard.writeText(mdTable)

        }, undefined,false);

    };

})(window, undefined);