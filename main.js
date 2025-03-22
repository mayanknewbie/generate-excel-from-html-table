let table = document.getElementById("gdvZone");
var tablelength = table.rows.length;
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function main() {
    for (let i = 8; i <= tablelength-3; i++) {
        await sleep(10000);
        var zoneID = 'gdvZone_lnkZone_'+(i);
        var zoneCol = document.getElementById(zoneID);
        console.log('Zone: ');
        console.log(zoneCol.innerText);
        // console.log(zoneCol.href);
        window.location.href = zoneCol.href;
        await sleep(10000);
        await rangeLoop();
    }
}

async function rangeLoop() {
    var tableR = document.getElementById("gvRange");
    var tablelengthR = tableR.rows.length;
    for (let r = 0;  r <= tablelengthR-3; r++) {
        await sleep(10000);
        var rangeID = 'gvRange_lnkRange_'+(r);
        var rangeCol = document.getElementById(rangeID);
        console.log('Range: ');
        console.log(rangeCol.innerText);
        // console.log(rangeCol.href);
        window.location.href = rangeCol.href;
        await sleep(10000); 
        await districtLoop();
    }
    console.log("Range se Bahar agaya mai to");
    await sleep(10000); 
    var btnReturnRange = document.getElementById('btnReturnRange');
    btnReturnRange.click();
}

async function districtLoop() {
    var tableD = document.getElementById("gvDistrict");
    var tablelengthD = tableD.rows.length;
    for (let d = 0;  d <= tablelengthD-3; d++) {
        await sleep(10000);
        console.log(d);
        var districtID = 'gvDistrict_lnkDistrict_'+(d);
        var districtCol = document.getElementById(districtID);
        console.log('District: ');
        console.log(districtCol.innerText);
        // console.log(districtCol.href);
        filename = districtCol.innerText
        window.location.href = districtCol.href;
        await sleep(10000); 
        exportTableToExcel('gvPS',filename); 
        var btnReturnPS = document.getElementById('btnReturnPS');
        btnReturnPS.click();
    }
    console.log("Bahar agaya mai to");
    await sleep(10000); 
    var btnReturnDistrict = document.getElementById('btnReturnDistrict');
    btnReturnDistrict.click();
}

main()

function exportTableToExcel(tableID, filename = ''){
    var downloadLink;
    var dataType = 'application/vnd.ms-excel';
    var tableSelect = document.getElementById(tableID);
    var tableHTML = tableSelect.outerHTML.replace(/ /g, '%20');
    
    // Specify file name
    filename = filename?filename+'.xls':'excel_data.xls';
    
    // Create download link element
    downloadLink = document.createElement("a");
    
    document.body.appendChild(downloadLink);
    
    if(navigator.msSaveOrOpenBlob){
        var blob = new Blob(['\ufeff', tableHTML], {
            type: dataType
        });
        navigator.msSaveOrOpenBlob( blob, filename);
    }else{
        // Create a link to the file
        downloadLink.href = 'data:' + dataType + ', ' + tableHTML;
    
        // Setting the file name
        downloadLink.download = filename;
        
        //triggering the function
        downloadLink.click();
    }
}
