attachEventToBtnExport()

function attachEventToBtnExport() {
    const bntExportElement = document.querySelector( '.btnExport' )
    bntExportElement.addEventListener( 'click', exportExcel )
}

async function exportExcel() {
    let response = await fetch( 'download_excel_api' )
    let jsonResponse = await response.json()
    bufferExcelFile = base64DecToArr( jsonResponse.data[ 0 ] ).buffer;
    const blobExcelFile = new Blob( [ bufferExcelFile ] )
    const fileName = 'dummy_excel.xlsx'
    downloadExcelSilently( blobExcelFile, fileName )
}
function downloadExcelSilently( blobExcelFile, filename ) {
    const url = window.URL.createObjectURL( blobExcelFile );
    const hiddenAnchor = document.createElement( "a" );
    hiddenAnchor.style.display = "none";
    hiddenAnchor.href = url;
    hiddenAnchor.download = filename;
    document.body.appendChild( hiddenAnchor );
    hiddenAnchor.click();
    window.URL.revokeObjectURL( url );
}


