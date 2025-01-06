// Fungsi untuk mengonversi file Excel menjadi XML
function excelToXml(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });

        // Sheet 1 untuk data TaxInvoice
        const sheet1 = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData1 = XLSX.utils.sheet_to_json(sheet1);

        // Sheet 2 untuk data GoodService
        const sheet2 = workbook.Sheets[workbook.SheetNames[1]];
        const jsonData2 = XLSX.utils.sheet_to_json(sheet2);

        let xmlString = '<?xml version="1.0" encoding="UTF-8"?>\n';
        xmlString += '<TaxInvoiceBulk xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n';
        xmlString += '  <TIN>1091031210912281</TIN>\n';
        xmlString += '  <ListOfTaxInvoice>\n';

        // Loop untuk sheet1 (TaxInvoice)
        jsonData1.forEach(row => {
            xmlString += '    <TaxInvoice>\n';
            xmlString += `      <TaxInvoiceDate>${row.TaxInvoiceDate || ''}</TaxInvoiceDate>\n`;
            xmlString += `      <TaxInvoiceOpt>${row.TaxInvoiceOpt || ''}</TaxInvoiceOpt>\n`;
            xmlString += `      <TrxCode>${row.TrxCode || ''}</TrxCode>\n`;
            xmlString += `      <AddInfo>${row.AddInfo || ''}</AddInfo>\n`;
            xmlString += `      <CustomDoc>${row.CustomDoc || ''}</CustomDoc>\n`;
            xmlString += `      <RefDesc>${row.RefDesc || ''}</RefDesc>\n`;
            xmlString += `      <FacilityStamp>${row.FacilityStamp || ''}</FacilityStamp>\n`;
            xmlString += `      <SellerIDTKU>${row.SellerIDTKU || ''}</SellerIDTKU>\n`;
            xmlString += `      <BuyerTin>${row.BuyerTin || ''}</BuyerTin>\n`;
            xmlString += `      <BuyerDocument>${row.BuyerDocument || ''}</BuyerDocument>\n`;
            xmlString += `      <BuyerCountry>${row.BuyerCountry || ''}</BuyerCountry>\n`;
            xmlString += `      <BuyerDocumentNumber>${row.BuyerDocumentNumber || ''}</BuyerDocumentNumber>\n`;
            xmlString += `      <BuyerName>${row.BuyerName || ''}</BuyerName>\n`;
            xmlString += `      <BuyerAdress>${row.BuyerAdress || ''}</BuyerAdress>\n`;
            xmlString += `      <BuyerEmail>${row.BuyerEmail || ''}</BuyerEmail>\n`;
            xmlString += `      <BuyerIDTKU>${row.BuyerIDTKU || ''}</BuyerIDTKU>\n`;

            // Tambahkan ListOfGoodService
            xmlString += '      <ListOfGoodService>\n';

            // Loop untuk sheet2 (GoodService) dan cocokkan dengan Baris yang sama
            jsonData2.forEach(item => {
                // Cek apakah Baris dari Sheet 2 sesuai dengan Baris dari Sheet 1
                if (item.Baris == row.Baris) {
                    xmlString += '        <GoodService>\n';
                    xmlString += `          <Opt>${item.Opt || ''}</Opt>\n`;
                    xmlString += `          <Code>${item.Code || ''}</Code>\n`;
                    xmlString += `          <Name>${item.Name || ''}</Name>\n`;
                    xmlString += `          <Unit>${item.Unit || ''}</Unit>\n`;
                    xmlString += `          <Price>${item.Price || ''}</Price>\n`;
                    xmlString += `          <Qty>${item.Qty || ''}</Qty>\n`;
                    xmlString += `          <TotalDiscount>${item.TotalDiscount || ''}</TotalDiscount>\n`;
                    xmlString += `          <TaxBase>${item.TaxBase || ''}</TaxBase>\n`;
                    xmlString += `          <OtherTaxBase>${item.OtherTaxBase || ''}</OtherTaxBase>\n`;
                    xmlString += `          <VATRate>${item.VATRate || ''}</VATRate>\n`;
                    xmlString += `          <VAT>${item.VAT || ''}</VAT>\n`;
                    xmlString += `          <STLGRate>${item.STLGRate || ''}</STLGRate>\n`;
                    xmlString += `          <STLG>${item.STLG || ''}</STLG>\n`;
                    xmlString += '        </GoodService>\n';
                }
            });

            xmlString += '      </ListOfGoodService>\n';
            xmlString += '    </TaxInvoice>\n';
        });

        xmlString += '  </ListOfTaxInvoice>\n';
        xmlString += '</TaxInvoiceBulk>';

        // Mengunduh file XML
        downloadXml(xmlString);
    };

    reader.readAsBinaryString(file);
}

// Fungsi untuk mengunduh file XML
function downloadXml(xmlString) {
    const blob = new Blob([xmlString], { type: 'application/xml' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'tax_invoice.xml';
    link.click();
}
