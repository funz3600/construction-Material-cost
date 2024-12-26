document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('upload').addEventListener('change', handleFile, false);
    document.getElementById('save').addEventListener('click', saveFile, false);

    let hot; // Declare Handsontable instance variable

    // Function to handle file upload
    function handleFile(e) {
        const file = e.target.files[0];
        const reader = new FileReader();
        reader.onload = function(event) {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Display the data using Handsontable
            displayTable(jsonData);
        };
        reader.readAsArrayBuffer(file);
    }

    // Function to display the data in Handsontable
    function displayTable(data) {
        const container = document.getElementById('excelTable');
        hot = new Handsontable(container, {
            data: data,
            rowHeaders: true,
            colHeaders: true,
            stretchH: 'all', // Stretch columns to fill the available width
            autoColumnSize: true, // Enable automatic column width adjustment
            manualColumnResize: true, // Allow manual column resizing
            manualRowResize: true, // Allow manual row resizing
            wordWrap: true, // Enable word wrapping
            formulas: {
                engine:[43dcd9a7-70db-4a1f-b0ae-981daa162054](https://github.com/krypton4149/darkgym.github.io/tree/c1682c20bec95d60b0f6f49feaf2d32dc72ccafd/README.md?citationMarker=43dcd9a7-70db-4a1f-b0ae-981daa162054 "1")[43dcd9a7-70db-4a1f-b0ae-981daa162054](https://github.com/alamtopan/maksima-production/tree/d2ea61957f20528de7daa2feedd4917278340065/shareds%2Ffooter.php?citationMarker=43dcd9a7-70db-4a1f-b0ae-981daa162054 "2")[43dcd9a7-70db-4a1f-b0ae-981daa162054](https://github.com/90sAlex/PEF/tree/bac690e47d8f2bc6d6a47f2df75de35168533773/cart.php?citationMarker=43dcd9a7-70db-4a1f-b0ae-981daa162054 "3")[43dcd9a7-70db-4a1f-b0ae-981daa162054](https://github.com/MyifanW/topdeck/tree/65196eb5d16685a29d620ae6aaa268516b51866b/www%2F424layout.php?citationMarker=43dcd9a7-70db-4a1f-b0ae-981daa162054 "4")[43dcd9a7-70db-4a1f-b0ae-981daa162054](https://github.com/BrandonHVX/FastlaneLogisticsGatsby/tree/4a0fa8b412b4289fe9633d03a1ddd827611d65ab/src%2Fcomponents%2FFooter.js?citationMarker=43dcd9a7-70db-4a1f-b0ae-981daa162054 "5")[43dcd9a7-70db-4a1f-b0ae-981daa162054](https://github.com/elifSamar/WordPress-v1-Crafts/tree/89a277faea2883db00b140aa60fa77f17ec5d0de/front-page.php?citationMarker=43dcd9a7-70db-4a1f-b0ae-981daa162054 "6")[43dcd9a7-70db-4a1f-b0ae-981daa162054](https://github.com/Sumonchandrashil/bdcare_magnetism/tree/e3bc1066f2f7becc58afc59a0aba778de0303fa5/public%2Fassets%2Ffrontend%2Ffooter.php?citationMarker=43dcd9a7-70db-4a1f-b0ae-981daa162054 "7")[43dcd9a7-70db-4a1f-b0ae-981daa162054](https://github.com/jcaanderson/jcaanderson.github.io/tree/6a89b321a454c643ff3dba83378b835c1f08d1c2/web%2Fapp%2Fthemes%2Fjessanderson-theme%2Findex.php?citationMarker=43dcd9a7-70db-4a1f-b0ae-981daa162054 "8")[43dcd9a7-70db-4a1f-b0ae-981daa162054](https://github.com/anthonyhdd/webdesign-meetharry/tree/bbf09dbd611f0d7aaaeb90e6128d59065a274c21/footer.php?citationMarker=43dcd9a7-70db-4a1f-b0ae-981daa162054 "9")[43dcd9a7-70db-4a1f-b0ae-981daa162054](https://github.com/raph0424/survey/tree/d5c2a3180e1ad12aae8fd23bf0b5be109a73b809/vue%2Fformulaire%2FformInscription.php?citationMarker=43dcd9a7-70db-4a1f-b0ae-981daa162054 "10")[43dcd9a7-70db-4a1f-b0ae-981daa162054](https://github.com/filipw2/System-rezerwacji/tree/e3fbc864c11535d0868a3fa5aee079a837655bec/login.php?citationMarker=43dcd9a7-70db-4a1f-b0ae-981daa162054 "11")[43dcd9a7-70db-4a1f-b0ae-981daa162054](https://github.com/CydoEntis/minima/tree/581b6551b7503b4c42a075999f1d1d0a6fc3e6d7/app%2Fviews%2Fminimalista%2Fsource%20code%2Fminima%2Fapp%2Fviews%2Fminima%2Fsignup.php?citationMarker=43dcd9a7-70db-4a1f-b0ae-981daa162054 "12")[43dcd9a7-70db-4a1f-b0ae-981daa162054](https://github.com/djade007/HackMD/tree/977294dcbf785d17697ff67ec53742cf94ef788f/login%2Findex.php?citationMarker=43dcd9a7-70db-4a1f-b0ae-981daa162054 "13")