/**
 * Triggers the .docx download using the custom FBP Parser.
 * Ensures branding and formatting are applied correctly.
 */
async function downloadReport() {
    if (!lastReport) {
        alert("No report data found. Please run the reconciliation first.");
        return;
    }

    try {
        const btn = document.querySelector('.dl-btn');
        const originalText = btn.innerText;
        btn.innerText = "Generating...";

        // Call your custom GitHub parser
        const blob = await ReportFormatterDocx.toBlob(lastReport);

        // Create the download link
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        const timestamp = new Date().toISOString().slice(0, 10);
        
        a.href = url;
        a.download = `FBP_Reconciliation_Report_${timestamp}.docx`;
        document.body.appendChild(a);
        a.click();

        // Cleanup
        setTimeout(() => {
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
            btn.innerText = originalText;
        }, 100);

    } catch (err) {
        console.error("Docx Generation Error:", err);
        alert("Failed to generate Word document. Check the console for details.");
    }
}
