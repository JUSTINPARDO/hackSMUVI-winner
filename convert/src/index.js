import addOnUISdk from "https://new.express.adobe.com/static/add-on-sdk/sdk.js";

// Wait for SDK to be ready
addOnUISdk.ready.then(() => {
    console.log("SDK is ready");
    const exportToPPTXButton = document.getElementById("exportToPPTXButton");
    const exportToDOCXButton = document.getElementById("exportToDOCXButton");
    const loadingSpinner = document.getElementById("loadingSpinner");
    const copyLinkButton = document.getElementById("copyLinkButton");
    const statusMessage = document.getElementById("statusMessage");
    exportToPPTXButton.disabled = false;
    exportToDOCXButton.disabled = false;

    function showStatus(message, duration = 3000) {
        statusMessage.textContent = message;
        statusMessage.style.display = "block";
        setTimeout(() => {
            statusMessage.style.display = "none";
        }, duration);
    }

    async function handleExport(format, serverUrl) {
        exportToPPTXButton.style.display = "none";
        exportToDOCXButton.style.display = "none";
        loadingSpinner.style.display = "block";

        try {
            const renditionOptions = {
                range: addOnUISdk.constants.Range.entireDocument,
                format: addOnUISdk.constants.RenditionFormat.pdf,
            };

            const renditions = await addOnUISdk.app.document.createRenditions(
                renditionOptions,
                addOnUISdk.constants.RenditionIntent.export
            );

            const pdfBlob = renditions[0].blob;

            if (format === "DOCX") {
                const formData = new FormData();
                formData.append('file', pdfBlob, 'presentation.pdf');

                const response = await fetch(serverUrl, {
                    method: 'POST',
                    body: formData
                });

                if (response.ok) {
                    const result = await response.json();
                    copyLinkButton.dataset.url = result.downloadUrl;
                    loadingSpinner.style.display = "none";
                    copyLinkButton.style.display = "block";
                    showStatus(`Conversion to ${format} complete. Click 'Copy Link' to get the download URL.`);
                } else {
                    console.error('Server response not OK:', response.status);
                    showStatus(`An error occurred during conversion to ${format}. Please try again.`);
                }
            } else {
                const response = await fetch(serverUrl, {
                    method: 'POST',
                    body: pdfBlob,
                    headers: {
                        'Content-Type': 'application/pdf'
                    }
                });

                if (response.ok) {
                    const result = await response.json();
                    copyLinkButton.dataset.url = result.downloadUrl;
                    loadingSpinner.style.display = "none";
                    copyLinkButton.style.display = "block";
                    showStatus(`Conversion to ${format} complete. Click 'Copy Link' to get the download URL.`);
                } else {
                    console.error('Server response not OK:', response.status);
                    showStatus(`An error occurred during conversion to ${format}. Please try again.`);
                }
            }
        } catch (error) {
            console.error(`Error during export or conversion to ${format}:`, error);
            showStatus('An error occurred. Please try again.');
        } finally {
            exportToPPTXButton.style.display = "block";
            exportToDOCXButton.style.display = "block";
            loadingSpinner.style.display = "none";
        }
    }

    exportToPPTXButton.addEventListener("click", () => handleExport("PPTX", 'http://localhost:3000/convert'));
    exportToDOCXButton.addEventListener("click", () => handleExport("DOCX", 'http://127.0.0.1:5000/convert'));

    copyLinkButton.addEventListener('click', () => {
        const downloadUrl = copyLinkButton.dataset.url;
        const tempInput = document.createElement("input");
        document.body.appendChild(tempInput);
        tempInput.value = downloadUrl;
        tempInput.select();
        document.execCommand("copy");
        document.body.removeChild(tempInput);
        showStatus('Download link copied to clipboard!');
        document.getElementById('copiedMessage').style.display = 'block';
        setTimeout(() => {
            document.getElementById('copiedMessage').style.display = 'none';
        }, 2000); 
    });
});