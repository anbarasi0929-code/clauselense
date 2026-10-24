
const fileInput = document.getElementById("contractFile");
const dropZone = document.getElementById("dropZone");
const fileInfo = document.getElementById("fileInfo");
const fileName = document.getElementById("fileName");
const uploadTitle = document.getElementById("uploadTitle");
const uploadText = document.getElementById("uploadText");
const form = document.getElementById("uploadForm");
const message = document.getElementById("formMessage");
const removeFileButton = document.getElementById("removeFile");

const allowedExtensions = ["pdf", "docx", "jpg", "jpeg", "png"];
const maxFileSize = 15 * 1024 * 1024; // 15 MB

function showMessage(text) {
    message.textContent = text;
}

function displayFile(file) {
    if (!file) return;

    const extension = file.name.split(".").pop().toLowerCase();

    if (!allowedExtensions.includes(extension)) {
        fileInput.value = "";
        showMessage("Unsupported format. Select a PDF, DOCX, JPG, JPEG, or PNG file.");
        return;
    }

    if (file.size > maxFileSize) {
        fileInput.value = "";
        showMessage("The file is too large. Please choose a file smaller than 15 MB.");
        return;
    }

    showMessage("");
    fileName.textContent =
        `${file.name} (${(file.size / (1024 * 1024)).toFixed(2)} MB)`;

    fileInfo.hidden = false;
    uploadTitle.textContent = "Contract selected";
    uploadText.textContent = "Choose another file by clicking here if needed";
}

fileInput.addEventListener("change", () => {
    displayFile(fileInput.files[0]);
});

removeFileButton.addEventListener("click", () => {
    fileInput.value = "";
    fileInfo.hidden = true;
    uploadTitle.textContent = "Drop your contract here";
    uploadText.textContent = "or click to browse your computer";
    showMessage("");
});

["dragenter", "dragover"].forEach(eventName => {
    dropZone.addEventListener(eventName, event => {
        event.preventDefault();
        dropZone.classList.add("dragover");
    });
});

["dragleave", "drop"].forEach(eventName => {
    dropZone.addEventListener(eventName, event => {
        event.preventDefault();
        dropZone.classList.remove("dragover");
    });
});

dropZone.addEventListener("drop", event => {
    const files = event.dataTransfer.files;

    if (!files.length) return;

    try {
        const transfer = new DataTransfer();
        transfer.items.add(files[0]);
        fileInput.files = transfer.files;
        displayFile(files[0]);
    } catch {
        showMessage("Please click the upload area and select your file.");
    }
});

form.addEventListener("submit", async event => {
    event.preventDefault();

    if (!fileInput.files.length) {
        showMessage("Please select a contract file first.");
        return;
    }

    const file = fileInput.files[0];
    const extension = file.name.split(".").pop().toLowerCase();

    if (!allowedExtensions.includes(extension) ||
        file.size > maxFileSize) {
        showMessage("Please select a supported file smaller than 15 MB.");
        return;
    }

    const formData = new FormData();
    formData.append("file", file);
    formData.append(
        "requirements",
        document.getElementById("requirements").value.trim()
    );

    const submitButton = document.getElementById("submitButton");
    submitButton.disabled = true;
    submitButton.textContent = "Analyzing...";

    showMessage("Please wait while ClauseLens analyzes your contract.");

    try {
        const response = await fetch("http://127.0.0.1:5000/api/analyze", {
            method: "POST",
            body: formData
        });

        const result = await response.json();

        if (!response.ok) {
            throw new Error(result.error || "Analysis failed.");
        }

        sessionStorage.setItem("clauseLensResult", JSON.stringify(result));
        window.location.href = "results.html";
    } catch (error) {
        showMessage("Could not analyze the contract: " + error.message);
    } finally {
        submitButton.disabled = false;
        submitButton.innerHTML = 'Analyze Contract <span>→</span>';
    }
});


