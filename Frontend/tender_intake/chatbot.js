document.addEventListener("DOMContentLoaded", () => {
    // Existing question-and-answer elements
    const form = document.getElementById("chatbot-form");
    const questionInput = document.getElementById("question");
    const categoryInput = document.getElementById("category");
    const askButton = document.getElementById("ask-button");

    const loading = document.getElementById("loading");
    const errorMessage = document.getElementById("error-message");
    const answerSection = document.getElementById("answer-section");
    const answerText = document.getElementById("answer-text");
    const sourcesContainer = document.getElementById("sources-container");
    const sourcesList = document.getElementById("sources-list");
    const retrievedCount = document.getElementById("retrieved-count");
    const questionCount = document.getElementById("question-count");

    // Vendor upload elements
    const uploadForm = document.getElementById("upload-form");
    const fileInput = document.getElementById("vendor-files");
    const uploadButton = document.getElementById("upload-button");
    const uploadLoading = document.getElementById("upload-loading");
    const uploadError = document.getElementById("upload-error");
    const uploadResult = document.getElementById("upload-result");
    const selectedFiles = document.getElementById("selected-files");
    const uploadFileCount = document.getElementById("upload-file-count");

    const questionEndpoint = form.dataset.endpoint;
    const uploadEndpoint = uploadForm.dataset.endpoint;
    // Keep the currently active vendor document submission.
    let activeSubmissionId = null;

    // Existing question character counter
    questionInput.addEventListener("input", () => {
        questionCount.textContent = `${questionInput.value.length} / 5000`;
    });

    // Existing question submission
    form.addEventListener("submit", async (event) => {
        event.preventDefault();

        const question = questionInput.value.trim();

        if (!question) {
            showError("Please enter a procurement question.");
            return;
        }

        loading.classList.remove("hidden");
        errorMessage.classList.add("hidden");
        answerSection.classList.add("hidden");
        askButton.disabled = true;
        askButton.textContent = "Processing...";

        try {
            const response = await fetch(questionEndpoint, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                question: question,
                category: categoryInput.value || null,
                top_k: 5,
                submission_id: activeSubmissionId
                })
            });

            const data = await response.json();

            if (!response.ok) {
                throw new Error(data.error || "Unable to process your question.");
            }

            displayAnswer(data);

        } catch (error) {
            showError(error.message || "An unexpected error occurred.");
        } finally {
            loading.classList.add("hidden");
            askButton.disabled = false;
            askButton.textContent = "Ask Question";
        }
    });

    // Display selected vendor files
    fileInput.addEventListener("change", () => {
        const files = Array.from(fileInput.files);

        uploadFileCount.textContent =
            `${files.length} file${files.length === 1 ? "" : "s"} selected`;

        if (files.length === 0) {
            selectedFiles.textContent = "No files selected.";
            return;
        }

        selectedFiles.replaceChildren();

        files.forEach((file) => {
            const item = document.createElement("div");
            item.textContent =
                `${file.name} (${(file.size / (1024 * 1024)).toFixed(2)} MB)`;
            selectedFiles.appendChild(item);
        });
    });

    // Vendor document upload and processing
    uploadForm.addEventListener("submit", async (event) => {
        event.preventDefault();

        const files = Array.from(fileInput.files);

        if (files.length === 0) {
            showUploadError("Please select at least one document.");
            return;
        }

        if (files.length > 20) {
            showUploadError("You can upload a maximum of 20 files.");
            return;
        }

        const allowedExtensions = [".pdf", ".docx", ".xlsx", ".xlsm"];
        const totalSize = files.reduce((total, file) => total + file.size, 0);

        for (const file of files) {
            const extension = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();

            if (!allowedExtensions.includes(extension)) {
                showUploadError(`Unsupported file type: ${file.name}`);
                return;
            }

            if (file.size > 10 * 1024 * 1024) {
                showUploadError(`${file.name} exceeds the 10 MB per-file limit.`);
                return;
            }

            if (file.size === 0) {
                showUploadError(`${file.name} is empty.`);
                return;
            }
        }

        if (totalSize > 50 * 1024 * 1024) {
            showUploadError("The combined upload size cannot exceed 50 MB.");
            return;
        }

        const formData = new FormData();

        files.forEach((file) => {
            formData.append("files", file);
        });

        formData.append("category", categoryInput.value);

        uploadLoading.classList.remove("hidden");
        uploadError.classList.add("hidden");
        uploadResult.classList.add("hidden");
        uploadButton.disabled = true;
        uploadButton.textContent = "Uploading...";

        try {
            const response = await fetch(uploadEndpoint, {
                method: "POST",
                body: formData
            });

            const data = await response.json();

            if (!response.ok) {
                throw new Error(data.error || "Unable to upload the documents.");
            }

            // Activate the new submission only after       successful processing.
if (data.submission_id) {
    activeSubmissionId = data.submission_id;
}

displayUploadResult(data);

        } catch (error) {
            showUploadError(
                error.message || "An unexpected upload error occurred."
            );
        } finally {
            uploadLoading.classList.add("hidden");
            uploadButton.disabled = false;
            uploadButton.textContent = "Upload and Process Documents";
        }
    });

    // Display existing chatbot answer
    function displayAnswer(data) {
        answerText.textContent = data.answer || "No answer was returned.";

        sourcesList.replaceChildren();

        const sources = Array.isArray(data.sources) ? data.sources : [];

        sources.forEach((source) => {
            const item = document.createElement("li");
            item.textContent = source;
            sourcesList.appendChild(item);
        });

        sourcesContainer.classList.toggle("hidden", sources.length === 0);

        retrievedCount.textContent =
            `Retrieved reference sections: ${data.retrieved_count ?? sources.length}`;

        answerSection.classList.remove("hidden");
    }

    // Display vendor submission processing results
    function displayUploadResult(data) {
        const summary = data.summary || {};
        const files = Array.isArray(data.files) ? data.files : [];
        const ocrPages = Array.isArray(data.pages_requiring_ocr)
            ? data.pages_requiring_ocr
            : [];

        document.getElementById("submission-id").textContent =
            `Submission ID: ${data.submission_id || "Not available"}`;

        document.getElementById("submission-status").textContent =
            `Processing status: ${data.status || "Unknown"}`;

        document.getElementById("files-processed").textContent =
            summary.files_processed ?? 0;

        document.getElementById("files-failed").textContent =
            summary.files_failed ?? 0;

        document.getElementById("sections-extracted").textContent =
            summary.total_sections ?? 0;

        document.getElementById("ocr-pages").textContent =
            summary.pages_requiring_ocr ?? ocrPages.length;

        const filesList = document.getElementById("uploaded-files-list");
        filesList.replaceChildren();

        files.forEach((file) => {
            const item = document.createElement("li");
            const fileName = file.file_name || "Unnamed document";
            const status = file.status || "Unknown";
            const sectionCount = file.section_count ?? 0;

            item.textContent =
                `${fileName} — ${status} (${sectionCount} sections)`;

            filesList.appendChild(item);
        });

        const reviewWarning = document.getElementById("review-warning");

        reviewWarning.classList.toggle(
            "hidden",
            ocrPages.length === 0 &&
            (summary.files_failed ?? 0) === 0
        );

        uploadResult.classList.remove("hidden");
        uploadResult.scrollIntoView({
            behavior: "smooth",
            block: "start"
        });
    }

    function showError(message) {
        errorMessage.textContent = message;
        errorMessage.classList.remove("hidden");
    }

    function showUploadError(message) {
        uploadError.textContent = message;
        uploadError.classList.remove("hidden");
    }
});
