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
    // Compliance visualization elements
const visualizationContainer =
    document.getElementById("visualization-container");

const visualizationTitle =
    document.getElementById("visualization-title");

const visualizationSubtitle =
    document.getElementById("visualization-subtitle");

const visualizationSummary =
    document.getElementById("visualization-summary");

const visualizationWarnings =
    document.getElementById("visualization-warnings");

const visualizationWarningsList =
    document.getElementById("visualization-warnings-list");

const visualizationNotice =
    document.getElementById("visualization-notice");

let complianceDoughnutChart = null;
let complianceBarChart = null;

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

        
// Existing question submission with visualization detection
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
        // Route compliance graph requests to the Python analyzer.
        if (
            isComplianceVisualizationRequest(question) &&
            !isWorkspaceComparisonRequest(question)
        ) {

            if (!activeSubmissionId) {
                throw new Error(
                    "Please upload vendor documents before requesting a compliance graph."
                );
            }

            loading.textContent =
                "Analysing vendor compliance and generating charts...";

            const response = await fetch("/chatbot/visualize", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    submission_id: activeSubmissionId
                })
            });

            const data = await response.json();

            if (!response.ok) {
                throw new Error(
                    data.error || "Unable to generate the compliance visualization."
                );
            }

            displayVisualization(data);

        } else {

            // Preserve the existing RAG question-and-answer flow.
            loading.textContent =
                "Searching contract documents and generating your answer...";

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
                throw new Error(
                    data.error || "Unable to process your question."
                );
            }

            if (data.comparison) {
                displayComparison(data);
            } else {
                displayAnswer(data);
            }
        }

    } catch (error) {
        showError(error.message || "An unexpected error occurred.");
    } finally {
        loading.classList.add("hidden");
        loading.textContent =
            "Searching contract documents and generating your answer...";

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

    // Safely convert supported Markdown into formatted HTML.
    function renderMarkdown(markdown) {
        function escapeHtml(value) {
            return value
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;")
                .replace(/'/g, "&#39;");
        }

        function formatInline(value) {
            return escapeHtml(value)
                .replace(/`([^`]+)`/g, "<code>$1</code>")
                .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
                .replace(/__(.+?)__/g, "<strong>$1</strong>")
                .replace(/\*(.+?)\*/g, "<em>$1</em>");
        }

        const lines = String(markdown).split(/\r?\n/);
        const output = [];

        let paragraph = [];
        let listType = null;

        function closeList() {
            if (listType) {
                output.push(listType === "ul" ? "</ul>" : "</ol>");
                listType = null;
            }
        }

        function flushParagraph() {
            if (paragraph.length) {
                output.push("<p>" + paragraph.map(formatInline).join("<br>") + "</p>");
                paragraph = [];
            }
        }

        for (const line of lines) {
            const trimmed = line.trim();

            if (!trimmed) {
                flushParagraph();
                closeList();
                continue;
            }

            const heading = trimmed.match(/^(#{1,6})\s+(.+)$/);

            if (heading) {
                flushParagraph();
                closeList();

                const level = Math.min(heading[1].length + 1, 6);
                output.push(
                    `<h${level}>${formatInline(heading[2])}</h${level}>`
                );
                continue;
            }

            if (/^(\*\s*\*\s*\*|-{3,}|_{3,})$/.test(trimmed)) {
                flushParagraph();
                closeList();
                output.push("<hr>");
                continue;
            }

            const bullet = trimmed.match(/^[-*+]\s+(.+)$/);
            const numbered = trimmed.match(/^\d+[.)]\s+(.+)$/);

            if (bullet || numbered) {
                flushParagraph();

                const requiredType = bullet ? "ul" : "ol";

                if (listType !== requiredType) {
                    closeList();
                    output.push(requiredType === "ul" ? "<ul>" : "<ol>");
                    listType = requiredType;
                }

                output.push("<li>" + formatInline((bullet || numbered)[1]) + "</li>");
                continue;
            }

            closeList();
            paragraph.push(trimmed);
        }

        flushParagraph();
        closeList();

        return output.join("\n");
    }


// Identify compliance visualization requests.
function isComplianceVisualizationRequest(question) {
    return /\b(graph|chart|visuali[sz](?:ation|e)|plot|graphical|diagram|distribution)\b/i.test(question);
}

function isWorkspaceComparisonRequest(question) {
    const comparisonIntent = /\b(compare|comparison|versus|vs\.?|against|which vendor|best vendor|meets? the tender requirements)\b/i;
    const workspaceSubject = /\b(documents?|vendors?|submissions?|uploaded|workspace|delivery|technical|specifications?|compliance|price|warranty|security|validity)\b/i;
    return comparisonIntent.test(question) && workspaceSubject.test(question);
}


// Destroy previous charts before rendering new ones.
function destroyComplianceCharts() {
    if (complianceDoughnutChart) {
        complianceDoughnutChart.destroy();
        complianceDoughnutChart = null;
    }

    if (complianceBarChart) {
        complianceBarChart.destroy();
        complianceBarChart = null;
    }
}


// Display the compliance visualization returned by Python.
function displayVisualization(data) {

    if (typeof Chart === "undefined") {
        throw new Error(
            "Chart.js could not be loaded. Please check your internet connection."
        );
    }

    const chartData = data.chart_data || {};
    const labels = chartData.labels || [];
    const counts = chartData.counts || [];
    const percentages = chartData.percentages || [];

    if (
        !labels.length ||
        labels.length !== counts.length ||
        labels.length !== percentages.length
    ) {
        throw new Error(
            "The compliance analyzer returned incomplete chart data."
        );
    }

    destroyComplianceCharts();
    document.getElementById("compliance-doughnut-chart").closest(".chart-card").classList.remove("hidden");
    document.getElementById("compliance-bar-chart").closest(".chart-card").classList.remove("hidden");
    document.getElementById("compliance-bar-chart").closest(".chart-card").querySelector("h4").textContent = "Clause Status Comparison";

    // Hide ordinary RAG answer elements.
    answerText.textContent = "";
    sourcesContainer.classList.add("hidden");
    retrievedCount.textContent = "";

    visualizationTitle.textContent =
        data.title || "Vendor Compliance Analysis";

    visualizationSubtitle.textContent =
        data.analysis_complete
            ? "All extracted declarations have been matched to the identified clauses."
            : "Preliminary analysis — review the warnings before relying on these results.";

    // Display summary figures.
    const summary = data.summary || {};

    const breakdown = summary.breakdown || {};
    const total = summary.total_clauses ?? counts.reduce(
        (sum, count) => sum + Number(count || 0), 0
    );

    const stats = [
        {
            label: "Total Clauses",
            value: total
        },
        {
            label: "Complied",
            value: breakdown["Complied"]?.count ?? summary.complied ?? 0
        },
        {
            label: "Partially Complied",
            value: breakdown["Partially Complied"]?.count ?? summary.partially_complied ?? 0
        },
        {
            label: "Not Complied",
            value: breakdown["Not Complied"]?.count ?? summary.not_complied ?? 0
        }
    ];

    visualizationSummary.replaceChildren();

    stats.forEach((stat) => {
        const item = document.createElement("div");
        item.className = "visualization-stat";

        const value = document.createElement("strong");
        value.textContent = stat.value;

        const label = document.createElement("span");
        label.textContent = stat.label;

        item.append(value, label);
        visualizationSummary.appendChild(item);
    });

    const statusColors = {
        "Complied": "#16A34A",
        "Not Complied": "#DC2626",
        "Partially Complied": "#F59E0B",
        "Not Specified": "#64748B",
        "Unclear": "#8B5CF6",
        "OCR Required": "#0EA5E9"
    };
    const chartColors = labels.map(label => statusColors[label] || "#64748B");

    // Doughnut chart.
    const doughnutContext = document
        .getElementById("compliance-doughnut-chart")
        .getContext("2d");

    complianceDoughnutChart = new Chart(doughnutContext, {
        type: "doughnut",

        data: {
            labels: labels,

            datasets: [{
                data: counts,
                backgroundColor: chartColors,
                borderColor: "#ffffff",
                borderWidth: 3,
                hoverOffset: 12
            }]
        },

        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: "65%",

            plugins: {
                legend: {
                    position: "bottom"
                },

                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const index = context.dataIndex;

                            return `${context.label}: ${context.raw} clauses (${percentages[index]}%)`;
                        }
                    }
                }
            }
        }
    });

    // Horizontal bar chart of the three explicit declaration categories.
    const barLabels = ["Complied", "Partially Complied", "Not Complied"]
        .filter(label => labels.includes(label));
    const barCounts = barLabels.map(label => counts[labels.indexOf(label)]);
    const barColors = barLabels.map(label => statusColors[label]);
    const barContext = document
        .getElementById("compliance-bar-chart")
        .getContext("2d");

    const barChartCard = document.getElementById("compliance-bar-chart").closest(".chart-card");
    barChartCard.classList.toggle("hidden", barLabels.length === 0);
    if (barLabels.length) complianceBarChart = new Chart(barContext, {
        type: "bar",

        data: {
            labels: barLabels,

            datasets: [{
                label: "Number of Clauses",
                data: barCounts,
                backgroundColor: barColors,
                borderRadius: 7,
                barThickness: 30
            }]
        },

        options: {
            indexAxis: "y",
            responsive: true,
            maintainAspectRatio: false,

            scales: {
                x: {
                    beginAtZero: true,
                    title: { display: true, text: "Number of clauses" },
                    ticks: {
                        precision: 0
                    }
                },

                y: {
                    title: { display: true, text: "Vendor declaration" },
                    grid: {
                        display: false
                    }
                }
            },

            plugins: {
                legend: {
                    display: false
                },

                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const index = context.dataIndex;

                            return `${context.raw} clauses (${percentages[index]}%)`;
                        }
                    }
                }
            },
            animation: { duration: 500 }
        },
        plugins: [{
            id: "complianceValueLabels",
            afterDatasetsDraw(chart) {
                const context = chart.ctx;
                context.save();
                context.font = "600 13px Arial";
                context.fillStyle = "#243247";
                context.textBaseline = "middle";
                chart.getDatasetMeta(0).data.forEach((bar, index) => {
                    context.fillText(String(barCounts[index]), bar.x + 8, bar.y);
                });
                context.restore();
            }
        }]
    });

    // Display analyzer warnings.
    const warnings = Array.isArray(data.warnings)
        ? data.warnings
        : [];

    visualizationWarningsList.replaceChildren();

    warnings.forEach((warning) => {
        const item = document.createElement("li");
        item.textContent = warning;
        visualizationWarningsList.appendChild(item);
    });

    visualizationWarnings.classList.toggle(
        "hidden",
        warnings.length === 0
    );

    visualizationNotice.textContent =
        data.notice ||
        "These figures represent vendor-declared compliance statuses. They do not independently establish contractual compliance.";

    visualizationContainer.classList.remove("hidden");

    answerSection.classList.remove("hidden");

    answerSection.scrollIntoView({
        behavior: "smooth",
        block: "start"
    });
}

function displayComparison(data) {
    const comparison = data.comparison || {};
    const submissions = comparison.submissions || [];
    if (typeof Chart === "undefined" || submissions.length === 0) {
        throw new Error("Comparison chart data is unavailable.");
    }
    destroyComplianceCharts();
    answerText.textContent = data.answer || "Submission comparison";
    sourcesContainer.classList.add("hidden");
    retrievedCount.textContent = "";
    visualizationTitle.textContent = "Submission Compliance Comparison";
    visualizationSubtitle.textContent = "Evidence excerpts are grouped by submission. Vendor identity and declared compliance remain subject to the review notes.";
    visualizationSummary.replaceChildren();
    const notes = [...(comparison.warnings || [])];
    const table = document.createElement("table");
    table.className = "comparison-table";
    const header = document.createElement("tr");
    ["Submission / vendor", "Documents", ...(comparison.topics || [])].forEach((label) => {
        const th = document.createElement("th"); th.textContent = label; header.appendChild(th);
    });
    table.appendChild(header);
    submissions.forEach((submission) => {
        const row = document.createElement("tr");
        const identityCell = document.createElement("td"); identityCell.textContent = `${submission.label} — ${submission.vendor_identity}`; row.appendChild(identityCell);
        const filesCell = document.createElement("td"); filesCell.textContent = (submission.documents || []).join(", ") || "Not found"; row.appendChild(filesCell);
        (comparison.topics || []).forEach((topic) => {
            const cell = document.createElement("td");
            const items = (submission.topic_evidence || {})[topic] || [];
            if (!items.length) { cell.textContent = "Not found in extracted text"; }
            else { items.forEach((item, index) => { const p = document.createElement("p"); p.textContent = `${item.file_name} — ${item.location}${item.requires_ocr ? " (OCR required)" : ""}: ${item.text}`; cell.appendChild(p); if (index < items.length - 1) cell.appendChild(document.createElement("hr")); }); }
            row.appendChild(cell);
        });
        table.appendChild(row);
    });
    const tableWrap = document.createElement("div"); tableWrap.className = "comparison-table-wrap"; tableWrap.appendChild(table); visualizationSummary.appendChild(tableWrap);
    visualizationWarningsList.replaceChildren();
    notes.forEach((note) => {
        const item = document.createElement("li");
        item.textContent = note;
        visualizationWarningsList.appendChild(item);
    });
    visualizationWarnings.classList.toggle("hidden", notes.length === 0);
    visualizationNotice.textContent = "Values are deterministic counts of extracted declarations, not independent verification of compliance.";
    const doughnutCard = document.getElementById("compliance-doughnut-chart").closest(".chart-card");
    doughnutCard.classList.add("hidden");
    const barCard = document.getElementById("compliance-bar-chart").closest(".chart-card");
    const deliveryMetric = comparison.metrics && comparison.metrics.delivery_days;
    barCard.querySelector("h4").textContent = deliveryMetric && deliveryMetric.length ? "Extracted Delivery Period (days)" : "Declaration Counts by Submission";
    const chartValues = submissions.flatMap(item => Object.values(item.counts || {}));
    barCard.classList.toggle("hidden", !chartValues.some(value => Number(value) > 0));
    if (chartValues.some(value => Number(value) > 0)) complianceBarChart = new Chart(document.getElementById("compliance-bar-chart").getContext("2d"), {
        type: "bar",
        data: deliveryMetric && deliveryMetric.length ? { labels: deliveryMetric.map(item => item.label), datasets: [{ label: "Days", data: deliveryMetric.map(item => item.value), backgroundColor: "#2563eb" }] } : { labels: submissions.map(item => item.label), datasets: (comparison.statuses || []).map((status, index) => ({ label: status, data: submissions.map(item => item.counts[status] || 0), backgroundColor: ["#16A34A", "#F59E0B", "#DC2626", "#64748B"][index] })) },
        options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, ticks: { precision: 0 } } } }
    });
    visualizationContainer.classList.remove("hidden");
    answerSection.classList.remove("hidden");
}

    // Display existing chatbot answer
    function displayAnswer(data) {
        visualizationContainer.classList.add("hidden");
        destroyComplianceCharts();
        document.getElementById("compliance-doughnut-chart").closest(".chart-card").classList.remove("hidden");
        document.getElementById("compliance-bar-chart").closest(".chart-card").querySelector("h4").textContent = "Clause Status Comparison";
        answerText.innerHTML = renderMarkdown(data.answer || "No answer was returned.");

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
