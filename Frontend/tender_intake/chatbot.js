document.addEventListener("DOMContentLoaded", () => {
    // Existing question-and-answer elements
    const form = document.getElementById("chatbot-form");
    const questionInput = document.getElementById("question");
    const categoryInput = document.getElementById("category");
    const askButton = document.getElementById("ask-button");

    const loading = document.getElementById("loading");
    const errorMessage = document.getElementById("error-message");
    const answerSection = document.getElementById("answer-section");
    const conversationMessages = document.getElementById("conversation-messages");
    const welcomeScreen = document.getElementById("welcome-screen");
    const newChatButton = document.getElementById("new-chat-button");
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

    // Vendor document attachment elements
    const fileInput = document.getElementById("vendor-files");
    const attachmentList = document.getElementById("attachment-list");
    const attachmentError = document.getElementById("attachment-error");

    const questionEndpoint = form.dataset.endpoint;
    const uploadEndpoint = form.dataset.uploadEndpoint;
    // Keep the currently active vendor document submission.
    let activeSubmissionId = null;
    let isProcessing = false;
    let selectedAttachments = [];
    let pendingAttachmentUpload = null;

    const allowedExtensions = new Set([".pdf", ".doc", ".docx", ".xlsx", ".xlsm"]);
    const maxFiles = 20;
    const maxFileSize = 10 * 1024 * 1024;
    const maxTotalSize = 50 * 1024 * 1024;

    // Existing question character counter
    questionInput.addEventListener("input", () => {
        questionCount.textContent = `${questionInput.value.length} / 5000`;
        questionInput.style.height = "auto";
        questionInput.style.height = `${Math.min(questionInput.scrollHeight, 180)}px`;
        updateSendButton();
    });

    questionInput.addEventListener("keydown", (event) => {
        if (event.key === "Enter" && !event.shiftKey && !event.isComposing) {
            event.preventDefault();
            if (!isProcessing && (questionInput.value.trim() || selectedAttachments.length)) form.requestSubmit();
        }
    });

    document.querySelectorAll(".suggestion-card").forEach((suggestion) => {
        suggestion.addEventListener("click", () => {
            questionInput.value = suggestion.dataset.prompt || suggestion.textContent.trim();
            questionInput.dispatchEvent(new Event("input", { bubbles: true }));
            form.requestSubmit();
        });
    });

    newChatButton.addEventListener("click", () => {
        conversationMessages.replaceChildren();
        welcomeScreen.classList.remove("hidden");
        questionInput.value = "";
        questionInput.style.height = "auto";
        questionCount.textContent = "0 / 5000";
        errorMessage.classList.add("hidden");
        loading.classList.add("hidden");
        clearSelectedAttachments();
        pendingAttachmentUpload = null;
        visualizationContainer.classList.add("hidden");
        if (visualizationContainer.parentElement !== answerSection) {
            answerSection.appendChild(visualizationContainer);
        }
        destroyComplianceCharts();
        updateSendButton();
        questionInput.focus();
    });

    function updateSendButton() {
        askButton.disabled = isProcessing || (!questionInput.value.trim() && selectedAttachments.length === 0);
        newChatButton.disabled = isProcessing;
        fileInput.disabled = isProcessing;
    }

    updateSendButton();

    fileInput.addEventListener("change", () => {
        const additions = [];
        const messages = [];
        const knownNames = new Set(selectedAttachments.map(file => file.name.toLocaleLowerCase()));
        let totalSize = selectedAttachments.reduce((total, file) => total + file.size, 0);

        Array.from(fileInput.files || []).forEach((file) => {
            const extension = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();
            const normalizedName = file.name.toLocaleLowerCase();

            if (!allowedExtensions.has(extension)) {
                messages.push(`${file.name}: choose a PDF, DOC, DOCX, XLSX or XLSM file.`);
                return;
            }
            if (file.size === 0) {
                messages.push(`${file.name}: the file is empty.`);
                return;
            }
            if (file.size > maxFileSize) {
                messages.push(`${file.name}: each file must be 10 MB or smaller.`);
                return;
            }
            if (knownNames.has(normalizedName)) {
                messages.push(`${file.name}: this filename is already attached.`);
                return;
            }
            if (selectedAttachments.length + additions.length >= maxFiles) {
                messages.push(`You can attach a maximum of ${maxFiles} files.`);
                return;
            }
            if (totalSize + file.size > maxTotalSize) {
                messages.push("The combined attachment size cannot exceed 50 MB.");
                return;
            }

            additions.push(file);
            knownNames.add(normalizedName);
            totalSize += file.size;
        });

        selectedAttachments = selectedAttachments.concat(additions);
        fileInput.value = "";
        renderAttachments();
        showAttachmentValidation(messages.join(" "));
        updateSendButton();
    });

    function renderAttachments() {
        attachmentList.replaceChildren();
        attachmentList.classList.toggle("hidden", selectedAttachments.length === 0);

        selectedAttachments.forEach((file, index) => {
            const chip = document.createElement("div");
            chip.className = "attachment-chip";
            const type = document.createElement("span");
            type.className = "attachment-type";
            type.setAttribute("aria-hidden", "true");
            type.textContent = file.name.substring(file.name.lastIndexOf(".") + 1).toUpperCase();
            const name = document.createElement("span");
            name.className = "attachment-name";
            name.textContent = file.name;
            name.title = `${file.name} (${(file.size / (1024 * 1024)).toFixed(2)} MB)`;
            const remove = document.createElement("button");
            remove.type = "button";
            remove.className = "attachment-remove";
            remove.setAttribute("aria-label", `Remove ${file.name}`);
            remove.textContent = "×";
            remove.addEventListener("click", () => {
                selectedAttachments.splice(index, 1);
                renderAttachments();
                updateSendButton();
            });
            chip.append(type, name, remove);
            attachmentList.appendChild(chip);
        });
    }

    function showAttachmentValidation(message) {
        attachmentError.textContent = message;
        attachmentError.classList.toggle("hidden", !message);
    }

    function clearSelectedAttachments() {
        selectedAttachments = [];
        fileInput.value = "";
        renderAttachments();
        showAttachmentValidation("");
    }

    function appendCopyAction(target, textToCopy) {
        const actions = document.createElement("div");
        actions.className = "message-actions";
        const button = document.createElement("button");
        button.type = "button";
        button.className = "copy-answer-button";
        button.textContent = "Copy answer";
        button.addEventListener("click", async () => {
            try {
                await navigator.clipboard.writeText(textToCopy);
                button.textContent = "Copied";
                setTimeout(() => { button.textContent = "Copy answer"; }, 1400);
            } catch {
                button.textContent = "Copy unavailable";
            }
        });
        actions.appendChild(button);
        target.appendChild(actions);
    }

    function appendUserMessage(text, files = []) {
        const row = document.createElement("article");
        row.className = "message-row user-message";
        const bubble = document.createElement("div");
        bubble.className = "message-bubble";
        const messageText = document.createElement("p");
        messageText.className = "user-message-text";
        messageText.textContent = text;
        bubble.appendChild(messageText);
        if (files.length) {
            const fileList = document.createElement("ul");
            fileList.className = "message-attachments";
            files.forEach((file) => {
                const item = document.createElement("li");
                item.textContent = file.name;
                fileList.appendChild(item);
            });
            bubble.appendChild(fileList);
        }
        const avatar = document.createElement("div");
        avatar.className = "user-avatar";
        avatar.setAttribute("aria-hidden", "true");
        avatar.textContent = "You";
        row.append(bubble, avatar);
        conversationMessages.appendChild(row);
        scrollConversationToEnd();
    }

    function createAssistantMessage() {
        const row = document.createElement("article");
        row.className = "message-row assistant-message";
        const avatar = document.createElement("img");
        avatar.className = "assistant-avatar";
        avatar.src = "/chatbot.jpg";
        avatar.alt = "ComOps";
        const bubble = document.createElement("div");
        bubble.className = "message-bubble";
        row.append(avatar, bubble);
        conversationMessages.appendChild(row);
        return bubble;
    }

    function appendTypingIndicator(message = "ComOps is preparing your answer…") {
        const bubble = createAssistantMessage();
        bubble.classList.add("typing-bubble");
        bubble.setAttribute("aria-label", message);
        const label = document.createElement("span");
        label.className = "typing-label";
        label.textContent = message;
        bubble.appendChild(label);
        for (let index = 0; index < 3; index += 1) {
            bubble.appendChild(document.createElement("span"));
        }
        scrollConversationToEnd();
        return bubble.closest(".message-row");
    }

    function appendAssistantText(text, isError = false) {
        const bubble = createAssistantMessage();
        bubble.classList.toggle("assistant-error", isError);
        const content = document.createElement("div");
        content.className = "answer-text";
        content.textContent = text;
        bubble.appendChild(content);
        scrollConversationToEnd();
    }

    function scrollConversationToEnd() {
        answerSection.scrollTop = answerSection.scrollHeight;
    }

        
// Existing question submission with visualization detection
form.addEventListener("submit", async (event) => {
    event.preventDefault();

    const attachments = selectedAttachments.slice();
    const question = questionInput.value.trim() ||
        (attachments.length ? "Please analyse the attached vendor documents." : "");

    if (!question) {
        showError("Please enter a procurement question.");
        return;
    }

    if (isProcessing) return;
    isProcessing = true;
    updateSendButton();
    welcomeScreen.classList.add("hidden");
    appendUserMessage(question, attachments);
    let typingIndicator = appendTypingIndicator(
        attachments.length ? "Uploading and processing your vendor documents…" : "ComOps is preparing your answer…"
    );
    errorMessage.classList.add("hidden");

    try {
        if (attachments.length) {
            let uploadData;
            const canReuseUpload = pendingAttachmentUpload &&
                pendingAttachmentUpload.category === (categoryInput.value || "") &&
                pendingAttachmentUpload.files.length === attachments.length &&
                pendingAttachmentUpload.files.every((file, index) => file === attachments[index]);

            if (canReuseUpload) {
                uploadData = pendingAttachmentUpload.data;
            } else {
                const formData = new FormData();
                attachments.forEach((file) => formData.append("files", file));
                formData.append("category", categoryInput.value || "");

                const uploadResponse = await fetch(uploadEndpoint, {
                    method: "POST",
                    body: formData
                });
                try {
                    uploadData = await uploadResponse.json();
                } catch {
                    throw new Error("The document upload service returned an unreadable response.");
                }
                if (!uploadResponse.ok) {
                    throw new Error(uploadData.error || "Unable to upload the documents.");
                }
                if (!uploadData.submission_id) {
                    throw new Error("The upload completed without a submission ID. Please try again.");
                }
            }

            activeSubmissionId = uploadData.submission_id;
            const processedFiles = Number(uploadData.summary?.files_processed ?? 0);
            if (processedFiles === 0) {
                typingIndicator.remove();
                displayUploadResult(uploadData);
                pendingAttachmentUpload = null;
                throw new Error("No attached documents could be extracted. Check the file status and try again.");
            }

            pendingAttachmentUpload = {
                files: attachments,
                category: categoryInput.value || "",
                data: uploadData
            };
            typingIndicator.remove();
            displayUploadResult(uploadData);
            typingIndicator = appendTypingIndicator("Reviewing the uploaded evidence…");
        }

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
                displayComparison(data, question);
            } else {
                displayAnswer(data);
            }
        }

        questionInput.value = "";
        questionInput.style.height = "auto";
        questionCount.textContent = "0 / 5000";
        if (attachments.length) {
            clearSelectedAttachments();
            pendingAttachmentUpload = null;
        }

    } catch (error) {
        appendAssistantText(error.message || "I couldn’t complete that request. Please try again.", true);
    } finally {
        typingIndicator.remove();
        loading.classList.add("hidden");
        isProcessing = false;
        updateSendButton();
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

    const chartData = data.chart_data || {};
    const labels = chartData.labels || [];
    const counts = chartData.counts || [];
    const percentages = chartData.percentages || [];
    const hasChartData = labels.length > 0 &&
        labels.length === counts.length &&
        labels.length === percentages.length;

    if (hasChartData && typeof Chart === "undefined") {
        throw new Error(
            "Chart.js could not be loaded. Please check your internet connection."
        );
    }

    destroyComplianceCharts();
    const messageBubble = createAssistantMessage();
    messageBubble.appendChild(visualizationContainer);
    visualizationContainer.classList.remove("hidden");
    const doughnutCard = document.getElementById("compliance-doughnut-chart").closest(".chart-card");
    const barChartCard = document.getElementById("compliance-bar-chart").closest(".chart-card");
    doughnutCard.classList.toggle("hidden", !hasChartData);
    barChartCard.classList.toggle("hidden", !hasChartData);
    doughnutCard.querySelector("h4").textContent = "Compliance Declaration Distribution";
    barChartCard.querySelector("h4").textContent = "Compliance Declaration Counts";

    visualizationTitle.textContent =
        data.title || "Vendor Compliance Analysis";

    visualizationSubtitle.textContent = data.clause_matching_performed === false
        ? "Clause-level matching was not performed. The figures and charts below count extracted vendor declarations only; they are not independently verified contractual compliance."
        : data.analysis_complete
            ? "Extracted vendor declarations were matched to the identified clauses. These are vendor-declared results, not independent verification."
            : "Preliminary analysis — review the warnings before relying on these results. Declaration counts are based on extracted vendor statements.";

    // Display summary figures.
    const summary = data.summary || {};
    const declarationSummary = data.declaration_summary || summary;
    const breakdown = declarationSummary.breakdown || {};
    const totalClauses = data.clause_count ?? summary.total_clauses ?? 0;
    const totalDeclarations = declarationSummary.total_declarations ?? data.declaration_count ?? 0;

    const stats = [
        {
            label: "Extracted Clauses",
            value: totalClauses
        },
        {
            label: "Compliance Declarations",
            value: totalDeclarations
        },
        {
            label: "Complied",
            value: breakdown["Complied"]?.count ?? declarationSummary.complied ?? 0
        },
        {
            label: "Partially Complied",
            value: breakdown["Partially Complied"]?.count ?? declarationSummary.partially_complied ?? 0
        },
        {
            label: "Not Complied",
            value: breakdown["Not Complied"]?.count ?? declarationSummary.not_complied ?? 0
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
    appendCopyAction(
        messageBubble,
        [data.title || "Vendor Compliance Analysis", ...stats.map(stat => `${stat.label}: ${stat.value}`), data.notice || ""].filter(Boolean).join("\n")
    );

    const statusColors = {
        "Complied": "#16A34A",
        "Not Complied": "#DC2626",
        "Partially Complied": "#F59E0B",
        "Not Specified": "#64748B",
        "Unclear": "#8B5CF6",
        "OCR Required": "#0EA5E9"
    };
    const chartColors = labels.map(label => statusColors[label] || "#64748B");

    // Display analyzer warnings even when no chart can be drawn.
    const warnings = Array.isArray(data.warnings) ? data.warnings : [];
    visualizationWarningsList.replaceChildren();
    warnings.forEach((warning) => {
        const item = document.createElement("li");
        item.textContent = warning;
        visualizationWarningsList.appendChild(item);
    });
    visualizationWarnings.classList.toggle("hidden", warnings.length === 0);
    visualizationNotice.textContent = data.notice ||
        "These figures represent vendor-declared compliance statuses. They do not independently establish contractual compliance.";

    // Doughnut chart.
    if (!hasChartData) {
        visualizationContainer.classList.remove("hidden");
        answerSection.classList.remove("hidden");
        scrollConversationToEnd();
        return;
    }

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

                            return `${context.label}: ${context.raw} declarations (${percentages[index]}%)`;
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

    barChartCard.classList.toggle("hidden", barLabels.length === 0);
    if (barLabels.length) complianceBarChart = new Chart(barContext, {
        type: "bar",

        data: {
            labels: barLabels,

            datasets: [{
                label: "Number of Compliance Declarations",
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
                    title: { display: true, text: "Number of declarations" },
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

                            return `${context.raw} declarations (${percentages[index]}%)`;
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

    visualizationContainer.classList.remove("hidden");

    answerSection.classList.remove("hidden");
    scrollConversationToEnd();
}

function displayComparison(data, question = "") {
    const comparison = data.comparison || {};
    const submissions = comparison.submissions || [];
    if (submissions.length === 0) {
        throw new Error("Comparison data is unavailable.");
    }
    const complianceChart = comparison.compliance_chart || {};
    const deliveryMetric = comparison.metrics && comparison.metrics.delivery_days;
    const deliveryChartDataValid = Array.isArray(deliveryMetric) && deliveryMetric.length > 0 &&
        deliveryMetric.every(item => item && typeof item.label === "string" && item.label.trim() !== "" &&
            typeof item.value === "number" && Number.isFinite(item.value) && item.value >= 0);
    const deliveryRequested = /\b(delivery|lead[ -]?time|days?|weeks?)\b/i.test(question);
    const useDeliveryChart = deliveryRequested && deliveryChartDataValid;
    const useComplianceChart = !deliveryRequested && complianceChart.available === true;
    const chartAvailable = useDeliveryChart || useComplianceChart;
    if (chartAvailable && typeof Chart === "undefined") {
        throw new Error("Chart.js could not be loaded. Please check your internet connection.");
    }
    destroyComplianceCharts();
    const messageBubble = createAssistantMessage();
    const answer = document.createElement("div");
    answer.className = "answer-text";
    answer.textContent = data.answer || "Submission comparison";
    messageBubble.appendChild(answer);
    appendCopyAction(messageBubble, answer.textContent);
    messageBubble.appendChild(visualizationContainer);
    visualizationContainer.classList.remove("hidden");
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
    if (deliveryRequested && !deliveryChartDataValid) {
        const warning = document.createElement("li");
        warning.textContent = "Delivery chart was suppressed because its data was missing or invalid.";
        visualizationWarningsList.appendChild(warning);
    }
    visualizationNotice.textContent = "Values are deterministic counts of extracted declarations, not independent verification of compliance.";
    const doughnutCard = document.getElementById("compliance-doughnut-chart").closest(".chart-card");
    doughnutCard.classList.add("hidden");
    const barCard = document.getElementById("compliance-bar-chart").closest(".chart-card");
    barCard.classList.toggle("hidden", !chartAvailable);
    if (useDeliveryChart) {
        barCard.querySelector("h4").textContent = "Extracted Delivery Period (days)";
        complianceBarChart = new Chart(document.getElementById("compliance-bar-chart").getContext("2d"), {
            type: "bar",
            data: {
                labels: deliveryMetric.map(item => item.label),
                datasets: [{ label: "Days", data: deliveryMetric.map(item => item.value), backgroundColor: "#2563EB" }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: { y: { beginAtZero: true, ticks: { precision: 0 }, title: { display: true, text: "Delivery period (days)" } }, x: { title: { display: true, text: "Submission / Vendor" } } },
                plugins: { legend: { display: true, position: "bottom" } }
            }
        });
    } else if (useComplianceChart) {
        const validChart = Array.isArray(complianceChart.labels) && complianceChart.labels.length >= 2 &&
            Array.isArray(complianceChart.datasets) && complianceChart.datasets.length === 2 &&
            complianceChart.datasets.every(dataset => Array.isArray(dataset.data) && dataset.data.length === complianceChart.labels.length && dataset.data.every(value => Number.isFinite(Number(value)) && Number(value) >= 0));
        if (!validChart) {
            barCard.classList.add("hidden");
            const warning = document.createElement("li");
            warning.textContent = "Compliance graph was suppressed because the comparison returned invalid chart values.";
            visualizationWarningsList.appendChild(warning);
            visualizationWarnings.classList.remove("hidden");
        } else {
            barCard.querySelector("h4").textContent = complianceChart.complete ? "Compliance Declarations by Submission" : "Extracted Compliance Declarations by Submission (Review Required)";
            complianceBarChart = new Chart(document.getElementById("compliance-bar-chart").getContext("2d"), {
        type: "bar",
                data: {
                    labels: complianceChart.labels,
                    datasets: complianceChart.datasets.map(dataset => ({
                        label: dataset.label,
                        data: dataset.data,
                        backgroundColor: dataset.label === "Complied" ? "#16A34A" : dataset.label === "Non-Complied" ? "#DC2626" : "#64748B",
                        borderRadius: 7
                    }))
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        y: { beginAtZero: true, ticks: { precision: 0 }, title: { display: true, text: complianceChart.y_axis_label || "Compliance declarations" } },
                        x: { title: { display: true, text: complianceChart.x_axis_label || "Submission / Vendor" } }
                    },
                    plugins: { legend: { display: true, position: "bottom" } }
                }
            });
        }
    } else {
        barCard.querySelector("h4").textContent = "Compliance Comparison";
    }
    visualizationWarnings.classList.toggle("hidden", visualizationWarningsList.children.length === 0);
    visualizationContainer.classList.remove("hidden");
    answerSection.classList.remove("hidden");
    scrollConversationToEnd();
}

    // Display existing chatbot answer
    function displayAnswer(data) {
        visualizationContainer.classList.add("hidden");
        destroyComplianceCharts();
        document.getElementById("compliance-doughnut-chart").closest(".chart-card").classList.remove("hidden");
        document.getElementById("compliance-bar-chart").closest(".chart-card").querySelector("h4").textContent = "Clause Status Comparison";
        const bubble = createAssistantMessage();
        const answerText = document.createElement("div");
        answerText.className = "answer-text";
        answerText.innerHTML = renderMarkdown(data.answer || "No answer was returned.");
        bubble.appendChild(answerText);

        const sources = Array.isArray(data.sources) ? data.sources : [];

        if (sources.length) {
            const sourcesContainer = document.createElement("details");
            sourcesContainer.className = "message-sources";
            const summary = document.createElement("summary");
            summary.textContent = `Sources (${sources.length})`;
            const list = document.createElement("ul");
            sources.forEach((source) => {
                const item = document.createElement("li");
                item.textContent = source;
                list.appendChild(item);
            });
            sourcesContainer.append(summary, list);
            bubble.appendChild(sourcesContainer);
        }

        appendCopyAction(bubble, data.answer || "");

        answerSection.classList.remove("hidden");
        scrollConversationToEnd();
    }

    // Display vendor submission processing results
    function displayUploadResult(data) {
        const summary = data.summary || {};
        const files = Array.isArray(data.files) ? data.files : [];
        const ocrPages = Array.isArray(data.pages_requiring_ocr)
            ? data.pages_requiring_ocr
            : [];
        const bubble = createAssistantMessage();
        bubble.classList.add("upload-summary-message");
        const heading = document.createElement("h3");
        heading.textContent = "Document processing results";
        const status = document.createElement("p");
        status.className = "upload-result-status";
        status.textContent = `Processing status: ${data.status || "Unknown"}`;
        bubble.append(heading, status);

        const metrics = document.createElement("div");
        metrics.className = "upload-result-metrics";
        [
            ["Files processed", summary.files_processed ?? 0],
            ["Files failed", summary.files_failed ?? 0],
            ["Sections extracted", summary.total_sections ?? 0],
            ["Pages requiring review", summary.pages_requiring_ocr ?? ocrPages.length]
        ].forEach(([label, value]) => {
            const item = document.createElement("div");
            const number = document.createElement("strong");
            number.textContent = value;
            const caption = document.createElement("span");
            caption.textContent = label;
            item.append(number, caption);
            metrics.appendChild(item);
        });
        bubble.appendChild(metrics);

        if (data.workspace_submission_count > 1) {
            const workspaceNote = document.createElement("p");
            workspaceNote.className = "upload-workspace-note";
            workspaceNote.textContent = `This tender workspace now contains ${data.workspace_submission_count} submissions.`;
            bubble.appendChild(workspaceNote);
        }

        const filesHeading = document.createElement("h4");
        filesHeading.textContent = "Uploaded documents";
        bubble.appendChild(filesHeading);
        const filesList = document.createElement("ul");
        filesList.className = "upload-result-files";

        files.forEach((file) => {
            const item = document.createElement("li");
            const fileName = file.file_name || "Unnamed document";
            const status = file.status || "Unknown";
            const sectionCount = file.section_count ?? 0;

            item.textContent =
                `${fileName} — ${status} (${sectionCount} sections)${file.error ? ` — ${file.error}` : ""}`;

            filesList.appendChild(item);
        });
        bubble.appendChild(filesList);

        if (ocrPages.length || (summary.pages_requiring_ocr ?? 0) > 0 || (summary.files_failed ?? 0) > 0) {
            const warning = document.createElement("p");
            warning.className = "upload-result-warning";
            warning.textContent = "Some files or pages require review. Check extraction status and OCR notes before relying on the analysis.";
            bubble.appendChild(warning);
        }
        scrollConversationToEnd();
    }

    function showError(message) {
        errorMessage.textContent = message;
        errorMessage.classList.remove("hidden");
    }

    // When embedded in the evaluation page drawer, reuse its panel controls.
    const expandButton = document.getElementById("comops-expand-button");
    const closeButton = document.getElementById("comops-close-button");
    try {
        const parentDocument = window.parent !== window ? window.parent.document : null;
        const parentExpand = parentDocument ? parentDocument.getElementById("comops-expand") : null;
        const parentClose = parentDocument ? parentDocument.getElementById("comops-close") : null;
        if (expandButton && closeButton && parentExpand && parentClose) {
            expandButton.hidden = false;
            closeButton.hidden = false;
            expandButton.addEventListener("click", () => parentExpand.click());
            closeButton.addEventListener("click", () => parentClose.click());
        }
    } catch (error) {
        // Standalone /chatbot page or cross-origin parent: keep buttons hidden.
    }

});
