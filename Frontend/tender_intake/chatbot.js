document.addEventListener("DOMContentLoaded", () => {
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

    const endpoint = form.dataset.endpoint;

    questionInput.addEventListener("input", () => {
        questionCount.textContent = `${questionInput.value.length} / 5000`;
    });

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
            const response = await fetch(endpoint, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    question: question,
                    category: categoryInput.value || null,
                    top_k: 5
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

    function showError(message) {
        errorMessage.textContent = message;
        errorMessage.classList.remove("hidden");
    }
});
