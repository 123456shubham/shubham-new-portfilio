(() => {
    "use strict";

    const root = document.documentElement;
    const themeToggle = document.querySelector("#theme-toggle");
    const menuToggle = document.querySelector("#menu-toggle");
    const mainNav = document.querySelector("#main-nav");

    function updateThemeIcon() {
        if (!themeToggle) return;
        const dark = root.dataset.theme === "dark";
        themeToggle.innerHTML = dark
            ? '<i class="fa-solid fa-sun"></i>'
            : '<i class="fa-solid fa-moon"></i>';
        themeToggle.setAttribute(
            "aria-label",
            dark ? "Switch to light mode" : "Switch to dark mode"
        );
    }

    updateThemeIcon();

    themeToggle?.addEventListener("click", () => {
        const next = root.dataset.theme === "dark" ? "light" : "dark";
        root.dataset.theme = next;
        localStorage.setItem("portfolio-theme", next);
        updateThemeIcon();
    });

    menuToggle?.addEventListener("click", () => {
        const open = mainNav.classList.toggle("open");
        menuToggle.setAttribute("aria-expanded", String(open));
        menuToggle.innerHTML = open
            ? '<i class="fa-solid fa-xmark"></i>'
            : '<i class="fa-solid fa-bars"></i>';
    });

    mainNav?.querySelectorAll("a").forEach((link) => {
        link.addEventListener("click", () => {
            mainNav.classList.remove("open");
            menuToggle?.setAttribute("aria-expanded", "false");
            if (menuToggle) menuToggle.innerHTML = '<i class="fa-solid fa-bars"></i>';
        });
    });

    const observer = new IntersectionObserver(
        (entries) => {
            entries.forEach((entry) => {
                if (entry.isIntersecting) {
                    entry.target.classList.add("visible");
                    observer.unobserve(entry.target);
                }
            });
        },
        { threshold: 0.12, rootMargin: "0px 0px -30px 0px" }
    );

    document.querySelectorAll(".reveal").forEach((element, index) => {
        element.style.transitionDelay = `${Math.min(index % 4, 3) * 65}ms`;
        observer.observe(element);
    });

    const launcher = document.querySelector("#chat-launcher");
    const panel = document.querySelector("#chat-panel");
    const closeChat = document.querySelector("#close-chat");
    const clearChat = document.querySelector("#clear-chat");
    const heroChat = document.querySelector("#open-chat-hero");
    const form = document.querySelector("#chat-form");
    const input = document.querySelector("#chat-input");
    const submit = document.querySelector("#chat-submit");
    const messages = document.querySelector("#chat-messages");
    const suggestions = document.querySelector("#chat-suggestions");

    const HISTORY_KEY = "shubham-portfolio-chat-history-v1";
    let history = [];

    try {
        const saved = JSON.parse(localStorage.getItem(HISTORY_KEY) || "[]");
        history = Array.isArray(saved) ? saved.slice(-14) : [];
    } catch {
        history = [];
    }

    function escapeText(text) {
        return String(text ?? "");
    }

    function scrollChat() {
        messages.scrollTop = messages.scrollHeight;
    }

    function createMessage(role, text = "") {
        const wrapper = document.createElement("div");
        wrapper.className = `message ${role}`;

        if (role === "assistant") {
            const avatar = document.createElement("div");
            avatar.className = "message-avatar";
            avatar.innerHTML = '<i class="fa-solid fa-sparkles"></i>';
            wrapper.appendChild(avatar);
        }

        const bubble = document.createElement("div");
        bubble.className = "message-bubble";
        bubble.textContent = escapeText(text);
        wrapper.appendChild(bubble);

        messages.appendChild(wrapper);
        scrollChat();
        return { wrapper, bubble };
    }

    function createTyping() {
        const wrapper = document.createElement("div");
        wrapper.className = "message assistant typing-wrapper";
        wrapper.innerHTML = `
            <div class="message-avatar"><i class="fa-solid fa-sparkles"></i></div>
            <div class="message-bubble typing"><i></i><i></i><i></i></div>
        `;
        messages.appendChild(wrapper);
        scrollChat();
        return wrapper;
    }

    function persistHistory() {
        localStorage.setItem(HISTORY_KEY, JSON.stringify(history.slice(-14)));
    }

    function restoreHistory() {
        if (!history.length) return;
        messages.innerHTML = "";
        history.forEach((entry) => createMessage(entry.role, entry.content));
    }

    restoreHistory();

    function setChatOpen(open) {
        panel.classList.toggle("open", open);
        panel.setAttribute("aria-hidden", String(!open));
        if (open) setTimeout(() => input.focus(), 120);
    }

    launcher?.addEventListener("click", () => setChatOpen(!panel.classList.contains("open")));
    closeChat?.addEventListener("click", () => setChatOpen(false));
    heroChat?.addEventListener("click", () => setChatOpen(true));

    document.querySelectorAll(".open-chat").forEach((button) => {
        button.addEventListener("click", () => {
            setChatOpen(true);
            input.value = `I want to build a project similar to ${button.dataset.project}. Please explain the recommended modules.`;
            input.focus();
        });
    });

    clearChat?.addEventListener("click", () => {
        history = [];
        persistHistory();
        messages.innerHTML = "";
        createMessage(
            "assistant",
            "Chat cleared. Tell me what you want to build."
        );
    });

    suggestions?.querySelectorAll("button").forEach((button) => {
        button.addEventListener("click", () => {
            input.value = button.textContent.trim();
            form.requestSubmit();
        });
    });

    input?.addEventListener("input", () => {
        input.style.height = "auto";
        input.style.height = `${Math.min(input.scrollHeight, 105)}px`;
    });

    input?.addEventListener("keydown", (event) => {
        if (event.key === "Enter" && !event.shiftKey) {
            event.preventDefault();
            form.requestSubmit();
        }
    });

    async function streamReply(message) {
        createMessage("user", message);
        const priorHistory = history.slice(-14);
        history.push({ role: "user", content: message });
        persistHistory();

        const typing = createTyping();
        submit.disabled = true;
        input.disabled = true;

        let assistantBubble = null;
        let assistantText = "";

        try {
            const response = await fetch("/api/chat-stream", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Accept": "text/event-stream",
                },
                body: JSON.stringify({
                    message,
                    history: priorHistory,
                }),
            });

            if (!response.ok || !response.body) {
                let reason = "Unable to connect to the assistant.";
                try {
                    const error = await response.json();
                    reason = error.error || reason;
                } catch {
                    // Keep default message.
                }
                throw new Error(reason);
            }

            const reader = response.body.getReader();
            const decoder = new TextDecoder("utf-8");
            let buffer = "";

            while (true) {
                const { value, done } = await reader.read();
                if (done) break;

                buffer += decoder.decode(value, { stream: true });
                const eventBlocks = buffer.split("\n\n");
                buffer = eventBlocks.pop() || "";

                for (const block of eventBlocks) {
                    const dataLine = block
                        .split("\n")
                        .find((line) => line.startsWith("data:"));

                    if (!dataLine) continue;

                    const payload = JSON.parse(dataLine.slice(5).trim());

                    if (payload.type === "notice") {
                        continue;
                    }

                    if (payload.type === "delta") {
                        if (!assistantBubble) {
                            typing.remove();
                            assistantBubble = createMessage("assistant", "").bubble;
                        }
                        assistantText += payload.text || "";
                        assistantBubble.textContent = assistantText;
                        scrollChat();
                    }
                }
            }

            typing.remove();

            if (!assistantText.trim()) {
                assistantText = "I could not generate a response. Please try again.";
                createMessage("assistant", assistantText);
            }

            history.push({ role: "assistant", content: assistantText });
            persistHistory();
        } catch (error) {
            typing.remove();
            assistantText = error.message || "The assistant is temporarily unavailable.";
            createMessage("assistant", assistantText);
            history.push({ role: "assistant", content: assistantText });
            persistHistory();
        } finally {
            submit.disabled = false;
            input.disabled = false;
            input.focus();
        }
    }

    form?.addEventListener("submit", async (event) => {
        event.preventDefault();
        const message = input.value.trim();
        if (!message || submit.disabled) return;

        input.value = "";
        input.style.height = "auto";
        suggestions.style.display = "none";
        await streamReply(message);
    });
})();
