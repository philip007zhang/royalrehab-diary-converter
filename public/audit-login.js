const elements = {
  form: document.querySelector("#audit-login-form"),
  message: document.querySelector("#audit-login-message"),
  password: document.querySelector("#audit-password"),
  username: document.querySelector("#audit-username")
};

elements.form.addEventListener("submit", async (event) => {
  event.preventDefault();
  elements.message.textContent = "Signing in...";

  try {
    const response = await fetch("/api/admin/login", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        username: elements.username.value.trim(),
        password: elements.password.value
      })
    });

    if (!response.ok) {
      const payload = await response.json().catch(() => ({ error: "Login failed" }));
      throw new Error(payload.error || "Login failed");
    }

    window.location.assign("/auditlog");
  } catch (error) {
    elements.message.textContent = error instanceof Error ? error.message : "Login failed";
  }
});
