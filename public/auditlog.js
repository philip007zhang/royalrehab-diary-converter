const elements = {
  body: document.querySelector("#auditlog-body"),
  logoutButton: document.querySelector("#audit-logout-btn"),
  message: document.querySelector("#auditlog-message"),
  refreshButton: document.querySelector("#audit-refresh-btn"),
  userPill: document.querySelector("#audit-user-pill")
};

elements.refreshButton.addEventListener("click", () => {
  void loadAuditLog();
});

elements.logoutButton.addEventListener("click", async () => {
  await fetch("/api/admin/logout", { method: "POST" });
  window.location.assign("/auditlogin");
});

void loadAuditLog();

async function loadAuditLog() {
  elements.message.textContent = "Loading audit log...";

  try {
    const response = await fetch("/api/auditlog", { cache: "no-store" });
    if (response.status === 401) {
      window.location.assign("/auditlogin");
      return;
    }

    if (!response.ok) {
      throw new Error(`Audit log request failed with status ${response.status}`);
    }

    const payload = await response.json();
    elements.userPill.textContent = payload.username ? `Admin: ${payload.username}` : "Admin";
    renderEntries(payload.entries ?? []);
    elements.message.textContent = `${(payload.entries ?? []).length} audit entr${(payload.entries ?? []).length === 1 ? "y" : "ies"} loaded.`;
  } catch (error) {
    elements.message.textContent = error instanceof Error ? error.message : "Failed to load audit log";
  }
}

function renderEntries(entries) {
  if (entries.length === 0) {
    elements.body.innerHTML = `
      <tr class="empty-row">
        <td colspan="5">No audit entries yet.</td>
      </tr>
    `;
    return;
  }

  elements.body.innerHTML = entries
    .map((entry) => {
      const timestamp = new Date(entry.timestamp).toLocaleString("en-AU");
      const details = JSON.stringify(entry.details ?? {});
      const actorIp = entry.actor?.ip ?? "";
      return `
        <tr>
          <td>${escapeHtml(timestamp)}</td>
          <td>${escapeHtml(entry.activity ?? "")}</td>
          <td>${escapeHtml(entry.status ?? "")}</td>
          <td><code>${escapeHtml(details)}</code></td>
          <td>${escapeHtml(actorIp)}</td>
        </tr>
      `;
    })
    .join("");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
