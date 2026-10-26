(() => {
  const $ = (id) => document.getElementById(id);
  const fileInput = $("file"), drop = $("drop"), dropText = $("drop-text");
  const urlInput = $("url"), target = $("target"), btn = $("convert");
  const sourceLabel = $("source-label"), status = $("status");

  let catalogue = null;
  let mode = "upload";
  let source = null; // detected input format key

  const GOOGLE_RE = /docs\.google\.com\/(document|spreadsheets)\/(?:u\/\d+\/)?d\/[\w-]+/;

  function setStatus(text, kind = "") {
    status.textContent = text;
    status.className = "status " + kind;
  }

  function extOf(name) {
    let ext = (name.split(".").pop() || "").toLowerCase();
    return ext === "xlsm" ? "xlsx" : ext;
  }

  function setSource(fmt) {
    source = fmt && catalogue.inputs[fmt] ? fmt : null;
    target.innerHTML = "";
    if (!source) {
      sourceLabel.textContent = fmt ? `.${fmt} not supported` : "—";
      sourceLabel.classList.add("muted");
      target.disabled = true;
      target.add(new Option(mode === "upload" ? "Pick a file first" : "Paste a link first"));
      btn.disabled = true;
      return;
    }
    sourceLabel.textContent = catalogue.labels[source];
    sourceLabel.classList.remove("muted");
    for (const out of catalogue.inputs[source].outputs) {
      target.add(new Option(catalogue.labels[out], out));
    }
    target.disabled = false;
    btn.disabled = false;
  }

  function refresh() {
    setStatus("");
    if (mode === "upload") {
      const f = fileInput.files[0];
      dropText.innerHTML = f ? `<strong>${escapeHtml(f.name)}</strong>` : "<strong>Choose a file</strong> or drag it here";
      setSource(f ? extOf(f.name) : null);
    } else {
      const m = urlInput.value.match(GOOGLE_RE);
      setSource(m ? (m[1] === "document" ? "gdoc" : "gsheet") : null);
      if (urlInput.value && !m) setStatus("Paste a docs.google.com document or spreadsheet link.", "error");
    }
  }

  function escapeHtml(s) {
    return s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  document.querySelectorAll(".tab").forEach((tab) => {
    tab.addEventListener("click", () => {
      mode = tab.dataset.tab;
      document.querySelectorAll(".tab").forEach((t) => {
        t.classList.toggle("active", t === tab);
        t.setAttribute("aria-selected", t === tab);
      });
      $("panel-upload").hidden = mode !== "upload";
      $("panel-google").hidden = mode !== "google";
      refresh();
    });
  });

  fileInput.addEventListener("change", refresh);
  urlInput.addEventListener("input", refresh);

  ["dragenter", "dragover"].forEach((ev) =>
    drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach((ev) =>
    drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", (e) => {
    if (e.dataTransfer.files.length) {
      fileInput.files = e.dataTransfer.files;
      refresh();
    }
  });

  btn.addEventListener("click", async () => {
    const form = new FormData();
    form.append("target", target.value);
    if (mode === "upload") form.append("file", fileInput.files[0]);
    else form.append("url", urlInput.value.trim());

    btn.disabled = true;
    setStatus("Converting…");
    try {
      const resp = await fetch("/api/convert", { method: "POST", body: form });
      if (!resp.ok) {
        let msg = `Conversion failed (HTTP ${resp.status}).`;
        try { msg = (await resp.json()).error || msg; } catch (_) {}
        throw new Error(msg);
      }
      const blob = await resp.blob();
      const disp = resp.headers.get("Content-Disposition") || "";
      const m = disp.match(/filename\*=UTF-8''([^;]+)|filename="?([^";]+)"?/);
      const name = m ? decodeURIComponent(m[1] || m[2]) : `converted.${target.value}`;
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = name;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 10000);
      setStatus(`Done — downloaded ${name}`, "ok");
    } catch (err) {
      setStatus(err.message, "error");
    } finally {
      btn.disabled = !source;
    }
  });

  fetch("/api/formats").then((r) => r.json()).then((data) => {
    catalogue = data;
    const table = $("matrix");
    for (const [fmt, info] of Object.entries(data.inputs)) {
      const row = table.insertRow();
      row.insertCell().textContent = info.label;
      row.insertCell().textContent = info.outputs.map((o) => data.labels[o]).join(", ");
    }
    $("lo-note").textContent = data.libreoffice
      ? "LibreOffice detected: .doc files and Office → PDF use high-fidelity conversion."
      : "LibreOffice not installed: .doc input/output is unavailable; Office → PDF uses a simplified layout.";
    refresh();
  }).catch(() => setStatus("Could not load supported formats from the server.", "error"));
})();
