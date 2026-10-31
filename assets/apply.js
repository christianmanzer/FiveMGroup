(() => {
  const form = document.querySelector(".apply-form");
  if (!form) return;

  const status = form.querySelector(".form-status");
  const button = form.querySelector("button[type=submit]");
  const file = form.querySelector("input[type=file]");
  const fileName = form.querySelector(".upload-name");
  const MAX_BYTES = 3 * 1024 * 1024;
  const EMPTY_LABEL = "Add a File";

  const setStatus = (message, state) => {
    status.textContent = message;
    status.dataset.state = state || "";
  };

  file.addEventListener("change", () => {
    const chosen = file.files[0];
    if (!chosen) {
      fileName.textContent = EMPTY_LABEL;
      return;
    }
    if (chosen.size > MAX_BYTES) {
      file.value = "";
      fileName.textContent = EMPTY_LABEL;
      setStatus("That file is over 3 MB. Please choose a smaller one.", "error");
      return;
    }
    fileName.textContent = chosen.name;
    setStatus("");
  });

  const nothingChecked = (name) => !form.querySelector('input[name="' + name + '"]:checked');

  form.addEventListener("submit", async (event) => {
    event.preventDefault();

    if (!form.checkValidity()) {
      form.reportValidity();
      return;
    }
    if (nothingChecked("position")) {
      setStatus("Choose at least one position.", "error");
      form.querySelector('input[name="position"]').focus();
      return;
    }
    if (nothingChecked("availability")) {
      setStatus("Choose at least one availability option.", "error");
      form.querySelector('input[name="availability"]').focus();
      return;
    }
    if (nothingChecked("timeline")) {
      setStatus("Choose a timeline.", "error");
      form.querySelector('input[name="timeline"]').focus();
      return;
    }

    button.disabled = true;
    setStatus("Sending…");

    try {
      const response = await fetch(form.getAttribute("action"), { method: "POST", body: new FormData(form) });
      const data = await response.json().catch(() => ({}));
      if (!response.ok || !data.ok) {
        throw new Error(data.error || "Something went wrong. Please try again or email hello@fivem.group.");
      }
      form.reset();
      fileName.textContent = EMPTY_LABEL;
      setStatus("Thank you! We received your application and will be in touch soon.", "ok");
    } catch (error) {
      setStatus(error.message || "Something went wrong. Please try again or email hello@fivem.group.", "error");
    } finally {
      button.disabled = false;
    }
  });
})();
