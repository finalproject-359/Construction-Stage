const AUTH_EMAIL = "costrack2026@gmail.com";
const AUTH_PASSWORD = "costrack2026";

const loginForm = document.getElementById("loginForm");
const emailInput = document.getElementById("email");
const passwordInput = document.getElementById("password");
const feedback = document.getElementById("loginFeedback");
const togglePasswordButton = document.getElementById("togglePassword");
const forgotPasswordButton = document.getElementById("forgotPasswordButton");
const capsLockHint = document.getElementById("capsLockHint");

if (sessionStorage.getItem("costrackAuth") === "authenticated") {
  window.location.replace("index.html");
}

togglePasswordButton?.addEventListener("click", () => {
  const isPassword = passwordInput.type === "password";
  passwordInput.type = isPassword ? "text" : "password";
  togglePasswordButton.setAttribute("aria-label", isPassword ? "Hide password" : "Show password");
});

forgotPasswordButton?.addEventListener("click", () => {
  feedback.textContent = "Please contact your system administrator to reset your credentials.";
  feedback.classList.add("info");
  feedback.classList.remove("error");
});

loginForm?.addEventListener("submit", (event) => {
  event.preventDefault();

  const email = emailInput.value.trim();
  const password = passwordInput.value;

  if (email === AUTH_EMAIL && password === AUTH_PASSWORD) {
    sessionStorage.setItem("costrackAuth", "authenticated");
    sessionStorage.setItem("costrackPlayDashboardIntro", "true");
    window.location.assign("index.html");
    return;
  }

  feedback.textContent = "Invalid credentials. Please use the assigned CosTrack account.";
  feedback.classList.add("error");
  feedback.classList.remove("info");
});


passwordInput?.addEventListener("keyup", (event) => {
  const capsLockOn = event.getModifierState && event.getModifierState("CapsLock");
  capsLockHint.textContent = capsLockOn ? "Caps Lock is on." : "";
});
