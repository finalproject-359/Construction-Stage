const AUTH_EMAIL = "costrack2026@gmail.com";
const AUTH_PASSWORD = "costrack2026";

const loginForm = document.getElementById("loginForm");
const emailInput = document.getElementById("email");
const passwordInput = document.getElementById("password");
const feedback = document.getElementById("loginFeedback");
const togglePasswordButton = document.getElementById("togglePassword");
const forgotPasswordButton = document.getElementById("forgotPasswordButton");
const capsLockHint = document.getElementById("capsLockHint");
const emailError = document.getElementById("emailError");
const emailWrap = document.getElementById("emailWrap");
const passwordWrap = document.getElementById("passwordWrap");
const submitButton = document.getElementById("submitButton");
let loginAttemptTimer = null;

if (sessionStorage.getItem("costrackAuth") === "authenticated") {
  window.location.replace("index.html");
}

const isValidEmail = (value) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);

const setFeedback = (message, type) => {
  feedback.textContent = message;
  feedback.classList.remove("error", "info", "success");
  if (type) feedback.classList.add(type);
};

const validateFields = () => {
  const email = emailInput.value.trim();
  const password = passwordInput.value;
  let valid = true;

  emailWrap.classList.remove("invalid");
  passwordWrap.classList.remove("invalid");
  emailError.textContent = "";

  if (!email) {
    emailWrap.classList.add("invalid");
    emailError.textContent = "Email is required.";
    valid = false;
  } else if (!isValidEmail(email)) {
    emailWrap.classList.add("invalid");
    emailError.textContent = "Enter a valid work email address.";
    valid = false;
  }

  if (!password) {
    passwordWrap.classList.add("invalid");
    setFeedback("Please enter your password.", "error");
    valid = false;
  } else if (password.length < 8) {
    passwordWrap.classList.add("invalid");
    setFeedback("Password must be at least 8 characters.", "error");
    valid = false;
  }

  return valid;
};

togglePasswordButton?.addEventListener("click", () => {
  const isPassword = passwordInput.type === "password";
  passwordInput.type = isPassword ? "text" : "password";
  togglePasswordButton.textContent = isPassword ? "🙈" : "👁";
  togglePasswordButton.setAttribute("aria-label", isPassword ? "Hide password" : "Show password");
});

forgotPasswordButton?.addEventListener("click", () => {
  setFeedback("Please contact your system administrator to reset your credentials.", "info");
});

emailInput?.addEventListener("input", () => {
  emailWrap.classList.remove("invalid");
  emailError.textContent = "";
});

passwordInput?.addEventListener("input", () => {
  passwordWrap.classList.remove("invalid");
  if (feedback.classList.contains("error")) {
    setFeedback("", "");
  }
});

loginForm?.addEventListener("submit", (event) => {
  event.preventDefault();

  if (!validateFields()) {
    return;
  }

  const email = emailInput.value.trim();
  const password = passwordInput.value;

  submitButton.disabled = true;
  submitButton.textContent = "Signing in...";

  if (loginAttemptTimer) {
    window.clearTimeout(loginAttemptTimer);
  }

  loginAttemptTimer = window.setTimeout(() => {
    if (email === AUTH_EMAIL && password === AUTH_PASSWORD) {
      sessionStorage.setItem("costrackAuth", "authenticated");
      sessionStorage.setItem("costrackPlayDashboardIntro", "true");
      setFeedback("Sign-in successful. Redirecting to dashboard...", "success");
      window.setTimeout(() => window.location.assign("index.html"), 350);
      return;
    }

    submitButton.disabled = false;
    submitButton.textContent = "Sign In";
    passwordWrap.classList.add("invalid");
    setFeedback("Invalid credentials. Please use the assigned CosTrack account.", "error");
  }, 500);
});

passwordInput?.addEventListener("keyup", (event) => {
  const capsLockOn = event.getModifierState && event.getModifierState("CapsLock");
  capsLockHint.textContent = capsLockOn ? "Caps Lock is on." : "";
});
