const AUTH_EMAIL = "costrack2026@gmail.com";
const AUTH_PASSWORD = "costrack2026";

const loginForm = document.getElementById("loginForm");
const emailInput = document.getElementById("email");
const passwordInput = document.getElementById("password");
const feedback = document.getElementById("loginFeedback");
const togglePasswordButton = document.getElementById("togglePassword");
const forgotPasswordButton = document.getElementById("forgotPasswordButton");
const rememberMeCheckbox = document.getElementById("rememberMe");
const capsLockHint = document.getElementById("capsLockHint");
const emailError = document.getElementById("emailError");
const emailWrap = document.getElementById("emailWrap");
const passwordWrap = document.getElementById("passwordWrap");
const submitButton = document.getElementById("submitButton");
let loginAttemptTimer = null;

const getAuthStorage = () => (localStorage.getItem("costrackRememberMe") === "true" ? localStorage : sessionStorage);

if (sessionStorage.getItem("costrackAuth") === "authenticated" || localStorage.getItem("costrackAuth") === "authenticated") {
  sessionStorage.setItem("costrackAuth", "authenticated");
  window.location.replace("index.html");
}

if (localStorage.getItem("costrackRememberMe") === "true" && rememberMeCheckbox) {
  rememberMeCheckbox.checked = true;
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
    emailError.textContent = "Enter a valid email address.";
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
  togglePasswordButton.innerHTML = isPassword
    ? '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17.94 17.94A10.95 10.95 0 0 1 12 19C5 19 1 12 1 12a21.77 21.77 0 0 1 5.06-5.94"></path><path d="M9.9 4.24A10.94 10.94 0 0 1 12 5c7 0 11 7 11 7a21.76 21.76 0 0 1-3.22 4.31"></path><path d="M14.12 14.12A3 3 0 0 1 9.88 9.88"></path><path d="M1 1l22 22"></path></svg>'
    : '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-7 11-7 11 7 11 7-4 7-11 7S1 12 1 12z"></path><circle cx="12" cy="12" r="3"></circle></svg>';
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
      const authStorage = getAuthStorage();
      authStorage.setItem("costrackAuth", "authenticated");
      localStorage.setItem("costrackRememberMe", String(Boolean(rememberMeCheckbox?.checked)));
      if (!(rememberMeCheckbox?.checked)) {
        localStorage.removeItem("costrackAuth");
      }
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

passwordInput?.addEventListener("keydown", (event) => {
  const capsLockOn = event.getModifierState && event.getModifierState("CapsLock");
  capsLockHint.textContent = capsLockOn ? "Caps Lock is on." : "";
});

passwordInput?.addEventListener("blur", () => {
  capsLockHint.textContent = "";
});
