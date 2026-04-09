// app.js
import { auth, db } from "./firebase-config.js";
import { jsPDF } from "https://cdn.jsdelivr.net/npm/jspdf@2.5.1/+esm";
import {
  signInWithEmailAndPassword,
  fetchSignInMethodsForEmail,
  EmailAuthProvider,
  GoogleAuthProvider,
  getRedirectResult,
  sendPasswordResetEmail,
  reauthenticateWithCredential,
  signInWithPopup,
  signInWithRedirect,
  signOut,
  updatePassword,
  onAuthStateChanged,
} from "https://www.gstatic.com/firebasejs/10.8.0/firebase-auth.js";
import {
  collection,
  addDoc,
  query,
  where,
  getDoc,
  getDocs,
  updateDoc,
  doc,
  deleteDoc,
  setDoc,
  onSnapshot,
  serverTimestamp,
} from "https://www.gstatic.com/firebasejs/10.8.0/firebase-firestore.js";

console.log("app.js loaded successfully");

let currentUser = null;
let borrowersUnsubscribe = null;
let appStatusTimeout = null;
let editingBorrowerId = null;
let borrowersCache = [];
let borrowerBindingsInitialized = false;
let currentDashboardSection = "borrowers";
let selectedAdminId = "admin_1";
const DEFAULT_ADMIN_PASSWORD = "1234567";
const SHARED_WORKSPACE_ID = "mambo-main";
const ADMIN_PROFILES = [
  {
    id: "admin_1",
    name: "Tulani",
    email: "tulani1mambo@gmail.com",
  },
  {
    id: "admin_2",
    name: "Kudzai",
    email: "kudzaimambo5@gmail.com",
  },
];
const kwachaFormatter = new Intl.NumberFormat("en-ZM", {
  style: "currency",
  currency: "ZMW",
  minimumFractionDigits: 2,
});

function formatCurrency(amount) {
  return kwachaFormatter.format(Number(amount) || 0);
}

function getSelectedAdminProfile() {
  return ADMIN_PROFILES.find((admin) => admin.id === selectedAdminId) || null;
}

function isAuthorizedAdminEmail(email) {
  if (!email) return false;
  const normalized = email.toLowerCase();
  return ADMIN_PROFILES.some(
    (admin) => admin.email.toLowerCase() === normalized,
  );
}

async function ensureAdminUserProfile(user, preferredName = "") {
  if (!user?.uid || !isAuthorizedAdminEmail(user.email)) return;
  await setDoc(
    doc(db, "users", user.uid),
    {
      uid: user.uid,
      email: user.email || "",
      name: preferredName || user.displayName || user.email || "Admin",
      role: "admin",
      updatedAt: serverTimestamp(),
    },
    { merge: true },
  );
}

async function enforceDefaultPasswordChange(enteredPassword) {
  if (enteredPassword !== DEFAULT_ADMIN_PASSWORD) return true;

  setAppStatus(
    "Default password detected. You must set a new password now.",
    "info",
    true,
  );

  let nextPassword = window.prompt(
    "Set a new password (minimum 6 characters):",
    "",
  );

  while (nextPassword !== null) {
    const trimmed = nextPassword.trim();
    if (trimmed.length < 6) {
      nextPassword = window.prompt(
        "Password too short. Enter at least 6 characters:",
        "",
      );
      continue;
    }
    if (trimmed === DEFAULT_ADMIN_PASSWORD) {
      nextPassword = window.prompt(
        "New password cannot be the default. Enter a different one:",
        "",
      );
      continue;
    }

    try {
      await updatePassword(auth.currentUser, trimmed);
      setAppStatus("Password updated successfully.", "ok");
      return true;
    } catch (error) {
      setAppStatus(`Password update failed: ${error.message}`, "error", true);
      return false;
    }
  }

  await signOut(auth);
  setAppStatus(
    "Password change canceled. Sign in again and set a new password.",
    "error",
    true,
  );
  return false;
}

function formatDateLabel(dateIso) {
  if (!dateIso) return "N/A";
  const date = new Date(`${dateIso}T00:00:00`);
  return date.toLocaleDateString("en-ZM", {
    year: "numeric",
    month: "short",
    day: "numeric",
  });
}

function formatTimestampLabel(ts) {
  if (!ts) return "-";
  try {
    if (typeof ts?.toDate === "function") {
      return ts.toDate().toLocaleString("en-ZM");
    }
    if (typeof ts === "string") {
      return new Date(ts).toLocaleString("en-ZM");
    }
    return "-";
  } catch (_error) {
    return "-";
  }
}

function openMailDraft(to, subject, body) {
  if (!to) {
    setAppStatus("Borrower email is missing.", "error", true);
    return;
  }
  const mailtoLink =
    `mailto:${encodeURIComponent(to)}` +
    `?subject=${encodeURIComponent(subject)}` +
    `&body=${encodeURIComponent(body)}`;
  window.location.href = mailtoLink;
}

function getDateIsoAfterDays(days) {
  const date = new Date();
  date.setDate(date.getDate() + days);
  const localDate = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
  return localDate.toISOString().slice(0, 10);
}

function getTodayIso() {
  const now = new Date();
  const localNow = new Date(now.getTime() - now.getTimezoneOffset() * 60000);
  return localNow.toISOString().slice(0, 10);
}

function getBorrowerIdFromUrl() {
  const params = new URLSearchParams(window.location.search);
  return params.get("id");
}

function saveBorrowerCacheForOffline(borrower) {
  try {
    if (!borrower?.id) return;
    const key = `borrowerCache:${borrower.id}`;
    sessionStorage.setItem(key, JSON.stringify(borrower));
  } catch (_error) {
    // Ignore session storage errors
  }
}

function getBorrowerCacheForOffline(borrowerId) {
  try {
    const key = `borrowerCache:${borrowerId}`;
    const raw = sessionStorage.getItem(key);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (_error) {
    return null;
  }
}

function normalizePhoneForWhatsApp(phone) {
  const cleaned = String(phone || "").replace(/[^\d+]/g, "").trim();
  if (!cleaned) return "";

  if (cleaned.startsWith("+")) {
    return cleaned.slice(1);
  }
  if (cleaned.startsWith("260")) {
    return cleaned;
  }
  if (cleaned.startsWith("0")) {
    return `260${cleaned.slice(1)}`;
  }
  if (cleaned.length === 9) {
    return `260${cleaned}`;
  }
  return cleaned;
}

function openWhatsAppDraft(phone, message) {
  const normalizedPhone = normalizePhoneForWhatsApp(phone);
  if (!normalizedPhone) {
    setAppStatus("Borrower phone number is missing or invalid.", "error", true);
    return;
  }
  const url = `https://wa.me/${encodeURIComponent(normalizedPhone)}?text=${encodeURIComponent(message)}`;
  window.open(url, "_blank");
}

async function copyTextToClipboard(text) {
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch (_error) {
    // fallback below
  }

  try {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand("copy");
    document.body.removeChild(textarea);
    return true;
  } catch (_error) {
    return false;
  }
}

async function recordCommunication(borrowerId, channel, category) {
  try {
    await updateDoc(doc(db, "borrowers", borrowerId), {
      lastCommunicationChannel: channel,
      lastCommunicationType: category,
      lastCommunicationAt: serverTimestamp(),
    });
  } catch (error) {
    console.error("Failed to record communication log:", error);
  }
}

function getFriendlyAuthError(error) {
  const code = error?.code || "";
  const fallback = error?.message || "Authentication failed. Please try again.";

  switch (code) {
    case "auth/email-already-in-use":
      return "This email already has an account. Please log in instead.";
    case "auth/invalid-email":
      return "Please enter a valid email address.";
    case "auth/weak-password":
      return "Password is too weak. Use at least 6 characters.";
    case "auth/invalid-credential":
    case "auth/wrong-password":
    case "auth/user-not-found":
      return "Invalid email or password.";
    case "auth/too-many-requests":
      return "Too many attempts. Wait a moment, then try again.";
    case "auth/network-request-failed":
      return "Network error. Check your internet and try again.";
    case "auth/missing-email":
      return "Enter your email first, then try again.";
    case "auth/popup-blocked":
      return "Google popup was blocked by your browser.";
    case "auth/popup-closed-by-user":
      return "Google sign-in was closed before completion.";
    case "auth/unauthorized-domain":
      return "This domain is not authorized in Firebase Auth settings.";
    case "auth/operation-not-allowed":
      return "Google sign-in is not enabled in Firebase Auth.";
    case "auth/account-exists-with-different-credential":
      return "This email already exists with another sign-in method.";
    default:
      return fallback;
  }
}

async function handleGoogleRedirectResult() {
  try {
    const result = await getRedirectResult(auth);
    if (!result || !result.user) return;

    const user = result.user;
    const displayName = user.displayName || user.email || "Google User";
    localStorage.setItem("userName", displayName);
    await setDoc(
      doc(db, "users", user.uid),
      {
        uid: user.uid,
        name: displayName,
        email: user.email,
        updatedAt: serverTimestamp(),
      },
      { merge: true },
    );

    setAppStatus("Google sign-in successful. Opening dashboard...", "ok");
    if (!window.location.pathname.includes("dashboard.html")) {
      window.location.href = "dashboard.html";
    }
  } catch (error) {
    const message = getFriendlyAuthError(error);
    const loginError = document.getElementById("loginError");
    if (loginError) loginError.textContent = message;
    setAppStatus(`Google sign-in error: ${message}`, "error", true);
  }
}

function setAppStatus(message, type = "info", persistent = false) {
  const statusElement = document.getElementById("appStatus");
  if (!statusElement) return;

  statusElement.textContent = message;
  statusElement.classList.remove("is-info", "is-ok", "is-error", "is-hidden");

  if (type === "ok") {
    statusElement.classList.add("is-ok");
  } else if (type === "error") {
    statusElement.classList.add("is-error");
  } else {
    statusElement.classList.add("is-info");
  }

  if (appStatusTimeout) {
    clearTimeout(appStatusTimeout);
    appStatusTimeout = null;
  }

  if (!persistent) {
    appStatusTimeout = setTimeout(() => {
      statusElement.classList.add("is-hidden");
    }, 3500);
  }
}

function setLoginLoading(isLoading) {
  const loginBtn = document.getElementById("loginSubmitBtn");
  const loginPassword = document.getElementById("loginPassword");
  if (!loginBtn) return;

  loginBtn.disabled = isLoading;
  loginBtn.textContent = isLoading ? "Signing in..." : "Sign In";
  if (loginPassword) loginPassword.disabled = isLoading;
}

function setChangePasswordLoading(isLoading) {
  const submitBtn = document.getElementById("changePasswordBtn");
  const clearBtn = document.querySelector(".change-password-actions .secondary-btn");
  const fieldIds = ["currentPassword", "newPassword", "confirmNewPassword"];

  if (submitBtn) {
    submitBtn.disabled = isLoading;
    submitBtn.textContent = isLoading ? "Updating..." : "Update Password";
  }

  if (clearBtn) clearBtn.disabled = isLoading;

  fieldIds.forEach((fieldId) => {
    const input = document.getElementById(fieldId);
    if (input) input.disabled = isLoading;
  });
}

function syncDashboardAdminSummary(user) {
  const adminName = document.getElementById("dashboardAdminName");
  const adminEmail = document.getElementById("dashboardAdminEmail");
  if (!adminName && !adminEmail) return;

  const matchedAdmin = ADMIN_PROFILES.find(
    (admin) => admin.email.toLowerCase() === (user?.email || "").toLowerCase(),
  );
  const displayName =
    matchedAdmin?.name || user?.displayName || localStorage.getItem("userName") || "Admin";

  if (adminName) adminName.textContent = displayName;
  if (adminEmail) adminEmail.textContent = user?.email || "Signed in email";
}

window.clearChangePasswordForm = function () {
  const fieldIds = ["currentPassword", "newPassword", "confirmNewPassword"];
  fieldIds.forEach((fieldId) => {
    const input = document.getElementById(fieldId);
    if (input) {
      input.value = "";
      if (input.type !== "password") input.type = "password";
    }
  });

  document
    .querySelectorAll(".change-password-card .toggle-password")
    .forEach((button) => {
      button.classList.remove("is-visible");
      button.setAttribute("aria-label", "Show password");
    });

  const errorDiv = document.getElementById("changePasswordError");
  if (errorDiv) errorDiv.textContent = "";
};

window.selectAdminProfile = function (adminId) {
  selectedAdminId = adminId;
  const admin = getSelectedAdminProfile();
  const selectedAdminLabel = document.getElementById("selectedAdminEmail");
  if (selectedAdminLabel) {
    selectedAdminLabel.textContent = admin
      ? `Selected: ${admin.name} (${admin.email})`
      : "No admin selected";
  }

  const profileButtons = document.querySelectorAll(".admin-profile-btn");
  profileButtons.forEach((btn) => {
    const isActive = btn.getAttribute("data-admin-id") === adminId;
    btn.classList.toggle("active", isActive);
  });
};

function initAdminProfileLoginUI() {
  const profilesContainer = document.getElementById("adminProfiles");
  if (!profilesContainer) return;

  profilesContainer.innerHTML = ADMIN_PROFILES.map(
    (admin) =>
      `<button type="button" class="admin-profile-btn" data-admin-id="${admin.id}" onclick="selectAdminProfile('${admin.id}')">${admin.name}</button>`,
  ).join("");

  const initialProfile = getSelectedAdminProfile() || ADMIN_PROFILES[0];
  if (initialProfile) {
    window.selectAdminProfile(initialProfile.id);
  }
}

window.switchDashboardSection = function (section) {
  currentDashboardSection = section;
  const sectionIds = {
    dashboard: "sectionDashboard",
    borrowers: "sectionBorrowers",
    reports: "sectionReports",
  };
  const tabIds = {
    dashboard: "tabDashboard",
    borrowers: "tabBorrowers",
    reports: "tabReports",
  };

  Object.entries(sectionIds).forEach(([key, id]) => {
    const element = document.getElementById(id);
    if (!element) return;
    element.classList.toggle("hidden-section", key !== section);
    if (key === section) {
      element.classList.remove("is-entering");
      void element.offsetWidth;
      element.classList.add("is-entering");
      setTimeout(() => element.classList.remove("is-entering"), 380);
    }
  });

  Object.entries(tabIds).forEach(([key, id]) => {
    const tab = document.getElementById(id);
    if (!tab) return;
    tab.classList.toggle("active", key === section);
  });
};

function runStartupHealthCheck() {
  try {
    if (auth && db) {
      setAppStatus("Connected to Firebase", "ok");
    } else {
      setAppStatus("Firebase is not initialized", "error", true);
    }
  } catch (error) {
    setAppStatus(`Startup error: ${error.message}`, "error", true);
  }

  if (!navigator.onLine) {
    setAppStatus("Offline: check your internet connection", "error", true);
  }
}

window.addEventListener("online", () => {
  setAppStatus("Back online", "ok");
});

window.addEventListener("offline", () => {
  setAppStatus("Offline: check your internet connection", "error", true);
});

runStartupHealthCheck();
handleGoogleRedirectResult();
initAdminProfileLoginUI();

// Authentication Functions
window.login = async function () {
  const selectedAdmin = getSelectedAdminProfile();
  const email = selectedAdmin?.email || "";
  const password = document.getElementById("loginPassword").value;
  const errorDiv = document.getElementById("loginError");

  if (!email) {
    const message = "Please select an admin profile first.";
    errorDiv.textContent = message;
    setAppStatus(message, "error", true);
    return;
  }

  try {
    setLoginLoading(true);
    const credential = await signInWithEmailAndPassword(auth, email, password);
    const changedPassword = await enforceDefaultPasswordChange(password);
    if (!changedPassword) return;
    await ensureAdminUserProfile(credential.user, selectedAdmin?.name || "");
    localStorage.setItem("userName", selectedAdmin?.name || email);
    setAppStatus("Login successful. Opening dashboard...", "ok");
    window.location.href = "dashboard.html";
  } catch (error) {
    const message = getFriendlyAuthError(error);
    errorDiv.textContent = message;
    setAppStatus(`Login error: ${message}`, "error", true);
  } finally {
    setLoginLoading(false);
  }
};

window.logout = async function () {
  try {
    await signOut(auth);
    window.location.href = "index.html";
  } catch (error) {
    console.error("Logout error:", error);
    setAppStatus(`Logout error: ${error.message}`, "error", true);
  }
};

window.loginWithGoogle = async function () {
  const loginError = document.getElementById("loginError");
  const provider = new GoogleAuthProvider();
  provider.setCustomParameters({ prompt: "select_account" });

  try {
    const userCredential = await signInWithPopup(auth, provider);
    const user = userCredential.user;
    if (!isAuthorizedAdminEmail(user.email)) {
      await signOut(auth);
      const message = "This Google account is not allowed as an admin.";
      if (loginError) loginError.textContent = message;
      setAppStatus(message, "error", true);
      return;
    }

    const displayName = user.displayName || user.email || "Google User";

    localStorage.setItem("userName", displayName);
    await ensureAdminUserProfile(user, displayName);

    setAppStatus("Google sign-in successful. Opening dashboard...", "ok");
    window.location.href = "dashboard.html";
  } catch (error) {
    if (error?.code === "auth/popup-blocked") {
      await signInWithRedirect(auth, provider);
      return;
    }

    if (error?.code === "auth/account-exists-with-different-credential") {
      const email = error?.customData?.email;
      if (email) {
        const methods = await fetchSignInMethodsForEmail(auth, email);
        const methodText = methods.length > 0 ? methods.join(", ") : "email/password";
        const message = `This email already exists. Use: ${methodText}.`;
        if (loginError) loginError.textContent = message;
        setAppStatus(message, "error", true);
        return;
      }
    }

    const message = getFriendlyAuthError(error);
    if (loginError) {
      loginError.textContent = message;
    }
    setAppStatus(`Google sign-in error: ${message}`, "error", true);
  }
};

window.forgotPassword = async function () {
  const loginError = document.getElementById("loginError");
  const selectedAdmin = getSelectedAdminProfile();
  const email = (selectedAdmin?.email || "").trim();

  if (!email) {
    const message = "Please select an admin profile first.";
    if (loginError) loginError.textContent = message;
    setAppStatus(message, "error", true);
    return;
  }

  try {
    await sendPasswordResetEmail(auth, email);
    if (loginError) loginError.textContent = "";
    setAppStatus("Password reset email sent. Check your inbox.", "ok", true);
  } catch (error) {
    const message = getFriendlyAuthError(error);
    if (loginError) loginError.textContent = message;
    setAppStatus(`Reset error: ${message}`, "error", true);
  }
};

window.togglePassword = function (inputId, buttonElement) {
  const input = document.getElementById(inputId);
  if (!input || !buttonElement) return;

  const isPassword = input.type === "password";
  input.type = isPassword ? "text" : "password";
  buttonElement.classList.toggle("is-visible", isPassword);
  buttonElement.setAttribute(
    "aria-label",
    isPassword ? "Hide password" : "Show password",
  );

  const toggleText = buttonElement.querySelector(".toggle-password-text");
  if (toggleText) {
    toggleText.textContent = isPassword ? "Hide" : "Show";
  }
};

window.changeAdminPassword = async function () {
  const errorDiv = document.getElementById("changePasswordError");
  const currentPassword = document.getElementById("currentPassword")?.value.trim() || "";
  const newPassword = document.getElementById("newPassword")?.value.trim() || "";
  const confirmPassword =
    document.getElementById("confirmNewPassword")?.value.trim() || "";

  if (errorDiv) errorDiv.textContent = "";

  if (!auth.currentUser?.email) {
    const message = "No signed-in admin account was found.";
    if (errorDiv) errorDiv.textContent = message;
    setAppStatus(message, "error", true);
    return;
  }

  if (!currentPassword || !newPassword || !confirmPassword) {
    const message = "Fill in all password fields first.";
    if (errorDiv) errorDiv.textContent = message;
    setAppStatus(message, "error", true);
    return;
  }

  if (newPassword.length < 6) {
    const message = "New password must be at least 6 characters.";
    if (errorDiv) errorDiv.textContent = message;
    setAppStatus(message, "error", true);
    return;
  }

  if (newPassword === DEFAULT_ADMIN_PASSWORD) {
    const message = "New password cannot be the default admin password.";
    if (errorDiv) errorDiv.textContent = message;
    setAppStatus(message, "error", true);
    return;
  }

  if (newPassword !== confirmPassword) {
    const message = "New password and confirmation do not match.";
    if (errorDiv) errorDiv.textContent = message;
    setAppStatus(message, "error", true);
    return;
  }

  if (currentPassword === newPassword) {
    const message = "Choose a new password that is different from the current one.";
    if (errorDiv) errorDiv.textContent = message;
    setAppStatus(message, "error", true);
    return;
  }

  try {
    setChangePasswordLoading(true);
    const credential = EmailAuthProvider.credential(
      auth.currentUser.email,
      currentPassword,
    );
    await reauthenticateWithCredential(auth.currentUser, credential);
    await updatePassword(auth.currentUser, newPassword);
    window.clearChangePasswordForm();
    setAppStatus("Password updated successfully.", "ok");
  } catch (error) {
    const message = getFriendlyAuthError(error);
    if (errorDiv) errorDiv.textContent = message;
    setAppStatus(`Password update failed: ${message}`, "error", true);
  } finally {
    setChangePasswordLoading(false);
  }
};

// Dashboard Functions
function calculateBorrowerFigures(amountBorrowed, interestPercentage) {
  const amount = Number(amountBorrowed) || 0;
  const interestRate = Number(interestPercentage) || 0;
  const interestAmount = amount * (interestRate / 100);
  const totalToPay = amount + interestAmount;
  return { interestAmount, totalToPay };
}

function focusBorrowerNameField() {
  const borrowerNameField = document.getElementById("borrowerName");
  if (borrowerNameField && typeof borrowerNameField.focus === "function") {
    borrowerNameField.focus();
  }
}

function setAutoBorrowerAmounts() {
  const amountInput = document.getElementById("amountBorrowed");
  const interestInput = document.getElementById("interestPercentage");
  const interestAmountInput = document.getElementById("interestAmount");
  const totalToPayInput = document.getElementById("totalToPay");

  if (!amountInput || !interestInput || !interestAmountInput || !totalToPayInput) {
    return;
  }

  const { interestAmount, totalToPay } = calculateBorrowerFigures(
    amountInput.value,
    interestInput.value,
  );

  interestAmountInput.value = formatCurrency(interestAmount);
  totalToPayInput.value = formatCurrency(totalToPay);
}

function getBorrowerFormValues() {
  const today = getTodayIso();
  const name = document.getElementById("borrowerName")?.value.trim() || "";
  const nrc = document.getElementById("borrowerNrc")?.value.trim() || "";
  const address = document.getElementById("borrowerAddress")?.value.trim() || "";
  const phone = document.getElementById("borrowerPhone")?.value.trim() || "";
  const email = document.getElementById("borrowerEmail")?.value.trim() || "";
  const amountBorrowed = parseFloat(
    document.getElementById("amountBorrowed")?.value || "0",
  );
  const interestPercentage = parseFloat(
    document.getElementById("interestPercentage")?.value || "0",
  );
  const dateBorrowed =
    document.getElementById("dateBorrowed")?.value || today;
  const dueDate = document.getElementById("dueDate")?.value || "";

  return {
    name,
    nrc,
    address,
    phone,
    email,
    amountBorrowed,
    interestPercentage,
    dateBorrowed,
    dueDate,
  };
}

function validateBorrowerForm(values) {
  const missing = [];
  if (!values.name) missing.push("Full Name");
  if (!values.nrc) missing.push("NRC Number");
  if (!values.dueDate) missing.push("Due Date");
  if (!values.dateBorrowed) missing.push("Date Borrowed");
  if (missing.length > 0) {
    return `Please complete: ${missing.join(", ")}.`;
  }
  if (Number.isNaN(values.amountBorrowed) || values.amountBorrowed <= 0) {
    return "Amount borrowed must be greater than 0.";
  }
  if (Number.isNaN(values.interestPercentage) || values.interestPercentage < 0) {
    return "Interest percentage must be 0 or higher.";
  }
  return "";
}

function resetBorrowerEditingState() {
  editingBorrowerId = null;
  const saveButton = document.getElementById("saveBorrowerBtn");
  if (saveButton) saveButton.textContent = "Add Borrower";
}

window.clearBorrowerForm = function () {
  const ids = [
    "borrowerName",
    "borrowerNrc",
    "borrowerAddress",
    "borrowerPhone",
    "borrowerEmail",
    "amountBorrowed",
    "interestPercentage",
    "dateBorrowed",
    "dueDate",
    "interestAmount",
    "totalToPay",
  ];
  ids.forEach((id) => {
    const element = document.getElementById(id);
    if (element) element.value = "";
  });
  setAutoBorrowerAmounts();
  resetBorrowerEditingState();
  focusBorrowerNameField();
};

window.saveBorrower = async function () {
  if (!currentUser?.uid) {
    setAppStatus(
      "Your session is not ready. Please refresh and sign in again.",
      "error",
      true,
    );
    return;
  }

  const values = getBorrowerFormValues();
  const validationMessage = validateBorrowerForm(values);
  if (validationMessage) {
    setAppStatus(validationMessage, "error", true);
    return;
  }

  const { interestAmount, totalToPay } = calculateBorrowerFigures(
    values.amountBorrowed,
    values.interestPercentage,
  );

  const payload = {
    workspaceId: SHARED_WORKSPACE_ID,
    userId: currentUser.uid,
    createdByUid: currentUser.uid,
    name: values.name,
    nrc: values.nrc,
    address: values.address,
    phone: values.phone,
    email: values.email,
    amountBorrowed: values.amountBorrowed,
    interestPercentage: values.interestPercentage,
    dateBorrowed: values.dateBorrowed,
    dueDate: values.dueDate,
    interestAmount,
    totalToPay,
    status: "pending",
    updatedAt: serverTimestamp(),
  };

  try {
    if (editingBorrowerId) {
      const existing = borrowersCache.find((b) => b.id === editingBorrowerId);
      await updateDoc(doc(db, "borrowers", editingBorrowerId), {
        ...payload,
        status: existing?.status || "pending",
      });
      setAppStatus("Borrower updated successfully.", "ok");
    } else {
      await addDoc(collection(db, "borrowers"), {
        ...payload,
        createdAt: serverTimestamp(),
      });
      setAppStatus("Borrower added successfully.", "ok");

      if (
        values.email &&
        window.confirm("Borrower added. Open a welcome email draft now?")
      ) {
        const body =
          `Hello ${values.name},\n\n` +
          `Your loan has been recorded in our system.\n` +
          `Amount Borrowed: ${formatCurrency(values.amountBorrowed)}\n` +
          `Interest: ${Number(values.interestPercentage).toFixed(2)}%\n` +
          `Total to Pay: ${formatCurrency(totalToPay)}\n` +
          `Due Date: ${formatDateLabel(values.dueDate)}\n\n` +
          "Thank you,\nMambo Finance";
        openMailDraft(values.email, "Welcome - Loan Created", body);
      }
    }

    window.clearBorrowerForm();
  } catch (error) {
    setAppStatus(`Error saving borrower: ${error.message}`, "error", true);
  }
};

window.sendWelcomeEmailDraft = function (borrowerId) {
  const borrower = borrowersCache.find((item) => item.id === borrowerId);
  if (!borrower) return;

  const body =
    `Hello ${borrower.name || "Borrower"},\n\n` +
    `Your loan is active in our system.\n` +
    `Amount Borrowed: ${formatCurrency(borrower.amountBorrowed)}\n` +
    `Interest: ${Number(borrower.interestPercentage || 0).toFixed(2)}%\n` +
    `Total to Pay: ${formatCurrency(borrower.totalToPay)}\n` +
    `Due Date: ${formatDateLabel(borrower.dueDate)}\n\n` +
    "Thank you,\nMambo Finance";

  recordCommunication(borrowerId, "email", "welcome");
  openMailDraft(borrower.email, "Welcome - Loan Created", body);
};

window.sendReminderEmailDraft = function (borrowerId) {
  const borrower = borrowersCache.find((item) => item.id === borrowerId);
  if (!borrower) return;

  const body =
    `Hello ${borrower.name || "Borrower"},\n\n` +
    `Friendly reminder: your loan is due on ${formatDateLabel(borrower.dueDate)}.\n` +
    `Total amount due: ${formatCurrency(borrower.totalToPay)}\n\n` +
    "Please plan your payment in time.\n\nThank you,\nMambo Finance";

  recordCommunication(borrowerId, "email", "reminder");
  openMailDraft(borrower.email, "Loan Due Reminder", body);
};

window.sendWelcomeWhatsAppDraft = function (borrowerId) {
  const borrower = borrowersCache.find((item) => item.id === borrowerId);
  if (!borrower) return;

  const message =
    `Hello ${borrower.name || "Borrower"}, your loan has been recorded. ` +
    `Amount: ${formatCurrency(borrower.amountBorrowed)}. ` +
    `Total to pay: ${formatCurrency(borrower.totalToPay)}. ` +
    `Due date: ${formatDateLabel(borrower.dueDate)}. ` +
    "Thank you - Mambo Finance.";

  recordCommunication(borrowerId, "whatsapp", "welcome");
  openWhatsAppDraft(borrower.phone, message);
};

window.sendReminderWhatsAppDraft = function (borrowerId) {
  const borrower = borrowersCache.find((item) => item.id === borrowerId);
  if (!borrower) return;

  const message =
    `Hello ${borrower.name || "Borrower"}, reminder that your loan is due on ` +
    `${formatDateLabel(borrower.dueDate)}. Amount due: ${formatCurrency(borrower.totalToPay)}. ` +
    "Please plan payment in time. - Mambo Finance";

  recordCommunication(borrowerId, "whatsapp", "reminder");
  openWhatsAppDraft(borrower.phone, message);
};

window.editBorrower = function (borrowerId) {
  const borrower = borrowersCache.find((item) => item.id === borrowerId);
  if (!borrower) return;

  const mappings = [
    ["borrowerName", borrower.name],
    ["borrowerNrc", borrower.nrc],
    ["borrowerAddress", borrower.address],
    ["borrowerPhone", borrower.phone],
    ["borrowerEmail", borrower.email],
    ["amountBorrowed", borrower.amountBorrowed],
    ["interestPercentage", borrower.interestPercentage],
    ["dateBorrowed", borrower.dateBorrowed],
    ["dueDate", borrower.dueDate],
  ];
  mappings.forEach(([id, value]) => {
    const element = document.getElementById(id);
    if (element) element.value = value ?? "";
  });

  editingBorrowerId = borrower.id;
  const saveButton = document.getElementById("saveBorrowerBtn");
  if (saveButton) saveButton.textContent = "Update Borrower";
  setAutoBorrowerAmounts();
};

window.startBorrowerEdit = function (borrowerId) {
  window.switchDashboardSection("borrowers");
  window.editBorrower(borrowerId);
  const formSection = document.querySelector(".add-loan-section");
  if (formSection && typeof formSection.scrollIntoView === "function") {
    formSection.scrollIntoView({ behavior: "smooth", block: "start" });
  }
};

window.openBorrowerActions = function (borrowerId) {
  const borrower = borrowersCache.find((item) => item.id === borrowerId);
  if (borrower) saveBorrowerCacheForOffline(borrower);
  window.location.href = `borrower-actions.html?id=${encodeURIComponent(borrowerId)}`;
};

window.markBorrowerPaid = async function (borrowerId) {
  try {
    await updateDoc(doc(db, "borrowers", borrowerId), {
      status: "paid",
      updatedAt: serverTimestamp(),
    });
    setAppStatus("Borrower marked as paid.", "ok");
  } catch (error) {
    setAppStatus(`Error updating status: ${error.message}`, "error", true);
  }
};

window.undoBorrowerPaid = async function (borrowerId) {
  try {
    await updateDoc(doc(db, "borrowers", borrowerId), {
      status: "pending",
      updatedAt: serverTimestamp(),
    });
    setAppStatus("Borrower moved back to pending.", "ok");
  } catch (error) {
    setAppStatus(`Error updating status: ${error.message}`, "error", true);
  }
};

window.deleteBorrowerRecord = async function (borrowerId) {
  const confirmed = window.confirm("Delete this borrower record?");
  if (!confirmed) return;

  try {
    await deleteDoc(doc(db, "borrowers", borrowerId));
    if (editingBorrowerId === borrowerId) window.clearBorrowerForm();
    setAppStatus("Borrower deleted successfully.", "ok");
    if (window.location.pathname.includes("borrower-actions.html")) {
      window.location.href = "dashboard.html";
    }
  } catch (error) {
    setAppStatus(`Error deleting borrower: ${error.message}`, "error", true);
  }
};

function formatIsoDate(isoDate) {
  if (!isoDate) return "-";
  return isoDate;
}

function isBorrowerDueOrOverdue(dueDateIso) {
  if (!dueDateIso) return false;
  const todayIso = getTodayIso();
  return dueDateIso <= todayIso;
}

function getBorrowerDisplayStatus(borrower) {
  if (borrower?.status === "paid") {
    return { label: "Paid", className: "paid", isOverdue: false };
  }

  const isOverdue =
    Boolean(borrower?.dueDate) && borrower.dueDate < getTodayIso();

  if (isOverdue) {
    return { label: "Over Due", className: "overdue", isOverdue: true };
  }

  return { label: "Pending", className: "pending", isOverdue: false };
}

function getFilteredBorrowers() {
  const search = (document.getElementById("searchBorrowers")?.value || "")
    .trim()
    .toLowerCase();
  const dueDateFilter = document.getElementById("filterDueDate")?.value || "";
  const statusFilter = document.getElementById("filterStatus")?.value || "";

  return borrowersCache.filter((borrower) => {
    const matchesSearch =
      !search ||
      borrower.name?.toLowerCase().includes(search) ||
      borrower.nrc?.toLowerCase().includes(search);

    const matchesDueDate = !dueDateFilter || borrower.dueDate === dueDateFilter;
    const displayStatus = getBorrowerDisplayStatus(borrower);
    const matchesStatus =
      !statusFilter ||
      (statusFilter === "paid" && displayStatus.className === "paid") ||
      (statusFilter === "pending" && displayStatus.className === "pending") ||
      (statusFilter === "overdue" && displayStatus.className === "overdue");
    return matchesSearch && matchesDueDate && matchesStatus;
  });
}

function renderBorrowersTable() {
  const tableBody = document.getElementById("borrowersTableBody");
  if (!tableBody) return;

  const filtered = getFilteredBorrowers();
  if (filtered.length === 0) {
    tableBody.innerHTML =
      '<tr><td colspan="10" class="empty-table">No borrowers found.</td></tr>';
    return;
  }

  tableBody.innerHTML = filtered
    .map((borrower) => {
      const displayStatus = getBorrowerDisplayStatus(borrower);
      const isPaid = displayStatus.className === "paid";
      return `
        <tr>
          <td>${escapeHtml(borrower.name || "")}</td>
          <td>${escapeHtml(borrower.nrc || "")}</td>
          <td>${escapeHtml(borrower.email || "")}</td>
          <td>${formatCurrency(borrower.amountBorrowed)}</td>
          <td>${Number(borrower.interestPercentage || 0).toFixed(2)}%</td>
          <td>${formatCurrency(borrower.totalToPay)}</td>
          <td>${formatIsoDate(borrower.dueDate)}</td>
          <td><span class="status-pill ${displayStatus.className}">${displayStatus.label}</span></td>
          <td>${escapeHtml(
            `${borrower.lastCommunicationChannel || "-"} / ${borrower.lastCommunicationType || "-"} / ${formatTimestampLabel(borrower.lastCommunicationAt)}`,
          )}</td>
          <td>
            <div class="table-actions">
              <button type="button" class="table-btn edit" onclick="startBorrowerEdit('${borrower.id}')">Edit</button>
              <button type="button" class="table-btn manage" onclick="openBorrowerActions('${borrower.id}')">Manage</button>
              ${
                isPaid
                  ? `<button type="button" class="table-btn paid" onclick="undoBorrowerPaid('${borrower.id}')">Undo Paid</button>`
                  : `<button type="button" class="table-btn paid" onclick="markBorrowerPaid('${borrower.id}')">Mark Paid</button>`
              }
            </div>
          </td>
        </tr>
      `;
    })
    .join("");
}

function renderBorrowerActionCard(borrower) {
  const card = document.getElementById("borrowerActionCard");
  if (!card) return;

  const displayStatus = getBorrowerDisplayStatus(borrower);
  const isPaid = displayStatus.className === "paid";
  const isDue = isBorrowerDueOrOverdue(borrower.dueDate);

  card.innerHTML = `
    <h2>${escapeHtml(borrower.name || "Borrower")}</h2>
    <p class="nav-subtitle">NRC: ${escapeHtml(borrower.nrc || "-")} | Status: ${displayStatus.label}</p>
    <div class="action-grid">
      <div><strong>Email:</strong> ${escapeHtml(borrower.email || "-")}</div>
      <div><strong>Phone:</strong> ${escapeHtml(borrower.phone || "-")}</div>
      <div><strong>Amount:</strong> ${formatCurrency(borrower.amountBorrowed)}</div>
      <div><strong>Total to Pay:</strong> ${formatCurrency(borrower.totalToPay)}</div>
      <div><strong>Due Date:</strong> ${formatIsoDate(borrower.dueDate)}</div>
      <div><strong>Last Comms:</strong> ${escapeHtml(
        `${borrower.lastCommunicationChannel || "-"} / ${borrower.lastCommunicationType || "-"} / ${formatTimestampLabel(borrower.lastCommunicationAt)}`,
      )}</div>
    </div>
    <div class="action-buttons-wrap">
      <button type="button" class="table-btn edit" onclick="window.location.href='dashboard.html'">Back to Edit</button>
      ${
        isPaid
          ? `<button type="button" class="table-btn paid" onclick="undoBorrowerPaid('${borrower.id}')">Undo Paid</button>`
          : `<button type="button" class="table-btn paid" onclick="markBorrowerPaid('${borrower.id}')">Mark Paid</button>`
      }
      <button type="button" class="table-btn welcome" onclick="sendWelcomeEmailDraft('${borrower.id}')">Email Welcome</button>
      <button type="button" class="table-btn whatsapp" onclick="sendWelcomeWhatsAppDraft('${borrower.id}')">WA Welcome</button>
      ${
        isDue
          ? `<button type="button" class="table-btn remind" onclick="sendReminderEmailDraft('${borrower.id}')">Email Reminder</button>
      <button type="button" class="table-btn whatsapp" onclick="sendReminderWhatsAppDraft('${borrower.id}')">WA Reminder</button>`
          : ""
      }
      <button type="button" class="table-btn delete" onclick="deleteBorrowerRecord('${borrower.id}')">Delete</button>
    </div>
  `;
}

async function loadBorrowerActionsPage() {
  if (!currentUser?.uid) return;

  const borrowerId = getBorrowerIdFromUrl();
  if (!borrowerId) {
    setAppStatus("Borrower ID missing in URL.", "error", true);
    return;
  }

  try {
    const borrowerRef = doc(db, "borrowers", borrowerId);
    const borrowerSnap = await getDoc(borrowerRef);

    if (!borrowerSnap.exists()) {
      setAppStatus("Borrower not found.", "error", true);
      return;
    }

    const borrower = { id: borrowerSnap.id, ...borrowerSnap.data() };
    if (borrower.userId !== currentUser.uid) {
      setAppStatus("Access denied for this borrower record.", "error", true);
      return;
    }

    borrowersCache = [borrower];
    saveBorrowerCacheForOffline(borrower);
    renderBorrowerActionCard(borrower);
  } catch (error) {
    const fallbackBorrower = getBorrowerCacheForOffline(borrowerId);
    if (fallbackBorrower) {
      borrowersCache = [fallbackBorrower];
      renderBorrowerActionCard(fallbackBorrower);
      setAppStatus(
        "Offline mode: showing cached borrower data.",
        "info",
        true,
      );
      return;
    }

    setAppStatus(`Error loading borrower: ${error.message}`, "error", true);
  }
}

window.draftFiveDayReminders = function () {
  const targetDueDate = getDateIsoAfterDays(5);
  const dueSoon = borrowersCache.filter(
    (borrower) =>
      borrower.status !== "paid" &&
      borrower.dueDate === targetDueDate &&
      borrower.email,
  );

  if (dueSoon.length === 0) {
    setAppStatus("No pending borrowers due in 5 days.", "info", true);
    return;
  }

  const recipients = dueSoon.map((borrower) => borrower.email).join(", ");
  const previewList = dueSoon
    .slice(0, 12)
    .map(
      (borrower) =>
        `- ${borrower.name || "Borrower"} (${formatCurrency(borrower.totalToPay)})`,
    )
    .join("\n");
  const extraCount = dueSoon.length > 12 ? `\n...and ${dueSoon.length - 12} more` : "";

  const body =
    `Hello,\n\n` +
    `This is a reminder that your loan payment is due on ${formatDateLabel(targetDueDate)}.\n\n` +
    `Borrowers included:\n${previewList}${extraCount}\n\n` +
    "Thank you,\nMambo Finance";

  dueSoon.forEach((borrower) => {
    recordCommunication(borrower.id, "email", "reminder");
  });
  openMailDraft(recipients, `Loan Due Reminder - ${formatDateLabel(targetDueDate)}`, body);
};

window.draftFiveDayWhatsApp = async function () {
  const targetDueDate = getDateIsoAfterDays(5);
  const dueSoon = borrowersCache.filter(
    (borrower) =>
      borrower.status !== "paid" &&
      borrower.dueDate === targetDueDate &&
      normalizePhoneForWhatsApp(borrower.phone),
  );

  if (dueSoon.length === 0) {
    setAppStatus("No pending borrowers with valid phones due in 5 days.", "info", true);
    return;
  }

  const links = dueSoon.map((borrower) => {
    const message =
      `Hello ${borrower.name || "Borrower"}, reminder that your loan is due on ` +
      `${formatDateLabel(borrower.dueDate)}. Amount due: ${formatCurrency(borrower.totalToPay)}. ` +
      "Please plan payment in time. - Mambo Finance";
    const phone = normalizePhoneForWhatsApp(borrower.phone);
    const url = `https://wa.me/${encodeURIComponent(phone)}?text=${encodeURIComponent(message)}`;
    return `${borrower.name || "Borrower"}: ${url}`;
  });

  const copied = await copyTextToClipboard(links.join("\n"));
  dueSoon.forEach((borrower) => {
    recordCommunication(borrower.id, "whatsapp", "reminder");
  });
  const firstBorrower = dueSoon[0];
  const firstMessage =
    `Hello ${firstBorrower.name || "Borrower"}, reminder that your loan is due on ` +
    `${formatDateLabel(firstBorrower.dueDate)}. Amount due: ${formatCurrency(firstBorrower.totalToPay)}. ` +
    "Please plan payment in time. - Mambo Finance";
  openWhatsAppDraft(firstBorrower.phone, firstMessage);
  setAppStatus(
    copied
      ? `Opened WhatsApp for first borrower. ${dueSoon.length} reminder links copied to clipboard.`
      : `Opened WhatsApp for first borrower. ${dueSoon.length} reminders ready.`,
    "ok",
    true,
  );
};

function renderSummaryAndReports() {
  const todayIso = getTodayIso();
  const dueSoonIso = getDateIsoAfterDays(5);
  const totalBorrowers = borrowersCache.length;
  const pendingBorrowers = borrowersCache.filter(
    (borrower) => getBorrowerDisplayStatus(borrower).className === "pending",
  ).length;
  const overdueBorrowers = borrowersCache.filter(
    (borrower) =>
      borrower.status !== "paid" &&
      Boolean(borrower.dueDate) &&
      borrower.dueDate < todayIso,
  ).length;
  const paidBorrowers = borrowersCache.filter(
    (borrower) => borrower.status === "paid",
  ).length;
  const dueSoonBorrowers = borrowersCache.filter(
    (borrower) =>
      borrower.status !== "paid" &&
      Boolean(borrower.dueDate) &&
      borrower.dueDate === dueSoonIso,
  );
  const borrowersWithFullContact = borrowersCache.filter(
    (borrower) => borrower.email?.trim() && borrower.phone?.trim(),
  ).length;
  const totalBorrowed = borrowersCache.reduce(
    (sum, borrower) => sum + Number(borrower.amountBorrowed || 0),
    0,
  );
  const totalExpected = borrowersCache.reduce(
    (sum, borrower) => sum + Number(borrower.totalToPay || 0),
    0,
  );
  const moneyRecovered = borrowersCache
    .filter((borrower) => borrower.status === "paid")
    .reduce((sum, borrower) => sum + Number(borrower.totalToPay || 0), 0);
  const totalInterest = borrowersCache.reduce(
    (sum, borrower) => sum + Number(borrower.interestAmount || 0),
    0,
  );
  const dueSoonValue = dueSoonBorrowers.reduce(
    (sum, borrower) => sum + Number(borrower.totalToPay || 0),
    0,
  );
  const collectionRate =
    totalExpected > 0 ? Math.round((moneyRecovered / totalExpected) * 100) : 0;
  const contactCoverage =
    totalBorrowers > 0
      ? Math.round((borrowersWithFullContact / totalBorrowers) * 100)
      : 0;

  const setText = (id, value) => {
    const element = document.getElementById(id);
    if (element) element.textContent = value;
  };

  setText("statTotalBorrowers", totalBorrowers);
  setText("statPendingBorrowers", pendingBorrowers);
  setText("statTotalBorrowed", formatCurrency(totalBorrowed));
  setText("statTotalExpected", formatCurrency(totalExpected));
  setText("statMoneyRecovered", formatCurrency(moneyRecovered));
  setText("statOverdueBorrowers", overdueBorrowers);

  setText("reportPaidBorrowers", paidBorrowers);
  setText("reportPendingBorrowers", pendingBorrowers);
  setText("reportOverdueBorrowers", overdueBorrowers);
  setText("reportMoneyRecovered", formatCurrency(moneyRecovered));
  setText("reportTotalInterest", formatCurrency(totalInterest));
  setText("reportPortfolioValue", formatCurrency(totalExpected));

  setText("collectionRate", `${collectionRate}%`);
  setText(
    "collectionRateMeta",
    totalExpected > 0
      ? `${formatCurrency(moneyRecovered)} recovered out of ${formatCurrency(totalExpected)} expected.`
      : "Paid loans compared to your full portfolio.",
  );
  setText("dueSoonCount", dueSoonBorrowers.length);
  setText(
    "dueSoonValue",
    dueSoonBorrowers.length > 0
      ? `${formatCurrency(dueSoonValue)} due on ${formatDateLabel(dueSoonIso)}.`
      : "No pending borrowers are due in 5 days.",
  );
  setText("contactCoverage", `${contactCoverage}%`);
  setText(
    "contactCoverageMeta",
    totalBorrowers > 0
      ? `${borrowersWithFullContact} of ${totalBorrowers} borrowers have both phone and email on file.`
      : "Borrowers with both phone and email available.",
  );
}

window.exportReportPdf = function () {
  try {
    const todayIso = getTodayIso();
    const totalBorrowers = borrowersCache.length;
    const pendingBorrowers = borrowersCache.filter(
      (borrower) => getBorrowerDisplayStatus(borrower).className === "pending",
    ).length;
    const overdueBorrowers = borrowersCache.filter(
      (borrower) =>
        borrower.status !== "paid" &&
        Boolean(borrower.dueDate) &&
        borrower.dueDate < todayIso,
    ).length;
    const paidBorrowers = borrowersCache.filter(
      (borrower) => borrower.status === "paid",
    ).length;
    const totalBorrowed = borrowersCache.reduce(
      (sum, borrower) => sum + Number(borrower.amountBorrowed || 0),
      0,
    );
    const totalExpected = borrowersCache.reduce(
      (sum, borrower) => sum + Number(borrower.totalToPay || 0),
      0,
    );
    const moneyRecovered = borrowersCache
      .filter((borrower) => borrower.status === "paid")
      .reduce((sum, borrower) => sum + Number(borrower.totalToPay || 0), 0);
    const totalInterest = borrowersCache.reduce(
      (sum, borrower) => sum + Number(borrower.interestAmount || 0),
      0,
    );

    const documentPdf = new jsPDF();
    const generatedAt = new Date().toLocaleString();
    const preparedBy = currentUser?.email || "User";

    documentPdf.setFontSize(18);
    documentPdf.text("Money Lending Management Report", 14, 20);
    documentPdf.setFontSize(11);
    documentPdf.text(`Generated: ${generatedAt}`, 14, 30);
    documentPdf.text(`Prepared by: ${preparedBy}`, 14, 36);

    let y = 48;
    const lineGap = 9;
    const addLine = (label, value) => {
      documentPdf.text(`${label}: ${value}`, 14, y);
      y += lineGap;
    };

    documentPdf.setFontSize(13);
    documentPdf.text("Portfolio Summary", 14, y);
    y += 8;
    documentPdf.setFontSize(11);
    addLine("Total Borrowers", String(totalBorrowers));
    addLine("Paid Borrowers", String(paidBorrowers));
    addLine("Pending Borrowers", String(pendingBorrowers));
    addLine("Over Due Borrowers", String(overdueBorrowers));
    addLine("Total Borrowed", formatCurrency(totalBorrowed));
    addLine("Money Recovered", formatCurrency(moneyRecovered));
    addLine("Total Interest", formatCurrency(totalInterest));
    addLine("Expected Collection", formatCurrency(totalExpected));

    y += 4;
    documentPdf.setFontSize(13);
    documentPdf.text("Borrower Breakdown", 14, y);
    y += 8;
    documentPdf.setFontSize(10);

    const header = "Name | NRC | Due Date | Status | Total";
    documentPdf.text(header, 14, y);
    y += 6;

    borrowersCache.slice(0, 25).forEach((borrower) => {
      if (y > 280) {
        documentPdf.addPage();
        y = 20;
      }
      const statusLabel = getBorrowerDisplayStatus(borrower).label;
      const row =
        `${(borrower.name || "-").slice(0, 16)} | ` +
        `${(borrower.nrc || "-").slice(0, 12)} | ` +
        `${borrower.dueDate || "-"} | ` +
        `${statusLabel} | ` +
        `${formatCurrency(borrower.totalToPay)}`;
      documentPdf.text(row, 14, y);
      y += 6;
    });

    const fileDate = new Date().toISOString().slice(0, 10);
    documentPdf.save(`mambo-report-${fileDate}.pdf`);
    setAppStatus("PDF report exported successfully.", "ok");
  } catch (error) {
    setAppStatus(`PDF export failed: ${error.message}`, "error", true);
  }
};

window.exportReportCsv = function () {
  try {
    const headers = [
      "Name",
      "NRC",
      "Email",
      "Phone",
      "Address",
      "Amount Borrowed (ZMW)",
      "Interest Percentage",
      "Interest Amount (ZMW)",
      "Total To Pay (ZMW)",
      "Date Borrowed",
      "Due Date",
      "Status",
      "Is Over Due",
      "Recovered Amount (ZMW)",
      "Last Communication Channel",
      "Last Communication Type",
      "Last Communication At",
    ];

    const escapeCsv = (value) => {
      const stringValue = String(value ?? "");
      if (/[",\n]/.test(stringValue)) {
        return `"${stringValue.replace(/"/g, '""')}"`;
      }
      return stringValue;
    };

    const rows = borrowersCache.map((borrower) => [
      borrower.name || "",
      borrower.nrc || "",
      borrower.email || "",
      borrower.phone || "",
      borrower.address || "",
      Number(borrower.amountBorrowed || 0).toFixed(2),
      Number(borrower.interestPercentage || 0).toFixed(2),
      Number(borrower.interestAmount || 0).toFixed(2),
      Number(borrower.totalToPay || 0).toFixed(2),
      borrower.dateBorrowed || "",
      borrower.dueDate || "",
      getBorrowerDisplayStatus(borrower).label,
      getBorrowerDisplayStatus(borrower).isOverdue
        ? "Yes"
        : "No",
      borrower.status === "paid"
        ? Number(borrower.totalToPay || 0).toFixed(2)
        : "0.00",
      borrower.lastCommunicationChannel || "",
      borrower.lastCommunicationType || "",
      formatTimestampLabel(borrower.lastCommunicationAt),
    ]);

    const csv = [headers, ...rows]
      .map((row) => row.map(escapeCsv).join(","))
      .join("\n");

    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const fileDate = new Date().toISOString().slice(0, 10);
    link.href = url;
    link.download = `mambo-report-${fileDate}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    setAppStatus("CSV report exported successfully.", "ok");
  } catch (error) {
    setAppStatus(`CSV export failed: ${error.message}`, "error", true);
  }
};

function bindBorrowerDashboardEvents() {
  if (borrowerBindingsInitialized) return;
  borrowerBindingsInitialized = true;

  const amountInput = document.getElementById("amountBorrowed");
  const interestInput = document.getElementById("interestPercentage");
  const searchInput = document.getElementById("searchBorrowers");
  const dueDateInput = document.getElementById("filterDueDate");
  const statusInput = document.getElementById("filterStatus");

  if (amountInput) amountInput.addEventListener("input", setAutoBorrowerAmounts);
  if (interestInput) interestInput.addEventListener("input", setAutoBorrowerAmounts);
  if (searchInput) searchInput.addEventListener("input", renderBorrowersTable);
  if (dueDateInput) dueDateInput.addEventListener("input", renderBorrowersTable);
  if (statusInput) statusInput.addEventListener("change", renderBorrowersTable);

  setAutoBorrowerAmounts();
  focusBorrowerNameField();
}

function loadBorrowersDashboard() {
  if (!currentUser) return;
  bindBorrowerDashboardEvents();

  if (borrowersUnsubscribe) {
    borrowersUnsubscribe();
    borrowersUnsubscribe = null;
  }

  const borrowersRef = collection(db, "borrowers");
  const borrowersQuery = query(
    borrowersRef,
    where("workspaceId", "==", SHARED_WORKSPACE_ID),
  );

  borrowersUnsubscribe = onSnapshot(
    borrowersQuery,
    async (snapshot) => {
      let allBorrowers = snapshot.docs.map((snapshotDoc) => ({
        id: snapshotDoc.id,
        ...snapshotDoc.data(),
      }));

      if (allBorrowers.length === 0) {
        const legacySnapshot = await getDocs(
          query(collection(db, "borrowers"), where("userId", "==", currentUser.uid)),
        );
        allBorrowers = legacySnapshot.docs.map((snapshotDoc) => ({
          id: snapshotDoc.id,
          ...snapshotDoc.data(),
        }));
      }

      borrowersCache = allBorrowers.sort((a, b) => {
        const aTs = a.createdAt?.toMillis ? a.createdAt.toMillis() : 0;
        const bTs = b.createdAt?.toMillis ? b.createdAt.toMillis() : 0;
        return bTs - aTs;
      });
      renderBorrowersTable();
      renderSummaryAndReports();
    },
    (error) => {
      if (error?.code === "permission-denied") {
        setAppStatus(
          "Firestore permissions blocked borrowers. Update your Firestore Rules.",
          "error",
          true,
        );
      } else {
        setAppStatus(`Error loading borrowers: ${error.message}`, "error", true);
      }
      renderSummaryAndReports();
    },
  );
}

// Helper function to escape HTML
function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

// Auth State Listener
onAuthStateChanged(auth, (user) => {
  if (user) {
    if (!isAuthorizedAdminEmail(user.email)) {
      setAppStatus(
        "Access denied: this account is not configured as an admin.",
        "error",
        true,
      );
      signOut(auth);
      return;
    }

    currentUser = user;
    ensureAdminUserProfile(user, localStorage.getItem("userName") || "")
      .then(() => {
        setAppStatus("Authenticated", "ok");
        if (window.location.pathname.includes("dashboard.html")) {
          syncDashboardAdminSummary(user);
          window.switchDashboardSection(currentDashboardSection);
          loadBorrowersDashboard();
        } else if (window.location.pathname.includes("borrower-actions.html")) {
          loadBorrowerActionsPage();
        }
      })
      .catch((error) => {
        setAppStatus(`Profile sync error: ${error.message}`, "error", true);
      });
  } else {
    setAppStatus("Signed out", "info");
    if (window.location.pathname.includes("dashboard.html")) {
      window.location.href = "index.html";
    }
  }
});
