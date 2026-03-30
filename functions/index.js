const admin = require("firebase-admin");
const { onDocumentCreated } = require("firebase-functions/v2/firestore");
const { onSchedule } = require("firebase-functions/v2/scheduler");
const { onRequest } = require("firebase-functions/v2/https");
const { defineSecret } = require("firebase-functions/params");
const { logger } = require("firebase-functions");

admin.initializeApp();

const RESEND_API_KEY = defineSecret("RESEND_API_KEY");
const RESEND_FROM_EMAIL = defineSecret("RESEND_FROM_EMAIL");
const TEST_EMAIL_TOKEN = defineSecret("TEST_EMAIL_TOKEN");
const TIME_ZONE = "Africa/Lusaka";

function formatDateForTimeZone(date, timeZone = TIME_ZONE) {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(date);
}

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

async function sendEmail({
  apiKey,
  fromEmail,
  toEmail,
  subject,
  html,
  text,
}) {
  const response = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      from: fromEmail,
      to: [toEmail],
      subject,
      html,
      text,
    }),
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Resend API error (${response.status}): ${body}`);
  }
}

exports.sendBorrowerWelcomeEmail = onDocumentCreated(
  {
    document: "borrowers/{borrowerId}",
    region: "us-central1",
    secrets: [RESEND_API_KEY, RESEND_FROM_EMAIL],
  },
  async (event) => {
    const borrower = event.data?.data();
    if (!borrower) return;

    const borrowerEmail = borrower.email;
    if (!borrowerEmail) {
      logger.info("Borrower has no email; skipping welcome email.");
      return;
    }

    const borrowerName = borrower.name || "Borrower";
    const totalToPay = Number(borrower.totalToPay || 0).toLocaleString(
      "en-ZM",
      {
        style: "currency",
        currency: "ZMW",
      },
    );
    const dueDate = borrower.dueDate || "N/A";

    await sendEmail({
      apiKey: RESEND_API_KEY.value(),
      fromEmail: RESEND_FROM_EMAIL.value(),
      toEmail: borrowerEmail,
      subject: "Loan Created - Mambo Finance",
      text: `Hello ${borrowerName}, your loan has been recorded. Total to pay: ${totalToPay}. Due date: ${dueDate}.`,
      html: `
        <div style="font-family: Arial, sans-serif; line-height: 1.5;">
          <h2>Loan Confirmation</h2>
          <p>Hello <strong>${borrowerName}</strong>,</p>
          <p>Your loan has been created in our system.</p>
          <ul>
            <li><strong>Total to Pay:</strong> ${totalToPay}</li>
            <li><strong>Due Date:</strong> ${dueDate}</li>
          </ul>
          <p>Thank you.</p>
        </div>
      `,
    });

    await event.data.ref.set(
      {
        welcomeEmailSentAt: admin.firestore.FieldValue.serverTimestamp(),
      },
      { merge: true },
    );
  },
);

exports.sendDueDateReminderEmails = onSchedule(
  {
    schedule: "0 8 * * *",
    timeZone: TIME_ZONE,
    region: "us-central1",
    secrets: [RESEND_API_KEY, RESEND_FROM_EMAIL],
  },
  async () => {
    const targetDate = formatDateForTimeZone(addDays(new Date(), 5));
    logger.info(`Checking reminders for due date ${targetDate}`);

    const snapshot = await admin
      .firestore()
      .collection("borrowers")
      .where("dueDate", "==", targetDate)
      .where("status", "==", "pending")
      .get();

    if (snapshot.empty) {
      logger.info("No borrowers due in 5 days.");
      return;
    }

    const tasks = snapshot.docs.map(async (docSnap) => {
      const borrower = docSnap.data();
      const borrowerEmail = borrower.email;
      if (!borrowerEmail) return;

      if (borrower.reminderSentForDueDate === targetDate) {
        return;
      }

      const borrowerName = borrower.name || "Borrower";
      const totalToPay = Number(borrower.totalToPay || 0).toLocaleString(
        "en-ZM",
        {
          style: "currency",
          currency: "ZMW",
        },
      );

      await sendEmail({
        apiKey: RESEND_API_KEY.value(),
        fromEmail: RESEND_FROM_EMAIL.value(),
        toEmail: borrowerEmail,
        subject: "Loan Due Reminder - 5 Days Remaining",
        text: `Hello ${borrowerName}, this is a reminder that your loan is due on ${targetDate}. Total due: ${totalToPay}.`,
        html: `
          <div style="font-family: Arial, sans-serif; line-height: 1.5;">
            <h2>Loan Due Reminder</h2>
            <p>Hello <strong>${borrowerName}</strong>,</p>
            <p>This is a reminder that your loan is due in <strong>5 days</strong>.</p>
            <ul>
              <li><strong>Due Date:</strong> ${targetDate}</li>
              <li><strong>Total to Pay:</strong> ${totalToPay}</li>
            </ul>
            <p>Please plan payment before the due date.</p>
          </div>
        `,
      });

      await docSnap.ref.set(
        {
          reminderSentForDueDate: targetDate,
          reminderSentAt: admin.firestore.FieldValue.serverTimestamp(),
        },
        { merge: true },
      );
    });

    await Promise.all(tasks);
    logger.info(`Processed ${snapshot.size} reminder candidates.`);
  },
);

exports.sendTestEmail = onRequest(
  {
    region: "us-central1",
    secrets: [RESEND_API_KEY, RESEND_FROM_EMAIL, TEST_EMAIL_TOKEN],
  },
  async (req, res) => {
    try {
      if (req.method !== "POST") {
        res.status(405).json({ ok: false, message: "Use POST." });
        return;
      }

      const token =
        req.get("x-test-token") ||
        req.query.token ||
        req.body?.token ||
        "";

      if (!token || token !== TEST_EMAIL_TOKEN.value()) {
        res.status(401).json({ ok: false, message: "Unauthorized." });
        return;
      }

      const toEmail = req.body?.to || req.query.to;
      if (!toEmail) {
        res
          .status(400)
          .json({ ok: false, message: "Missing 'to' email address." });
        return;
      }

      await sendEmail({
        apiKey: RESEND_API_KEY.value(),
        fromEmail: RESEND_FROM_EMAIL.value(),
        toEmail,
        subject: "Mambo Finance Test Email",
        text: "This is a test email from your Firebase Functions setup.",
        html: `
          <div style="font-family: Arial, sans-serif; line-height: 1.5;">
            <h2>Mambo Finance - Test Email</h2>
            <p>Your email integration is working correctly.</p>
            <p>Sent at: ${new Date().toISOString()}</p>
          </div>
        `,
      });

      res.status(200).json({
        ok: true,
        message: `Test email sent to ${toEmail}`,
      });
    } catch (error) {
      logger.error("sendTestEmail failed", error);
      res.status(500).json({
        ok: false,
        message: error.message || "Failed to send test email.",
      });
    }
  },
);
