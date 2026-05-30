const express    = require("express");
const nodemailer = require("nodemailer");
const multer     = require("multer");
const fetch      = require("node-fetch");
const path       = require("path");

const app  = express();
const PORT = process.env.PORT || 3000;

// ── Multer: store file in memory (PDF, JPG, PNG) ───────────
const upload = multer({
  storage: multer.memoryStorage(),
  limits : { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowed = ["application/pdf", "image/jpeg", "image/png"];
    if (allowed.includes(file.mimetype)) cb(null, true);
    else cb(new Error("Only PDF, JPG or PNG files are allowed"));
  }
});

// ── Middleware ─────────────────────────────────────────────
app.use(express.json({ limit: "15mb" }));
app.use(express.urlencoded({ extended: true, limit: "15mb" }));
app.use(express.static(path.join(__dirname, "public")));

// ── Outlook SMTP Transporter ───────────────────────────────
function createTransporter() {
  return nodemailer.createTransport({
    host  : "smtp.office365.com",
    port  : 587,
    secure: false,
    auth  : {
      user: process.env.SMTP_USER,
      pass: process.env.SMTP_PASS
    },
    tls: { rejectUnauthorized: false }
  });
}


// ============================================================
//  ROUTE 1: Serve the form
// ============================================================
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});


// ============================================================
//  ROUTE 2: Form Submission  POST /submit
// ============================================================
app.post("/submit", upload.single("FaceAuthenticationPdf"), async (req, res) => {
  try {
    const { employeeEmail, employeeId, employeeName, completionStatus, reason } = req.body;

    if (!employeeEmail || !employeeId || !employeeName || !completionStatus) {
      return res.status(400).json({ success: false, message: "All fields are required." });
    }

    // ── Forward to Google Apps Script ──────────────────────
    const pdfBase64   = req.file ? req.file.buffer.toString("base64") : "";
    const pdfFileName = req.file ? req.file.originalname : "";

    const gasPayload = {
      employeeEmail,
      employeeId,
      employeeName,
      completionStatus,
      reason        : reason || "",
      pdfFileName,
      pdfBase64     : pdfBase64 ? `data:${req.file.mimetype};base64,${pdfBase64}` : ""
    };

    const GAS_URL = process.env.GAS_URL;
    if (GAS_URL) {
      await fetch(GAS_URL, {
        method : "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body   : JSON.stringify(gasPayload)
      }).catch(err => console.error("GAS forward error:", err.message));
    }

    // ── Send confirmation email via Outlook SMTP ────────────
    const transporter = createTransporter();
    const mailOptions = {
      from    : `"HR Department" <${process.env.SMTP_USER}>`,
      to      : employeeEmail,
      cc      : process.env.CC_EMAILS || "",
      subject : `PF Face Authentication – Submission Received – ${employeeName}`,
      html    : buildConfirmationEmail(employeeName, completionStatus, reason)
    };

    if (req.file) {
      mailOptions.attachments = [{
        filename   : req.file.originalname,
        content    : req.file.buffer,
        contentType: req.file.mimetype
      }];
    }

    await transporter.sendMail(mailOptions);

    return res.json({
      success: true,
      message: completionStatus === "Completed"
        ? "Status marked as Completed. Confirmation email sent."
        : "Status recorded as Not Completed. Please complete soon."
    });

  } catch (err) {
    console.error("Submit error:", err.message);
    return res.status(500).json({ success: false, message: "Server error: " + err.message });
  }
});


// ============================================================
//  ROUTE 3: Send Daily Reminders  POST /send-reminders
//  Returns both sent and failed results back to GAS
// ============================================================
app.post("/send-reminders", async (req, res) => {
  console.log("POST /send-reminders hit");

  const token = req.headers["x-auth-token"];
  if (token !== process.env.REMINDER_SECRET) {
    console.log("Unauthorized - token mismatch");
    return res.status(401).json({ success: false, message: "Unauthorized" });
  }

  const { employees, guideBase64, guideFileName } = req.body;

  if (!employees || !Array.isArray(employees)) {
    return res.status(400).json({ success: false, message: "No employee data provided" });
  }

  console.log(`Sending reminders to ${employees.length} employees...`);

  const transporter = createTransporter();
  const results     = [];

  for (const emp of employees) {
    try {
      const mailOptions = {
        from    : `"HR Department" <${process.env.SMTP_USER}>`,
        to      : emp.email,
        cc      : process.env.CC_EMAILS || "",
        subject : `Urgent: Action Required - PF Face Authentication Pending - ${emp.name}`,
        html    : buildReminderEmail(emp.name)
      };

      if (guideBase64 && guideFileName) {
        mailOptions.attachments = [{
          filename   : guideFileName,
          content    : Buffer.from(guideBase64, "base64"),
          contentType: "application/pdf"
        }];
      }

      await transporter.sendMail(mailOptions);
      console.log(`✅ Sent to: ${emp.email}`);
      results.push({ email: emp.email, name: emp.name, sent: true, error: "" });

    } catch (err) {
      console.error(`❌ Failed → ${emp.email}: ${err.message}`);
      results.push({ email: emp.email, name: emp.name, sent: false, error: err.message });
    }
  }

  return res.json({ success: true, results });
});


// ============================================================
//  ROUTE 4: Send Salary Hold Emails  POST /send-salary-hold
//  Called on 9th June for Not Completed + Blank employees
// ============================================================
app.post("/send-salary-hold", async (req, res) => {
  console.log("POST /send-salary-hold hit");

  const token = req.headers["x-auth-token"];
  if (token !== process.env.REMINDER_SECRET) {
    return res.status(401).json({ success: false, message: "Unauthorized" });
  }

  const { employees, guideBase64, guideFileName } = req.body;

  if (!employees || !Array.isArray(employees)) {
    return res.status(400).json({ success: false, message: "No employee data provided" });
  }

  const transporter = createTransporter();
  const results     = [];

  for (const emp of employees) {
    try {
      const mailOptions = {
        from    : `"HR Department" <${process.env.SMTP_USER}>`,
        to      : emp.email,
        cc      : process.env.CC_EMAILS || "",
        subject : `Salary on Hold – PF Face Authentication Incomplete – ${emp.name}`,
        html    : buildSalaryHoldEmail(emp.name)
      };

      if (guideBase64 && guideFileName) {
        mailOptions.attachments = [{
          filename   : guideFileName,
          content    : Buffer.from(guideBase64, "base64"),
          contentType: "application/pdf"
        }];
      }

      await transporter.sendMail(mailOptions);
      console.log(`✅ Salary hold sent to: ${emp.email}`);
      results.push({ email: emp.email, name: emp.name, sent: true, error: "" });

    } catch (err) {
      console.error(`❌ Failed → ${emp.email}: ${err.message}`);
      results.push({ email: emp.email, name: emp.name, sent: false, error: err.message });
    }
  }

  return res.json({ success: true, results });
});


// ============================================================
//  EMAIL TEMPLATES
// ============================================================
function buildReminderEmail(name) {
  const formLink = process.env.FORM_LINK || "https://pf-auth-i8w8.onrender.com/";
  return `
  <div style="font-family:Arial,sans-serif;color:#333;max-width:600px;
              border:1px solid #eee;padding:20px;border-radius:10px;">
    <h2 style="color:#2563eb;margin-top:0;">Action Required: PF Face Authentication</h2>
    <p>Dear <b>${name}</b>,</p>
    <p>Our records show that you have not yet Completed your <b>PF Face Authentication</b>.</p>
    <p>Please complete the process immediately. We have attached a step-by-step guide to this email to assist you.</p>
    <div style="background:#f8fafc;padding:15px;border-left:4px solid #2563eb;margin:20px 0;">
      <p style="margin:0;"><b>Step 1:</b> Follow the attached PDF guide to complete authentication.</p>
      <p style="margin:10px 0 0;"><b>Step 2:</b> Click the button below to submit your confirmation.</p>
    </div>
    <div style="text-align:center;margin:30px 0;">
      <a href="${formLink}"
         style="background:#2563eb;color:#fff;padding:14px 28px;text-decoration:none;
                border-radius:8px;font-weight:bold;display:inline-block;">
        Confirm Completion Now
      </a>
    </div>
    <div style="background:#fff5f5;border-left:4px solid #dc2626;padding:12px 16px;margin:20px 0;border-radius:4px;">
      <p style="margin:0;color:#dc2626;font-weight:bold;font-size:14px;">⚠️ Action Required: PF Face Authentication</p>
      <p style="margin:8px 0 0;color:#7f1d1d;font-size:13px;">If you do not complete your PF Face Authentication before <b>8th June 2026</b>, your salary will be placed on <b>Hold</b>. Please take immediate action to avoid any inconvenience.</p>
    </div>
    <p>Regards,<br><b>HR Department</b></p>
  </div>`;
}

function buildSalaryHoldEmail(name) {
  const formLink = process.env.FORM_LINK || "https://pf-auth-i8w8.onrender.com/";
  return `
  <div style="font-family:Arial,sans-serif;color:#333;max-width:600px;
              border:1px solid #eee;padding:20px;border-radius:10px;">

    <div style="background:#fff1f2;border:2px solid #dc2626;border-radius:8px;
                padding:16px 20px;margin-bottom:24px;text-align:center;">
      <p style="margin:0;font-size:22px;font-weight:bold;color:#991b1b;">YOUR SALARY IS ON HOLD</p>
      <p style="margin:6px 0 0;font-size:13px;color:#b91c1c;">PF Face Authentication Not Completed</p>
    </div>

    <p>Dear <b>${name}</b>,</p>

    <p>Despite multiple reminders sent to you over the past few weeks, your
       <b>PF Face Authentication</b> remains incomplete. As communicated earlier,
       failure to complete this process would result in your salary being placed on hold.</p>

    <p>Effective immediately, <b style="color:#dc2626;">your salary has been placed on hold</b>
       until you complete the PF Face Authentication process and submit your confirmation.</p>

    <div style="background:#f8fafc;padding:15px;border-left:4px solid #dc2626;margin:20px 0;">
      <p style="margin:0;font-weight:bold;">To release your salary, complete these steps immediately:</p>
      <p style="margin:10px 0 0;"><b>Step 1:</b> Follow the attached PDF guide to complete PF Face Authentication.</p>
      <p style="margin:10px 0 0;"><b>Step 2:</b> Click the button below to submit your confirmation.</p>
    </div>

    <div style="text-align:center;margin:30px 0;">
      <a href="${formLink}"
         style="background:#dc2626;color:#fff;padding:14px 28px;text-decoration:none;
                border-radius:8px;font-weight:bold;display:inline-block;">
        Complete Authentication Now
      </a>
    </div>

    <p style="color:#666;font-size:12px;border-top:1px solid #eee;padding-top:10px;">
      For any assistance, please contact HR immediately. Your salary will be released
      as soon as your authentication is verified.
    </p>
    <p>Regards,<br><b>HR Department</b></p>
  </div>`;
}

function buildConfirmationEmail(name, status, reason) {
  const isCompleted = status === "Completed";
  return `
  <div style="font-family:Arial,sans-serif;color:#333;max-width:600px;
              border:1px solid #eee;padding:20px;border-radius:10px;">
    <h2 style="color:${isCompleted ? "#166534" : "#92400e"};margin-top:0;">
      PF Face Authentication – ${isCompleted ? "Submission Received ✅" : "Status Recorded ⚠️"}
    </h2>
    <p>Dear <b>${name}</b>,</p>
    ${isCompleted
      ? `<p>Thank you! Your PF Face Authentication status has been recorded as <b>Completed</b>.</p>
         <p>HR will verify your uploaded document. Daily reminder emails will stop once verified.</p>`
      : `<p>Your status has been recorded as <b>Not Completed</b>.</p>
         ${reason ? `<p><b>Reason:</b> ${reason}</p>` : ""}
         <p>Please complete the PF Face Authentication before <b>8th June 2026</b> to avoid your salary being placed on hold.</p>`
    }
    <p>Regards,<br><b>HR Department</b></p>
  </div>`;
}


// ============================================================
//  START SERVER
// ============================================================
app.listen(PORT, () => {
  console.log(`✅ PF Auth server running on port ${PORT}`);
});
