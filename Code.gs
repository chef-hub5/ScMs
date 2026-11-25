const USERS = "Users";
const MESSAGES = "Inbox";
const PAYMENTS = "Payments";
const SCHOOL_FEE_ATTEMPTS = "School Fee Payment Attempts";
const MESSAGE_READS = "Message Reads";
const NOTIFICATION_CLEARS = "Notification Clears";
const RECORDS = "Records";
const TENANTS = "Schools";
const SUBSCRIPTION_PAYMENTS = "Subscription Payments";
const CONTACT_MESSAGES = "Contact Messages";
const SUPPORT_FEEDBACK = "Support Feedback";
const DEFAULT_CAMPUS_ID = "campus-default";
const CBT_EXAMS_SHEET = "CBT Exams";
const CBT_CODES_SHEET = "CBT Access Codes";
const CBT_ATTEMPTS_SHEET = "CBT Attempts";
const CBT_ATTEMPT_SHEET = CBT_ATTEMPTS_SHEET;
const CBT_QUESTION_BANK_SHEET = "CBT Question Bank";
const ID_CARDS_SHEET = "ID Cards";
const CBT_EXAM_HEADERS = ["id", "campusId", "title", "className", "subject", "term", "academicYear", "opensAt", "closesAt", "durationMinutes", "status", "questionsJson", "createdByEmail", "createdByName", "reviewNotes", "resultsReleased", "createdAt", "updatedAt"];
const CBT_CODE_HEADERS = ["id", "examId", "studentId", "codeHash", "issuedAt", "usedAt", "revokedAt"];
const CBT_ATTEMPT_HEADERS = ["id", "examId", "studentId", "attemptTokenHash", "startedAt", "expiresAt", "submittedAt", "answersJson", "score", "maxScore", "writtenMarksJson", "status", "integrityEventsJson"];
const CBT_QUESTION_BANK_HEADERS = ["id", "campusId", "className", "subject", "topic", "difficulty", "questionJson", "createdByEmail", "createdAt"];
const ID_CARD_HEADERS = ["id", "recordType", "recordId", "campusId", "cardNumber", "validFrom", "validUntil", "issuedAt", "revision", "status"];
const CBT_JSON_COLUMNS = {
  [CBT_EXAMS_SHEET]: ["questionsJson"],
  [CBT_ATTEMPTS_SHEET]: ["answersJson", "writtenMarksJson", "integrityEventsJson"],
  [CBT_QUESTION_BANK_SHEET]: ["questionJson"]
};
const RECORD_SHEETS = {
  campus: "Campuses",
  student: "Students",
  employee: "Employees",
  parent: "Parents",
  application: "Applications",
  academic: "Academic",
  attendance: "Attendance",
  employeeAttendance: "Employee Attendance",
  employeeLeave: "Employee Leave Requests",
  timetable: "Timetable",
  reportCard: "Report Cards",
  gradeScale: "Grading Scales",
  gradeEntry: "Teacher Grade Entries",
  followUp: "Student Follow-ups",
  fee: "Fees",
  feeInvoice: "Fee Invoices",
  feePlan: "Fee Types",
  event: "Calendar",
  announcement: "Inbox"
};
const SYSTEM_SHEETS = ["Dashboard", "Reports"];
const TERMS_VERSION = "scholaro-terms-v4";
const SCHOOL_DATA_EDIT_NOTICE = "Manage school records through Clasten. Changes made directly in Google Sheets will not synchronize with Clasten.";
let requestSpreadsheetId = "";
let requestSpreadsheetCache = new Map();
let requestTenantRows = null;

function setup() {
  const properties = PropertiesService.getScriptProperties();
  let registryId = String(properties.getProperty("MASTER_SPREADSHEET_ID") || "");
  if (!registryId) {
    registryId = SpreadsheetApp.create("Clasten School Registry").getId();
    properties.setProperty("MASTER_SPREADSHEET_ID", registryId);
  }
  const registry = openSpreadsheetById_(registryId);
  ensureHeadersIn_(registry, TENANTS, [
    "tenantId", "schoolName", "schoolCode", "spreadsheetId", "adminEmail",
    "plan", "billingPeriod", "subscriptionStatus", "trialEndsAt", "paidThrough",
    "createdAt", "portalLogo", "schoolProfilePicture", "portalBackgroundImage", "brandColor", "brandingSetupCompleted",
    "termsVersion", "termsAcceptedAt"
  ]);
  ensureHeadersIn_(registry, SUBSCRIPTION_PAYMENTS, [
    "reference", "schoolCode", "adminEmail", "plan", "billingPeriod",
    "amountNaira", "currency", "paidAt", "paidThrough", "paymentChannel",
    "ownerEmailStatus", "scholaroEmailStatus", "annualDiscountApplied"
  ]);
  ensureHeadersIn_(registry, CONTACT_MESSAGES, [
    "submittedAt", "name", "email", "subject", "message", "emailStatus"
  ]);
  ensureHeadersIn_(registry, SUPPORT_FEEDBACK, [
    "submittedAt", "ticketId", "tenantId", "schoolName", "schoolCode",
    "userName", "userEmail", "userRole", "replyEmail", "category",
    "currentPage", "message", "emailStatus"
  ]);
  const legacySpreadsheet = SpreadsheetApp.getActive();
  SYSTEM_SHEETS.forEach(name => ensureHeadersIn_(legacySpreadsheet, name, ["message"]));
}

function authorizeMailService() {
  const remainingQuota = MailApp.getRemainingDailyQuota();
  console.log(`Mail service authorized. Remaining daily recipient quota: ${remainingQuota}.`);
  return remainingQuota;
}

function unregisterLegacySchool() {
  const schoolCode = normalizeSchoolCode_(
    PropertiesService.getScriptProperties().getProperty("DEFAULT_TENANT_CODE")
  );
  if (!schoolCode) {
    throw new Error("Set DEFAULT_TENANT_CODE to the old school's exact code before running this cleanup.");
  }
  const sheet = registrySheet_();
  const rows = readRowsFrom_(sheet);
  const index = rows.findIndex(row => normalizeSchoolCode_(row.schoolCode) === schoolCode);
  if (index < 0) {
    console.log(`No registered school found for code "${schoolCode}". No changes made.`);
    return;
  }
  sheet.deleteRow(index + 2);
  console.log(`Unregistered school "${schoolCode}". Its original spreadsheet was not deleted.`);
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.action === "paystackReturn") return json_(verifySubscriptionPayment_(params.reference, params.tenant));
  return json_({
    ok: true,
    service: "Clasten School Management API"
  });
}

function doPost(e) {
  requestSpreadsheetId = "";
  requestSpreadsheetCache = new Map();
  requestTenantRows = null;
  const startedAt = Date.now();
  let action = "unknown";
  try {
    const body = JSON.parse(e.postData.contents || "{}");
    action = String(body.action || (body.event ? "webhook" : "unknown"));
    if (body.event === "charge.success" && body.data && body.data.reference) {
      const cachedPayment = CacheService.getScriptCache().get(`pay:${body.data.reference}`);
      if (cachedPayment) {
        const pendingPayment = JSON.parse(cachedPayment);
        return json_(verifySubscriptionPayment_(body.data.reference, pendingPayment.tenantId));
      }
      const cachedSchoolFee = CacheService.getScriptCache().get(`schoolfee:${body.data.reference}`);
      if (cachedSchoolFee) {
        const pendingPayment = JSON.parse(cachedSchoolFee);
        const result = verifySchoolFeePayment_(body.data.reference, pendingPayment.tenantId);
        delete result.payment;
        delete result.invoice;
        return json_(result);
      }
      return json_({ ok: true, ignored: true });
    }
    if (body.action === "verifySchoolFeePayment") {
      const result = verifySchoolFeePayment_(body.reference, body.tenant);
      if (result.payment) {
        const authContext = getSessionContext_(body.authToken);
        const role = String(authContext && authContext.account.role || "").toLowerCase();
        const sameTenant = authContext &&
          normalizeSchoolCode_(authContext.session.tenantId) === normalizeSchoolCode_(body.tenant);
        const sameParent = role === "parent" &&
          String(authContext.account.email || "").trim().toLowerCase() ===
            String(result.payment.parentEmail || "").trim().toLowerCase();
        if (!sameTenant || (role !== "admin" && !sameParent)) {
          delete result.payment;
          delete result.invoice;
        }
      }
      return json_(result);
    }
    if (["cbtStart", "cbtResume", "cbtSave", "cbtSubmit", "cbtResult", "cbtIntegrity"].includes(body.action)) {
      return json_(handlePublicCbtAction_(body));
    }
    if (body.action === "signupRequest") return json_(signupRequest_(body));
    if (body.action === "verifySignup") return json_(verifySignup_(body));
    if (body.action === "activateCampusAdmin") return json_(activateCampusAdmin_(body));
    if (body.action === "requestPasswordReset") return json_(requestPasswordReset_(body));
    if (body.action === "resetPassword") return json_(resetPassword_(body));
    if (body.action === "schoolPortal") return json_(getSchoolPortal_(body.schoolCode));
    if (body.action === "submitApplication") return json_(submitApplication_(body));
    if (body.action === "getPlacementOffer") return json_(getPlacementOffer_(body));
    if (body.action === "respondPlacement") return json_(respondPlacement_(body));
    if (body.action === "activateParent") return json_(activateParent_(body));
    if (body.action === "activateEmployee") return json_(activateEmployee_(body));
    if (body.action === "contactMessage") return json_(sendContactMessage_(body));
    if (body.action === "plans") {
      let annualDiscountEligible = true;
      const session = body.authToken ? getSession_(body.authToken) : null;
      if (session) {
        const tenant = findTenant_(session.tenantId);
        if (!tenant) return json_({ ok: false, error: "School account could not be found." });
        const reservationKey = annualDiscountReservationKey_(tenant.tenantId);
        annualDiscountEligible =
          !hasUsedAnnualDiscount_(tenant.tenantId) &&
          !CacheService.getScriptCache().get(reservationKey);
      }
      return json_({ ok: true, plans: getPlans_({ annualDiscountEligible: annualDiscountEligible }) });
    }
    if (body.action === "login") return json_(login_(body));
    if (body.action === "completeAdminCampusLogin") return json_(completeAdminCampusLogin_(body));

    if (body.action === "paystackReturn") {
      return json_(verifySubscriptionPayment_(body.reference, body.tenant));
    }

    if (body.action === "startSubscription") {
      const authContext = getSessionContext_(body.authToken);
      if (!authContext) return json_({ ok: false, error: "Your session has expired. Please sign in again." });
      const session = authContext.session;
      requestSpreadsheetId = session.spreadsheetId;
      body.authUser = authContext.account;
      if (!body.authUser || String(body.authUser.role).toLowerCase() !== "admin") {
        return json_({ ok: false, error: "Only a school administrator can manage billing." });
      }
      return json_(startSubscription_(body, session.tenantId));
    }

    let session;
    {
      const authContext = getSessionContext_(body.authToken);
      if (!authContext) {
        return json_({ ok: false, error: "Your session has expired. Please sign in again." });
      }
      body.authUser = authContext.account;
      session = authContext.session;
    }
    const tenant = session && findTenant_(session.tenantId);
    if (!tenant) return json_({ ok: false, error: "School account could not be found. Contact Clasten support." });
    body.schoolTenant = tenant;
    const actorRole = String(body.authUser.role || body.authUser.type || "").toLowerCase();
    if (actorRole === "admin" && !isSchoolOwnerAdmin_(body.authUser, tenant)) {
      const assignedCampusId = String(body.authUser.campusId || "").trim();
      const campuses = ensureDefaultCampus_();
      if (!assignedCampusId || !campuses.some(campus => String(campus.id || "") === assignedCampusId)) {
        return json_({ ok: false, error: "This administrator is not assigned to an active campus. Contact the school owner." });
      }
      const requestedCampusId = String(body.campusId || body.record && body.record.campusId || "").trim();
      if (requestedCampusId && requestedCampusId !== assignedCampusId) {
        return json_({ ok: false, error: "You are not authorized to access that campus." });
      }
      const ownerOnlyActions = ["startSubscription", "schoolPaystackSettings", "saveSchoolPaystackSettings", "updateProfile", "getAuditLog", "inviteCampusAdmin", "saveUser", "completeBrandingSetup", "announcement", "editMessage", "deleteMessage"];
      if (ownerOnlyActions.includes(String(body.action || ""))) {
        return json_({ ok: false, error: "Only the school owner can manage this school-wide setting." });
      }
      if (body.action === "saveRecord" && ["campus", "feePlan", "gradeScale"].includes(String(body.type || "")) ||
          body.action === "deleteRecord" && ["campus", "feePlan", "gradeScale"].includes(String(body.type || ""))) {
        return json_({ ok: false, error: "Only the school owner can manage this school-wide record." });
      }
      body.campusId = assignedCampusId;
      if (body.record && typeof body.record === "object" &&
          !["parent", "feePlan", "gradeScale", "campus"].includes(String(body.type || ""))) {
        body.record.campusId = assignedCampusId;
      }
    }
    const readOnlyActions = ["getRecords", "messages", "markMessagesRead", "clearNotifications", "updateProfile", "updateNotificationPreferences", "subscription", "schoolPaystackSettings", "getAuditLog", "supportFeedback", "cbtListExams", "cbtListAttempts", "cbtListQuestionBank", "cbtGetExamForAuthor", "getIdCardRoster"];
    const cbtActions = ["cbtListExams", "cbtListAttempts", "cbtListQuestionBank", "cbtGetExamForAuthor", "cbtSaveQuestionBank", "cbtCreateExam", "cbtUpdateExam", "cbtDeleteExam", "cbtReviewExam", "cbtGenerateAccessCodes", "cbtMarkWritten", "cbtReleaseResults"];
    const requiredFeature = body.action === "getIdCardRoster" || body.action === "issueIdCards"
      ? "idCards"
      : cbtActions.includes(body.action) ? "cbt" : "";
    if (requiredFeature && !planSupportsFeature_(tenant.plan, requiredFeature)) {
      const featureName = requiredFeature === "idCards" ? "Printable student and employee ID cards" : "CBT exams";
      return json_({ ok: false, error: `${featureName} require the Enterprise plan. Upgrade the school plan to continue.` });
    }
    if (!isSubscriptionUsable_(tenant) && !readOnlyActions.includes(body.action)) {
      return json_({ ok: false, error: "This school's trial or subscription has ended. Renew the plan to make changes." });
    }

    switch (body.action) {
      case "saveUser":
        return json_(saveUser_(body));
      case "updateProfile":
        return json_(updateProfile_(body));
      case "updateNotificationPreferences":
        return json_(updateNotificationPreferences_(body));
      case "completeBrandingSetup":
        return json_(completeBrandingSetup_(body));
      case "announcement":
        if (!planSupportsFeature_(tenant.plan, "announcements")) {
          return json_({ ok: false, error: "School announcements require the Enterprise plan. Upgrade the school plan to continue." });
        }
        return json_(saveAnnouncement_(body));
      case "sendReportCard":
        if (!planSupportsFeature_(tenant.plan, "reportCards")) {
          return json_({ ok: false, error: "Student report cards require the Enterprise plan. Upgrade the school plan to continue." });
        }
        return json_(sendReportCard_(body));
      case "messages":
        return json_(getMessages_(body));
      case "markMessagesRead":
        return json_(markMessagesRead_(body));
      case "replyToMessage":
        return json_(replyToMessage_(body));
      case "messageAdmin":
        return json_(messageSchoolAdmin_(body));
      case "editMessage":
        return json_(editMessage_(body));
      case "deleteMessage":
        return json_(deleteMessage_(body));
      case "clearNotifications":
        return json_(clearNotifications_(body));
      case "supportFeedback":
        return json_(submitSupportFeedback_(body));
      case "subscription":
        return json_({
          ok: true, plan: tenant.plan, billingPeriod: tenant.billingPeriod,
          subscriptionStatus: tenant.subscriptionStatus,
          accessActive: isSubscriptionUsable_(tenant),
          trialEndsAt: tenant.trialEndsAt, paidThrough: tenant.paidThrough,
          portalLogo: tenant.portalLogo || "", brandColor: tenant.brandColor || "",
          schoolProfilePicture: tenant.schoolProfilePicture || findUserByEmail_(tenant.adminEmail)?.profilePicture || "",
          brandingSetupCompleted: String(tenant.brandingSetupCompleted || "").toLowerCase() === "true"
        });
      case "payment":
        if (!planSupportsFeature_(tenant.plan, "fees")) {
          return json_({ ok: false, error: "Fee payments require the Standard or Enterprise plan. Ask the school administrator to upgrade." });
        }
        return json_({ ok: false, error: "Parent fee payments must be completed through verified Paystack checkout." });
      case "schoolPaystackSettings":
        return json_(getSchoolPaystackSettings_(body, tenant.tenantId));
      case "saveSchoolPaystackSettings":
        return json_(saveSchoolPaystackSettings_(body, tenant.tenantId));
      case "startSchoolFeePayment":
        if (!planSupportsFeature_(tenant.plan, "fees")) {
          return json_({ ok: false, error: "Online fee payments require the Standard or Enterprise plan." });
        }
        return json_(startSchoolFeePayment_(body, tenant));
      case "getRecords":
        return json_(getRecords_(body));
      case "cbtListExams":
        return json_(listCbtExams_(body));
      case "cbtGetExamForAuthor":
        return json_(getCbtExamForAuthor_(body));
      case "cbtListQuestionBank":
        return json_(listCbtQuestionBank_(body));
      case "cbtSaveQuestionBank":
        return json_(saveCbtQuestionBank_(body));
      case "cbtCreateExam":
        return json_(createCbtExam_(body));
      case "cbtUpdateExam":
        return json_(updateCbtExam_(body));
      case "cbtDeleteExam":
        return json_(deleteCbtExam_(body));
      case "cbtReviewExam":
        return json_(reviewCbtExam_(body));
      case "cbtGenerateAccessCodes":
        return json_(generateCbtAccessCodes_(body));
      case "cbtListAttempts":
        return json_(listCbtAttempts_(body));
      case "cbtMarkWritten":
        return json_(markCbtWritten_(body));
      case "cbtReleaseResults":
        return json_(releaseCbtResults_(body));
      case "getIdCardRoster":
        return json_(getIdCardRoster_(body));
      case "issueIdCards":
        return json_(issueIdCards_(body));
      case "createFeeInvoices":
        if (!planSupportsFeature_(tenant.plan, "fees")) {
          return json_({ ok: false, error: "Fee invoices require the Standard or Enterprise plan." });
        }
        return json_(createFeeInvoices_(body));
      case "recordFeePayment":
        if (!planSupportsFeature_(tenant.plan, "fees")) {
          return json_({ ok: false, error: "Fee payments require the Standard or Enterprise plan." });
        }
        return json_(recordFeePayment_(body));
      case "reconcileFeePayment":
        if (!planSupportsFeature_(tenant.plan, "fees")) {
          return json_({ ok: false, error: "Fee reconciliation requires the Standard or Enterprise plan." });
        }
        return json_(reconcileFeePayment_(body));
      case "sendFeeReminders":
        if (!planSupportsFeature_(tenant.plan, "fees")) {
          return json_({ ok: false, error: "Fee reminders require the Standard or Enterprise plan." });
        }
        return json_(sendFeeReminders_(body));
      case "getAuditLog":
        return json_(getAuditLog_(body));
      case "saveRecord": {
        const recordType = String(body.type || "");
        if (!["campus", "parent", "feePlan"].includes(recordType) && body.record && typeof body.record === "object") {
          const campusId = validatedCampusId_(tenant, body.record.campusId || body.campusId || DEFAULT_CAMPUS_ID);
          if (!campusId) return json_({ ok: false, error: "Select an active campus in this school workspace." });
          body.record.campusId = campusId;
        }
        if (recordType === "fee" && !planSupportsFeature_(tenant.plan, "fees")) {
          return json_({ ok: false, error: "Fee payments require the Standard or Enterprise plan. Upgrade the school plan to continue." });
        }
        if (recordType === "reportCard" && !planSupportsFeature_(tenant.plan, "reportCards")) {
          return json_({ ok: false, error: "Student report cards require the Enterprise plan. Upgrade the school plan to continue." });
        }
        if (["gradeScale", "gradeEntry"].includes(recordType) && !planSupportsFeature_(tenant.plan, "reportCards")) {
          return json_({ ok: false, error: "Grading scales and teacher grade entry require the Enterprise plan." });
        }
        let accountSaved = false;
        if (["employee", "parent"].includes(String(body.type || "")) && body.user) {
          const accountResult = saveUser_({
            authUser: body.authUser,
            user: body.user,
            previousEmail: body.previousEmail
          });
          if (!accountResult.ok) return json_(accountResult);
          accountSaved = true;
        }
        return json_(saveRecord_(body, accountSaved));
      }
      case "saveAttendanceRegister":
        return json_(saveAttendanceRegister_(body));
      case "saveEmployeeLeave":
        return json_(saveEmployeeLeave_(body));
      case "reviewEmployeeLeave":
        return json_(reviewEmployeeLeave_(body));
      case "deleteRecord":
        if (String(body.type || "") === "campus" && !planSupportsFeature_(tenant.plan, "multiCampus")) {
          return json_({ ok: false, error: "Multiple campuses require the Enterprise plan. Upgrade the school plan to continue." });
        }
        return json_(deleteRecord_(body));
      case "reviewApplication":
        return json_(reviewApplication_(body));
      case "inviteParent":
        return json_(inviteParent_(body));
      case "inviteEmployee":
        return json_(inviteEmployee_(body));
      case "inviteCampusAdmin":
        return json_(inviteCampusAdmin_(body));
      default:
        return json_({ ok: false, error: "Unknown action" });
    }
  } catch (error) {
    return json_({ ok: false, error: error.message });
  } finally {
    const durationMs = Date.now() - startedAt;
    if (durationMs >= 500) console.log(`Slow Apps Script request: ${action} (${durationMs} ms)`);
  }
}

function cbtRows_(sheetName, headers) {
  const sheet = ensureHeaders_(sheetName, headers);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const columnNames = values[0].map(value => String(value).trim());
  const idIndex = columnNames.indexOf("id");
  return values.slice(1).filter(row => String(row[idIndex] || "").trim())
    .map(row => rowToObject_(columnNames, row));
}

function cbtSheetCellValue_(sheetName, columnName, value) {
  if (value && typeof value === "object" &&
      (CBT_JSON_COLUMNS[sheetName] || []).includes(columnName)) {
    return JSON.stringify(value);
  }
  return value;
}

function cbtWrite_(sheetName, headers, record) {
  const sheet = ensureHeaders_(sheetName, headers);
  const values = sheet.getDataRange().getValues();
  const columnNames = values[0].map(value => String(value).trim());
  const idIndex = columnNames.indexOf("id");
  const rowIndex = values.slice(1).findIndex(row => String(row[idIndex] || "") === String(record.id));
  const rowValues = columnNames.map(name => cbtSheetCellValue_(
    sheetName, name, Object.prototype.hasOwnProperty.call(record, name) ? record[name] : ""
  ));
  if (rowIndex < 0) writeSafeRows_(sheet, Math.max(2, sheet.getLastRow() + 1), [rowValues]);
  else writeSafeRows_(sheet, rowIndex + 2, [rowValues]);
}

function cbtWriteMany_(sheetName, headers, records) {
  if (!records.length) return;
  const sheet = ensureHeaders_(sheetName, headers);
  const values = sheet.getDataRange().getValues();
  const columnNames = values[0].map(value => String(value).trim());
  const idIndex = columnNames.indexOf("id");
  const rowById = new Map();
  values.slice(1).forEach((row, index) => {
    const id = String(row[idIndex] || "");
    if (id) rowById.set(id, index + 2);
  });
  const updatesByRow = new Map(), newRecords = [];
  records.forEach(record => {
    const rowValues = columnNames.map(name => cbtSheetCellValue_(
      sheetName, name, Object.prototype.hasOwnProperty.call(record, name) ? record[name] : ""
    ));
    const rowNumber = rowById.get(String(record.id || ""));
    if (rowNumber) updatesByRow.set(rowNumber, rowValues);
    else newRecords.push(rowValues);
  });
  const updates = [...updatesByRow.entries()].sort((a, b) => a[0] - b[0]);
  let startRow = 0, group = [];
  updates.forEach(([rowNumber, rowValues]) => {
    if (!group.length) {
      startRow = rowNumber;
      group = [rowValues];
    } else if (rowNumber === startRow + group.length) {
      group.push(rowValues);
    } else {
      writeSafeRows_(sheet, startRow, group);
      startRow = rowNumber;
      group = [rowValues];
    }
  });
  if (group.length) writeSafeRows_(sheet, startRow, group);
  if (newRecords.length) writeSafeRows_(sheet, Math.max(2, sheet.getLastRow() + 1), newRecords);
}

function cbtAdmin_(body) {
  return String(body.authUser && body.authUser.role || "").toLowerCase() === "admin";
}

function cbtTeacherContext_(body) {
  if (!body.authUser || String(body.authUser.role || "").toLowerCase() !== "employee") return null;
  const email = String(body.authUser.email || "").trim().toLowerCase();
  const employee = readSchoolRecords_("employee").find(item =>
    String(item.email || "").trim().toLowerCase() === email
  );
  if (!employee || !/teacher/i.test(String(employee.role || body.authUser.jobRole || ""))) return null;
  const campusId = effectiveCampusId_(employee);
  const assignments = [];
  const add = (className, subject) => {
    className = String(className || "").trim();
    subject = String(subject || "").trim();
    if (className && subject && !assignments.some(item =>
      item.className.toLowerCase() === className.toLowerCase() &&
      item.subject.toLowerCase() === subject.toLowerCase()
    )) assignments.push({ className: className, subject: subject });
  };
  const subjects = String(employee.subjects || "").split(/[,;\n]/).map(value => value.trim()).filter(Boolean);
  subjects.forEach(subject => add(employee.className, subject));
  readSchoolRecords_("timetable").filter(item =>
    String(item.teacherEmail || "").trim().toLowerCase() === email &&
    effectiveCampusId_(item) === campusId
  ).forEach(item => add(item.className, item.subject));
  return { email: email, employee: employee, campusId: campusId, assignments: assignments };
}

function cbtTeacherAssigned_(context, className, subject, campusId) {
  return Boolean(context && context.campusId === String(campusId || DEFAULT_CAMPUS_ID) &&
    context.assignments.some(item =>
      item.className.toLowerCase() === String(className || "").trim().toLowerCase() &&
      item.subject.toLowerCase() === String(subject || "").trim().toLowerCase()
    ));
}

function cbtQuestionSet_(input) {
  if (!Array.isArray(input) || !input.length || input.length > 100) {
    return { error: "Add between 1 and 100 exam questions." };
  }
  const ids = new Set();
  const questions = [];
  for (let index = 0; index < input.length; index += 1) {
    const raw = input[index] || {};
    const type = String(raw.type || "multipleChoice");
    const prompt = String(raw.prompt || "").trim();
    const marks = Number(raw.marks);
    if (!["multipleChoice", "trueFalse", "written"].includes(type) ||
        !prompt || prompt.length > 2000 || !Number.isFinite(marks) || marks <= 0 || marks > 100) {
      return { error: `Question ${index + 1} needs a valid type, prompt, and mark value.` };
    }
    const questionId = Utilities.getUuid();
    if (ids.has(questionId)) return { error: "Question identifiers must be unique." };
    ids.add(questionId);
    const question = {
      id: questionId,
      type: type,
      prompt: prompt,
      marks: marks,
      topic: String(raw.topic || "").trim().slice(0, 100),
      difficulty: ["easy", "medium", "hard"].includes(String(raw.difficulty || "").toLowerCase())
        ? String(raw.difficulty).toLowerCase() : "medium"
    };
    if (type === "multipleChoice") {
      const options = Array.isArray(raw.options) ? raw.options.map(value => String(value).trim()) : [];
      const correctIndex = String(raw.correctIndex == null ? "" : raw.correctIndex).trim() === ""
        ? NaN : Number(raw.correctIndex);
      if (options.length < 2 || options.length > 8 || options.some(value => !value || value.length > 300) ||
          new Set(options.map(value => value.toLowerCase())).size !== options.length ||
          !Number.isInteger(correctIndex) || correctIndex < 0 || correctIndex >= options.length) {
        return { error: `Question ${index + 1} needs 2 to 8 unique options and one correct option.` };
      }
      question.options = options;
      question.correctIndex = correctIndex;
    } else if (type === "trueFalse") {
      if (!["true", "false"].includes(String(raw.correctAnswer || "").toLowerCase())) {
        return { error: `Question ${index + 1} needs a true or false answer key.` };
      }
      question.correctAnswer = String(raw.correctAnswer).toLowerCase();
    }
    questions.push(question);
  }
  return { questions: questions };
}

function cbtNormalizeSavedQuestion_(raw, examId, index) {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
    throw new Error(`Question ${index + 1} is not a valid question record.`);
  }
  const nested = raw.question && typeof raw.question === "object" && !Array.isArray(raw.question)
    ? raw.question : {};
  const source = Object.assign({}, raw, nested);
  const prompt = String(source.prompt || source.questionText || source.text ||
    (typeof source.question === "string" ? source.question : "")).trim();
  if (!prompt || prompt.length > 2000) {
    throw new Error(`Question ${index + 1} is missing a valid question prompt.`);
  }

  let rawOptions = Array.isArray(source.options) ? source.options
    : Array.isArray(source.choices) ? source.choices
      : Array.isArray(source.answerOptions) ? source.answerOptions
        : Array.isArray(source.answers) ? source.answers : [];
  if (!rawOptions.length && source.options && typeof source.options === "object") {
    rawOptions = Object.keys(source.options).map(key => source.options[key]);
  }
  rawOptions = rawOptions.map(option => {
    if (option && typeof option === "object") return String(option.text || option.label || option.value || "").trim();
    return String(option == null ? "" : option).trim();
  });
  let type = String(source.type || source.questionType || "").trim().toLowerCase()
    .replace(/[\s_-]+/g, "");
  if (["multiplechoice", "mcq", "objective"].includes(type)) type = "multipleChoice";
  else if (["truefalse", "true/false", "tf", "boolean"].includes(type)) type = "trueFalse";
  else if (["written", "essay", "shortanswer", "subjective", "descriptive"].includes(type)) type = "written";
  else if (!type) {
    const savedAnswer = String(source.correctAnswer || source.answerKey || source.answer || "").trim().toLowerCase();
    type = rawOptions.length ? "multipleChoice"
      : ["true", "false", "yes", "no"].includes(savedAnswer) ? "trueFalse" : "written";
  }
  else throw new Error(`Question ${index + 1} has an unsupported question type.`);

  const marksValue = source.marks !== undefined ? source.marks : source.points;
  const marks = marksValue === undefined || marksValue === "" ? 1 : Number(marksValue);
  if (!Number.isFinite(marks) || marks <= 0 || marks > 100) {
    throw new Error(`Question ${index + 1} has an invalid mark value.`);
  }
  const savedId = String(source.id || source.questionId || source.questionID || source.uid || "").trim();
  const questionId = savedId || `legacy-${String(examId || "exam")}-${index + 1}`;
  const question = {
    id: questionId, type: type, prompt: prompt, marks: marks,
    topic: String(source.topic || "").trim().slice(0, 100),
    difficulty: ["easy", "medium", "hard"].includes(String(source.difficulty || "").toLowerCase())
      ? String(source.difficulty).toLowerCase() : "medium"
  };
  if (type === "multipleChoice") {
    if (rawOptions.length < 2 || rawOptions.length > 8 || rawOptions.some(option => !option || option.length > 300) ||
        new Set(rawOptions.map(option => option.toLowerCase())).size !== rawOptions.length) {
      throw new Error(`Question ${index + 1} needs 2 to 8 unique answer options.`);
    }
    let correctIndexValue = source.correctIndex !== undefined ? source.correctIndex
      : source.answerIndex !== undefined ? source.answerIndex : source.correctOptionIndex;
    let correctIndex = correctIndexValue == null || String(correctIndexValue).trim() === ""
      ? NaN : Number(correctIndexValue);
    if (!Number.isInteger(correctIndex) || correctIndex < 0 || correctIndex >= rawOptions.length) {
      const correctAnswer = String(source.correctAnswer || source.answerKey || source.correctOption || source.answer || "")
        .trim().replace(/^option\s+/i, "").replace(/[.)]$/, "");
      const answerIndex = rawOptions.findIndex((option, optionIndex) =>
        option.toLowerCase() === correctAnswer.toLowerCase() ||
        String.fromCharCode(65 + optionIndex).toLowerCase() === correctAnswer.toLowerCase()
      );
      correctIndex = answerIndex;
    }
    if (!Number.isInteger(correctIndex) || correctIndex < 0 || correctIndex >= rawOptions.length) {
      throw new Error(`Question ${index + 1} is missing a valid correct answer.`);
    }
    question.options = rawOptions;
    question.correctIndex = correctIndex;
  } else if (type === "trueFalse") {
    let answer = String(source.correctAnswer || source.answerKey || source.answer || "").trim().toLowerCase();
    if (answer === "yes") answer = "true";
    if (answer === "no") answer = "false";
    if (!["true", "false"].includes(answer)) {
      throw new Error(`Question ${index + 1} is missing its true-or-false answer.`);
    }
    question.correctAnswer = answer;
  }
  return question;
}

function cbtQuestionJson_(exam) {
  let stored = exam && (exam.questionsJson || exam.questions || exam.questionJson || exam.questionData);
  if (stored === undefined || stored === null || stored === "") {
    throw new Error("Saved exam questions could not be read: no question data is stored for this exam.");
  }
  let questions = stored, parseError = null;
  for (let depth = 0; depth < 3 && typeof questions === "string"; depth++) {
    try {
      questions = JSON.parse(questions);
    } catch (error) {
      parseError = error;
      break;
    }
  }
  for (let depth = 0; depth < 3 && questions && typeof questions === "object" && !Array.isArray(questions); depth++) {
    const wrappers = ["questions", "items", "question", "data", "questionSet", "questionData", "exam"];
    const wrapper = wrappers.find(key => {
      const value = questions[key];
      if (value && typeof value === "object") return true;
      if (typeof value !== "string" || !value.trim()) return false;
      try {
        JSON.parse(value);
        return true;
      } catch (error) {
        return key !== "question";
      }
    });
    if (wrapper) {
      questions = questions[wrapper];
      if (typeof questions === "string") {
        try {
          questions = JSON.parse(questions);
        } catch (error) {
          parseError = error;
          break;
        }
      }
    } else {
      questions = [questions];
      break;
    }
  }
  if (!Array.isArray(questions) || !questions.length) {
    if (parseError) console.error(`Invalid CBT question data for ${exam && exam.id}: ${parseError.message}`);
    throw new Error("Saved exam questions could not be read: the saved question set is not a non-empty list.");
  }
  try {
    const normalized = questions.map((question, index) =>
      cbtNormalizeSavedQuestion_(question, exam && exam.id, index)
    );
    if (new Set(normalized.map(question => question.id)).size !== normalized.length) {
      throw new Error("The saved question set contains duplicate question IDs.");
    }
    return normalized;
  } catch (error) {
    console.error(`Invalid CBT question data for ${exam && exam.id}: ${error.message}`);
    throw new Error(`Saved exam questions could not be read: ${error.message}`);
  }
}

function cbtPublicQuestion_(question) {
  return {
    id: String(question.id),
    type: String(question.type),
    prompt: String(question.prompt),
    marks: Number(question.marks) || 0,
    options: Array.isArray(question.options) ? question.options.slice() : []
  };
}

function cbtScore_(exam, answers, writtenMarks) {
  const questions = cbtQuestionJson_(exam);
  let score = 0, maxScore = 0, writtenCount = 0, markedWritten = 0;
  questions.forEach(question => {
    const marks = Number(question.marks) || 0;
    maxScore += marks;
    if (question.type === "written") {
      if (!String(answers && answers[question.id] || "").trim()) return;
      writtenCount += 1;
      if (Object.prototype.hasOwnProperty.call(writtenMarks || {}, question.id)) {
        const value = Number(writtenMarks[question.id]);
        if (Number.isFinite(value) && value >= 0 && value <= marks) {
          score += value;
          markedWritten += 1;
        }
      }
    } else if (question.type === "multipleChoice" && answers &&
        answers[question.id] !== undefined && answers[question.id] !== null &&
        String(answers[question.id]).trim() !== "" &&
        Number(answers[question.id]) === Number(question.correctIndex)) {
      score += marks;
    } else if (question.type === "trueFalse" &&
        String(answers && answers[question.id] || "").toLowerCase() === String(question.correctAnswer)) {
      score += marks;
    }
  });
  return { score: score, maxScore: maxScore, pendingWritten: markedWritten < writtenCount };
}

function cbtAttemptByToken_(token) {
  const tokenHash = hashAdmissionToken_(token);
  return cbtRows_(CBT_ATTEMPTS_SHEET, CBT_ATTEMPT_HEADERS).find(item =>
    String(item.attemptTokenHash || "") === tokenHash
  );
}

function cbtSavedAnswers_(attempt) {
  return cbtJsonObject_(attempt.answersJson, "Saved exam answers could not be read.");
}

function cbtJsonObject_(value, errorMessage) {
  if (value && typeof value === "object" && !Array.isArray(value)) return value;
  let parsed;
  try {
    parsed = JSON.parse(String(value || "{}"));
  } catch (error) {
    throw new Error(errorMessage);
  }
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new Error(errorMessage);
  }
  return parsed;
}

function cbtWrittenMarks_(attempt) {
  return cbtJsonObject_(attempt.writtenMarksJson, "Saved written marks could not be read.");
}

function cbtFinalizeAttempt_(exam, attempt, expired, writeRecord = true) {
  let answers = cbtSavedAnswers_(attempt);
  const scoring = cbtScore_(exam, answers, cbtWrittenMarks_(attempt));
  attempt.score = scoring.score;
  attempt.maxScore = scoring.maxScore;
  attempt.submittedAt = new Date().toISOString();
  attempt.status = scoring.pendingWritten ? "pending_mark" : "submitted";
  attempt.attemptTokenHash = "";
  attempt.expired = Boolean(expired);
  if (writeRecord) cbtWrite_(CBT_ATTEMPTS_SHEET, CBT_ATTEMPT_HEADERS, attempt);
  return attempt;
}

function handlePublicCbtAction_(body) {
  const tenant = findTenant_(body.schoolCode);
  if (!tenant || !tenant.spreadsheetId || !isSubscriptionUsable_(tenant) ||
      !planSupportsFeature_(tenant.plan, "cbt")) {
    return { ok: false, error: "This CBT exam link is invalid or unavailable. Contact your school administrator." };
  }
  requestSpreadsheetId = String(tenant.spreadsheetId);
  const exam = cbtRows_(CBT_EXAMS_SHEET, CBT_EXAM_HEADERS).find(item =>
    String(item.id) === String(body.examId || "").trim()
  );
  if (body.action === "cbtStart") return startCbtAttempt_(body, exam);
  if (body.action === "cbtResume") return resumeCbtAttempt_(body, exam);
  if (body.action === "cbtResult") return getCbtCandidateResult_(body, exam);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const attempt = cbtAttemptByToken_(body.attemptToken);
    if (!attempt || String(attempt.examId) !== String(body.examId || "")) {
      return { ok: false, error: "This exam session has expired. Enter your access code to continue." };
    }
    if (body.action === "cbtSave") return saveCbtAnswers_(body, exam, attempt);
    if (body.action === "cbtSubmit") return submitCbtAttempt_(body, exam, attempt);
    if (body.action === "cbtIntegrity") return recordCbtIntegrityEvents_(body, exam, attempt);
    return { ok: false, error: "Unknown CBT action." };
  } finally {
    lock.releaseLock();
  }
}

function startCbtAttempt_(body, exam) {
  const studentId = String(body.studentId || "").trim();
  const accessCode = String(body.accessCode || "").trim();
  const examId = String(body.examId || exam && exam.id || "").trim();
  const validAccessCode = /^\d{8}$/.test(accessCode) || accessCode.length >= 12 && accessCode.length <= 80;
  if (!examId || !studentId || studentId.length > 100 ||
      !validAccessCode) {
    return { ok: false, error: "Enter a valid student ID and individual exam access code." };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    exam = cbtRows_(CBT_EXAMS_SHEET, CBT_EXAM_HEADERS).find(item => String(item.id) === examId);
    if (!exam || String(exam.status) !== "published") {
      return { ok: false, error: "This exam is not available." };
    }
    if (String(exam.resultsReleased).toLowerCase() === "true") {
      return { ok: false, error: "This exam is closed to new attempts because results have been released." };
    }
    const now = Date.now(), opensAt = new Date(exam.opensAt || "").getTime(), closesAt = new Date(exam.closesAt || "").getTime();
    if (!Number.isFinite(opensAt) || !Number.isFinite(closesAt) || now < opensAt || now >= closesAt) {
      return { ok: false, error: "This exam is not currently open." };
    }
    const campusId = String(exam.campusId || DEFAULT_CAMPUS_ID);
    const student = readSchoolRecords_("student").find(item =>
      String(item.studentId || "").trim().toLowerCase() === studentId.toLowerCase() &&
      isActiveStudent_(item) && String(item.className || "").trim() === String(exam.className || "").trim() &&
      effectiveCampusId_(item) === campusId
    );
    if (!student) return { ok: false, error: "The student ID or exam access details are invalid." };
    const cache = CacheService.getScriptCache();
    const accessAttemptsKey = `cbt-access-attempts:${cacheKey_(`${body.schoolCode || ""}:${exam.id}:${student.studentId}`)}`;
    const accessAttempts = Number(cache.get(accessAttemptsKey) || 0);
    if (accessAttempts >= 5) {
      return { ok: false, error: "Too many incorrect access-code attempts. Wait 15 minutes before trying again." };
    }
    const code = cbtRows_(CBT_CODES_SHEET, CBT_CODE_HEADERS).find(item =>
      String(item.examId) === String(exam.id) &&
      String(item.studentId).toLowerCase() === String(student.studentId).toLowerCase() &&
      String(item.codeHash) === hashAdmissionToken_(accessCode.toUpperCase()) && !item.revokedAt
    );
    if (!code) {
      cache.put(accessAttemptsKey, String(accessAttempts + 1), 900);
      return { ok: false, error: "The student ID or exam access details are invalid." };
    }
    cache.remove(accessAttemptsKey);
    const existing = cbtRows_(CBT_ATTEMPT_SHEET, CBT_ATTEMPT_HEADERS).find(item =>
      String(item.examId) === String(exam.id) && String(item.studentId).toLowerCase() === String(student.studentId).toLowerCase()
    );
    if (!existing && code.usedAt) {
      return { ok: false, error: "This exam access code has already been used." };
    }
    const questions = cbtQuestionJson_(exam);
    if (!questions.length) {
      return { ok: false, error: "This exam has no available questions. Contact your school administrator; your access code has not been used." };
    }
    if (existing) {
      if (String(existing.status) !== "in_progress") {
        return { ok: false, error: "An attempt has already been submitted for this exam." };
      }
      const attemptExpiresAt = new Date(existing.expiresAt).getTime();
      if (!Number.isFinite(attemptExpiresAt) || Date.now() >= attemptExpiresAt || Date.now() >= closesAt) {
        if (Number.isFinite(attemptExpiresAt) && Date.now() >= attemptExpiresAt || Date.now() >= closesAt) {
          cbtFinalizeAttempt_(exam, existing, true);
        }
        return { ok: false, error: "This exam attempt has ended. Contact your school administrator." };
      }
      const answers = cbtSavedAnswers_(existing), token = newAdmissionToken_();
      existing.attemptTokenHash = hashAdmissionToken_(token);
      cbtWrite_(CBT_ATTEMPT_SHEET, CBT_ATTEMPT_HEADERS, existing);
      return {
        ok: true, resumed: true, attemptToken: token, expiresAt: existing.expiresAt,
        studentName: `${student.firstName || ""} ${student.lastName || ""}`.trim(),
        studentId: String(student.studentId),
        exam: { title: exam.title, className: exam.className, subject: exam.subject },
        questions: questions.map(cbtPublicQuestion_), answers: answers
      };
    }
    const token = newAdmissionToken_(), duration = Math.max(1, Number(exam.durationMinutes) || 1) * 60000;
    const expiresAt = Math.min(now + duration, closesAt);
    const attempt = {
      id: Utilities.getUuid(),
      examId: String(exam.id),
      studentId: String(student.studentId),
      attemptTokenHash: hashAdmissionToken_(token),
      startedAt: new Date(now).toISOString(),
      expiresAt: new Date(expiresAt).toISOString(),
      submittedAt: "",
      answersJson: "{}",
      score: 0,
      maxScore: 0,
      writtenMarksJson: "{}",
      status: "in_progress"
    };
    cbtWrite_(CBT_ATTEMPT_SHEET, CBT_ATTEMPT_HEADERS, attempt);
    code.usedAt = attempt.startedAt;
    cbtWrite_(CBT_CODES_SHEET, CBT_CODE_HEADERS, code);
    return {
      ok: true, attemptToken: token, expiresAt: attempt.expiresAt,
      studentName: `${student.firstName || ""} ${student.lastName || ""}`.trim(),
      studentId: String(student.studentId),
      exam: { title: exam.title, className: exam.className, subject: exam.subject },
      questions: questions.map(cbtPublicQuestion_)
    };
  } finally {
    lock.releaseLock();
  }
}

function resumeCbtAttempt_(body, exam) {
  const token = String(body.attemptToken || "").trim();
  if (!exam || String(exam.status) !== "published" || !token || token.length > 200) {
    return { ok: false, error: "This exam session cannot be resumed. Contact your school administrator." };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const attempt = cbtAttemptByToken_(token);
    if (!attempt || String(attempt.examId) !== String(exam.id) || String(attempt.status) !== "in_progress") {
      return { ok: false, error: "This exam session has expired. Enter your access code or contact your school administrator." };
    }
    if (Date.now() >= new Date(attempt.expiresAt).getTime() ||
        Date.now() >= new Date(exam.closesAt).getTime()) {
      return { ok: false, error: "The exam time has ended. Contact your school administrator." };
    }
    const questions = cbtQuestionJson_(exam);
    if (!questions.length) {
      return { ok: false, error: "This exam has no available questions. Contact your school administrator." };
    }
    return {
      ok: true, studentId: String(attempt.studentId), expiresAt: attempt.expiresAt,
      exam: { title: exam.title, className: exam.className, subject: exam.subject },
      questions: questions.map(cbtPublicQuestion_),
      answers: cbtSavedAnswers_(attempt)
    };
  } finally {
    lock.releaseLock();
  }
}

function saveCbtAnswers_(body, exam, attempt) {
  if (!exam || String(exam.status) !== "published" || String(attempt.status) !== "in_progress") {
    return { ok: false, error: "This exam attempt is no longer active." };
  }
  if (Date.now() >= new Date(attempt.expiresAt).getTime() ||
      Date.now() >= new Date(exam.closesAt).getTime()) {
    cbtFinalizeAttempt_(exam, attempt, true);
    return { ok: true, submitted: true, message: "The exam deadline has passed. Your saved answers were submitted." };
  }
  if (!body.answers || typeof body.answers !== "object" || Array.isArray(body.answers)) {
    return { ok: false, error: "Submitted answers are invalid." };
  }
  const questions = cbtQuestionJson_(exam), allowed = {};
  if (!questions.length) return { ok: false, error: "This exam has no available questions. Contact your school administrator." };
  const questionIds = new Set(questions.map(question => String(question.id)));
  if (Object.keys(body.answers).some(questionId => !questionIds.has(questionId))) {
    return { ok: false, error: "Submitted answers include an unknown question." };
  }
  for (const question of questions) {
    if (!Object.prototype.hasOwnProperty.call(body.answers, question.id)) continue;
    let value = body.answers[question.id];
    if (question.type === "written") {
      if (typeof value !== "string" || value.length > 8000) {
        return { ok: false, error: "Written answers must be text with no more than 8,000 characters." };
      }
    } else if (question.type === "multipleChoice") {
      if (value === null || value === undefined || !/^\d+$/.test(String(value).trim())) {
        return { ok: false, error: "A multiple-choice answer is invalid." };
      }
      const selectedIndex = Number(value);
      if (!Number.isInteger(selectedIndex) || selectedIndex < 0 || selectedIndex >= question.options.length) {
        return { ok: false, error: "A selected option is outside this question's answer choices." };
      }
      value = selectedIndex;
    } else {
      value = String(value || "").toLowerCase();
      if (!["true", "false"].includes(value)) return { ok: false, error: "A true-or-false answer is invalid." };
    }
    allowed[question.id] = value;
  }
  attempt.answersJson = JSON.stringify(allowed);
  cbtWrite_(CBT_ATTEMPTS_SHEET, CBT_ATTEMPT_HEADERS, attempt);
  return { ok: true, savedAt: new Date().toISOString() };
}

function cbtIntegrityEvents_(attempt) {
  if (Array.isArray(attempt.integrityEventsJson)) return attempt.integrityEventsJson;
  try {
    const events = JSON.parse(String(attempt.integrityEventsJson || "[]"));
    if (!Array.isArray(events)) throw new Error("The incident log is not a list.");
    return events;
  } catch (error) {
    throw new Error(`Integrity event data could not be read for attempt ${attempt.id}.`);
  }
}

function recordCbtIntegrityEvents_(body, exam, attempt) {
  if (!exam || String(exam.status) !== "published" || String(attempt.status) !== "in_progress") {
    return { ok: false, error: "This exam attempt is no longer active." };
  }
  const currentTime = Date.now();
  const startTime = new Date(attempt.startedAt).getTime();
  const expiresAt = new Date(attempt.expiresAt).getTime();
  const closesAt = new Date(exam.closesAt).getTime();
  if (![startTime, expiresAt, closesAt].every(Number.isFinite)) {
    return { ok: false, error: "This exam attempt has invalid timing information." };
  }
  if (currentTime >= expiresAt || currentTime >= closesAt) {
    return { ok: false, error: "The exam has ended. Integrity activity can no longer be recorded." };
  }
  const allowedTypes = new Set(["focus_lost", "tab_hidden", "copy", "cut", "paste", "context_menu"]);
  const events = body.events;
  if (!Array.isArray(events) || !events.length || events.length > 20) {
    return { ok: false, error: "Integrity activity must contain between 1 and 20 events." };
  }
  const savedEvents = cbtIntegrityEvents_(attempt);
  const knownIds = new Set(savedEvents.map(event => String(event.id || "")));
  const now = new Date(currentTime);
  const acceptedEventIds = [];
  let droppedEvents = 0, addedEvents = 0;
  events.forEach(event => {
    const id = String(event && event.id || "");
    const type = String(event && event.type || "");
    const occurredAt = new Date(event && event.occurredAt || "").getTime();
    if (!/^[A-Za-z0-9._:-]{1,80}$/.test(id) || !allowedTypes.has(type) ||
        !Number.isFinite(occurredAt) || occurredAt > now.getTime() + 5 * 60 * 1000) {
      droppedEvents += 1;
      return;
    }
    if (knownIds.has(id)) {
      acceptedEventIds.push(id);
      return;
    }
    if (savedEvents.length >= 200) {
      droppedEvents += 1;
      return;
    }
    savedEvents.push({
      id: id,
      type: type,
      occurredAt: new Date(occurredAt).toISOString(),
      receivedAt: now.toISOString()
    });
    knownIds.add(id);
    acceptedEventIds.push(id);
    addedEvents += 1;
  });
  if (addedEvents) {
    attempt.integrityEventsJson = JSON.stringify(savedEvents);
    cbtWrite_(CBT_ATTEMPTS_SHEET, CBT_ATTEMPT_HEADERS, attempt);
  }
  return {
    ok: true,
    acceptedEventIds: acceptedEventIds,
    eventCount: savedEvents.length,
    droppedEvents: droppedEvents
  };
}

function submitCbtAttempt_(body, exam, attempt) {
  if (!exam || String(exam.status) !== "published" || String(attempt.status) !== "in_progress") {
    return { ok: false, error: "This exam attempt is no longer active." };
  }
  if (!cbtQuestionJson_(exam).length) return { ok: false, error: "This exam has no available questions. Contact your school administrator." };
  if (body.answers && typeof body.answers === "object" && !Array.isArray(body.answers)) {
    const saveResult = saveCbtAnswers_(body, exam, attempt);
    if (!saveResult.ok) return saveResult;
    if (saveResult.submitted) return saveResult;
    attempt = cbtRows_(CBT_ATTEMPT_SHEET, CBT_ATTEMPT_HEADERS).find(item => String(item.id) === String(attempt.id));
  }
  cbtFinalizeAttempt_(exam, attempt, Date.now() >= new Date(attempt.expiresAt).getTime());
  return { ok: true, submitted: true, message: "Your exam has been submitted." };
}

function getCbtCandidateResult_(body, exam) {
  if (!exam || String(exam.status) !== "published" ||
      String(exam.resultsReleased).toLowerCase() !== "true") {
    return { ok: false, error: "Results have not been released by your school." };
  }
  const studentId = String(body.studentId || "").trim();
  const codeHash = hashAdmissionToken_(String(body.accessCode || "").trim().toUpperCase());
  const code = cbtRows_(CBT_CODES_SHEET, CBT_CODE_HEADERS).find(item =>
    String(item.examId) === String(exam.id) &&
    String(item.studentId).toLowerCase() === studentId.toLowerCase() &&
    String(item.codeHash) === codeHash && !item.revokedAt
  );
  const attempt = cbtRows_(CBT_ATTEMPT_SHEET, CBT_ATTEMPT_HEADERS).find(item =>
    String(item.examId) === String(exam.id) && String(item.studentId).toLowerCase() === studentId.toLowerCase()
  );
  if (!code || !attempt || String(attempt.status) === "in_progress") {
    return { ok: false, error: "No released result was found for these details." };
  }
  return { ok: true, score: Number(attempt.score) || 0, maxScore: Number(attempt.maxScore) || 0, status: attempt.status };
}

function cbtExamContext_(body, examId) {
  const admin = cbtAdmin_(body);
  const teacher = admin ? null : cbtTeacherContext_(body);
  if (!admin && !teacher) return { error: "Only an administrator or assigned teacher can manage CBT exams." };
  const exam = cbtRows_(CBT_EXAMS_SHEET, CBT_EXAM_HEADERS).find(item =>
    String(item.id) === String(examId) &&
    effectiveCampusId_(item) === String(body.campusId || DEFAULT_CAMPUS_ID)
  );
  if (!exam) return { error: "The exam could not be found in the active campus." };
  if (!admin && !cbtTeacherAssigned_(teacher, exam.className, exam.subject, effectiveCampusId_(exam))) {
    return { error: "This exam is outside your assigned class or subject." };
  }
  return { exam: exam, admin: admin, teacher: teacher };
}

function listCbtExams_(body) {
  const admin = cbtAdmin_(body), teacher = admin ? null : cbtTeacherContext_(body);
  if (!admin && !teacher) return { ok: false, error: "Only administrators and assigned teachers can access CBT exams." };
  const campusId = validatedCampusId_(body.schoolTenant, body.campusId || DEFAULT_CAMPUS_ID);
  if (!campusId) return { ok: false, error: "Select an active campus." };
  const attemptCounts = new Map();
  cbtRows_(CBT_ATTEMPT_SHEET, CBT_ATTEMPT_HEADERS).forEach(attempt => {
    const examId = String(attempt.examId);
    attemptCounts.set(examId, (attemptCounts.get(examId) || 0) + 1);
  });
  const exams = cbtRows_(CBT_EXAMS_SHEET, CBT_EXAM_HEADERS)
    .filter(exam => effectiveCampusId_(exam) === campusId &&
      (admin || String(exam.createdByEmail || "").toLowerCase() === teacher.email ||
        cbtTeacherAssigned_(teacher, exam.className, exam.subject, campusId)))
    .map(exam => {
      let questionCount = 0, questionDataError = "";
      try {
        questionCount = cbtQuestionJson_(exam).length;
      } catch (error) {
        questionDataError = error.message || "Saved exam questions could not be read.";
      }
      return {
        id: exam.id, title: exam.title, className: exam.className, subject: exam.subject,
        term: exam.term, academicYear: exam.academicYear, opensAt: exam.opensAt, closesAt: exam.closesAt,
        durationMinutes: Number(exam.durationMinutes), status: exam.status, reviewNotes: exam.reviewNotes || "",
        resultsReleased: String(exam.resultsReleased).toLowerCase() === "true",
        questionCount: questionCount, questionDataError: questionDataError,
        attemptCount: attemptCounts.get(String(exam.id)) || 0, createdByEmail: exam.createdByEmail
      };
    });
  return { ok: true, exams: exams, assignments: teacher ? teacher.assignments : [] };
}

function getCbtExamForAuthor_(body) {
  const context = cbtExamContext_(body, body.examId);
  if (context.error) return { ok: false, error: context.error };
  let questions = [], questionDataError = "";
  try {
    questions = cbtQuestionJson_(context.exam);
  } catch (error) {
    questionDataError = error.message || "Saved exam questions could not be read.";
  }
  return { ok: true, exam: {
    id: context.exam.id, title: context.exam.title, className: context.exam.className,
    subject: context.exam.subject, term: context.exam.term, academicYear: context.exam.academicYear,
    opensAt: context.exam.opensAt, closesAt: context.exam.closesAt,
    durationMinutes: Number(context.exam.durationMinutes), status: context.exam.status,
    questions: questions, questionDataError: questionDataError
  } };
}

function cbtQuestionBankContext_(body, className, subject) {
  const admin = cbtAdmin_(body), teacher = admin ? null : cbtTeacherContext_(body);
  if (!admin && !teacher) return { error: "Only administrators and assigned teachers can use the question bank." };
  if (!String(className || "").trim() || String(className).trim().length > 100 ||
      !String(subject || "").trim() || String(subject).trim().length > 100) {
    return { error: "Enter a valid class and subject for this question bank." };
  }
  const campusId = validatedCampusId_(body.schoolTenant, body.campusId || DEFAULT_CAMPUS_ID);
  if (!campusId) return { error: "Select an active campus." };
  if (!admin && !cbtTeacherAssigned_(teacher, className, subject, campusId)) {
    return { error: "This question bank is outside your assigned class or subject." };
  }
  return { campusId: campusId, admin: admin, teacher: teacher };
}

function listCbtQuestionBank_(body) {
  const context = cbtQuestionBankContext_(body, body.className, body.subject);
  if (context.error) return { ok: false, error: context.error };
  const questions = cbtRows_(CBT_QUESTION_BANK_SHEET, CBT_QUESTION_BANK_HEADERS)
    .filter(item => effectiveCampusId_(item) === context.campusId &&
      String(item.className).toLowerCase() === String(body.className || "").toLowerCase() &&
      String(item.subject).toLowerCase() === String(body.subject || "").toLowerCase())
    .map(item => ({ id: item.id, topic: item.topic, difficulty: item.difficulty, question: cbtJsonObject_(item.questionJson, "Saved question bank item could not be read.") }));
  return { ok: true, questions: questions };
}

function saveCbtQuestionBank_(body) {
  const context = cbtQuestionBankContext_(body, body.className, body.subject);
  if (context.error) return { ok: false, error: context.error };
  const parsed = cbtQuestionSet_([body.question]);
  if (parsed.error) return { ok: false, error: parsed.error };
  const question = parsed.questions[0];
  const record = {
    id: Utilities.getUuid(), campusId: context.campusId,
    className: String(body.className).trim(), subject: String(body.subject).trim(),
    topic: question.topic, difficulty: question.difficulty,
    questionJson: JSON.stringify(question),
    createdByEmail: String(body.authUser.email || "").toLowerCase(),
    createdAt: new Date().toISOString()
  };
  cbtWrite_(CBT_QUESTION_BANK_SHEET, CBT_QUESTION_BANK_HEADERS, record);
  return { ok: true, question: { id: record.id, topic: record.topic, difficulty: record.difficulty, question: question } };
}

function cbtExamDetails_(body) {
  const title = String(body.title || "").trim(), className = String(body.className || "").trim();
  const subject = String(body.subject || "").trim(), term = String(body.term || "").trim();
  const academicYear = String(body.academicYear || "").trim();
  const opensAt = new Date(body.opensAt || "").getTime(), closesAt = new Date(body.closesAt || "").getTime();
  const durationMinutes = Number(body.durationMinutes);
  const questions = cbtQuestionSet_(body.questions);
  if (title.length < 2 || title.length > 160 || !className || className.length > 100 ||
      !subject || subject.length > 100 || !term || !academicYear || !Number.isFinite(opensAt) ||
      !Number.isFinite(closesAt) || closesAt <= opensAt || !Number.isInteger(durationMinutes) ||
      durationMinutes < 1 || durationMinutes > 600 || questions.error) {
    return { error: questions.error || "Enter valid exam details, schedule, duration, and questions." };
  }
  return {
    title: title, className: className, subject: subject, term: term, academicYear: academicYear,
    opensAt: new Date(opensAt).toISOString(), closesAt: new Date(closesAt).toISOString(),
    durationMinutes: durationMinutes, questions: questions.questions
  };
}

function createCbtExam_(body) {
  const admin = cbtAdmin_(body), teacher = admin ? null : cbtTeacherContext_(body);
  if (!admin && !teacher) return { ok: false, error: "Only an administrator or assigned teacher can create exams." };
  const campusId = validatedCampusId_(body.schoolTenant, body.campusId || DEFAULT_CAMPUS_ID);
  const details = cbtExamDetails_(body);
  if (!campusId || details.error) return { ok: false, error: details.error || "Enter valid exam details, schedule, duration, and questions." };
  const { title, className, subject, term, academicYear, opensAt, closesAt, durationMinutes, questions } = details;
  if (!admin && !cbtTeacherAssigned_(teacher, className, subject, campusId)) {
    return { ok: false, error: "Teachers can only submit exams for their assigned classes and subjects." };
  }
  const record = {
    id: Utilities.getUuid(), campusId: campusId, title: title, className: className,
    subject: subject, term: term, academicYear: academicYear,
    opensAt: opensAt, closesAt: closesAt,
    durationMinutes: durationMinutes, status: admin ? "published" : "pending_review",
    questionsJson: JSON.stringify(questions),
    createdByEmail: String(body.authUser.email || "").trim().toLowerCase(),
    createdByName: String(body.authUser.name || ""),
    reviewNotes: "", resultsReleased: false,
    createdAt: new Date().toISOString(), updatedAt: new Date().toISOString()
  };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    cbtWrite_(CBT_EXAMS_SHEET, CBT_EXAM_HEADERS, record);
  } finally {
    lock.releaseLock();
  }
  logAudit_(body.authUser, admin ? "created CBT exam" : "submitted CBT exam for review", "cbtExam", record.id, null, {
    title: title, className: className, subject: subject
  });
  return { ok: true, examId: record.id, status: record.status };
}

function updateCbtExam_(body) {
  const admin = cbtAdmin_(body), teacher = admin ? null : cbtTeacherContext_(body);
  if (!admin && !teacher) return { ok: false, error: "Only an administrator or assigned teacher can edit exams." };
  const examId = String(body.examId || "").trim();
  if (!examId || examId.length > 100) return { ok: false, error: "Select a valid exam to edit." };
  const campusId = validatedCampusId_(body.schoolTenant, body.campusId || DEFAULT_CAMPUS_ID);
  const details = cbtExamDetails_(body);
  if (!campusId || details.error) return { ok: false, error: details.error || "Enter valid exam details, schedule, duration, and questions." };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let exam, revokedCodeCount = 0;
  try {
    const context = cbtExamContext_(body, examId);
    if (context.error) return { ok: false, error: context.error };
    if (!admin && String(context.exam.createdByEmail || "").trim().toLowerCase() !== teacher.email) {
      return { ok: false, error: "Teachers can only edit exams they submitted." };
    }
    if (String(context.exam.resultsReleased).toLowerCase() === "true") {
      return { ok: false, error: "An exam cannot be edited after results have been released." };
    }
    const attempts = cbtRows_(CBT_ATTEMPT_SHEET, CBT_ATTEMPT_HEADERS)
      .filter(item => String(item.examId) === String(context.exam.id));
    if (attempts.length) {
      return { ok: false, error: "This exam already has candidate attempts and cannot be edited, to preserve saved answers and grading records." };
    }

    exam = Object.assign({}, context.exam, details, {
      status: admin && String(context.exam.status) === "published" ? "published" : "pending_review",
      reviewNotes: !admin || String(context.exam.status) === "rejected" ? "" : context.exam.reviewNotes || "",
      questionsJson: JSON.stringify(details.questions),
      updatedAt: new Date().toISOString()
    });
    const codes = cbtRows_(CBT_CODES_SHEET, CBT_CODE_HEADERS)
      .filter(item => String(item.examId) === String(exam.id) && !item.revokedAt);
    if (codes.length) {
      const revokedAt = new Date().toISOString();
      codes.forEach(item => { item.revokedAt = revokedAt; });
      cbtWriteMany_(CBT_CODES_SHEET, CBT_CODE_HEADERS, codes);
      revokedCodeCount = codes.length;
    }
    cbtWrite_(CBT_EXAMS_SHEET, CBT_EXAM_HEADERS, exam);
  } finally {
    lock.releaseLock();
  }
  logAudit_(body.authUser, "updated CBT exam", "cbtExam", exam.id, null, {
    title: exam.title, className: exam.className, subject: exam.subject,
    status: exam.status, revokedCodeCount: revokedCodeCount
  });
  return { ok: true, examId: exam.id, status: exam.status, revokedCodeCount: revokedCodeCount };
}

function deleteCbtExam_(body) {
  const examId = String(body.examId || "").trim();
  if (!examId || examId.length > 100) return { ok: false, error: "Select a valid exam to delete." };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let exam, revokedCodeCount = 0;
  try {
    const context = cbtExamContext_(body, examId);
    if (context.error) return { ok: false, error: context.error };
    if (!context.admin && String(context.exam.createdByEmail || "").trim().toLowerCase() !== context.teacher.email) {
      return { ok: false, error: "Teachers can only delete exams they created." };
    }
    const attempts = cbtRows_(CBT_ATTEMPT_SHEET, CBT_ATTEMPT_HEADERS)
      .filter(item => String(item.examId) === String(context.exam.id));
    if (attempts.length) {
      return { ok: false, error: "This exam has candidate attempts and cannot be deleted. Exams with attempts must be retained for marking and records." };
    }
    exam = context.exam;
    const sheet = ensureHeaders_(CBT_EXAMS_SHEET, CBT_EXAM_HEADERS);
    const values = sheet.getDataRange().getValues();
    const idIndex = values[0].map(value => String(value).trim()).indexOf("id");
    const rowIndex = values.findIndex((row, index) => index > 0 && String(row[idIndex] || "") === String(exam.id));
    if (rowIndex < 1) return { ok: false, error: "The exam is no longer available. Refresh and try again." };
    const codes = cbtRows_(CBT_CODES_SHEET, CBT_CODE_HEADERS)
      .filter(item => String(item.examId) === String(exam.id) && !item.revokedAt);
    if (codes.length) {
      const revokedAt = new Date().toISOString();
      codes.forEach(item => { item.revokedAt = revokedAt; });
      cbtWriteMany_(CBT_CODES_SHEET, CBT_CODE_HEADERS, codes);
      revokedCodeCount = codes.length;
    }
    sheet.deleteRow(rowIndex + 1);
  } finally {
    lock.releaseLock();
  }
  logAudit_(body.authUser, "deleted CBT exam", "cbtExam", exam.id, {
    title: exam.title, className: exam.className, subject: exam.subject,
    status: exam.status, createdByEmail: exam.createdByEmail
  }, null, `Deleted duplicate CBT exam "${exam.title}"`);
  return { ok: true, revokedCodeCount: revokedCodeCount };
}

function reviewCbtExam_(body) {
  if (!cbtAdmin_(body)) return { ok: false, error: "Only an administrator can review CBT exams." };
  const decision = String(body.decision || "");
  if (!["approve", "reject"].includes(decision)) return { ok: false, error: "Choose approve or reject." };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let exam;
  try {
    exam = cbtRows_(CBT_EXAMS_SHEET, CBT_EXAM_HEADERS).find(item =>
      String(item.id) === String(body.examId) && effectiveCampusId_(item) === String(body.campusId || DEFAULT_CAMPUS_ID)
    );
    if (!exam || String(exam.status) !== "pending_review") return { ok: false, error: "This exam is not awaiting review." };
    if (decision === "approve" && !cbtQuestionJson_(exam).length) {
      return { ok: false, error: "Add at least one valid question before publishing this exam." };
    }
    exam.status = decision === "approve" ? "published" : "rejected";
    exam.reviewNotes = String(body.reviewNotes || "").trim().slice(0, 1000);
    exam.updatedAt = new Date().toISOString();
    cbtWrite_(CBT_EXAMS_SHEET, CBT_EXAM_HEADERS, exam);
  } finally {
    lock.releaseLock();
  }
  logAudit_(body.authUser, `${decision}d CBT exam`, "cbtExam", exam.id, null, { reviewNotes: exam.reviewNotes });
  return { ok: true, status: exam.status };
}

function generateCbtAccessCodes_(body) {
  if (!cbtAdmin_(body)) return { ok: false, error: "Only an administrator can issue exam access codes." };
  const codes = [], skipped = [], writes = [];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let exam;
  try {
    const context = cbtExamContext_(body, body.examId);
    if (context.error) return { ok: false, error: context.error };
    exam = context.exam;
    if (String(exam.status) !== "published") return { ok: false, error: "Approve and publish the exam before issuing access codes." };
    if (String(exam.resultsReleased).toLowerCase() === "true") {
      return { ok: false, error: "Access codes cannot be issued after results have been released." };
    }
    if (!cbtQuestionJson_(exam).length) return { ok: false, error: "Add at least one valid question before issuing access codes." };
    const students = readSchoolRecords_("student").filter(item =>
      isActiveStudent_(item) && String(item.className || "").trim() === String(exam.className).trim() &&
      effectiveCampusId_(item) === effectiveCampusId_(exam)
    );
    if (!students.length) return { ok: false, error: "No active students are assigned to this exam class." };
    const existingCodes = cbtRows_(CBT_CODES_SHEET, CBT_CODE_HEADERS);
    const attempts = cbtRows_(CBT_ATTEMPT_SHEET, CBT_ATTEMPT_HEADERS);
    students.forEach(student => {
      const hasAttempt = attempts.some(item => String(item.examId) === String(exam.id) &&
        String(item.studentId).toLowerCase() === String(student.studentId).toLowerCase());
      if (hasAttempt) {
        skipped.push(String(student.studentId));
        return;
      }
      existingCodes.filter(item => String(item.examId) === String(exam.id) &&
        String(item.studentId).toLowerCase() === String(student.studentId).toLowerCase() && !item.revokedAt
      ).forEach(item => { item.revokedAt = new Date().toISOString(); writes.push(item); });
      const code = newCbtAccessCode_();
      const record = {
        id: Utilities.getUuid(), examId: String(exam.id), studentId: String(student.studentId),
        codeHash: hashAdmissionToken_(code), issuedAt: new Date().toISOString(), usedAt: "", revokedAt: ""
      };
      writes.push(record);
      codes.push({ studentId: String(student.studentId), studentName: `${student.firstName || ""} ${student.lastName || ""}`.trim(), code: code });
    });
    cbtWriteMany_(CBT_CODES_SHEET, CBT_CODE_HEADERS, writes);
  } finally {
    lock.releaseLock();
  }
  logAudit_(body.authUser, "generated CBT access codes", "cbtExam", exam.id, null, {
    issuedCodeCount: codes.length, skippedStudentCount: skipped.length
  });
  return { ok: true, codes: codes, skippedStudentIds: skipped };
}

function listCbtAttempts_(body) {
  const context = cbtExamContext_(body, body.examId);
  if (context.error) return { ok: false, error: context.error };
  const students = readSchoolRecords_("student");
  let attemptRows = cbtRows_(CBT_ATTEMPT_SHEET, CBT_ATTEMPT_HEADERS)
    .filter(item => String(item.examId) === String(context.exam.id));
  if (attemptRows.some(item =>
    String(item.status) === "in_progress" && Date.now() >= new Date(item.expiresAt).getTime()
  )) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      attemptRows = cbtRows_(CBT_ATTEMPT_SHEET, CBT_ATTEMPT_HEADERS)
        .filter(item => String(item.examId) === String(context.exam.id));
      const expiredAttempts = [];
      const now = Date.now();
      attemptRows.forEach(item => {
        if (String(item.status) === "in_progress" && now >= new Date(item.expiresAt).getTime()) {
          expiredAttempts.push(cbtFinalizeAttempt_(context.exam, item, true, false));
        }
      });
      cbtWriteMany_(CBT_ATTEMPT_SHEET, CBT_ATTEMPT_HEADERS, expiredAttempts);
    } finally {
      lock.releaseLock();
    }
  }
  const attempts = attemptRows
    .map(item => {
      const student = students.find(record => String(record.studentId).toLowerCase() === String(item.studentId).toLowerCase());
      return {
        id: item.id, studentId: item.studentId,
        studentName: student ? `${student.firstName || ""} ${student.lastName || ""}`.trim() : "Student",
        startedAt: item.startedAt, submittedAt: item.submittedAt, expiresAt: item.expiresAt,
        status: item.status, answers: cbtSavedAnswers_(item),
        score: Number(item.score) || 0, maxScore: Number(item.maxScore) || 0,
        writtenMarks: cbtWrittenMarks_(item),
        integrityEvents: cbtIntegrityEvents_(item)
      };
    });
  return { ok: true, exam: { id: context.exam.id, title: context.exam.title, questions: cbtQuestionJson_(context.exam).map(cbtPublicQuestion_) }, attempts: attempts };
}

function markCbtWritten_(body) {
  const context = cbtExamContext_(body, body.examId);
  if (context.error) return { ok: false, error: context.error };
  const marks = body.marks;
  if (!marks || typeof marks !== "object" || Array.isArray(marks)) return { ok: false, error: "Written marks are invalid." };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const attempt = cbtRows_(CBT_ATTEMPT_SHEET, CBT_ATTEMPT_HEADERS).find(item =>
      String(item.id) === String(body.attemptId) && String(item.examId) === String(context.exam.id)
    );
    if (!attempt || !["pending_mark", "submitted"].includes(String(attempt.status))) {
      return { ok: false, error: "This submitted attempt is not available for marking." };
    }
    const questions = cbtQuestionJson_(context.exam).filter(question => question.type === "written");
    const submittedAnswers = cbtSavedAnswers_(attempt), savedMarks = cbtWrittenMarks_(attempt);
    for (const question of questions) {
      if (!Object.prototype.hasOwnProperty.call(marks, question.id)) continue;
      if (!String(submittedAnswers[question.id] || "").trim()) return { ok: false, error: "A written response cannot be marked while blank." };
      const score = Number(marks[question.id]);
      if (!Number.isFinite(score) || score < 0 || score > Number(question.marks)) {
        return { ok: false, error: `Marks for a written response must be between 0 and ${question.marks}.` };
      }
      savedMarks[question.id] = score;
    }
    attempt.writtenMarksJson = JSON.stringify(savedMarks);
    const scoring = cbtScore_(context.exam, submittedAnswers, savedMarks);
    attempt.score = scoring.score;
    attempt.maxScore = scoring.maxScore;
    attempt.status = scoring.pendingWritten ? "pending_mark" : "marked";
    cbtWrite_(CBT_ATTEMPT_SHEET, CBT_ATTEMPT_HEADERS, attempt);
    return { ok: true, status: attempt.status, score: attempt.score, maxScore: attempt.maxScore };
  } finally {
    lock.releaseLock();
  }
}

function releaseCbtResults_(body) {
  if (!cbtAdmin_(body)) return { ok: false, error: "Only an administrator can release exam results." };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let attemptCount = 0, examId = "";
  try {
    const exam = cbtRows_(CBT_EXAMS_SHEET, CBT_EXAM_HEADERS).find(item =>
      String(item.id) === String(body.examId) && effectiveCampusId_(item) === String(body.campusId || DEFAULT_CAMPUS_ID)
    );
    if (!exam || String(exam.status) !== "published") return { ok: false, error: "Only a published exam can have results released." };
    if (String(exam.resultsReleased).toLowerCase() === "true") {
      return { ok: false, error: "Results have already been released for this exam." };
    }
    if (!cbtQuestionJson_(exam).length) return { ok: false, error: "Add at least one valid question before releasing results." };
    const attempts = cbtRows_(CBT_ATTEMPT_SHEET, CBT_ATTEMPT_HEADERS).filter(item => String(item.examId) === String(exam.id));
    if (!attempts.length) return { ok: false, error: "There are no candidate attempts to release results for." };
    const expiredAttempts = [];
    attempts.forEach(item => {
      if (String(item.status) === "in_progress" && Date.now() >= new Date(item.expiresAt).getTime()) {
        expiredAttempts.push(cbtFinalizeAttempt_(exam, item, true, false));
      }
    });
    cbtWriteMany_(CBT_ATTEMPT_SHEET, CBT_ATTEMPT_HEADERS, expiredAttempts);
    if (attempts.some(item => ["pending_mark", "in_progress"].includes(String(item.status)))) {
      return { ok: false, error: "Submit or mark all attempts before releasing results." };
    }
    const now = new Date().toISOString();
    const unusedCodes = cbtRows_(CBT_CODES_SHEET, CBT_CODE_HEADERS)
      .filter(item => String(item.examId) === String(exam.id) && !item.revokedAt && !item.usedAt);
    unusedCodes.forEach(item => { item.revokedAt = now; });
    cbtWriteMany_(CBT_CODES_SHEET, CBT_CODE_HEADERS, unusedCodes);
    exam.resultsReleased = true;
    exam.updatedAt = new Date().toISOString();
    cbtWrite_(CBT_EXAMS_SHEET, CBT_EXAM_HEADERS, exam);
    attemptCount = attempts.length;
    examId = exam.id;
  } finally {
    lock.releaseLock();
  }
  logAudit_(body.authUser, "released CBT results", "cbtExam", examId, null, { attemptCount: attemptCount });
  return { ok: true };
}

function getIdCardRoster_(body) {
  if (!cbtAdmin_(body)) return { ok: false, error: "Only an administrator can manage school ID cards." };
  const campusId = validatedCampusId_(body.schoolTenant, body.campusId || DEFAULT_CAMPUS_ID);
  const type = String(body.type || ""), className = String(body.className || "").trim();
  if (!campusId || !["student", "employee"].includes(type)) return { ok: false, error: "Choose a valid card type and campus." };
  const cards = cbtRows_(ID_CARDS_SHEET, ID_CARD_HEADERS).filter(item =>
    effectiveCampusId_(item) === campusId && String(item.recordType) === type && String(item.status) === "active"
  );
  const records = readSchoolRecords_(type).filter(item =>
    effectiveCampusId_(item) === campusId &&
    (type !== "student" || isActiveStudent_(item) && (!className || String(item.className || "") === className)) &&
    (type !== "employee" || !["inactive", "archived"].includes(String(item.status || item.dashboardStatus || "").toLowerCase()))
  ).map(item => {
    const card = cards.find(value => String(value.recordId) === String(item.id));
    return {
      id: String(item.id),
      name: type === "student" ? `${item.firstName || ""} ${item.lastName || ""}`.trim() : String(item.name || ""),
      identifier: type === "student" ? String(item.studentId || "") : String(item.employeeId || item.email || ""),
      className: String(item.className || ""),
      role: String(item.role || ""),
      department: String(item.department || ""),
      photo: idCardPhoto_(item),
      card: card ? { cardNumber: card.cardNumber, validFrom: card.validFrom, validUntil: card.validUntil, issuedAt: card.issuedAt, revision: Number(card.revision) || 1 } : null
    };
  });
  return { ok: true, records: records };
}

function idCardPhoto_(record) {
  const photo = String(record && (record.profilePicture || record.passport) || "");
  return photo.length <= 45000 && /^data:image\/(?:jpeg|png);base64,[A-Za-z0-9+/=]+$/i.test(photo)
    ? photo
    : "";
}

function issueIdCards_(body) {
  if (!cbtAdmin_(body)) return { ok: false, error: "Only an administrator can issue or reissue ID cards." };
  const campusId = validatedCampusId_(body.schoolTenant, body.campusId || DEFAULT_CAMPUS_ID);
  const type = String(body.type || "");
  const ids = Array.isArray(body.recordIds) ? [...new Set(body.recordIds.map(String))] : [];
  const validFrom = String(body.validFrom || ""), validUntil = String(body.validUntil || "");
  const validFromDate = new Date(`${validFrom}T00:00:00Z`);
  const validUntilDate = new Date(`${validUntil}T00:00:00Z`);
  if (!campusId || !["student", "employee"].includes(type) || !ids.length || ids.length > 500 ||
      !/^\d{4}-\d{2}-\d{2}$/.test(validFrom) || !/^\d{4}-\d{2}-\d{2}$/.test(validUntil) ||
      !Number.isFinite(validFromDate.getTime()) || !Number.isFinite(validUntilDate.getTime()) ||
      validFromDate.toISOString().slice(0, 10) !== validFrom ||
      validUntilDate.toISOString().slice(0, 10) !== validUntil || validUntil < validFrom) {
    return { ok: false, error: "Select records and valid card dates. The end date must be on or after the start date." };
  }
  let records = [], cards = [];
  let issued = 0;
  const missingPhotoNames = [];
  const writes = [];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    records = readSchoolRecords_(type);
    cards = cbtRows_(ID_CARDS_SHEET, ID_CARD_HEADERS);
    ids.forEach(id => {
      const record = records.find(item => String(item.id) === id && effectiveCampusId_(item) === campusId &&
        (type !== "student" ? !["inactive", "archived"].includes(String(item.status || item.dashboardStatus || "").toLowerCase()) : isActiveStudent_(item)));
      if (!record) return;
      if (!idCardPhoto_(record)) {
        missingPhotoNames.push(type === "student"
          ? `${record.firstName || ""} ${record.lastName || ""}`.trim() || String(record.studentId || id)
          : String(record.name || record.email || id));
        return;
      }
      let card = cards.find(item => String(item.recordId) === id && String(item.recordType) === type && String(item.status) === "active");
      const now = new Date().toISOString();
      if (card) {
        card.revision = (Number(card.revision) || 1) + 1;
        card.validFrom = validFrom;
        card.validUntil = validUntil;
        card.issuedAt = now;
      } else {
        const number = cards.reduce((max, item) => Math.max(max, Number(String(item.cardNumber || "").split("-").pop()) || 0), 0) + 1;
        card = {
          id: Utilities.getUuid(), recordType: type, recordId: id, campusId: campusId,
          cardNumber: `CLT-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy")}-${String(number).padStart(6, "0")}`,
          validFrom: validFrom, validUntil: validUntil, issuedAt: now, revision: 1, status: "active"
        };
        cards.push(card);
      }
      writes.push(card);
      issued += 1;
    });
    cbtWriteMany_(ID_CARDS_SHEET, ID_CARD_HEADERS, writes);
  } finally {
    lock.releaseLock();
  }
  if (!issued && missingPhotoNames.length) {
    return { ok: false, error: `Upload a valid JPG or PNG photo before issuing cards for: ${missingPhotoNames.join(", ")}.` };
  }
  if (!issued) return { ok: false, error: "No active records matched the selected IDs and campus." };
  logAudit_(body.authUser, "issued or reissued ID cards", "idCard", campusId, null, {
    recordType: type, issuedCount: issued, validFrom: validFrom,
    validUntil: validUntil, missingPhotoCount: missingPhotoNames.length
  });
  return { ok: true, issued: issued, missingPhotoNames: missingPhotoNames };
}

function sendContactMessage_(body) {
  const name = String(body.name || "").trim();
  const email = String(body.email || "").trim().toLowerCase();
  const subject = String(body.subject || "").trim();
  const message = String(body.message || "").trim();
  if (name.length < 2 || name.length > 120) {
    return { ok: false, error: "Enter your name (2 to 120 characters)." };
  }
  if (email.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { ok: false, error: "Enter a valid email address." };
  }
  if (subject.length < 2 || subject.length > 160) {
    return { ok: false, error: "Enter a subject (2 to 160 characters)." };
  }
  if (message.length < 10 || message.length > 5000) {
    return { ok: false, error: "Your message must be between 10 and 5,000 characters." };
  }

  const properties = PropertiesService.getScriptProperties();
  const supportEmail = supportEmailAddress_();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(supportEmail)) {
    return { ok: false, error: "The contact form email is not configured. Please email support directly." };
  }

  const cache = CacheService.getScriptCache();
  const throttleKey = `contact:${cacheKey_(email)}`;
  const lock = LockService.getScriptLock();
  lock.waitLock(3000);
  try {
    if (cache.get(throttleKey)) {
      return { ok: false, error: "Please wait a minute before sending another message." };
    }
    cache.put(throttleKey, "pending", 60);
  } finally {
    lock.releaseLock();
  }

  const safeSubject = subject.replace(/[\r\n\x00-\x1f\x7f]/g, " ").trim();
  let sheet;
  let headers;
  let rowNumber;
  let sheetLockAcquired = false;
  try {
    lock.waitLock(3000);
    sheetLockAcquired = true;
    const registryId = String(properties.getProperty("MASTER_SPREADSHEET_ID") || "");
    if (!registryId) throw new Error("Run setup() to configure the Clasten registry.");
    sheet = ensureHeadersIn_(
      openSpreadsheetById_(registryId),
      CONTACT_MESSAGES,
      ["submittedAt", "name", "email", "subject", "message", "emailStatus"]
    );
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    const contactRecord = {
      submittedAt: new Date(),
      name: name,
      email: email,
      subject: safeSubject,
      message: message,
      emailStatus: "Pending"
    };
    rowNumber = sheet.getLastRow() + 1;
    const row = headers.map(header => contactRecord[header] === undefined ? "" : contactRecord[header]);
    writeSafeRows_(sheet, rowNumber, [row]);
  } catch (error) {
    cache.remove(throttleKey);
    throw error;
  } finally {
    if (sheetLockAcquired) lock.releaseLock();
  }

  const plainBody = [
    "New message from the Clasten website",
    "",
    `Name: ${name}`,
    `Email: ${email}`,
    `Subject: ${safeSubject}`,
    "",
    message
  ].join("\n");
  const htmlBody = `<div style="font-family:Arial,sans-serif;color:#17233b"><h2>New message from the Clasten website</h2><p><strong>Name:</strong> ${escapeEmailHtml_(name)}</p><p><strong>Email:</strong> ${escapeEmailHtml_(email)}</p><p><strong>Subject:</strong> ${escapeEmailHtml_(safeSubject)}</p><p style="white-space:pre-wrap">${escapeEmailHtml_(message)}</p></div>`;

  try {
    MailApp.sendEmail({
      to: supportEmail,
      replyTo: email,
      name: "Clasten contact form",
      subject: `Website contact: ${safeSubject}`,
      body: plainBody,
      htmlBody: htmlBody
    });
  } catch (error) {
    sheet.getRange(rowNumber, headers.indexOf("emailStatus") + 1).setValue("Failed");
    console.error(`Contact message was saved, but support email failed: ${error.message}`);
    cache.put(throttleKey, "saved", 60);
    return {
      ok: true,
      message: "Your message was saved, but we couldn't email the support team. Please contact support directly if your request is urgent."
    };
  }

  sheet.getRange(rowNumber, headers.indexOf("emailStatus") + 1).setValue("Sent");
  cache.put(throttleKey, "sent", 60);
  return { ok: true, message: "Your message was saved and sent. We'll reply to your email address." };
}

function supportEmailAddress_() {
  const properties = PropertiesService.getScriptProperties();
  return String(
    properties.getProperty("ADMIN_EMAIL") ||
    properties.getProperty("SCHOLARO_BILLING_EMAIL") ||
    "jayomoluwa@gmail.com"
  ).trim().toLowerCase();
}

function submitSupportFeedback_(body) {
  const account = body.authUser || {};
  const tenant = body.schoolTenant || {};
  const categories = {
    issue: "Something is broken",
    howto: "I need help using something",
    idea: "Suggest an improvement",
    other: "Other"
  };
  const category = String(body.category || "");
  const categoryLabel = categories[category];
  const message = String(body.message || "").trim();
  const replyEmail = String(body.replyEmail || "").trim().toLowerCase();
  const userEmail = String(account.email || "").trim().toLowerCase();
  const currentPage = String(body.currentPage || "School workspace")
    .replace(/[\r\n\x00-\x1f\x7f]/g, " ")
    .trim()
    .slice(0, 120) || "School workspace";

  if (!account.email || !tenant.tenantId) {
    return { ok: false, error: "Your school session is incomplete. Sign in again and retry." };
  }
  if (!categoryLabel) {
    return { ok: false, error: "Choose a feedback category." };
  }
  if (message.length < 10 || message.length > 3000) {
    return { ok: false, error: "Your message must be between 10 and 3,000 characters." };
  }
  if (replyEmail && (replyEmail.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(replyEmail))) {
    return { ok: false, error: "Enter a valid reply email address or leave it blank." };
  }

  const supportEmail = supportEmailAddress_();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(supportEmail)) {
    return { ok: false, error: "Support email is not configured. Please contact Clasten support directly." };
  }

  const cache = CacheService.getScriptCache();
  const throttleKey = `support-feedback:${tenant.tenantId}:${cacheKey_(userEmail)}`;
  const lock = LockService.getScriptLock();
  let sheet;
  let headers;
  let rowNumber;
  let ticketId;
  lock.waitLock(3000);
  try {
    if (cache.get(throttleKey)) {
      return { ok: false, error: "Please wait a minute before sending another feedback message." };
    }
    cache.put(throttleKey, "pending", 60);
    const registryId = String(PropertiesService.getScriptProperties().getProperty("MASTER_SPREADSHEET_ID") || "");
    if (!registryId) throw new Error("Run setup() to configure the Clasten registry.");
    sheet = ensureHeadersIn_(openSpreadsheetById_(registryId), SUPPORT_FEEDBACK, [
      "submittedAt", "ticketId", "tenantId", "schoolName", "schoolCode",
      "userName", "userEmail", "userRole", "replyEmail", "category",
      "currentPage", "message", "emailStatus"
    ]);
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    ticketId = `SUP-${Utilities.getUuid().replace(/-/g, "").slice(0, 12).toUpperCase()}`;
    const record = {
      submittedAt: new Date(),
      ticketId: ticketId,
      tenantId: String(tenant.tenantId || ""),
      schoolName: String(tenant.schoolName || ""),
      schoolCode: String(tenant.schoolCode || tenant.tenantId || ""),
      userName: String(account.name || userEmail),
      userEmail: userEmail,
      userRole: String(account.role || account.type || ""),
      replyEmail: replyEmail,
      category: categoryLabel,
      currentPage: currentPage,
      message: message,
      emailStatus: "Pending"
    };
    rowNumber = sheet.getLastRow() + 1;
    writeSafeRows_(sheet, rowNumber, [headers.map(header => record[header] === undefined ? "" : record[header])]);
  } catch (error) {
    cache.remove(throttleKey);
    throw error;
  } finally {
    lock.releaseLock();
  }

  const schoolName = String(tenant.schoolName || tenant.tenantId);
  const plainBody = [
    `New Clasten help request · ${ticketId}`,
    "",
    `School: ${schoolName} (${tenant.tenantId})`,
    `School code: ${tenant.schoolCode || tenant.tenantId}`,
    `User: ${account.name || userEmail} (${userEmail})`,
    `Role: ${account.role || account.type || "unknown"}`,
    `Reply email: ${replyEmail || "Not provided"}`,
    `Category: ${categoryLabel}`,
    `Current page: ${currentPage}`,
    "",
    message
  ].join("\n");
  const htmlBody = `<div style="font-family:Arial,sans-serif;color:#17233b"><h2>New Clasten help request</h2><p><strong>Reference:</strong> ${escapeEmailHtml_(ticketId)}</p><p><strong>School:</strong> ${escapeEmailHtml_(schoolName)} (${escapeEmailHtml_(tenant.tenantId)})</p><p><strong>School code:</strong> ${escapeEmailHtml_(tenant.schoolCode || tenant.tenantId)}</p><p><strong>User:</strong> ${escapeEmailHtml_(account.name || userEmail)} (${escapeEmailHtml_(userEmail)})</p><p><strong>Role:</strong> ${escapeEmailHtml_(account.role || account.type || "unknown")}</p><p><strong>Reply email:</strong> ${escapeEmailHtml_(replyEmail || "Not provided")}</p><p><strong>Category:</strong> ${escapeEmailHtml_(categoryLabel)}</p><p><strong>Current page:</strong> ${escapeEmailHtml_(currentPage)}</p><p style="white-space:pre-wrap">${escapeEmailHtml_(message)}</p></div>`;
  let emailWarning = "";
  let emailSent = false;
  try {
    const emailOptions = {
      to: supportEmail,
      name: "Clasten Help & Feedback",
      subject: `Help request ${ticketId}: ${categoryLabel}`,
      body: plainBody,
      htmlBody: htmlBody
    };
    if (replyEmail) emailOptions.replyTo = replyEmail;
    MailApp.sendEmail(emailOptions);
    emailSent = true;
  } catch (error) {
    emailWarning = "Your feedback was saved, but the support email notification failed. If this is urgent, email Clasten support directly.";
    console.error(`Support ticket ${ticketId} was saved, but notification failed: ${error.message}`);
  }
  try {
    sheet.getRange(rowNumber, headers.indexOf("emailStatus") + 1).setValue(emailSent ? "Sent" : "Failed");
  } catch (statusError) {
    const statusWarning = `Ticket ${ticketId} was saved${emailSent ? " and emailed" : ""}, but its delivery status could not be recorded.`;
    emailWarning = emailWarning ? `${emailWarning} ${statusWarning}` : statusWarning;
    console.error(statusWarning + ` ${statusError.message}`);
  }
  cache.put(throttleKey, "sent", 60);
  return { ok: true, ticketId: ticketId, message: "Your feedback was saved and sent to the support team.", emailWarning: emailWarning };
}

function getSchoolPortal_(schoolCode) {
  const tenant = findTenant_(schoolCode);
  if (!tenant) {
    return { ok: false, error: "This school portal could not be found. Check the link or contact your school." };
  }
  requestSpreadsheetId = tenant.spreadsheetId;
  const campuses = ensureDefaultCampus_();
  return {
    ok: true,
    school: {
      schoolCode: tenant.schoolCode || tenant.tenantId,
      schoolName: tenant.schoolName || "School",
      logo: tenant.portalLogo || "",
      backgroundImage: tenant.portalBackgroundImage || "",
      brandColor: tenant.brandColor || "",
      campuses: campuses
        .filter(campus => planSupportsFeature_(tenant.plan, "multiCampus") || String(campus.id) === DEFAULT_CAMPUS_ID)
        .map(campus => ({ id: String(campus.id), name: String(campus.name || "Campus") }))
    }
  };
}

function submitApplication_(body) {
  const tenant = findTenant_(body.schoolCode);
  if (!tenant) return { ok: false, error: "School code not found. Check the application link or contact the school." };
  requestSpreadsheetId = String(tenant.spreadsheetId);
  const submitted = body.application || {};
  if (!submitted || typeof submitted !== "object" || Array.isArray(submitted)) {
    return { ok: false, error: "Application details are invalid." };
  }
  const allowedFields = [
    "firstName", "middleName", "lastName", "preferredName", "dob", "gender",
    "nationality", "previousSchool", "className", "entryTerm", "guardianName",
    "relationship", "guardianEmail", "phone", "phoneCountry", "alternatePhone", "alternatePhoneCountry", "occupation",
    "address", "emergencyContact", "emergencyPhone", "emergencyPhoneCountry", "medicalConditions",
    "specialNeeds", "documents", "campusId"
  ];
  const application = Object.fromEntries(allowedFields.map(key => [
    key,
    String(submitted[key] == null ? "" : submitted[key]).trim().slice(0, 3000)
  ]));
  const required = ["firstName", "lastName", "dob", "gender", "className", "entryTerm", "guardianName", "relationship", "guardianEmail", "phone", "address"];
  if (required.some(key => !String(application[key] || "").trim())) {
    return { ok: false, error: "Complete all required parent and student details before submitting." };
  }
  const email = String(application.guardianEmail).trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: "Enter a valid parent or guardian email address." };
  application.campusId = validatedCampusId_(tenant, application.campusId);
  if (!application.campusId) return { ok: false, error: "Choose an active campus for this application." };
  if (Object.keys(application).some(key => String(application[key] || "").length > 3000)) {
    return { ok: false, error: "An application field is too long. Shorten the details and try again." };
  }
  const cache = CacheService.getScriptCache();
  const rateKey = `application-submit:${tenant.tenantId}:${cacheKey_(email)}`;
  if (cache.get(rateKey)) return { ok: false, error: "An application was recently submitted from this email. Please wait a minute before trying again." };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let record;
  try {
    const year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy");
    const reference = `APP-${year}-${Utilities.getUuid().replace(/-/g, "").slice(0, 8).toUpperCase()}`;
    record = {
      ...application,
      id: `application-${reference}`,
      reference: reference,
      status: "Submitted",
      reviewNotes: "",
      placementClass: "",
      placementTerm: "",
      placementStartDate: "",
      recordsCreated: false,
      invitationStatus: "Not sent",
      parentEmailStatus: "Pending",
      adminEmailStatus: "Pending",
      submittedAt: new Date().toISOString(),
      activityHistory: JSON.stringify([applicationEvent_("Application submitted", "Parent/guardian", "Application received.")])
    };
    writeSchoolRecord_("application", record);
    cache.put(rateKey, "1", 60);
  } finally {
    lock.releaseLock();
  }

  const emailErrors = [];
  try {
    sendAdmissionsEmail_(
      email,
      `Application received · ${record.reference}`,
      "Application received",
      `<p>Dear ${admissionsHtml_(record.guardianName)},</p><p>We have received the application for <strong>${admissionsHtml_(`${record.firstName} ${record.lastName}`)}</strong> at ${admissionsHtml_(tenant.schoolName)}.</p><p>Your application reference is <strong>${admissionsHtml_(record.reference)}</strong>. The school will contact you after reviewing the application.</p>`,
      null,
      tenant.schoolName
    );
    record.parentEmailStatus = "Sent";
  } catch (error) {
    console.error(`Application confirmation email failed for ${record.reference}: ${error.message}`);
    record.parentEmailStatus = "Failed";
    emailErrors.push("parent confirmation");
  }
  try {
    sendAdmissionsEmail_(
      tenant.adminEmail,
      `New application · ${record.reference}`,
      "A new application needs review",
      `<p>A parent/guardian has submitted an application to ${admissionsHtml_(tenant.schoolName)}.</p><p><strong>Application reference:</strong> ${admissionsHtml_(record.reference)}<br><strong>Student:</strong> ${admissionsHtml_(`${record.firstName} ${record.lastName}`)}<br><strong>Applying for:</strong> ${admissionsHtml_(record.className)}<br><strong>Parent/guardian:</strong> ${admissionsHtml_(record.guardianName)} · ${admissionsHtml_(email)}</p><p>Sign in to the school dashboard and open Applications to review it.</p>`,
      null,
      tenant.schoolName
    );
    record.adminEmailStatus = "Sent";
  } catch (error) {
    console.error(`Administrator application notification failed for ${record.reference}: ${error.message}`);
    record.adminEmailStatus = "Failed";
    emailErrors.push("school notification");
  }
  writeSchoolRecord_("application", record);
  return {
    ok: true,
    reference: record.reference,
    emailWarning: emailErrors.length ? `Application saved, but the ${emailErrors.join(" and ")} email could not be sent. Contact the school with reference ${record.reference}.` : ""
  };
}

function reviewApplication_(body) {
  if (!body.authUser || String(body.authUser.role || "").toLowerCase() !== "admin") {
    return { ok: false, error: "Only a school administrator can review applications." };
  }
  const reference = String(body.reference || "").trim();
  const decision = String(body.decision || "");
  if (!reference || !["Under Review", "Placement Offered", "Rejected"].includes(decision)) {
    return { ok: false, error: "Choose an application and a valid review decision." };
  }
  const applications = readSchoolRecords_("application");
  const application = applications.find(item => String(item.reference || item.id) === reference);
  if (!application) return { ok: false, error: "Application not found. Refresh the application list and try again." };
  const tenant = body.schoolTenant || {};
  if (!isSchoolOwnerAdmin_(body.authUser, tenant) &&
      effectiveCampusId_(application) !== String(body.authUser.campusId || "").trim()) {
    return { ok: false, error: "You are not authorized to review applications for another campus." };
  }
  if (["Placement Accepted", "Placement Declined", "Enrolled"].includes(String(application.status))) {
    return { ok: false, error: "This placement has already been accepted or declined and cannot be changed." };
  }
  application.status = decision;
  application.reviewNotes = String(body.reviewNotes || "").trim().slice(0, 5000);
  application.reviewedAt = new Date().toISOString();
  application.reviewedBy = String(body.authUser.name || body.authUser.email || "School administrator");
  application.placementClass = String(body.placementClass || application.className || "").trim().slice(0, 150);
  application.placementTerm = String(body.placementTerm || application.entryTerm || "").trim().slice(0, 80);
  application.placementStartDate = String(body.placementStartDate || "").trim().slice(0, 30);
  application.activityHistory = appendApplicationEvent_(application, `Application ${decision.toLowerCase()}`, application.reviewedBy, application.reviewNotes || "Application reviewed.");
  if (decision !== "Placement Offered") {
    application.placementTokenHash = "";
    application.placementTokenExpiresAt = "";
  }
  if (decision === "Placement Offered" && (!application.placementClass || !application.placementTerm)) {
    return { ok: false, error: "Enter the offered class and term before approving this application." };
  }
  writeSchoolRecord_("application", application);
  if (decision === "Under Review") return { ok: true, application: safeApplication_(application) };

  try {
    if (decision === "Rejected") {
      sendAdmissionsEmail_(
        application.guardianEmail,
        `Application decision · ${application.reference}`,
        "Application update",
        `<p>Dear ${admissionsHtml_(application.guardianName)},</p><p>Thank you for applying to ${admissionsHtml_(body.schoolTenant && body.schoolTenant.schoolName || "our school")} for <strong>${admissionsHtml_(`${application.firstName} ${application.lastName}`)}</strong>.</p><p>After careful review, we are unable to offer placement at this time. Please contact the school if you would like to discuss the decision.</p><p><strong>Reference:</strong> ${admissionsHtml_(application.reference)}</p>`,
        null,
        body.schoolTenant && body.schoolTenant.schoolName
      );
      application.parentEmailStatus = "Decision sent";
    } else {
      const token = newAdmissionToken_();
      application.placementTokenHash = hashAdmissionToken_(token);
      application.placementTokenExpiresAt = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000).toISOString();
      application.activityHistory = appendApplicationEvent_(application, "Placement offer created", application.reviewedBy, "Secure placement link sent to parent/guardian.");
      const placementUrl = admissionsLink_(body.publicUrl, `placement=${encodeURIComponent(body.schoolTenant && body.schoolTenant.schoolCode || body.schoolTenant && body.schoolTenant.tenantId || "")}&ref=${encodeURIComponent(application.reference)}&token=${encodeURIComponent(token)}`);
      sendAdmissionsEmail_(
        application.guardianEmail,
        `Placement offered · ${application.reference}`,
        "A placement has been offered",
        `<p>Dear ${admissionsHtml_(application.guardianName)},</p><p>${admissionsHtml_(body.schoolTenant && body.schoolTenant.schoolName || "Our school")} is pleased to offer <strong>${admissionsHtml_(`${application.firstName} ${application.lastName}`)}</strong> a placement.</p><p><strong>Class:</strong> ${admissionsHtml_(application.placementClass)}<br><strong>Term:</strong> ${admissionsHtml_(application.placementTerm)}${application.placementStartDate ? `<br><strong>Proposed start date:</strong> ${admissionsHtml_(application.placementStartDate)}` : ""}</p><p>Use the secure link below to review and accept or decline this placement. The link expires in 14 days.</p><p><strong>Application reference:</strong> ${admissionsHtml_(application.reference)}</p>`,
        { label: "Review placement offer", url: placementUrl },
        body.schoolTenant && body.schoolTenant.schoolName
      );
      application.parentEmailStatus = "Placement offer sent";
    }
    application.activityHistory = appendApplicationEvent_(application, decision === "Rejected" ? "Decision emailed" : "Placement offer emailed", body.schoolTenant && body.schoolTenant.schoolName || "School", "Parent notification sent.");
  } catch (error) {
    console.error(`Application decision email failed for ${reference}: ${error.message}`);
    application.parentEmailStatus = `Failed: ${error.message}`.slice(0, 500);
    writeSchoolRecord_("application", application);
    return { ok: true, application: safeApplication_(application), emailWarning: `Decision saved, but the parent email could not be sent: ${error.message}` };
  }
  writeSchoolRecord_("application", application);
  return { ok: true, application: safeApplication_(application) };
}

function getPlacementOffer_(body) {
  const tenant = findTenant_(body.schoolCode);
  if (!tenant) return { ok: false, error: "This placement link is invalid or expired. Contact the school." };
  requestSpreadsheetId = String(tenant.spreadsheetId);
  const application = readSchoolRecords_("application").find(item =>
    String(item.reference || "") === String(body.reference || "")
  );
  if (!validApplicationToken_(application, body.token, "placementTokenHash", "placementTokenExpiresAt") ||
      String(application.status) !== "Placement Offered") {
    return { ok: false, error: "This placement link is invalid, expired, or has already been used. Contact the school." };
  }
  return {
    ok: true,
    schoolName: tenant.schoolName || "School",
    application: {
      reference: application.reference,
      studentName: `${application.firstName || ""} ${application.middleName || ""} ${application.lastName || ""}`.trim(),
      className: application.placementClass || application.className,
      term: application.placementTerm || application.entryTerm,
      startDate: application.placementStartDate || "",
      guardianName: application.guardianName || ""
    }
  };
}

function respondPlacement_(body) {
  const tenant = findTenant_(body.schoolCode);
  if (!tenant) return { ok: false, error: "This placement link is invalid or expired. Contact the school." };
  requestSpreadsheetId = String(tenant.spreadsheetId);
  const decision = String(body.decision || "");
  if (!["accept", "decline"].includes(decision)) return { ok: false, error: "Choose accept or decline." };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let application;
  try {
    application = readSchoolRecords_("application").find(item => String(item.reference || "") === String(body.reference || ""));
    if (!validApplicationToken_(application, body.token, "placementTokenHash", "placementTokenExpiresAt") ||
        String(application.status) !== "Placement Offered") {
      return { ok: false, error: "This placement link is invalid, expired, or has already been used. Contact the school." };
    }
    application.status = decision === "accept" ? "Placement Accepted" : "Placement Declined";
    application.placementTokenHash = "";
    application.placementTokenExpiresAt = "";
    application.respondedAt = new Date().toISOString();
    application.activityHistory = appendApplicationEvent_(
      application,
      decision === "accept" ? "Placement accepted" : "Placement declined",
      application.guardianName || "Parent/guardian",
      decision === "accept" ? "Parent/guardian accepted the placement offer." : "Parent/guardian declined the placement offer."
    );
    if (decision === "accept") {
      const limitError = studentLimitError_(tenant.plan, activeStudentCount_(readSchoolRecords_("student")));
      if (limitError) return { ok: false, error: limitError };
      const created = createEnrollmentRecords_(application);
      application.studentRecordId = created.student.id;
      application.parentRecordId = created.parent.id;
      application.recordsCreated = true;
      application.enrolledAt = new Date().toISOString();
      application.activityHistory = appendApplicationEvent_(application, "Student and parent records created", "Clasten", "Linked records were created in the school dashboard.");
      if (created.parentAccountLinked) {
        application.status = "Enrolled";
        application.dashboardStatus = "Active";
        application.invitationStatus = "Activated";
        application.activityHistory = appendApplicationEvent_(application, "Linked to existing parent account", "Clasten", "The existing parent dashboard can access this newly enrolled child.");
      }
    }
    writeSchoolRecord_("application", application);
  } finally {
    lock.releaseLock();
  }

  const notificationText = decision === "accept"
    ? `Placement accepted for ${application.firstName} ${application.lastName} (${application.reference}). Linked student and parent records were created.`
    : `Placement declined for ${application.firstName} ${application.lastName} (${application.reference}).`;
  const warnings = [];
  try {
    sendAdmissionsEmail_(
      tenant.adminEmail,
      decision === "accept" ? `Placement accepted · ${application.reference}` : `Placement declined · ${application.reference}`,
      decision === "accept" ? "Placement accepted" : "Placement declined",
      `<p>${admissionsHtml_(notificationText)}</p><p>Parent/guardian: ${admissionsHtml_(application.guardianName)} · ${admissionsHtml_(application.guardianEmail)}</p>${decision === "accept" ? (application.dashboardStatus === "Active" ? "<p>The existing parent dashboard is linked to the newly enrolled child.</p>" : "<p>The parent and student records are ready in the admin dashboard. Send the parent a dashboard invitation when you are ready.</p>") : ""}`,
      null,
      tenant.schoolName
    );
  } catch (error) {
    console.error(`Administrator placement response notification failed for ${application.reference}: ${error.message}`);
    warnings.push("The school administrator email could not be sent.");
  }
  if (decision === "accept") {
    try {
      sendAdmissionsEmail_(
        tenant.adminEmail,
        `Enrollment records created · ${application.reference}`,
        "Enrollment records created",
        `<p>Student <strong>${admissionsHtml_(`${application.firstName} ${application.lastName}`)}</strong> and parent/guardian <strong>${admissionsHtml_(application.guardianName)}</strong> are now recorded and linked.</p><p>Application reference: ${admissionsHtml_(application.reference)}.${application.dashboardStatus === "Active" ? " The existing parent account is already active." : " The dashboard invitation can be sent from Applications."}</p>`,
        null,
        tenant.schoolName
      );
    } catch (error) {
      console.error(`Enrollment confirmation email failed for ${application.reference}: ${error.message}`);
      warnings.push("The enrollment confirmation email could not be sent.");
    }
  }
  return {
    ok: true,
    status: application.status,
    recordsCreated: Boolean(application.recordsCreated),
    emailWarning: warnings.join(" ")
  };
}

function inviteParent_(body) {
  if (!body.authUser || String(body.authUser.role || "").toLowerCase() !== "admin") {
    return { ok: false, error: "Only a school administrator can send parent invitations." };
  }
  const reference = String(body.reference || "").trim();
  const application = readSchoolRecords_("application").find(item => String(item.reference || "") === reference);
  if (!application || String(application.status) !== "Placement Accepted" || !application.recordsCreated) {
    return { ok: false, error: "Accept the placement before inviting this parent." };
  }
  if (!isSchoolOwnerAdmin_(body.authUser, body.schoolTenant || {}) &&
      effectiveCampusId_(application) !== String(body.authUser.campusId || "").trim()) {
    return { ok: false, error: "You are not authorized to invite parents for another campus." };
  }
  const parent = readSchoolRecords_("parent").find(item => String(item.id) === String(application.parentRecordId));
  if (!parent) return { ok: false, error: "The linked parent record was not found. Refresh the application and try again." };
  const existingUser = findUserByEmail_(parent.email);
  const existingRole = String(existingUser && (existingUser.role || existingUser.type) || "").toLowerCase();
  if (existingUser && existingRole !== "parent") {
    return { ok: false, error: "This email address is already assigned to a non-parent account in this school." };
  }
  if (existingUser && (existingUser.passwordHash || existingUser.password)) {
    parent.dashboardStatus = "Active";
    parent.invitationStatus = "Activated";
    parent.activationTokenHash = "";
    parent.activationTokenExpiresAt = "";
    parent.activatedAt = parent.activatedAt || new Date().toISOString();
    writeSchoolRecord_("parent", parent);
    application.status = "Enrolled";
    application.dashboardStatus = "Active";
    application.invitationStatus = "Activated";
    application.activityHistory = appendApplicationEvent_(application, "Linked to existing parent account", body.authUser.name || body.authUser.email, "The existing parent dashboard can access this newly enrolled child.");
    writeSchoolRecord_("application", application);
    return {
      ok: true,
      invitationStatus: "Activated",
      existingAccount: true,
      application: safeApplication_(application)
    };
  }
  if (String(parent.dashboardStatus || "").toLowerCase() === "active") {
    return { ok: false, error: "This parent already has an active dashboard account." };
  }
  const token = newAdmissionToken_();
  parent.activationTokenHash = hashAdmissionToken_(token);
  parent.activationTokenExpiresAt = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();
  parent.dashboardStatus = "Invitation Sent";
  parent.invitationStatus = "Sent";
  parent.invitedAt = new Date().toISOString();
  writeSchoolRecord_("parent", parent);
  application.invitationStatus = "Sent";
  application.activityHistory = appendApplicationEvent_(application, "Parent dashboard invitation sent", body.authUser.name || body.authUser.email, "Activation link sent to parent/guardian.");
  writeSchoolRecord_("application", application);
  const activationUrl = admissionsLink_(body.publicUrl, `activate=${encodeURIComponent(body.schoolTenant && body.schoolTenant.schoolCode || body.schoolTenant && body.schoolTenant.tenantId || "")}&token=${encodeURIComponent(token)}`);
  try {
    sendAdmissionsEmail_(
      parent.email,
      `Activate your ${body.schoolTenant && body.schoolTenant.schoolName || "school"} parent dashboard`,
      "Your parent dashboard is ready",
      `<p>Dear ${admissionsHtml_(parent.name)},</p><p>Your child's placement has been accepted and your parent/guardian record is ready. Set a password to activate your ${admissionsHtml_(body.schoolTenant && body.schoolTenant.schoolName || "school")} dashboard.</p><p>This secure activation link expires in 7 days and can only be used once.</p><p>Application reference: ${admissionsHtml_(reference)}<br>Student: ${admissionsHtml_(`${application.firstName} ${application.lastName}`)}</p>`,
      { label: "Activate parent account", url: activationUrl },
      body.schoolTenant && body.schoolTenant.schoolName
    );
  } catch (error) {
    parent.invitationStatus = `Failed: ${error.message}`.slice(0, 500);
    parent.dashboardStatus = "Invitation Failed";
    writeSchoolRecord_("parent", parent);
    application.invitationStatus = "Failed";
    application.activityHistory = appendApplicationEvent_(application, "Parent invitation email failed", body.schoolTenant && body.schoolTenant.schoolName || "School", error.message);
    writeSchoolRecord_("application", application);
    return { ok: false, error: `The invitation could not be emailed: ${error.message}` };
  }
  return {
    ok: true,
    invitationStatus: "Sent",
    application: safeApplication_(application)
  };
}

function inviteEmployee_(body) {
  if (!body.authUser || String(body.authUser.role || "").toLowerCase() !== "admin") {
    return { ok: false, error: "Only a school administrator can invite employees." };
  }
  const campusId = validatedCampusId_(body.schoolTenant || {}, body.campusId || body.authUser.campusId || DEFAULT_CAMPUS_ID);
  if (!campusId) return { ok: false, error: "Choose an active campus in this school workspace." };
  const employeeId = String(body.employeeId || "").trim();
  const input = body.record && typeof body.record === "object" ? body.record : {};
  const name = String(input.name || "").trim();
  const email = String(input.email || "").trim().toLowerCase();
  const jobRole = String(input.role || "").trim();
  if (!employeeId && (!name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))) {
    return { ok: false, error: "Enter the employee's name and a valid work email address." };
  }
  if (!employeeId && !jobRole) {
    return { ok: false, error: "Choose an employee role before sending the invitation." };
  }
  if (!employeeId && email.length > 200) {
    return { ok: false, error: "The employee email address is too long." };
  }
  const profilePicture = String(input.profilePicture || input.passport || "");
  if (profilePicture && (!/^data:image\/(?:jpeg|png);base64,/i.test(profilePicture) || profilePicture.length > 45000)) {
    return { ok: false, error: "The employee profile image must be a JPG or PNG image under 45,000 characters." };
  }
  const phoneCountry = String(input.phoneCountry || "").trim().toUpperCase();
  if (phoneCountry && !/^[A-Z]{2}$/.test(phoneCountry)) return { ok: false, error: "Choose a valid phone country code." };
  const permissions = (Array.isArray(input.permissions)
    ? input.permissions
    : String(input.permissions || "").split(/[,|]/))
    .map(permission => String(permission || "").trim())
    .filter((permission, index, values) => permission && values.indexOf(permission) === index);
  if (permissions.some(permission => !["fees", "grades"].includes(permission))) {
    return { ok: false, error: "Choose only supported employee permissions." };
  }
  const subjects = String(input.subjects || "").split(/[,;\n]/).map(subject => subject.trim()).filter(Boolean).join(", ");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let employee;
  let before = null;
  let activationUrl;
  try {
    const employees = readSchoolRecords_("employee");
    if (employeeId) {
      employee = employees.find(record => String(record.id) === employeeId);
      if (!employee) return { ok: false, error: "Employee not found. Refresh the employee list and try again." };
      if (effectiveCampusId_(employee) !== campusId) {
        return { ok: false, error: "You are not authorized to invite employees for another campus." };
      }
      before = { ...employee };
    } else {
      if (employees.some(record => String(record.email || "").trim().toLowerCase() === email)) {
        return { ok: false, error: "An employee record already uses this email address." };
      }
      employee = {
        id: `employee-${Utilities.getUuid()}`,
        name: name,
        email: email,
        role: jobRole,
        department: String(input.department || "").trim(),
        className: String(input.className || "").trim(),
        subjects: subjects,
        permissions: permissions,
        phone: String(input.phone || "").trim(),
        phoneCountry: phoneCountry,
        startDate: String(input.startDate || "").trim(),
        address: String(input.address || "").trim(),
        profilePicture: profilePicture,
        campusId: campusId,
        createdAt: new Date().toISOString(),
        createdBy: String(body.authUser.email || "")
      };
    }
    employee.campusId = campusId;

    const account = findUserByEmail_(employee.email);
    if (account && String(account.role || "").toLowerCase() !== "employee") {
      return { ok: false, error: "This email address is already assigned to another school account." };
    }
    if (account && (account.passwordHash || account.password)) {
      return { ok: false, error: "This employee already has an active account. Ask them to sign in or reset their password." };
    }
    const token = newAdmissionToken_();
    const schoolCode = body.schoolTenant && (body.schoolTenant.schoolCode || body.schoolTenant.tenantId);
    activationUrl = admissionsLink_(
      body.publicUrl,
      `activate=${encodeURIComponent(schoolCode || "")}&type=employee&token=${encodeURIComponent(token)}`
    );
    employee.activationTokenHash = hashAdmissionToken_(token);
    employee.activationTokenExpiresAt = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();
    employee.dashboardStatus = "Invitation Sent";
    employee.invitationStatus = "Sent";
    employee.invitedAt = new Date().toISOString();
    writeSchoolRecord_("employee", employee);
  } finally {
    lock.releaseLock();
  }

  let emailError = null;
  try {
    sendAdmissionsEmail_(
      employee.email,
      `Activate your ${body.schoolTenant && body.schoolTenant.schoolName || "school"} employee account`,
      "Your employee account is ready",
      `<p>Dear ${admissionsHtml_(employee.name)},</p><p>${admissionsHtml_(body.schoolTenant && body.schoolTenant.schoolName || "Your school")} has invited you to its secure employee workspace.</p><p>Use the secure link below to set your own password and activate your account. The link expires in 7 days and can only be used once.</p>`,
      { label: "Activate employee account", url: activationUrl },
      body.schoolTenant && body.schoolTenant.schoolName
    );
  } catch (error) {
    emailError = error;
  }
  if (emailError) {
    const error = emailError;
    console.error(`Employee invitation failed for ${employee.email}: ${error.message}`);
    employee.activationTokenHash = "";
    employee.activationTokenExpiresAt = "";
    employee.dashboardStatus = "Invitation Failed";
    employee.invitationStatus = "Failed";
    writeSchoolRecord_("employee", employee);
    logAudit_(
      body.authUser,
      "employee invitation failed",
      "employee",
      employee.id,
      before,
      employee,
      `Employee record saved, but the activation email failed: ${error.message}`
    );
    return {
      ok: true,
      record: safeEmployee_(employee),
      invitationStatus: "Failed",
      notificationWarning: `Employee record saved, but the invitation email could not be sent: ${error.message}`
    };
  }
  logAudit_(
    body.authUser,
    employeeId ? "resent employee invitation" : "invited employee",
    "employee",
    employee.id,
    before,
    employee
  );
  return { ok: true, record: safeEmployee_(employee), invitationStatus: "Sent" };
}

function inviteCampusAdmin_(body) {
  const tenant = body.schoolTenant || {};
  if (!isSchoolOwnerAdmin_(body.authUser, tenant)) {
    return { ok: false, error: "Only the school owner can invite campus administrators." };
  }
  const name = String(body.name || "").trim();
  const email = String(body.email || "").trim().toLowerCase();
  const campusId = String(body.campusId || "").trim();
  const campus = ensureDefaultCampus_().find(item => String(item.id || "") === campusId);
  if (!campus) return { ok: false, error: "Choose a campus in this school workspace." };
  if (!campusPinState_(tenant.tenantId, campusId)) {
    return { ok: false, error: "Set a campus PIN before inviting an administrator to this campus." };
  }
  if (name.length < 2 || name.length > 100 ||
      !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 200) {
    return { ok: false, error: "Enter the campus administrator's name and a valid email address." };
  }
  if (email === String(tenant.adminEmail || "").trim().toLowerCase()) {
    return { ok: false, error: "The school owner already has access to every campus." };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let account;
  let activationUrl;
  try {
    const sheet = ensureHeaders_(USERS, [
      "id", "name", "email", "password", "passwordHash", "passwordSalt", "role", "type",
      "jobRole", "department", "className", "childStudentId", "profilePicture", "sessionVersion", "campusId",
      "campusId", "campusAdminActivationTokenHash", "campusAdminActivationTokenExpiresAt"
    ]);
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(String);
    const emailColumn = headers.indexOf("email");
    const rowIndex = values.findIndex((row, index) =>
      index > 0 && String(row[emailColumn] || "").trim().toLowerCase() === email
    );
    const existing = rowIndex > 0 ? rowToObject_(headers, values[rowIndex]) : null;
    if (existing && (String(existing.role || "").toLowerCase() !== "admin" ||
        String(existing.campusId || "") !== campusId || existing.passwordHash || existing.password)) {
      return { ok: false, error: "That email already belongs to an active or differently assigned school account." };
    }
    const token = newAdmissionToken_();
    account = {
      ...(existing || {}),
      id: existing && existing.id || email,
      name: name,
      email: email,
      password: "",
      passwordHash: "",
      passwordSalt: "",
      role: "admin",
      type: "admin",
      campusId: campusId,
      campusAdminActivationTokenHash: hashAdmissionToken_(token),
      campusAdminActivationTokenExpiresAt: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString()
    };
    writeRecord_(sheet, rowIndex > 0 ? rowIndex + 1 : sheet.getLastRow() + 1, account, headers);
    const schoolCode = tenant.schoolCode || tenant.tenantId;
    activationUrl = admissionsLink_(
      body.publicUrl,
      `campusAdminActivate=${encodeURIComponent(schoolCode || "")}&token=${encodeURIComponent(token)}`
    );
  } finally {
    lock.releaseLock();
  }
  try {
    sendAdmissionsEmail_(
      email,
      `Activate your ${tenant.schoolName || "school"} campus administrator account`,
      "Your campus administrator account is ready",
      `<p>Dear ${admissionsHtml_(name)},</p><p>${admissionsHtml_(tenant.schoolName || "Your school")} assigned you as an administrator for ${admissionsHtml_(campus.name || "your campus")}.</p><p>Your access is restricted to that campus. Use the secure link below to set your password. The link expires in 7 days and can only be used once.</p>`,
      { label: "Activate campus administrator account", url: activationUrl },
      tenant.schoolName
    );
  } catch (error) {
    account.campusAdminActivationTokenHash = "";
    account.campusAdminActivationTokenExpiresAt = "";
    const sheet = ensureHeaders_(USERS, ["id"]);
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(String);
    const emailColumn = headers.indexOf("email");
    const rowIndex = rows.findIndex((row, index) =>
      index > 0 && String(row[emailColumn] || "").trim().toLowerCase() === email
    );
    if (rowIndex > 0) writeRecord_(sheet, rowIndex + 1, account, headers);
    console.error(`Campus administrator invitation failed for ${email}: ${error.message}`);
    return { ok: false, error: `The campus-admin account was prepared, but its activation email could not be sent: ${error.message}` };
  }
  logAudit_(body.authUser, "invited campus administrator", "campus", campusId, null, {
    name: name,
    email: email,
    campusId: campusId
  });
  return { ok: true, campusId: campusId, email: email };
}

function activateCampusAdmin_(body) {
  const tenant = findTenant_(body.schoolCode);
  if (!tenant) return { ok: false, error: "This activation link is invalid or expired. Contact the school." };
  requestSpreadsheetId = String(tenant.spreadsheetId);
  const password = String(body.password || "");
  if (password.length < 12 || password.length > 256) {
    return { ok: false, error: "Choose a password with at least 12 characters." };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let account;
  try {
    const sheet = ensureHeaders_(USERS, [
      "id", "name", "email", "password", "passwordHash", "passwordSalt", "role", "type",
      "campusId", "campusAdminActivationTokenHash", "campusAdminActivationTokenExpiresAt",
      "sessionVersion"
    ]);
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(String);
    const tokenColumn = headers.indexOf("campusAdminActivationTokenHash");
    const rowIndex = values.findIndex((row, index) =>
      index > 0 && validApplicationToken_({
        campusAdminActivationTokenHash: tokenColumn >= 0 ? row[tokenColumn] : "",
        campusAdminActivationTokenExpiresAt: row[headers.indexOf("campusAdminActivationTokenExpiresAt")]
      }, body.token, "campusAdminActivationTokenHash", "campusAdminActivationTokenExpiresAt")
    );
    if (rowIndex <= 0) {
      return { ok: false, error: "This activation link is invalid, expired, or has already been used. Ask the school owner to send another invitation." };
    }
    account = rowToObject_(headers, values[rowIndex]);
    if (String(account.role || "").toLowerCase() !== "admin" ||
        !account.campusId ||
        String(account.email || "").trim().toLowerCase() === String(tenant.adminEmail || "").trim().toLowerCase()) {
      return { ok: false, error: "This campus administrator invitation is not valid. Contact the school owner." };
    }
    if (!ensureDefaultCampus_().some(campus => String(campus.id || "") === String(account.campusId))) {
      return { ok: false, error: "The assigned campus is no longer available. Contact the school owner." };
    }
    const salt = Utilities.getUuid();
    account.password = "";
    account.passwordHash = hashPassword_(password, salt);
    account.passwordSalt = salt;
    account.campusAdminActivationTokenHash = "";
    account.campusAdminActivationTokenExpiresAt = "";
    account.sessionVersion = Number(account.sessionVersion) || 0;
    writeRecord_(sheet, rowIndex + 1, account, headers);
  } finally {
    lock.releaseLock();
  }
  logAudit_(
    { name: account.name, email: account.email },
    "activated campus administrator account",
    "campus",
    account.campusId,
    null,
    { name: account.name, email: account.email, campusId: account.campusId }
  );
  const campus = ensureDefaultCampus_().find(item => String(item.id || "") === String(account.campusId));
  return { ok: true, email: account.email, campusName: campus && campus.name || "your assigned campus" };
}

function activateParent_(body) {
  const tenant = findTenant_(body.schoolCode);
  if (!tenant) return { ok: false, error: "This activation link is invalid or expired. Contact the school." };
  requestSpreadsheetId = String(tenant.spreadsheetId);
  const password = String(body.password || "");
  if (password.length < 12) return { ok: false, error: "Choose a password with at least 12 characters." };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let parent;
  let application;
  try {
    parent = readSchoolRecords_("parent").find(item =>
      validApplicationToken_(item, body.token, "activationTokenHash", "activationTokenExpiresAt")
    );
    if (!parent || String(parent.dashboardStatus || "") !== "Invitation Sent") {
      return { ok: false, error: "This activation link is invalid, expired, or has already been used. Ask the school to send a new invitation." };
    }
    const email = String(parent.email || "").toLowerCase().trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: "The parent email on this invitation is invalid. Contact the school." };
    const sheet = ensureHeaders_(USERS, [
      "id", "name", "email", "password", "passwordHash", "passwordSalt", "role", "type",
      "jobRole", "department", "className", "childStudentId", "profilePicture", "sessionVersion", "campusId",
      "phone", "phoneCountry", "address", "relationship", "occupation",
      "emailAbsenceNotifications", "emailPaymentNotifications",
      "emailTimetableNotifications", "emailReportCardNotifications"
    ]);
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(String);
    const emailColumn = headers.indexOf("email");
    const existingIndex = rows.findIndex((row, index) => index > 0 && String(row[emailColumn] || "").toLowerCase().trim() === email);
    const existing = existingIndex > 0 ? rowToObject_(headers, rows[existingIndex]) : null;
    if (existing && String(existing.role || "").toLowerCase() !== "parent") {
      return { ok: false, error: "This email address is already used by another account. Contact the school administrator." };
    }
    const salt = Utilities.getUuid();
    writeRecord_(sheet, existing ? existingIndex + 1 : sheet.getLastRow() + 1, {
      id: existing && existing.id || email,
      name: parent.name,
      email: email,
      password: "",
      passwordHash: hashPassword_(password, salt),
      passwordSalt: salt,
      role: "parent",
      type: "parent",
      jobRole: "",
      department: "",
      className: "",
      childStudentId: parent.childStudentId || "",
      profilePicture: existing && existing.profilePicture || ""
    }, headers);
    parent.activationTokenHash = "";
    parent.activationTokenExpiresAt = "";
    parent.dashboardStatus = "Active";
    parent.invitationStatus = "Activated";
    parent.activatedAt = new Date().toISOString();
    writeSchoolRecord_("parent", parent);
    application = readSchoolRecords_("application").find(item => String(item.reference || "") === String(parent.applicationReference || ""));
    if (application) {
      application.status = "Enrolled";
      application.dashboardStatus = "Active";
      application.invitationStatus = "Activated";
      application.activityHistory = appendApplicationEvent_(application, "Parent dashboard activated", parent.name, "Parent set a password and activated their account.");
      writeSchoolRecord_("application", application);
    }
  } finally {
    lock.releaseLock();
  }
  try {
    sendAdmissionsEmail_(
      tenant.adminEmail,
      `Parent dashboard activated · ${application && application.reference || parent.applicationReference || ""}`,
      "Parent account activated",
      `<p>${admissionsHtml_(parent.name)} has activated their parent dashboard.</p><p>Email: ${admissionsHtml_(parent.email)}${application ? `<br>Student: ${admissionsHtml_(`${application.firstName} ${application.lastName}`)}<br>Application reference: ${admissionsHtml_(application.reference)}` : ""}</p>`,
      null,
      tenant.schoolName
    );
  } catch (error) {
    console.error(`Parent activation notification failed for ${parent.email}: ${error.message}`);
    return { ok: true, emailWarning: "Your account is active, but the school could not be notified by email." };
  }
  return { ok: true };
}

function activateEmployee_(body) {
  const tenant = findTenant_(body.schoolCode);
  if (!tenant) return { ok: false, error: "This activation link is invalid or expired. Contact the school." };
  requestSpreadsheetId = String(tenant.spreadsheetId);
  const password = String(body.password || "");
  if (password.length < 12) return { ok: false, error: "Choose a password with at least 12 characters." };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let employee;
  try {
    employee = readSchoolRecords_("employee").find(record =>
      validApplicationToken_(record, body.token, "activationTokenHash", "activationTokenExpiresAt")
    );
    if (!employee || String(employee.dashboardStatus || "") !== "Invitation Sent") {
      return { ok: false, error: "This activation link is invalid, expired, or has already been used. Ask the school to send a new invitation." };
    }
    const email = String(employee.email || "").toLowerCase().trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return { ok: false, error: "The employee email on this invitation is invalid. Contact the school." };
    }
    const sheet = ensureHeaders_(USERS, [
      "id", "name", "email", "password", "passwordHash", "passwordSalt", "role", "type",
      "jobRole", "department", "className", "childStudentId", "profilePicture", "sessionVersion", "campusId",
      "emailAbsenceNotifications", "emailPaymentNotifications",
      "emailTimetableNotifications", "emailReportCardNotifications"
    ]);
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(String);
    const emailColumn = headers.indexOf("email");
    const existingIndex = rows.findIndex((row, index) =>
      index > 0 && String(row[emailColumn] || "").toLowerCase().trim() === email
    );
    const existing = existingIndex > 0 ? rowToObject_(headers, rows[existingIndex]) : null;
    if (existing && String(existing.role || "").toLowerCase() !== "employee") {
      return { ok: false, error: "This email address is already used by another account. Contact the school administrator." };
    }
    if (existing && (existing.passwordHash || existing.password)) {
      return { ok: false, error: "This employee account is already active. Sign in or reset the password instead." };
    }
    const salt = Utilities.getUuid();
    writeRecord_(sheet, existing ? existingIndex + 1 : sheet.getLastRow() + 1, {
      id: existing && existing.id || email,
      name: employee.name,
      email: email,
      password: "",
      passwordHash: hashPassword_(password, salt),
      passwordSalt: salt,
      role: "employee",
      type: "employee",
      jobRole: employee.role || "",
      department: employee.department || "",
      className: employee.className || "",
      campusId: effectiveCampusId_(employee),
      childStudentId: "",
      profilePicture: employee.profilePicture || existing && existing.profilePicture || ""
    }, headers);
    const before = { ...employee };
    employee.activationTokenHash = "";
    employee.activationTokenExpiresAt = "";
    employee.dashboardStatus = "Active";
    employee.invitationStatus = "Activated";
    employee.activatedAt = new Date().toISOString();
    writeSchoolRecord_("employee", employee);
    logAudit_(
      { name: employee.name, email: employee.email },
      "activated employee account",
      "employee",
      employee.id,
      before,
      employee
    );
  } finally {
    lock.releaseLock();
  }
  try {
    sendAdmissionsEmail_(
      tenant.adminEmail,
      `Employee account activated · ${employee.name}`,
      "Employee account activated",
      `<p>${admissionsHtml_(employee.name)} has activated their employee dashboard.</p><p>Email: ${admissionsHtml_(employee.email)}</p>`,
      null,
      tenant.schoolName
    );
  } catch (error) {
    console.error(`Employee activation notification failed for ${employee.email}: ${error.message}`);
    return { ok: true, emailWarning: "Your account is active, but the school could not be notified by email." };
  }
  return { ok: true };
}

function createEnrollmentRecords_(application) {
  const students = readSchoolRecords_("student");
  const studentId = nextStudentId_(students);
  const parentEmail = String(application.guardianEmail || "").toLowerCase();
  const existingAccount = findUserByEmail_(parentEmail);
  const parentAccountLinked = Boolean(
    existingAccount &&
    String(existingAccount.role || existingAccount.type || "").toLowerCase() === "parent" &&
    (existingAccount.passwordHash || existingAccount.password)
  );
  const student = {
    id: `student-${application.reference}`,
    firstName: application.firstName || "",
    middleName: application.middleName || "",
    lastName: application.lastName || "",
    preferredName: application.preferredName || "",
    studentId: studentId,
    dob: application.dob || "",
    gender: application.gender || "",
    nationality: application.nationality || "",
    previousSchool: application.previousSchool || "",
    className: application.placementClass || application.className || "",
    campusId: effectiveCampusId_(application),
    academicYear: application.academicYear || "",
    term: application.placementTerm || application.entryTerm || "",
    guardian: application.guardianName || "",
    guardianName: application.guardianName || "",
    guardianEmail: parentEmail,
    parentEmail: parentEmail,
    phone: application.phone || "",
    phoneCountry: application.phoneCountry || "",
    alternatePhone: application.alternatePhone || "",
    alternatePhoneCountry: application.alternatePhoneCountry || "",
    address: application.address || "",
    emergencyContact: application.emergencyContact || "",
    emergencyPhone: application.emergencyPhone || "",
    emergencyPhoneCountry: application.emergencyPhoneCountry || "",
    medicalConditions: application.medicalConditions || "",
    specialNeeds: application.specialNeeds || "",
    applicationReference: application.reference,
    enrollmentDate: new Date().toISOString().slice(0, 10),
    status: "Active"
  };
  const parent = {
    id: `parent-${application.reference}`,
    name: application.guardianName || "",
    relationship: application.relationship || "Guardian",
    phone: application.phone || "",
    phoneCountry: application.phoneCountry || "",
    alternatePhone: application.alternatePhone || "",
    alternatePhoneCountry: application.alternatePhoneCountry || "",
    email: parentEmail,
    occupation: application.occupation || "",
    address: application.address || "",
    childStudentId: studentId,
    childStudentIds: studentId,
    applicationReference: application.reference,
    dashboardStatus: parentAccountLinked ? "Active" : "Not invited",
    invitationStatus: parentAccountLinked ? "Activated" : "Not sent",
    ...(parentAccountLinked ? { activatedAt: new Date().toISOString() } : {}),
    status: "Active"
  };
  writeSchoolRecord_("student", student);
  writeSchoolRecord_("parent", parent);
  return { student: student, parent: parent, parentAccountLinked: parentAccountLinked };
}

function safeApplication_(application) {
  const safe = { ...application };
  delete safe.placementTokenHash;
  delete safe.placementTokenExpiresAt;
  return safe;
}

function safeEmployee_(employee) {
  const safe = { ...employee };
  delete safe.activationTokenHash;
  delete safe.activationTokenExpiresAt;
  return safe;
}

function validApplicationToken_(record, token, hashField, expiryField) {
  if (!record || !token || !record[hashField] || hashAdmissionToken_(token) !== String(record[hashField])) return false;
  const expiry = new Date(record[expiryField] || "").getTime();
  return Number.isFinite(expiry) && expiry >= Date.now();
}

function newAdmissionToken_() {
  return `${Utilities.getUuid().replace(/-/g, "")}${Utilities.getUuid().replace(/-/g, "")}`;
}

function newCbtAccessCode_() {
  const randomHex = Utilities.getUuid().replace(/-/g, "").slice(0, 8);
  return String((parseInt(randomHex, 16) % 90000000) + 10000000);
}

function hashAdmissionToken_(token) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(token || ""),
    Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, "");
}

function applicationEvent_(event, actor, details) {
  return { at: new Date().toISOString(), event: event, actor: actor, details: String(details || "").slice(0, 1000) };
}

function appendApplicationEvent_(application, event, actor, details) {
  let history = [];
  try {
    history = Array.isArray(application.activityHistory)
      ? application.activityHistory
      : JSON.parse(application.activityHistory || "[]");
  } catch (error) {
    console.error(`Invalid application activity history for ${application.reference}: ${error.message}`);
  }
  history.push(applicationEvent_(event, actor, details));
  return JSON.stringify(history.slice(-100));
}

function admissionsHtml_(value) {
  return String(value == null ? "" : value).replace(/[&<>"']/g, character => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  })[character]);
}

function admissionsLink_(baseUrl, query) {
  const base = String(baseUrl || "").trim().replace(/[?#].*$/, "");
  if (!/^https?:\/\/[^/]+/i.test(base)) throw new Error("The school application website URL is unavailable. Reopen the dashboard from the public school website and try again.");
  return `${base}?${query}`;
}

function sendAdmissionsEmail_(to, subject, heading, introduction, action, schoolName) {
  const recipient = String(to || "").trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(recipient)) throw new Error("A valid recipient email address is not configured.");
  if (action && (typeof action !== "object" || !action.url || !action.label)) {
    throw new Error("The admissions email action must include a label and URL.");
  }
  if (action && !/^https?:\/\/\S+$/i.test(String(action.url))) {
    throw new Error("The admissions email action URL must be an absolute HTTP or HTTPS link.");
  }
  const plainIntroduction = String(introduction || "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p\s*>/gi, "\n\n")
    .replace(/<[^>]*>/g, "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
  const plainAction = action
    ? `\n\n${action.label}:\n${action.url}\n\nIf the link is not clickable, copy and paste it into your browser.`
    : "";
  const brand = String(schoolName || "School").replace(/[\r\n\x00-\x1f\x7f]/g, " ").trim().slice(0, 100) || "School";
  const plainBody = `${subject}\n\n${heading}\n\n${plainIntroduction}${plainAction}\n\nThis is an automated message from ${brand}. If you have questions, contact the school directly.`;
  const actionUrl = action ? admissionsHtml_(action.url) : "";
  const button = action
    ? `<p style="margin:28px 0"><a href="${actionUrl}" style="display:inline-block;background:#315bc7;color:#fff;text-decoration:none;padding:13px 22px;border-radius:8px;font-weight:bold">${admissionsHtml_(action.label)}</a></p><p style="font-size:12px;color:#758096;word-break:break-all">If the button does not work, use this link:<br><a href="${actionUrl}" style="color:#315bc7">${actionUrl}</a></p>`
    : "";
  const htmlBody = `<div style="background:#f3f6fb;padding:30px 12px;font-family:Arial,sans-serif;color:#172033"><div style="max-width:560px;margin:0 auto;background:#fff;border:1px solid #e4e8ef;border-radius:16px;overflow:hidden"><div style="height:7px;background:linear-gradient(90deg,#3762ea,#17b890)"></div><div style="padding:28px"><div style="font-size:11px;font-weight:bold;letter-spacing:2px;color:#3762ea">${admissionsHtml_(brand)} · SCHOOL ADMISSIONS</div><h2 style="margin:22px 0 14px;font-size:23px">${admissionsHtml_(subject)}</h2><p style="margin:0 0 14px;font-size:16px;font-weight:bold">${admissionsHtml_(heading)}</p><div style="font-size:14px;line-height:1.75">${introduction}${button}</div><p style="border-top:1px solid #e4e8ef;padding-top:16px;margin:24px 0 0;color:#758096;font-size:11px">This is an automated message from ${admissionsHtml_(brand)}. If you have questions, contact the school directly.</p></div></div></div>`;
  MailApp.sendEmail({ to: recipient, subject: subject, name: `${brand} Admissions`, body: plainBody, htmlBody: htmlBody });
}

function login_(body) {
  const tenant = findTenant_(body.schoolCode);
  if (!tenant) return { ok: false, error: "School code not found. Check the code or contact your school." };
  requestSpreadsheetId = String(tenant.spreadsheetId);
  const email = String(body.email || "").toLowerCase().trim();
  const password = String(body.password || "");
  ensureHeaders_(USERS, ["sessionVersion"]);
  const users = readRows_(USERS);
  const account = users.filter(row =>
    String(row.email || "").toLowerCase().trim() === email
  ).pop();

  const passwordValid = Boolean(password) && account && account.passwordHash
    ? hashPassword_(password, String(account.passwordSalt || "")) === String(account.passwordHash)
    : Boolean(password) && account && String(account.password) === password;
  if (!passwordValid) {
    return { ok: false, error: "Incorrect email or password." };
  }
  if (!account.passwordHash) {
    if (password.length < 12) return { ok: false, error: "This account needs a password reset before it can be used securely. Contact your school administrator." };
    const salt = Utilities.getUuid();
    account.password = "";
    account.passwordSalt = salt;
    account.passwordHash = hashPassword_(password, salt);
    writeRecord_(ensureHeaders_(USERS, [
      "id", "name", "email", "password", "passwordHash", "passwordSalt", "role", "type",
      "jobRole", "department", "className", "childStudentId", "profilePicture"
    ]), users.findIndex(row => String(row.email || "").toLowerCase().trim() === email) + 2, account);
  }

  account.email = email;
  const role = String(account.role || "").trim().toLowerCase();
  const type = String(account.type || "").trim().toLowerCase();
  const validRoles = ["admin", "employee", "parent"];
  account.role = validRoles.includes(role) ? role : validRoles.includes(type) ? type : role;
  account.type = validRoles.includes(type) ? type : account.role;
  const isOwner = isSchoolOwnerAdmin_(account, tenant);
  const loginEntry = body.loginEntry === "schoolPortal" || body.loginEntry === "campusPortal"
    ? body.loginEntry
    : "landing";
  if (loginEntry === "landing" && !isOwner) {
    return { ok: false, error: "Only the school owner can sign in from the Clasten landing page. Use your school's portal link to sign in." };
  }
  if (account.role === "admin" && !isOwner && loginEntry !== "campusPortal") {
    return { ok: false, error: "Campus administrators must use the dedicated campus administrator sign-in link provided by the school owner." };
  }
  if (loginEntry === "campusPortal" && account.role !== "admin") {
    return { ok: false, error: "The campus administrator sign-in is only available to administrator accounts. Use your school's portal link to sign in." };
  }
  if (account.role === "admin") return startAdminCampusLogin_(account, tenant);
  if (["employee", "parent"].includes(account.role)) {
    const profileRecord = readSchoolRecords_(account.role).find(record =>
      String(record.email || "").trim().toLowerCase() === email
    );
    if (profileRecord) {
      account.phone = String(profileRecord.phone || "");
      account.phoneCountry = String(profileRecord.phoneCountry || "");
      account.address = String(profileRecord.address || "");
      if (account.role === "parent") {
        account.relationship = String(profileRecord.relationship || "");
        account.occupation = String(profileRecord.occupation || "");
      }
    }
  }
  return createLoginSession_(account, tenant, account.campusId || account.activeCampusId || DEFAULT_CAMPUS_ID);
}

function isSchoolOwnerAdmin_(account, tenant) {
  return Boolean(account && tenant &&
    String(account.role || account.type || "").trim().toLowerCase() === "admin" &&
    String(account.email || "").trim().toLowerCase() ===
      String(tenant.adminEmail || "").trim().toLowerCase());
}

function campusPinPropertyKey_(tenantId, campusId) {
  return `CAMPUS_PIN:${cacheKey_(tenantId)}:${cacheKey_(campusId)}`;
}

function campusPinState_(tenantId, campusId) {
  const raw = PropertiesService.getScriptProperties().getProperty(
    campusPinPropertyKey_(tenantId, campusId)
  );
  if (!raw) return null;
  try {
    const stored = JSON.parse(raw);
    return stored && stored.salt && stored.hash ? stored : null;
  } catch (error) {
    console.error(`Could not read campus PIN configuration: ${error.message}`);
    throw new Error("The campus PIN configuration could not be read. Contact the school owner.");
  }
}

function saveCampusPin_(tenantId, campusId, pin) {
  const value = String(pin || "");
  if (value.length < 6 || value.length > 64) {
    throw new Error("Campus PINs must be between 6 and 64 characters.");
  }
  const salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty(
    campusPinPropertyKey_(tenantId, campusId),
    JSON.stringify({ salt: salt, hash: hashPassword_(value, salt) })
  );
}

function startAdminCampusLogin_(account, tenant) {
  const campuses = ensureDefaultCampus_()
    .filter(campus => planSupportsFeature_(tenant.plan, "multiCampus") ||
      String(campus.id || "") === DEFAULT_CAMPUS_ID);
  const isOwner = isSchoolOwnerAdmin_(account, tenant);
  const assignedCampusId = isOwner ? "" : String(account.campusId || "").trim();
  const allowedCampuses = isOwner
    ? campuses
    : campuses.filter(campus => String(campus.id || "") === assignedCampusId);
  if (!allowedCampuses.length) {
    return { ok: false, error: "This administrator is not assigned to an available campus. Contact the school owner." };
  }
  if (!isOwner && !campusPinState_(tenant.tenantId, assignedCampusId)) {
    return { ok: false, error: "The school owner must set a campus PIN before this campus administrator can sign in." };
  }
  const challenge = Utilities.getUuid();
  CacheService.getScriptCache().put(`admin-campus-login:${challenge}`, JSON.stringify({
    email: String(account.email || "").trim().toLowerCase(),
    tenantId: tenant.tenantId,
    authVersion: Number(account.sessionVersion) || 0,
    allowedCampusIds: allowedCampuses.map(campus => String(campus.id || ""))
  }), 300);
  return {
    ok: true,
    requiresCampusSelection: true,
    challenge: challenge,
    isCampusOwner: isOwner,
    campuses: allowedCampuses.map(campus => ({
      id: String(campus.id || ""),
      name: String(campus.name || "Campus"),
      pinRequired: Boolean(campusPinState_(tenant.tenantId, campus.id))
    }))
  };
}

function completeAdminCampusLogin_(body) {
  const challenge = String(body.challenge || "").trim();
  if (!challenge || challenge.length > 100) {
    return { ok: false, error: "Your campus sign-in expired. Enter your account details and try again." };
  }
  const cache = CacheService.getScriptCache();
  const challengeKey = `admin-campus-login:${challenge}`;
  const raw = cache.get(challengeKey);
  if (!raw) return { ok: false, error: "Your campus sign-in expired. Enter your account details and try again." };
  const pending = JSON.parse(raw);
  const tenant = findTenant_(pending.tenantId);
  if (!tenant) {
    cache.remove(challengeKey);
    return { ok: false, error: "School account could not be found. Contact Clasten support." };
  }
  requestSpreadsheetId = String(tenant.spreadsheetId);
  const account = findUserByEmail_(pending.email);
  if (!account || String(account.role || "").trim().toLowerCase() !== "admin" ||
      (Number(pending.authVersion) || 0) !== (Number(account.sessionVersion) || 0)) {
    cache.remove(challengeKey);
    return { ok: false, error: "Your campus sign-in expired. Enter your account details and try again." };
  }
  const campusId = String(body.campusId || "").trim();
  const isOwner = isSchoolOwnerAdmin_(account, tenant);
  const assignedCampusId = isOwner ? "" : String(account.campusId || "").trim();
  const allowedCampusIds = Array.isArray(pending.allowedCampusIds)
    ? pending.allowedCampusIds.map(String)
    : [];
  const allowedCampuses = ensureDefaultCampus_().filter(campus =>
    (planSupportsFeature_(tenant.plan, "multiCampus") ||
      String(campus.id || "") === DEFAULT_CAMPUS_ID) &&
    (isOwner || String(campus.id || "") === assignedCampusId)
  );
  if (!allowedCampusIds.includes(campusId) ||
      !allowedCampuses.some(campus => String(campus.id || "") === campusId)) {
    return { ok: false, error: "You are not authorized to access that campus." };
  }
  if (!isOwner && String(account.campusId || "").trim() !== campusId) {
    cache.remove(challengeKey);
    return { ok: false, error: "You are not authorized to access that campus." };
  }
  const attemptKey = `admin-campus-login-attempts:${cacheKey_(challenge)}`;
  const attempts = Number(cache.get(attemptKey) || 0);
  if (attempts >= 5) {
    cache.remove(challengeKey);
    cache.remove(attemptKey);
    return { ok: false, error: "Too many incorrect campus PIN attempts. Enter your account details and try again." };
  }
  const pin = String(body.campusPin || "");
  const pinState = campusPinState_(tenant.tenantId, campusId);
  if (pinState) {
    if (hashPassword_(pin, pinState.salt) !== pinState.hash) {
      cache.put(attemptKey, String(attempts + 1), 300);
      if (attempts + 1 >= 5) cache.remove(challengeKey);
      return { ok: false, error: "Could not verify access to that campus. Check the selected campus and PIN." };
    }
  } else if (!isOwner || pin) {
    return { ok: false, error: "A campus PIN is required. Ask the school owner to configure one." };
  }
  cache.remove(challengeKey);
  cache.remove(attemptKey);
  const session = createLoginSession_(account, tenant, campusId, !pinState);
  session.activeCampusName = allowedCampuses.find(campus => String(campus.id || "") === campusId)?.name || "Campus";
  return session;
}

function createLoginSession_(account, tenant, activeCampusId, campusPinSetupRequired) {
  delete account.password;
  delete account.passwordHash;
  delete account.passwordSalt;
  delete account.campusAdminActivationTokenHash;
  delete account.campusAdminActivationTokenExpiresAt;
  const isCampusOwner = isSchoolOwnerAdmin_(account, tenant);
  account.isCampusOwner = isCampusOwner;
  const ownerAccount = isCampusOwner ? account : findUserByEmail_(tenant.adminEmail);
  const authToken = Utilities.getUuid();
  CacheService.getScriptCache().put(`session:${authToken}`, JSON.stringify({
    email: String(account.email || "").trim().toLowerCase(),
    tenantId: tenant.tenantId,
    spreadsheetId: tenant.spreadsheetId,
    authVersion: Number(account.sessionVersion) || 0,
    campusId: activeCampusId,
    subscriptionStatus: tenant.subscriptionStatus,
    trialEndsAt: tenant.trialEndsAt,
    paidThrough: tenant.paidThrough
  }), 21600);
  return {
    ok: true, user: account, authToken: authToken,
    schoolName: tenant.schoolName, tenantId: tenant.tenantId,
    subscriptionStatus: tenant.subscriptionStatus,
    accessActive: isSubscriptionUsable_(tenant),
    billingPeriod: tenant.billingPeriod,
    portalLogo: tenant.portalLogo || "",
    schoolProfilePicture: tenant.schoolProfilePicture || ownerAccount && ownerAccount.profilePicture || "",
    brandColor: tenant.brandColor || "",
    brandingSetupCompleted: String(tenant.brandingSetupCompleted || "").toLowerCase() === "true",
    trialEndsAt: tenant.trialEndsAt,
    paidThrough: tenant.paidThrough,
    plan: tenant.plan,
    activeCampusId: activeCampusId || DEFAULT_CAMPUS_ID,
    campusPinSetupRequired: campusPinSetupRequired === true
  };
}

function employeeHasPermission_(employee, permission) {
  if (!employee) return false;
  const permissions = Array.isArray(employee.permissions)
    ? employee.permissions
    : String(employee.permissions || "").split(/[,|]/).map(value => value.trim());
  return permissions.includes(permission);
}

function getRecords_(body) {
  const account = body.authUser;
  if (!account) {
    return { ok: false, error: "User account was not found. Please sign in again." };
  }

  const tenant = body.schoolTenant || {};
  const campuses = ensureDefaultCampus_();
  const role = String(account.role || account.type || "").toLowerCase();
  const isCampusOwner = role === "admin" && isSchoolOwnerAdmin_(account, tenant);
  const assignedAdminCampusId = role === "admin" && !isCampusOwner
    ? String(account.campusId || "").trim()
    : "";
  if (role === "admin" && !isCampusOwner &&
      (!assignedAdminCampusId || !campuses.some(campus => String(campus.id || "") === assignedAdminCampusId))) {
    return { ok: false, error: "This administrator is not assigned to an available campus. Contact the school owner." };
  }
  const migrationKey = `records-migrated:${requestSpreadsheetId}`;
  const properties = PropertiesService.getScriptProperties();
  const needsMigration = properties.getProperty(migrationKey) !== "true";
  const allRecordTypes = Object.keys(RECORD_SHEETS);
  const employeeRecord = role === "employee"
    ? readSchoolRecords_("employee").find(record =>
      String(record.email || "").trim().toLowerCase() === String(account.email || "").trim().toLowerCase()
    )
    : null;
  const isTeacher = role === "employee" && String(employeeRecord && employeeRecord.role || account.jobRole || "").toLowerCase().includes("teacher");
  const canEnterGrades = role === "employee" && employeeHasPermission_(employeeRecord, "grades");
  const canManageFees = role === "employee" && employeeHasPermission_(employeeRecord, "fees");
  const roleVisibleRecordTypes = role === "admin"
    ? allRecordTypes.filter(type => type !== "employeeAttendance")
    : role === "parent"
      ? ["student", "attendance", "reportCard", "fee", "feeInvoice", "feePlan"]
      : role === "employee"
        ? ["employee", "employeeLeave"]
          .concat(isTeacher ? ["student", "attendance", "timetable"] : [])
          .concat(canEnterGrades ? ["student", "academic", "gradeEntry"] : [])
          .concat(canManageFees ? ["student", "fee", "feeInvoice", "feePlan"] : [])
        : [];
  const requestedRecordTypes = Array.isArray(body.recordTypes)
    ? [...new Set(body.recordTypes.map(type => String(type || "")))]
      .filter(type => allRecordTypes.includes(type) && roleVisibleRecordTypes.includes(type))
    : roleVisibleRecordTypes;
  const visibleRecordTypes = needsMigration
    ? allRecordTypes.filter(type => type !== "employeeAttendance")
    : requestedRecordTypes;
  const records = Object.fromEntries(allRecordTypes.map(type => [type, []]));
  visibleRecordTypes.forEach(type => {
    records[type] = readSchoolRecords_(type);
  });
  if (needsMigration) {
    readRows_(RECORDS).forEach(row => {
      if (!row.type || !row.dataJson || !RECORD_SHEETS[row.type] || row.type === "employeeAttendance") return;
      try {
        const record = JSON.parse(row.dataJson);
        if (!records[row.type].some(existing => String(existing.id) === String(record.id))) {
          writeSchoolRecord_(row.type, record);
          records[row.type].push(record);
        }
      } catch (error) {
        console.error(`Skipped invalid ${row.type} record ${row.id}: ${error.message}`);
      }
    });
    properties.setProperty(migrationKey, "true");
  }
  if (role === "admin" && requestedRecordTypes.includes("feePlan")) {
    records.feePlan = schoolFeePlans_(records.feePlan, account);
  } else if (role === "parent" && requestedRecordTypes.includes("feePlan")) {
    records.feePlan = schoolFeePlans_(records.feePlan);
  }
  const needsStudentContext = role === "parent" ||
    role === "admin" && !isCampusOwner && requestedRecordTypes.includes("parent") ||
    requestedRecordTypes.includes("student") ||
    requestedRecordTypes.some(type =>
      ["attendance", "reportCard", "gradeEntry", "followUp", "fee", "feeInvoice"].includes(type)
    );
  const allStudents = visibleRecordTypes.includes("student")
    ? records.student
    : needsStudentContext ? readSchoolRecords_("student") : [];
  records.timetable = (records.timetable || []).filter(record =>
    ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"].includes(String(record.day || "").trim())
  );
  if (role === "admin") {
    const userByEmail = records.employee.length
      ? new Map(readRows_(USERS).map(user => [
        String(user.email || "").toLowerCase().trim(),
        user
      ]))
      : new Map();
    records.application = (records.application || []).map(safeApplication_);
    records.employee = (records.employee || []).map(record => {
      const safe = { ...record };
      const activationExpiresAt = safe.activationTokenExpiresAt;
      delete safe.activationTokenHash;
      delete safe.activationTokenExpiresAt;
      const account = userByEmail.get(String(record.email || "").toLowerCase().trim());
      if (account && (account.passwordHash || account.password)) {
        safe.dashboardStatus = "Active";
      } else if (safe.dashboardStatus === "Invitation Sent" &&
          (!Number.isFinite(new Date(activationExpiresAt || "").getTime()) ||
           new Date(activationExpiresAt || "").getTime() < Date.now())) {
        safe.dashboardStatus = "Invitation Expired";
      } else if (!safe.dashboardStatus) {
        safe.dashboardStatus = "Invitation Needed";
      }
      return safe;
    });
    records.parent = (records.parent || []).map(record => {
      const safe = { ...record };
      delete safe.activationTokenHash;
      delete safe.activationTokenExpiresAt;
      return safe;
    });
    if (!isCampusOwner && requestedRecordTypes.includes("parent")) {
      const studentCampusById = new Map(allStudents.map(student => [
        String(student.studentId || "").trim(),
        effectiveCampusId_(student)
      ]));
      records.parent = records.parent.map(parent => {
        const linkedIds = new Set();
        [parent.childStudentId, parent.childStudentIds].forEach(value =>
          String(value || "").split(/[,;\s]+/).forEach(id => { if (id.trim()) linkedIds.add(id.trim()); })
        );
        if (!linkedIds.size) {
          const parentEmail = String(parent.email || "").trim().toLowerCase();
          allStudents.forEach(student => {
            const linkedEmail = String(student.parentEmail || student.guardianEmail || "").trim().toLowerCase();
            if (parentEmail && linkedEmail === parentEmail && student.studentId) {
              linkedIds.add(String(student.studentId).trim());
            }
          });
        }
        const visibleIds = [...linkedIds].filter(id => studentCampusById.get(id) === assignedAdminCampusId);
        if (!visibleIds.length) return null;
        return {
          ...parent,
          childStudentId: visibleIds.includes(String(parent.childStudentId || "").trim())
            ? String(parent.childStudentId || "").trim()
            : visibleIds[0],
          childStudentIds: visibleIds.join(",")
        };
      }).filter(Boolean);
    }
  } else {
    records.employee = (records.employee || []).map(record => {
      const safe = { ...record };
      delete safe.activationTokenHash;
      delete safe.activationTokenExpiresAt;
      return safe;
    });
  }
  const email = String(account.email || "").toLowerCase();
  const linkedStudentIds = role === "parent"
    ? linkedStudentIdsForParent_(account, allStudents)
    : new Set();
  const paymentRecords = requestedRecordTypes.includes("fee") &&
    (role === "admin" || role === "parent" || role === "employee" && canManageFees)
    ? readRows_(PAYMENTS).map(payment => ({
    id: payment.id,
    studentId: payment.studentId,
    student: payment.studentName,
    amount: payment.amount,
    paidFor: payment.paidFor,
    method: payment.paymentMethod,
    reference: payment.reference,
    date: payment.paymentDate,
    academicYear: payment.academicYear,
    term: payment.term,
    feePlanId: payment.feePlanId || "",
    invoiceId: payment.invoiceId || "",
    source: "payments",
    parentEmail: String(payment.parentEmail || "").toLowerCase()
    })).filter(payment => role === "admin" || role === "employee" && canManageFees ||
      payment.parentEmail === email || linkedStudentIds.has(String(payment.studentId || "").trim()))
    : [];
  if (requestedRecordTypes.includes("fee")) {
    records.fee = [...(records.fee || []), ...paymentRecords];
  }

  if (role === "parent") {
    records.student = (records.student || []).filter(record =>
      linkedStudentIds.has(String(record.studentId || "").trim())
    );
    records.employee = [];
    records.parent = [];
    records.application = [];
    records.academic = [];
    records.attendance = (records.attendance || []).filter(record =>
      linkedStudentIds.has(String(record.studentId || "").trim())
    );
    records.reportCard = (records.reportCard || []).filter(record =>
      linkedStudentIds.has(String(record.studentId || "").trim()) &&
      String(record.parentEmail || "").toLowerCase() === email
    );
    records.fee = records.fee.filter(record =>
      record.parentEmail === email || linkedStudentIds.has(String(record.studentId || "").trim())
    );
    records.feeInvoice = (records.feeInvoice || []).filter(record =>
      linkedStudentIds.has(String(record.studentId || "").trim())
    );
    records.feePlan = (records.feePlan || []).filter(record =>
      record.active !== false && String(record.active || "Active").toLowerCase() !== "inactive"
    );
    records.event = [];
    records.gradeScale = [];
    records.gradeEntry = [];
    records.timetable = String(account.jobRole || "").toLowerCase().includes("teacher")
      ? (records.timetable || []).filter(record =>
        String(record.teacherEmail || "").toLowerCase() === email
      )
      : [];
  } else if (role !== "admin") {
    const email = String(account.email || "").toLowerCase();
    records.employee = (records.employee || []).filter(record =>
      String(record.email || "").toLowerCase() === email
    );
    records.employeeLeave = (records.employeeLeave || []).filter(record =>
      String(record.employeeEmail || "").trim().toLowerCase() === email
    );
    const assignedClass = String(employeeRecord && employeeRecord.className || account.className || "");
    records.timetable = isTeacher
      ? (records.timetable || []).filter(record => String(record.teacherEmail || "").toLowerCase() === email)
      : [];
    records.attendance = isTeacher
      ? (records.attendance || []).filter(record => String(record.className || "") === assignedClass)
      : [];
    records.student = (isTeacher || canEnterGrades || canManageFees)
      ? (records.student || []).filter(record => {
        if (canManageFees) {
          return effectiveCampusId_(record) === effectiveCampusId_(employeeRecord || account);
        }
        return (isTeacher || canEnterGrades) && String(record.className || "") === assignedClass;
      })
      : [];
    const assignedStudentIds = new Set((records.student || []).map(record => String(record.studentId || "")));
    records.parent = [];
    records.application = [];
    records.academic = [];
    if (!canEnterGrades) records.academic = [];
    if (!canManageFees) {
      records.fee = [];
      records.feeInvoice = [];
      records.feePlan = [];
    }
    records.reportCard = [];
    records.event = [];
    records.gradeScale = [];
    records.gradeEntry = canEnterGrades
      ? (records.gradeEntry || []).filter(record =>
        String(record.enteredByEmail || "").toLowerCase() === String(account.email || "").toLowerCase() ||
        assignedStudentIds.has(String(record.studentId || ""))
      )
      : [];
    if (canManageFees) {
      records.fee = (records.fee || []).filter(record =>
        !record.studentId || (records.student || []).some(student => String(student.studentId) === String(record.studentId))
      );
      records.feeInvoice = (records.feeInvoice || []).filter(invoice =>
        (records.student || []).some(student => String(student.studentId) === String(invoice.studentId))
      );
      records.feePlan = (records.feePlan || []).filter(plan =>
        plan.active !== false && String(plan.active || "Active").toLowerCase() !== "inactive"
      );
    }
  }
  const visibleCampuses = role === "admin"
    ? campuses.filter(campus => (isCampusOwner || String(campus.id || "") === assignedAdminCampusId) &&
      (planSupportsFeature_(tenant.plan, "multiCampus") || String(campus.id || "") === DEFAULT_CAMPUS_ID))
    : [];
  if (role !== "admin") records.campus = [];
  else records.campus = (records.campus || []).filter(campus =>
    visibleCampuses.some(visible => String(visible.id || "") === String(campus.id || ""))
  );

  let activeCampusId = DEFAULT_CAMPUS_ID;
  if (role === "admin") {
    activeCampusId = validatedCampusId_(tenant, body.campusId || assignedAdminCampusId || DEFAULT_CAMPUS_ID);
    if (!activeCampusId) return { ok: false, error: "That campus is not available in this school workspace." };
    if (!isCampusOwner && activeCampusId !== assignedAdminCampusId) {
      return { ok: false, error: "You are not authorized to access that campus." };
    }
  } else if (role === "employee") {
    activeCampusId = effectiveCampusId_(employeeRecord || account);
  }
  if (role === "admin" || role === "employee") {
    const campusByStudentId = new Map(allStudents.map(student => [
      String(student.studentId || "").trim(),
      effectiveCampusId_(student)
    ]));
    const scopedTypes = ["student", "employee", "application", "academic", "attendance", "employeeLeave", "timetable", "reportCard", "gradeEntry", "followUp", "fee", "feeInvoice", "event"];
    scopedTypes.forEach(type => {
      records[type] = (records[type] || []).filter(record => {
        const campusByStudent = record.studentId && campusByStudentId.get(String(record.studentId).trim());
        const campusId = ["fee", "gradeEntry"].includes(type) && campusByStudent
          ? campusByStudent
          : effectiveCampusId_(record);
        return campusId === activeCampusId;
      }).map(record => {
        const campusByStudent = record.studentId && campusByStudentId.get(String(record.studentId).trim());
        return {
          ...record,
          campusId: ["fee", "gradeEntry"].includes(type) && campusByStudent
            ? campusByStudent
            : effectiveCampusId_(record)
        };
      });
    });
  }

  const studentLimit = studentLimitForPlan_(tenant.plan);
  const campusAdminNames = new Map();
  let schoolOwnerName = "";
  if (role === "admin" && requestedRecordTypes.includes("campus")) {
    const visibleCampusIds = new Set(visibleCampuses.map(campus => String(campus.id || "")));
    readRows_(USERS).forEach(user => {
      if (String(user.role || user.type || "").toLowerCase() !== "admin") return;
      if (isSchoolOwnerAdmin_(user, tenant)) {
        schoolOwnerName = String(user.name || "").trim();
        return;
      }
      const campusId = String(user.campusId || "").trim();
      const name = String(user.name || "").trim();
      if (!name || !visibleCampusIds.has(campusId)) return;
      if (!campusAdminNames.has(campusId)) campusAdminNames.set(campusId, []);
      const names = campusAdminNames.get(campusId);
      if (!names.includes(name)) names.push(name);
    });
  }
  const campusSummaries = role === "admin"
    ? visibleCampuses.map(campus => {
      const campusId = String(campus.id || "");
      const names = campusAdminNames.get(campusId) || [];
      return {
        ...(requestedRecordTypes.includes("campus")
          ? {
            ...campus,
            activeStudentCount: allStudents.filter(student =>
              effectiveCampusId_(student) === campusId
            ).filter(isActiveStudent_).length
          }
          : campus),
        campusAdminName: names.length
          ? names.join(", ")
          : campusId === DEFAULT_CAMPUS_ID && schoolOwnerName
            ? schoolOwnerName
            : "Not assigned",
        pinConfigured: Boolean(campusPinState_(tenant.tenantId, campus.id))
      };
    })
    : [];
  allRecordTypes.forEach(type => {
    if (!requestedRecordTypes.includes(type)) records[type] = [];
  });
  const studentUsage = role === "admin" &&
    (requestedRecordTypes.includes("student") || requestedRecordTypes.includes("campus")) ? {
    activeStudents: activeStudentCount_(isCampusOwner ? allStudents : allStudents.filter(student =>
      effectiveCampusId_(student) === assignedAdminCampusId
    )),
    studentLimit: studentLimit,
    campusCount: visibleCampuses.length
  } : null;
  return {
    ok: true,
    records: records,
    campuses: role === "admin" ? campusSummaries : [],
    activeCampusId: activeCampusId,
    serverToday: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
    studentUsage: studentUsage
  };
}

function linkedStudentIdsForParent_(account, students) {
  const email = String(account && account.email || "").trim().toLowerCase();
  const ids = new Set();
  const addIds = value => String(value || "").split(/[,;\s]+/).forEach(id => {
    if (id.trim()) ids.add(id.trim());
  });
  addIds(account && account.childStudentId);
  addIds(account && account.childStudentIds);
  readSchoolRecords_("parent")
    .filter(parent => String(parent.email || "").trim().toLowerCase() === email)
    .forEach(parent => {
      addIds(parent.childStudentId);
      addIds(parent.childStudentIds);
    });
  (students || []).forEach(student => {
    const linkedEmail = String(student.parentEmail || student.guardianEmail || "").trim().toLowerCase();
    if (email && linkedEmail === email && student.studentId) ids.add(String(student.studentId).trim());
  });
  return ids;
}

function requireSchoolAdmin_(body, message) {
  if (!body.authUser || String(body.authUser.role || "").toLowerCase() !== "admin") {
    return { ok: false, error: message || "Only a school administrator can manage this feature." };
  }
  return null;
}

function feeInvoiceHeaders_() {
  return [
    "id", "studentId", "studentName", "parentName", "parentEmail", "className",
    "feePlanId", "feeName", "amount", "discount", "amountDue", "dueDate",
    "academicYear", "term", "status", "createdAt", "createdBy", "paidAmount", "balance", "lastReminderAt"
  ];
}

function feePaymentHeaders_() {
  return [
    "id", "studentId", "studentName", "parentName", "parentEmail", "amount",
    "paidFor", "feePlanId", "invoiceId", "paymentMethod", "reference",
    "paymentDate", "academicYear", "term", "status", "createdAt"
  ];
}

function feePaymentTypeMatchesInvoice_(payment, invoice) {
  const paymentPlanId = String(payment.feePlanId || "").trim();
  const invoicePlanId = String(invoice.feePlanId || "").trim();
  if (paymentPlanId && invoicePlanId && paymentPlanId !== invoicePlanId) return false;
  const paymentFeeName = String(payment.paidFor || "").trim().toLowerCase();
  const invoiceFeeName = String(invoice.feeName || "").trim().toLowerCase();
  return !paymentFeeName || !invoiceFeeName || paymentFeeName === invoiceFeeName;
}

function feePaymentMatchesInvoice_(payment, invoice) {
  return String(payment.invoiceId || "") === String(invoice.id || "") &&
    feePaymentTypeMatchesInvoice_(payment, invoice);
}

function feeInvoiceBalance_(invoice) {
  const totalDue = Number(invoice.amountDue || (Number(invoice.amount || 0) - Number(invoice.discount || 0)));
  const payments = readRows_(PAYMENTS).concat(readSchoolRecords_("fee"))
    .filter(payment => feePaymentMatchesInvoice_(payment, invoice));
  const paid = payments.reduce((sum, payment) => sum + (Number(payment.amount) || 0), 0);
  return { totalDue: totalDue, paid: paid, balance: Math.max(0, totalDue - paid) };
}

function feeInvoiceSummaryMap_(invoices) {
  const invoiceById = new Map((invoices || []).map(invoice => [String(invoice.id), invoice]));
  const summaries = new Map((invoices || []).map(invoice => [String(invoice.id), {
    totalDue: Number(invoice.amountDue || (Number(invoice.amount || 0) - Number(invoice.discount || 0))),
    paid: 0,
    balance: 0
  }]));
  readRows_(PAYMENTS).concat(readSchoolRecords_("fee")).forEach(payment => {
    const invoiceId = String(payment.invoiceId || "");
    const summary = summaries.get(invoiceId);
    const invoice = invoiceById.get(invoiceId);
    if (summary && invoice && feePaymentMatchesInvoice_(payment, invoice)) {
      summary.paid += Number(payment.amount) || 0;
    }
  });
  summaries.forEach(summary => {
    summary.balance = Math.max(0, summary.totalDue - summary.paid);
  });
  return summaries;
}

function updateFeeInvoiceStatus_(invoiceId, actor) {
  const invoice = readSchoolRecords_("feeInvoice").find(item => String(item.id) === String(invoiceId));
  if (!invoice) return null;
  const before = { ...invoice };
  const summary = feeInvoiceBalance_(invoice);
  const date = String(invoice.dueDate || "");
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  invoice.paidAmount = summary.paid;
  invoice.balance = summary.balance;
  invoice.status = summary.balance <= 0
    ? "Paid"
    : date && date < today ? "Overdue"
      : summary.paid > 0 ? "Partially paid" : "Open";
  writeSchoolRecord_("feeInvoice", invoice);
  if (actor) {
    logAudit_(actor, "updated invoice balance", "feeInvoice", invoice.id, before, invoice);
  }
  return invoice;
}

function parentContactsForStudent_(student, knownParents) {
  const studentId = String(student && student.studentId || "").trim();
  const parents = knownParents || readSchoolRecords_("parent");
  const linkedEmails = new Set([student && student.parentEmail, student && student.guardianEmail]
    .map(email => String(email || "").trim().toLowerCase()).filter(Boolean));
  const contacts = parents.filter(parent => {
    const linkedIds = `${parent.childStudentId || ""} ${parent.childStudentIds || ""}`
      .split(/[,;\s]+/).map(id => id.trim()).filter(Boolean);
    return linkedIds.includes(studentId) ||
      linkedEmails.has(String(parent.email || "").trim().toLowerCase());
  }).map(parent => ({
    name: String(parent.name || parent.email || ""),
    email: String(parent.email || "").trim().toLowerCase()
  })).filter(parent => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(parent.email));
  [student && student.parentEmail, student && student.guardianEmail].forEach(email => {
    const normalized = String(email || "").trim().toLowerCase();
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalized) &&
        !contacts.some(contact => contact.email === normalized)) {
      contacts.push({ name: String(student.guardian || ""), email: normalized });
    }
  });
  return contacts;
}

function parentEmailPreferenceEnabled_(email, preference) {
  const account = findUserByEmail_(email);
  if (!account || String(account.role || account.type || "").toLowerCase() !== "parent") return false;
  return String(account[preference] || "").toLowerCase() === "true";
}

function updateNotificationPreferences_(body) {
  const actor = body.authUser || {};
  if (String(actor.role || actor.type || "").toLowerCase() !== "parent") {
    return { ok: false, error: "Only a parent account can change parent email preferences." };
  }
  const fields = [
    "emailAbsenceNotifications",
    "emailPaymentNotifications",
    "emailTimetableNotifications",
    "emailReportCardNotifications"
  ];
  const preferences = body.preferences || {};
  if (fields.some(field => typeof preferences[field] !== "boolean")) {
    return { ok: false, error: "Choose an email preference for each type of school update." };
  }
  const sheet = ensureHeaders_(USERS, [
    "id", "name", "email", "password", "passwordHash", "passwordSalt", "role", "type",
    "jobRole", "department", "className", "childStudentId", "profilePicture", "sessionVersion", "campusId",
    ...fields
  ]);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(header => String(header).trim());
  const emailColumn = headers.indexOf("email");
  const rowIndex = values.findIndex((row, index) =>
    index > 0 && String(row[emailColumn] || "").trim().toLowerCase() === String(actor.email || "").trim().toLowerCase()
  );
  if (rowIndex < 1) return { ok: false, error: "Parent account was not found. Sign in again and retry." };
  const roleColumn = headers.indexOf("role");
  const typeColumn = headers.indexOf("type");
  const storedRole = String(values[rowIndex][roleColumn] || values[rowIndex][typeColumn] || "").toLowerCase();
  if (storedRole !== "parent") return { ok: false, error: "Only a parent account can change parent email preferences." };
  fields.forEach(field => sheet.getRange(rowIndex + 1, headers.indexOf(field) + 1).setValue(preferences[field]));
  return { ok: true, preferences: fields.reduce((result, field) => {
    result[field] = preferences[field];
    return result;
  }, {}) };
}

function appendInboxMessage_(recipient, subject, message, sentBy, actionView) {
  const sheet = ensureHeaders_(MESSAGES, [
    "id", "audience", "recipient", "subject", "message",
    "sendDate", "sendTime", "sentBy", "sentByEmail", "createdAt", "actionView",
    "replyToId", "recipientType", "recipientId", "editedAt", "sentByRole", "messageType"
  ]);
  appendRecord_(sheet, {
    id: Utilities.getUuid(),
    audience: "Specific parent",
    recipient: String(recipient || "").trim().toLowerCase(),
    subject: subject,
    message: message,
    sendDate: new Date(),
    sentBy: sentBy || "School",
    sentByEmail: "",
    createdAt: new Date(),
    actionView: actionView || "",
    sentByRole: "system",
    messageType: "system"
  });
}

function notifyFeePayment_(payment, portalUrl) {
  const warnings = [];
  let emailSent = false;
  let inboxSaved = false;
  let parentLookupFailed = false;
  const receiptMessage = `Payment received: NGN ${Number(payment.amount).toLocaleString("en-NG")} for ${payment.studentName}. ${payment.paidFor}. Receipt: ${payment.reference}.`;
  if (!payment.parentEmail && payment.studentId) {
    try {
      const student = readSchoolRecords_("student").find(record =>
        String(record.studentId || "").trim() === String(payment.studentId).trim()
      );
      const parent = student ? parentContactsForStudent_(student)[0] : null;
      if (parent) {
        payment.parentName = payment.parentName || parent.name;
        payment.parentEmail = parent.email;
      }
    } catch (error) {
      parentLookupFailed = true;
      console.error(`Could not find a parent for fee payment ${payment.reference}: ${error.message}`);
      warnings.push("The parent receipt notification could not be sent because parent contact details could not be loaded.");
    }
  }
  if (payment.parentEmail) {
    try {
      appendInboxMessage_(payment.parentEmail, "Fee payment receipt", receiptMessage, "School", "parentFees");
      inboxSaved = true;
    } catch (error) {
      console.error(`Could not save the parent payment notification: ${error.message}`);
      warnings.push("The in-app parent receipt could not be saved.");
    }
    if (parentEmailPreferenceEnabled_(payment.parentEmail, "emailPaymentNotifications")) {
      try {
        MailApp.sendEmail({
          to: payment.parentEmail,
          subject: `Fee payment receipt · ${payment.studentName}`,
          body: `${receiptMessage}\nDate: ${payment.paymentDate || new Date()}\nMethod: ${payment.paymentMethod}.\nSign in to your school portal to view your invoice balance and download the receipt.${portalUrl ? `\n${portalUrl}` : ""}`
        });
        emailSent = true;
      } catch (error) {
        console.error(`Fee payment ${payment.reference} was recorded, but the parent email failed: ${error.message}`);
        warnings.push("The payment was recorded, but the parent receipt email could not be sent.");
      }
    } else if (inboxSaved) {
      warnings.push("Parent email updates are off; the receipt was saved to the parent's inbox.");
    } else {
      warnings.push("Parent email updates are off and the in-app receipt fallback could not be saved.");
    }
  } else if (!parentLookupFailed) {
    warnings.push("The payment was recorded, but no linked parent email was found for the receipt.");
  }
  try {
    const tenant = findTenantBySpreadsheetId_(requestSpreadsheetId) || {};
    const adminEmail = String(tenant.adminEmail || "").trim().toLowerCase();
    if (adminEmail) appendInboxMessage_(adminEmail, "Fee payment received", receiptMessage, payment.parentName || "Parent", "fees");
  } catch (error) {
    console.error(`Fee payment ${payment.reference} was recorded, but the school inbox notification failed: ${error.message}`);
    warnings.push("The school inbox notification could not be saved.");
  }
  if (emailSent) warnings.push(inboxSaved
    ? "Parent receipt email sent; the in-app inbox copy was saved."
    : "Parent receipt email sent, but the in-app inbox fallback could not be saved.");
  return warnings;
}

function requireFeeStaff_(body, invoice) {
  const actor = body && body.authUser || {};
  const role = String(actor.role || actor.type || "").toLowerCase();
  if (role === "admin") {
    if (invoice && !isSchoolOwnerAdmin_(actor, body.schoolTenant || {}) &&
        effectiveCampusId_(invoice) !== String(actor.campusId || "").trim()) {
      return { ok: false, error: "You may only manage fee payments for your assigned campus." };
    }
    return null;
  }
  const employee = role === "employee" ? readSchoolRecords_("employee").find(item =>
    String(item.email || "").trim().toLowerCase() === String(actor.email || "").trim().toLowerCase()
  ) : null;
  if (!employeeHasPermission_(employee, "fees")) {
    return { ok: false, error: "Your employee account is not authorized to manage fee payments." };
  }
  if (invoice && effectiveCampusId_(invoice) !== effectiveCampusId_(employee)) {
    return { ok: false, error: "You may only manage fee payments for your assigned campus." };
  }
  return null;
}

function notifyAttendanceAbsences_(absences, actor, portalUrl) {
  const warnings = [];
  let emailSent = 0;
  let inboxSaved = 0;
  let parents;
  try {
    parents = readSchoolRecords_("parent");
  } catch (error) {
    console.error(`Attendance was saved, but parent contacts could not be loaded: ${error.message}`);
    return "Attendance was saved, but parent absence notifications could not be sent.";
  }
  (absences || []).forEach(absence => {
    let contacts;
    try {
      contacts = parentContactsForStudent_(absence.student, parents);
    } catch (error) {
      console.error(`Could not find a parent for absence notification: ${error.message}`);
      warnings.push("parent lookup failed");
      return;
    }
    const studentName = `${absence.student.firstName || ""} ${absence.student.lastName || ""}`.trim();
    const message = `${studentName} was marked absent on ${absence.date}. Contact the school office if you have questions.`;
    contacts.forEach(parent => {
      try {
        appendInboxMessage_(parent.email, "Student absence notification", message, actor && actor.name || "School", "attendance");
        inboxSaved++;
      } catch (error) {
        console.error(`Could not save an absence notification for ${parent.email}: ${error.message}`);
        warnings.push(parent.email);
      }
      if (parentEmailPreferenceEnabled_(parent.email, "emailAbsenceNotifications")) {
        try {
          MailApp.sendEmail({ to: parent.email, subject: `Absence notification · ${studentName}`, body: `${message}\nSign in to your school portal to review your child's attendance.${portalUrl ? `\n${portalUrl}` : ""}` });
          emailSent++;
        } catch (error) {
          console.error(`Could not email an absence notification to ${parent.email}: ${error.message}`);
          warnings.push(parent.email);
        }
      } else {
        warnings.push(`${parent.email} (email updates off; inbox fallback saved)`);
      }
    });
  });
  if (warnings.length) {
    return `Attendance saved. Absence emails sent to ${emailSent} parent(s); ${[...new Set(warnings)].length} recipient(s) opted out or could not be notified. ${inboxSaved} inbox copy/copies saved where available.`;
  }
  return absences && absences.length
    ? `Attendance saved and absence emails sent to ${emailSent} parent(s); ${inboxSaved} inbox copy/copies saved.`
    : "";
}

function auditSafeValue_(value) {
  if (Array.isArray(value)) return value.map(auditSafeValue_);
  if (!value || typeof value !== "object") return value;
  const hidden = /password|secret|token|credential|passport|profilepicture|image|photo/i;
  return Object.keys(value).reduce((safe, key) => {
    if (!hidden.test(key)) safe[key] = auditSafeValue_(value[key]);
    return safe;
  }, {});
}

function logAudit_(actor, action, recordType, recordId, before, after, summary) {
  const sheet = ensureHeaders_("Audit Log", [
    "id", "timestamp", "actorName", "actorEmail", "action",
    "recordType", "recordId", "summary", "beforeJson", "afterJson"
  ]);
  appendRecord_(sheet, {
    id: Utilities.getUuid(),
    timestamp: new Date(),
    actorName: String(actor && actor.name || "School administrator"),
    actorEmail: String(actor && actor.email || "").toLowerCase(),
    action: action,
    recordType: recordType,
    recordId: String(recordId || ""),
    summary: String(summary || `${action} ${recordType}`),
    beforeJson: before == null ? "" : JSON.stringify(auditSafeValue_(before)).slice(0, 30000),
    afterJson: after == null ? "" : JSON.stringify(auditSafeValue_(after)).slice(0, 30000)
  });
}

function getAuditLog_(body) {
  const denied = requireSchoolAdmin_(body, "Only a school administrator can view the audit log.");
  if (denied) return denied;
  return { ok: true, entries: readRows_("Audit Log").slice(-500).reverse() };
}

function createFeeInvoices_(body) {
  const denied = requireSchoolAdmin_(body);
  if (denied) return denied;
  const campusId = validatedCampusId_(body.schoolTenant || {}, body.campusId || DEFAULT_CAMPUS_ID);
  if (!campusId) return { ok: false, error: "Select an active campus in this school workspace." };
  const className = String(body.className || "").trim();
  const feePlanId = String(body.feePlanId || "").trim();
  const term = String(body.term || "").trim();
  const academicYear = String(body.academicYear || "").trim();
  const dueDate = String(body.dueDate || "").trim();
  const discount = Number(body.discount || 0);
  const validDate = isValidDateKey_(dueDate);
  if (!className || !feePlanId || !term || !academicYear || !validDate ||
      !Number.isFinite(discount) || discount < 0) {
    return { ok: false, error: "Choose a class, fee type, term, academic year, valid due date, and non-negative discount." };
  }
  const feePlan = schoolFeePlans_(readSchoolRecords_("feePlan"), body.authUser).find(plan =>
    String(plan.id) === feePlanId && plan.active !== false &&
    String(plan.active || "Active").toLowerCase() !== "inactive"
  );
  if (!feePlan || !Number.isFinite(Number(feePlan.amount)) || Number(feePlan.amount) <= 0) {
    return { ok: false, error: "Select an active fee type with a valid amount." };
  }
  if (discount >= Number(feePlan.amount)) {
    return { ok: false, error: "The discount for each student must be less than the fee amount." };
  }
  const students = readSchoolRecords_("student").filter(student =>
    String(student.className || "").trim() === className &&
    String(student.studentId || "").trim() &&
    isActiveStudent_(student) &&
    effectiveCampusId_(student) === campusId
  );
  if (!students.length) return { ok: false, error: "There are no students with valid IDs in this class." };
  const parents = readSchoolRecords_("parent");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
  const existing = readSchoolRecords_("feeInvoice");
  const existingKeys = new Set(existing.map(invoice =>
    [invoice.studentId, invoice.feePlanId, invoice.term, invoice.academicYear].join("|").toLowerCase()
  ));
  const createdAt = new Date();
  const invoices = [];
  let skipped = 0;
  students.forEach(student => {
    const key = [student.studentId, feePlanId, term, academicYear].join("|").toLowerCase();
    if (existingKeys.has(key)) {
      skipped++;
      return;
    }
    existingKeys.add(key);
    const parent = parentContactsForStudent_(student, parents)[0] || {};
    const amount = Number(feePlan.amount);
    invoices.push({
      id: Utilities.getUuid(),
      studentId: String(student.studentId).trim(),
      studentName: `${student.firstName || ""} ${student.lastName || ""}`.trim(),
      parentName: parent.name || "",
      parentEmail: parent.email || "",
      className: className,
      campusId: campusId,
      feePlanId: feePlanId,
      feeName: String(feePlan.name || ""),
      amount: amount,
      discount: discount,
      amountDue: amount - discount,
      dueDate: dueDate,
      academicYear: academicYear,
      term: term,
      status: "Open",
      createdAt: createdAt,
      createdBy: String(body.authUser.email || ""),
      paidAmount: 0,
      balance: amount - discount
    });
  });
  if (invoices.length) {
    const sheet = ensureHeaders_(RECORD_SHEETS.feeInvoice, feeInvoiceHeaders_());
    const headers = sheet.getDataRange().getValues()[0].map(String);
    const rows = invoices.map(invoice => headers.map(header =>
      Object.prototype.hasOwnProperty.call(invoice, header) ? invoice[header] : ""
    ));
    writeSafeRows_(sheet, sheet.getLastRow() + 1, rows);
  }
  logAudit_(body.authUser, "created bulk invoices", "feeInvoice", className, null, {
    className: className,
    feePlanId: feePlanId,
    term: term,
    academicYear: academicYear,
    dueDate: dueDate,
    invoiceCount: invoices.length,
    skipped: skipped
  }, `${invoices.length} invoices created for ${className}; ${skipped} duplicates skipped.`);
  return { ok: true, invoices: invoices, skipped: skipped };
  } finally {
    lock.releaseLock();
  }
}

function recordFeePayment_(body) {
  const denied = requireFeeStaff_(body);
  if (denied) return denied;
  const invoiceId = String(body.invoiceId || "").trim();
  const amount = Number(body.amount);
  const method = String(body.paymentMethod || "").trim();
  const paymentDate = String(body.paymentDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"));
  const dateParts = paymentDate.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!dateParts || !isValidDateKey_(paymentDate)) {
    return { ok: false, error: "Enter a valid payment date." };
  }
  if (!["Cash", "Bank transfer", "Card", "Other"].includes(method)) {
    return { ok: false, error: "Choose a valid payment method." };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let payment, updatedInvoice;
  try {
    const invoice = readSchoolRecords_("feeInvoice").find(item => String(item.id) === invoiceId);
    if (!invoice) return { ok: false, error: "Select a valid invoice for this payment." };
    const invoiceDenied = requireFeeStaff_(body, invoice);
    if (invoiceDenied) return invoiceDenied;
    const summary = feeInvoiceBalance_(invoice);
    if (!Number.isFinite(amount) || amount <= 0 || amount > summary.balance + 0.000001) {
      return { ok: false, error: `Enter a payment greater than zero and no more than the remaining balance of NGN ${summary.balance.toLocaleString("en-NG")}.` };
    }
    const reference = String(body.reference || `RCPT-${Utilities.getUuid().slice(0, 8).toUpperCase()}`).trim();
    if (reference.length > 100 || readRows_(PAYMENTS).some(record =>
      String(record.reference || "").toLowerCase() === reference.toLowerCase()
    )) {
      return { ok: false, error: "That receipt reference is invalid or already in use." };
    }
    payment = {
      id: Utilities.getUuid(),
      studentId: invoice.studentId,
      studentName: invoice.studentName,
      parentName: invoice.parentName || "",
      parentEmail: invoice.parentEmail || "",
      amount: amount,
      paidFor: invoice.feeName,
      feePlanId: invoice.feePlanId,
      invoiceId: invoice.id,
      paymentMethod: method,
      reference: reference,
      paymentDate: paymentDate,
      academicYear: invoice.academicYear,
      term: invoice.term,
      status: "paid",
      createdAt: new Date()
    };
    appendRecord_(ensureHeaders_(PAYMENTS, feePaymentHeaders_()), payment);
    updatedInvoice = updateFeeInvoiceStatus_(invoice.id, body.authUser);
  } finally {
    lock.releaseLock();
  }
  const warnings = notifyFeePayment_(payment, body.portalUrl);
  logAudit_(body.authUser, "recorded fee payment", "payment", payment.id, null, payment, `Recorded NGN ${amount} for invoice ${payment.invoiceId}.`);
  return { ok: true, payment: payment, invoice: updatedInvoice, notificationWarning: warnings.join(" ") };
}

function reconcileFeePayment_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
  const paymentId = String(body.paymentId || "").trim();
  const invoiceId = String(body.invoiceId || "").trim();
  const invoice = readSchoolRecords_("feeInvoice").find(item => String(item.id) === invoiceId);
  if (!paymentId || !invoice) return { ok: false, error: "Select a valid payment and invoice to reconcile." };
  const denied = requireFeeStaff_(body, invoice);
  if (denied) return denied;
  const sheet = ensureHeaders_(PAYMENTS, feePaymentHeaders_());
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const idColumn = headers.indexOf("id");
  const rowIndex = values.findIndex((row, index) => index > 0 && String(row[idColumn]) === paymentId);
  const paymentSheet = rowIndex > 0 ? sheet : null;
  const legacyPayments = paymentSheet ? [] : readSchoolRecords_("fee");
  const legacyIndex = legacyPayments.findIndex(record => String(record.id) === paymentId);
  if (!paymentSheet && legacyIndex < 0) return { ok: false, error: "Payment not found. Refresh and try again." };
  const payment = paymentSheet ? rowToObject_(headers, values[rowIndex]) : legacyPayments[legacyIndex];
  const currentInvoice = payment.invoiceId
    ? readSchoolRecords_("feeInvoice").find(item => String(item.id) === String(payment.invoiceId))
    : null;
  if (currentInvoice && feePaymentMatchesInvoice_(payment, currentInvoice)) {
    return { ok: false, error: "This payment is already linked to an invoice." };
  }
  const paymentStudentId = String(payment.studentId || "").trim();
  const sameStudent = paymentStudentId
    ? paymentStudentId === String(invoice.studentId || "").trim()
    : String(payment.student || payment.studentName || "").trim().toLowerCase() ===
      String(invoice.studentName || "").trim().toLowerCase();
  if (!sameStudent) {
    return { ok: false, error: "The payment and invoice belong to different students." };
  }
  if (!feePaymentTypeMatchesInvoice_(payment, invoice)) {
    return { ok: false, error: "The payment and invoice belong to different fee types." };
  }
  const summary = feeInvoiceBalance_(invoice);
  if (!(Number(payment.amount) > 0) || Number(payment.amount) > summary.balance + 0.000001) {
    return { ok: false, error: "The payment amount exceeds the invoice's remaining balance." };
  }
  const previousLink = {
    invoiceId: String(payment.invoiceId || ""),
    feePlanId: String(payment.feePlanId || "")
  };
  payment.invoiceId = invoice.id;
  payment.feePlanId = invoice.feePlanId;
  payment.academicYear = invoice.academicYear;
  payment.term = invoice.term;
  payment.studentId = invoice.studentId;
  payment.studentName = payment.studentName || payment.student || invoice.studentName;
  if (paymentSheet) writeRecord_(paymentSheet, rowIndex + 1, payment, headers);
  else writeSchoolRecord_("fee", payment);
  const updatedInvoice = updateFeeInvoiceStatus_(invoice.id, body.authUser);
  logAudit_(
    body.authUser,
    "reconciled payment",
    "payment",
    payment.id,
    previousLink,
    { invoiceId: invoice.id, feePlanId: invoice.feePlanId },
    `Linked payment ${payment.id} to invoice ${invoice.id}.`
  );
  return { ok: true, payment: payment, invoice: updatedInvoice };
  } finally {
    lock.releaseLock();
  }
}

function sendFeeReminders_(body) {
  const denied = requireSchoolAdmin_(body);
  if (denied) return denied;
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const allInvoices = readSchoolRecords_("feeInvoice");
  const summaries = feeInvoiceSummaryMap_(allInvoices);
  const invoices = allInvoices.filter(invoice => {
    const dueDate = invoice.dueDate instanceof Date
      ? Utilities.formatDate(invoice.dueDate, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(invoice.dueDate || "").slice(0, 10);
    return isValidDateKey_(dueDate) && dueDate < today &&
    (invoice.lastReminderAt instanceof Date
      ? Utilities.formatDate(invoice.lastReminderAt, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(invoice.lastReminderAt || "").slice(0, 10)) !== today &&
    summaries.get(String(invoice.id)).balance > 0;
  });
  const warnings = [];
  let sent = 0;   
  const students = readSchoolRecords_("student");
  const parents = readSchoolRecords_("parent");
  invoices.forEach(invoice => {
    const student = students.find(item => String(item.studentId) === String(invoice.studentId));
    const contacts = student ? parentContactsForStudent_(student, parents) : [];
    if (!contacts.length && invoice.parentEmail) contacts.push({ name: invoice.parentName, email: invoice.parentEmail });
    const balance = summaries.get(String(invoice.id)).balance;
    const sentBefore = sent;
    contacts.forEach(parent => {
      const message = `Reminder: NGN ${balance.toLocaleString("en-NG")} remains due for ${invoice.studentName}'s ${invoice.feeName} (${invoice.term}, ${invoice.academicYear}). Due date: ${invoice.dueDate}.`;
      try {
        appendInboxMessage_(parent.email, "Overdue school fee reminder", message, body.authUser.name || "School", "parentFees");
        MailApp.sendEmail({ to: parent.email, subject: "Overdue school fee reminder", body: message });
        sent++;
      } catch (error) {
        console.error(`Could not send fee reminder to ${parent.email}: ${error.message}`);
        warnings.push(parent.email);
      }
    });
    if (sent > sentBefore) {
      invoice.lastReminderAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
      writeSchoolRecord_("feeInvoice", invoice);
    }
  });
  logAudit_(body.authUser, "sent fee reminders", "feeInvoice", "", null, { invoiceCount: invoices.length, recipientCount: sent }, `Sent ${sent} overdue fee reminders.`);
  return {
    ok: true,
    sent: sent,
    overdueInvoices: invoices.length,
    warning: warnings.length ? `Could not notify ${[...new Set(warnings)].length} parent email address(es).` : ""
  };
}

function gradeScaleBand_(className, average, scales) {
  const classKey = String(className || "").trim().toLowerCase();
  const matching = (scales || []).find(scale =>
    String(scale.className || "").trim().toLowerCase() === classKey && classKey
  ) || (scales || []).find(scale => !String(scale.className || "").trim());
  if (!matching) return null;
  const bands = Array.isArray(matching.bands) ? matching.bands : [];
  const band = bands
    .filter(item => Number(average) >= Number(item.minScore))
    .sort((left, right) => Number(right.minScore) - Number(left.minScore))[0];
  if (!band) throw new Error(`The grading scale "${matching.name}" has no grade for ${Number(average).toFixed(1)}%.`);
  return { grade: String(band.grade), description: String(band.description || "") };
}

function reportGradeBand_(className, average, scales) {
  const level = String(className || "").trim();
  if (!level) throw new Error("The selected student does not have a class assigned. Add a class before processing the report card.");
  const configured = gradeScaleBand_(level, average, scales);
  if (configured) return configured;
  const tokens = level.toLowerCase().match(/[a-z]+|\d+/g) || [];
  const classIndex = tokens.findIndex(token => ["grade", "class", "year", "level", "form", "standard"].includes(token));
  const classLevel = classIndex >= 0 ? Number(tokens[classIndex + 1]) : NaN;
  const junior = tokens.includes("jss") ||
    (tokens.includes("junior") && (tokens.includes("secondary") || tokens.includes("sec"))) ||
    [7, 8, 9].includes(classLevel);
  const senior = tokens.includes("ss") || tokens.includes("sss") ||
    (tokens.includes("senior") && (tokens.includes("secondary") || tokens.includes("sec"))) ||
    [10, 11, 12].includes(classLevel);
  const earlyYears = tokens.some(token =>
    ["nursery", "creche", "kindergarten", "reception", "kg", "preschool", "preprimary", "eyfs"].includes(token)
  );
  const primary = tokens.some(token => ["primary", "basic"].includes(token)) ||
    earlyYears ||
    [1, 2, 3, 4, 5, 6].includes(classLevel);

  if (junior) {
    if (average >= 80) return { grade: "A", description: "Excellent / Distinction" };
    if (average >= 60) return { grade: "C", description: "Credit" };
    if (average >= 50) return { grade: "P", description: "Pass" };
    return { grade: "F", description: "Fail" };
  }
  if (senior) {
    if (average >= 75) return { grade: "A1", description: "Excellent" };
    if (average >= 70) return { grade: "B2", description: "Very Good" };
    if (average >= 65) return { grade: "B3", description: "Good" };
    if (average >= 60) return { grade: "C4", description: "Credit" };
    if (average >= 55) return { grade: "C5", description: "Credit" };
    if (average >= 50) return { grade: "C6", description: "Credit" };
    if (average >= 45) return { grade: "D7", description: "Pass" };
    if (average >= 40) return { grade: "E8", description: "Pass" };
    return { grade: "F9", description: "Fail" };
  }
  if (primary) {
    if (average >= 70) return { grade: "A", description: "Excellent / Mastery" };
    if (average >= 60) return { grade: "B", description: "Very Good" };
    if (average >= 50) return { grade: "C", description: "Good / Credit" };
    if (average >= 45) return { grade: "D", description: "Fair / Pass" };
    if (average >= 40) return { grade: "E", description: "Weak Pass" };
    return { grade: "F", description: "Fail" };
  }
  throw new Error(`The student's class "${level.slice(0, 100)}" could not be assigned a grading scale.`);
}

function reportCardAverage_(subjectsAndResults) {
  const lines = String(subjectsAndResults || "").split(/\r?\n/).map(line => line.trim()).filter(Boolean);
  if (!lines.length) throw new Error("Enter at least one subject result before processing the grade.");
  const scores = lines.map((line, index) => {
    const match = line.match(/^(.+?)\s*(?::|\s+[–—-]\s+)\s*(.+)$/);
    if (!match) {
      throw new Error(`Line ${index + 1}: use the format Subject: score (for example Mathematics: 85%).`);
    }
    if (!match[1].trim()) throw new Error(`Line ${index + 1}: enter a subject name.`);
    const result = match[2].trim();
    const fraction = result.match(/^(\d+(?:\.\d+)?)\s*\/\s*(\d+(?:\.\d+)?)$/);
    const number = result.match(/^(\d+(?:\.\d+)?)\s*%?$/);
    let score;
    if (fraction) {
      const earned = Number(fraction[1]);
      const possible = Number(fraction[2]);
      if (possible <= 0) throw new Error(`Line ${index + 1}: the total marks must be greater than zero.`);
      score = earned / possible * 100;
    } else if (number) {
      score = Number(number[1]);
    } else {
      throw new Error(`Line ${index + 1}: enter a percentage, score out of 100, or marks as earned/total.`);
    }
    if (!Number.isFinite(score) || score < 0 || score > 100) {
      throw new Error(`Line ${index + 1}: scores must be between 0 and 100%.`);
    }
    return score;
  });
  return scores.reduce((sum, score) => sum + score, 0) / scores.length;
}

function saveRecord_(body, accountSaved = false) {
  const actor = body.authUser;
  if (!actor) {
    return { ok: false, error: "User account was not found. Please sign in again." };
  }
  const type = String(body.type || "");
  const record = body.record;
  const allowedTypes = ["campus", "student", "employee", "parent", "application", "academic", "attendance", "timetable", "reportCard", "gradeScale", "gradeEntry", "followUp", "fee", "feePlan", "event"];
  if (!allowedTypes.includes(type) || !record || typeof record !== "object") {
    return { ok: false, error: "The school record is invalid." };
  }
  const role = String(actor.role || actor.type || "").toLowerCase();
  const tenant = body.schoolTenant || findTenantBySpreadsheetId_(requestSpreadsheetId) || {};
  const recordData = { ...record };
  let campusPinToSet = "";
  let creatingCampus = false;
  let campusMetadataChanged = false;
  let existingCampusRecord = null;
  const attendanceEmployee = role === "employee" && type === "attendance"
    ? readSchoolRecords_("employee").find(employee =>
      String(employee.email || "").trim().toLowerCase() === String(actor.email || "").trim().toLowerCase()
    )
    : null;
  if (type === "attendance" && role === "employee") {
    if (!attendanceEmployee || !String(attendanceEmployee.role || "").toLowerCase().includes("teacher")) {
      return { ok: false, error: "Only an assigned teacher can record attendance." };
    }
    recordData.className = String(attendanceEmployee.className || "").trim();
    recordData.campusId = effectiveCampusId_(attendanceEmployee);
  }
  if (type === "campus") {
    const currentCampuses = readSchoolRecords_("campus");
    existingCampusRecord = currentCampuses.find(campus =>
      String(campus.id || "") === String(recordData.id || "")
    );
    creatingCampus = !existingCampusRecord;
    if (!planSupportsFeature_(tenant.plan, "multiCampus") &&
        (!existingCampusRecord || String(recordData.id || "") !== DEFAULT_CAMPUS_ID)) {
      return { ok: false, error: "Multiple campuses require the Enterprise plan. Upgrade the school plan to continue." };
    }
    const campusName = String(recordData.name || "").trim();
    if (!campusName || campusName.length > 100) {
      return { ok: false, error: "Enter a campus name up to 100 characters long." };
    }
    campusPinToSet = String(recordData.campusPin || "");
    delete recordData.campusPin;
    delete recordData.pinConfigured;
    if (!existingCampusRecord && campusPinToSet.length < 6) {
      return { ok: false, error: "Set a campus PIN of at least 6 characters before creating the campus." };
    }
    if (campusPinToSet && (campusPinToSet.length < 6 || campusPinToSet.length > 64)) {
      return { ok: false, error: "Campus PINs must be between 6 and 64 characters." };
    }
    const duplicateCampus = currentCampuses.find(campus =>
      String(campus.id || "") !== String(recordData.id || "") &&
      String(campus.name || "").trim().toLowerCase() === campusName.toLowerCase()
    );
    if (duplicateCampus) return { ok: false, error: "A campus with that name already exists." };
    recordData.name = campusName;
    recordData.address = String(recordData.address || "").trim().slice(0, 500);
    campusMetadataChanged = !existingCampusRecord ||
      String(existingCampusRecord.name || "").trim() !== recordData.name ||
      String(existingCampusRecord.address || "").trim() !== recordData.address;
  } else if (type !== "parent" && type !== "feePlan" && type !== "gradeScale") {
    const campusId = validatedCampusId_(tenant, recordData.campusId || body.campusId || DEFAULT_CAMPUS_ID);
    if (!campusId) return { ok: false, error: "Select an active campus in this school workspace." };
    recordData.campusId = campusId;
  }
  if (type === "feePlan") delete recordData.campusId;
  if (type === "gradeScale") delete recordData.campusId;
  if (["attendance", "reportCard", "gradeEntry"].includes(type) && recordData.studentId) {
    const linkedStudent = readSchoolRecords_("student").find(student =>
      String(student.studentId || "").trim() === String(recordData.studentId || "").trim()
    );
    if (!linkedStudent) return { ok: false, error: "Select a valid student in the active campus." };
    if (attendanceEmployee &&
        String(linkedStudent.className || "").trim() !== String(attendanceEmployee.className || "").trim()) {
      return { ok: false, error: "Teachers can only record attendance for students in their assigned class." };
    }
    const gradeEmployee = type === "gradeEntry" && role === "employee"
      ? readSchoolRecords_("employee").find(employee =>
        String(employee.email || "").trim().toLowerCase() === String(actor.email || "").trim().toLowerCase()
      )
      : null;
    const expectedCampusId = gradeEmployee
      ? effectiveCampusId_(gradeEmployee)
      : recordData.campusId;
    if (expectedCampusId && effectiveCampusId_(linkedStudent) !== expectedCampusId) {
      return { ok: false, error: "The selected student belongs to a different campus." };
    }
  }
  if (type === "timetable") return saveTimetableRecord_(actor, recordData, body);
  if (type === "followUp") {
    if (role !== "admin") return { ok: false, error: "Only a school administrator can manage student follow-ups." };
    const student = readSchoolRecords_("student").find(item =>
      String(item.studentId || "").trim() === String(recordData.studentId || "").trim()
    );
    if (!student || effectiveCampusId_(student) !== String(recordData.campusId || "")) {
      return { ok: false, error: "Select a valid student in the active campus." };
    }
    const signalType = String(recordData.signalType || "");
    const signalKey = String(recordData.signalKey || "").trim();
    const status = String(recordData.status || "Open");
    if (!["attendance", "academic"].includes(signalType) || !signalKey || signalKey.length > 200 ||
        !["Open", "Completed"].includes(status)) {
      return { ok: false, error: "The student follow-up details are invalid." };
    }
    recordData.studentName = `${student.firstName || ""} ${student.lastName || ""}`.trim();
    recordData.className = String(student.className || "");
    recordData.status = status;
    recordData.updatedAt = new Date().toISOString();
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const existing = readSchoolRecords_("followUp").find(item =>
        String(item.signalKey || "") === signalKey
      );
      recordData.id = existing && existing.id || String(recordData.id || Utilities.getUuid());
      if (existing && existing.status === "Completed" && status === "Open") {
        return { ok: false, error: "This follow-up is already complete. Refresh the radar before making changes." };
      }
      const updated = writeSchoolRecord_("followUp", recordData);
      logAudit_(actor, updated ? "updated student follow-up" : "created student follow-up", "followUp", recordData.id, existing, recordData);
      return { ok: true, updated: updated, record: recordData };
    } finally {
      lock.releaseLock();
    }
  }
  if (type === "reportCard") {
    const student = readSchoolRecords_("student").find(item =>
      String(item.studentId || "").trim() === String(recordData.studentId || "").trim()
    );
    if (!student) return { ok: false, error: "Select a valid student before saving this report card." };
    const parentEmail = String(recordData.parentEmail || "").trim().toLowerCase();
    const parent = readSchoolRecords_("parent").find(item =>
      String(item.email || "").trim().toLowerCase() === parentEmail
    );
    if (!parent || !parentEmail) return { ok: false, error: "Select a parent with a saved email address." };
    if (!linkedStudentIdsForParent_(parent, [student]).has(String(student.studentId || "").trim())) {
      return { ok: false, error: "The selected parent is linked to a different student." };
    }
    recordData.studentName = `${student.firstName || ""} ${student.lastName || ""}`.trim();
    recordData.className = String(student.className || "");
    recordData.parentName = String(parent.name || parent.email || "");
    try {
      const average = reportCardAverage_(recordData.subjectsAndResults);
      const band = reportGradeBand_(recordData.className, average, readSchoolRecords_("gradeScale"));
      recordData.overallGrade = `${band.grade} — ${band.description} (${average.toFixed(1)}%)`;
    } catch (error) {
      return { ok: false, error: error.message };
    }
  }
  if (type === "attendance") {
    recordData.date = String(recordData.date || "").trim();
    if (!isWeekdayDateKey_(recordData.date)) {
      return { ok: false, error: "Attendance can only be recorded Monday through Friday. Choose a weekday date." };
    }
    const studentId = String(recordData.studentId || "").trim();
    const student = readSchoolRecords_("student").find(item =>
      String(item.studentId || "").trim() === studentId
    );
    if (!student) return { ok: false, error: "Select a valid student before recording attendance." };
    if (attendanceEmployee &&
        String(student.className || "").trim() !== String(attendanceEmployee.className || "").trim()) {
      return { ok: false, error: "Teachers can only record attendance for students in their assigned class." };
    }
    const status = String(recordData.status || "").trim().toLowerCase();
    const statuses = { present: "Present", absent: "Absent", late: "Late", excused: "Excused" };
    if (!statuses[status]) return { ok: false, error: "Choose Present, Absent, Late, or Excused for this student." };
    recordData.status = statuses[status];
    recordData.studentName = `${student.firstName || ""} ${student.lastName || ""}`.trim();
    recordData.className = String(student.className || "");
    recordData.present = status === "present" ? 1 : 0;
    recordData.absent = status === "absent" ? 1 : 0;
  }
  if (type === "feePlan") {
    const name = String(recordData.name || "").trim();
    const amount = Number(recordData.amount);
    if (!name || !Number.isFinite(amount) || amount <= 0) {
      return { ok: false, error: "Enter a fee name and an amount greater than zero." };
    }
    const feePlans = readSchoolRecords_("feePlan");
    const existingFeePlan = feePlans.find(plan => String(plan.id) === String(recordData.id || ""));
    const duplicateFeePlan = feePlans.find(plan =>
      String(plan.id) !== String(recordData.id || "") &&
      String(plan.name || "").trim().toLowerCase() === name.toLowerCase()
    );
    if (duplicateFeePlan && (!existingFeePlan ||
        String(existingFeePlan.name || "").trim().toLowerCase() !== name.toLowerCase())) {
      return { ok: false, error: "Fee types are shared across the school. Edit the existing fee type instead of creating another with the same name." };
    }
    recordData.name = name;
    recordData.amount = amount;
    recordData.description = String(recordData.description || "").trim();
    recordData.active = String(recordData.active || "Active").toLowerCase() !== "inactive";
  }
  if (type === "gradeScale") {
    if (role !== "admin") return { ok: false, error: "Only a school administrator can manage grading scales." };
    const name = String(recordData.name || "").trim();
    const className = String(recordData.className || "").trim();
    const bands = Array.isArray(recordData.bands) ? recordData.bands : [];
    if (!name || name.length > 100 || className.length > 100 || bands.length < 2 || bands.length > 20) {
      return { ok: false, error: "Enter a scale name and between 2 and 20 grade bands." };
    }
    const normalizedBands = [];
    const seenScores = new Set();
    for (const band of bands) {
      const minScore = Number(band && band.minScore);
      const grade = String(band && band.grade || "").trim();
      const description = String(band && band.description || "").trim();
      if (!Number.isFinite(minScore) || minScore < 0 || minScore > 100 || !grade || grade.length > 20 || description.length > 100 || seenScores.has(minScore)) {
        return { ok: false, error: "Each grade band needs a unique minimum score from 0 to 100 and a grade label." };
      }
      seenScores.add(minScore);
      normalizedBands.push({ minScore: minScore, grade: grade, description: description });
    }
    if (!seenScores.has(0)) return { ok: false, error: "Add a grade band with a minimum score of 0 so every result has a grade." };
    const duplicate = readSchoolRecords_("gradeScale").find(scale =>
      String(scale.id || "") !== String(recordData.id || "") &&
      String(scale.className || "").trim().toLowerCase() === className.toLowerCase()
    );
    if (duplicate) return { ok: false, error: className ? `A grading scale is already configured for ${className}. Edit that scale instead.` : "A school-wide default grading scale already exists. Edit that scale instead." };
    recordData.name = name;
    recordData.className = className;
    recordData.bands = normalizedBands.sort((left, right) => right.minScore - left.minScore);
  }
  if (type === "gradeEntry") {
    const employee = role === "employee" ? readSchoolRecords_("employee").find(item =>
      String(item.email || "").trim().toLowerCase() === String(actor.email || "").trim().toLowerCase()
    ) : null;
    if (role !== "admin" && !employeeHasPermission_(employee, "grades")) {
      return { ok: false, error: "Your employee account is not authorized to enter grades." };
    }
    const student = readSchoolRecords_("student").find(item =>
      String(item.studentId || "").trim() === String(recordData.studentId || "").trim()
    );
    if (!student) return { ok: false, error: "Select a valid student before entering a grade." };
    const subject = String(recordData.subject || "").trim();
    const score = Number(recordData.score);
    const term = String(recordData.term || "").trim();
    const academicYear = String(recordData.academicYear || "").trim();
    if (!subject || subject.length > 120 || !Number.isFinite(score) || score < 0 || score > 100 || !term || !academicYear) {
      return { ok: false, error: "Enter a subject, score from 0 to 100, term, and academic year." };
    }
    if (employee) {
      const assignedClass = String(employee.className || "").trim();
      const assignedSubjects = String(employee.subjects || "").split(/[,;\n]/).map(value => value.trim().toLowerCase()).filter(Boolean);
      if (String(student.className || "").trim() !== assignedClass || !assignedSubjects.includes(subject.toLowerCase())) {
        return { ok: false, error: "You may only enter scores for your assigned class and subjects." };
      }
      const existingEntry = readSchoolRecords_("gradeEntry").find(item => String(item.id) === String(recordData.id || ""));
      if (existingEntry && (String(existingEntry.enteredByEmail || "").toLowerCase() !== String(actor.email || "").toLowerCase() ||
          String(existingEntry.status || "Pending") !== "Pending")) {
        return { ok: false, error: "This grade entry is not editable by your account." };
      }
      recordData.enteredByEmail = String(actor.email || "").toLowerCase();
      recordData.enteredBy = String(employee.name || actor.name || "");
      recordData.status = "Pending";
    }
    recordData.studentName = `${student.firstName || ""} ${student.lastName || ""}`.trim();
    recordData.className = String(student.className || "");
    recordData.subject = subject;
    recordData.score = score;
    recordData.term = term;
    recordData.academicYear = academicYear;
  }
  const canRecordAttendance = Boolean(attendanceEmployee);
  const canEnterGrade = role === "employee" && type === "gradeEntry";
  if (role !== "admin" && !canRecordAttendance && !canEnterGrade) {
    return { ok: false, error: "You do not have permission to save this school record." };
  }

  const id = String(record.id || "");
  if (!id) {
    return { ok: false, error: "The school record is missing its ID." };
  }
  const oldRecord = (type === "campus"
    ? existingCampusRecord
    : readSchoolRecords_(type).find(item => String(item.id) === id)) || null;
  if (role === "admin" && oldRecord && !isSchoolOwnerAdmin_(actor, tenant) &&
      ["student", "employee", "application", "academic", "attendance", "timetable", "reportCard", "gradeEntry", "followUp", "fee", "feeInvoice", "event", "employeeLeave"].includes(type) &&
      effectiveCampusId_(oldRecord) !== String(actor.campusId || "").trim()) {
    return { ok: false, error: "You are not authorized to update records for another campus." };
  }
  if (role === "admin" && type === "parent" && !isSchoolOwnerAdmin_(actor, tenant)) {
    const linkedIds = new Set();
    const parentStudents = readSchoolRecords_("student");
    [oldRecord, recordData].filter(Boolean).forEach(parent => {
      [parent.childStudentId, parent.childStudentIds].forEach(value =>
        String(value || "").split(/[,;\s]+/).forEach(studentId => { if (studentId.trim()) linkedIds.add(studentId.trim()); })
      );
    });
    if (!linkedIds.size) {
      const email = String(recordData.email || oldRecord && oldRecord.email || "").trim().toLowerCase();
      parentStudents.forEach(student => {
        const studentEmail = String(student.parentEmail || student.guardianEmail || "").trim().toLowerCase();
        if (email && email === studentEmail && student.studentId) linkedIds.add(String(student.studentId).trim());
      });
    }
    const linkedStudents = parentStudents.filter(student =>
      linkedIds.has(String(student.studentId || "").trim())
    );
    if (!linkedIds.size || linkedStudents.length !== linkedIds.size ||
        linkedStudents.some(student => effectiveCampusId_(student) !== String(actor.campusId || "").trim())) {
      return { ok: false, error: "You may only manage parent records linked exclusively to your assigned campus." };
    }
  }
  if (type === "reportCard" && oldRecord) {
    const changed = ["studentId", "term", "academicYear", "subjectsAndResults", "overallGrade", "teacherComment", "principalComment", "parentEmail"]
      .some(key => String(oldRecord[key] || "") !== String(recordData[key] || ""));
    recordData.sentAt = changed ? "" : recordData.sentAt || oldRecord.sentAt || "";
  }
  if (["student", "employee"].includes(type)) {
    const submittedPhoto = String(recordData.passport || "");
    const recordPhoto = String(recordData.profilePicture || "");
    const previousPhoto = String(oldRecord && oldRecord.profilePicture || "");
    const changedRecordPhoto = recordPhoto && recordPhoto !== previousPhoto ? recordPhoto : "";
    const newPhoto = submittedPhoto || changedRecordPhoto;
    if (newPhoto && (newPhoto.length > 45000 ||
        !/^data:image\/(?:jpeg|png);base64,[A-Za-z0-9+/=]+$/i.test(newPhoto))) {
      return { ok: false, error: "The profile photo must be a valid JPG or PNG image under 45,000 characters." };
    }
    recordData.profilePicture = newPhoto || idCardPhoto_(oldRecord);
  }
  if (type === "employee" && !oldRecord && !accountSaved) {
    return { ok: false, error: "Create new employee records through the invitation workflow." };
  }
  if (type === "employee" && oldRecord && !accountSaved &&
      String(oldRecord.email || "").trim().toLowerCase() !== String(recordData.email || "").trim().toLowerCase()) {
    recordData.activationTokenHash = "";
    recordData.activationTokenExpiresAt = "";
    recordData.dashboardStatus = "Invitation Needed";
    recordData.invitationStatus = "Not sent";
  }

  if (type === "student") {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const students = readSchoolRecords_("student");
      const existingStudent = students.find(item => String(item.id) === id);
      if (!recordData.status) recordData.status = existingStudent && existingStudent.status || "Active";
      if (!["active", "withdrawn", "archived"].includes(String(recordData.status).trim().toLowerCase())) {
        return { ok: false, error: "Choose Active, Withdrawn, or Archived for the student status." };
      }
      recordData.status = String(recordData.status).trim().replace(/^./, first => first.toUpperCase());
      if (existingStudent && !String(recordData.studentId || "").trim()) {
        recordData.studentId = String(existingStudent.studentId || "");
      }
      if (!String(recordData.studentId || "").trim()) {
        recordData.studentId = nextStudentId_(students);
      }
      const currentActiveCount = activeStudentCount_(students);
      const wasActive = existingStudent ? isActiveStudent_(existingStudent) : false;
      if (isActiveStudent_(recordData) && !wasActive) {
        const limitError = studentLimitError_(tenant.plan, currentActiveCount);
        if (limitError) return { ok: false, error: limitError };
      }
      const duplicate = students.find(item =>
        String(item.id) !== id &&
        String(item.studentId || "").trim().toLowerCase() === String(recordData.studentId).trim().toLowerCase()
      );
      if (duplicate) return { ok: false, error: "That Student ID is already in use. Try saving again." };
      delete recordData.password;
      delete recordData.passport;
      const updated = writeSchoolRecord_(type, recordData);
      logAudit_(actor, updated ? "updated record" : "created record", type, id, existingStudent || null, recordData);
      return {
        ok: true,
        updated: updated,
        record: recordData,
        studentUsage: {
          activeStudents: activeStudentCount_(readSchoolRecords_("student")),
          studentLimit: studentLimitForPlan_(tenant.plan),
          campusCount: readSchoolRecords_("campus").length
        }
      };
    } finally {
      lock.releaseLock();
    }
  }

  delete recordData.password;
  delete recordData.passport;
  let updated;
  const campusLock = type === "campus" && campusMetadataChanged
    ? LockService.getScriptLock()
    : null;
  if (campusLock) campusLock.waitLock(10000);
  try {
    if (type === "campus" && campusMetadataChanged) {
      const currentCampuses = readSchoolRecords_("campus");
      const duplicate = currentCampuses.find(campus =>
        String(campus.id || "") !== String(recordData.id || "") &&
        String(campus.name || "").trim().toLowerCase() === String(recordData.name || "").trim().toLowerCase()
      );
      if (duplicate) return { ok: false, error: "A campus with that name already exists." };
      const currentCampus = currentCampuses.find(campus => String(campus.id || "") === String(recordData.id || ""));
      if (creatingCampus && currentCampus) {
        return { ok: false, error: "This campus ID was created by another request. Refresh and try again." };
      }
      if (!currentCampus && !creatingCampus) {
        return { ok: false, error: "This campus is no longer available. Refresh and try again." };
      }
      if (!currentCampus && campusPinToSet.length < 6) {
        return { ok: false, error: "Set a campus PIN of at least 6 characters before creating the campus." };
      }
      updated = writeSchoolRecord_(type, recordData);
    } else if (type === "campus") {
      updated = true;
    } else {
      updated = writeSchoolRecord_(type, recordData);
    }
    if (type === "campus" && campusPinToSet) {
      saveCampusPin_(tenant.tenantId, recordData.id, campusPinToSet);
    }
  } finally {
    if (campusLock) campusLock.releaseLock();
  }
  logAudit_(actor, updated ? "updated record" : "created record", type, id, oldRecord, recordData);
  if (type === "feePlan") {
    const feePlan = schoolFeePlans_(readSchoolRecords_("feePlan"), actor)
      .find(item => String(item.id) === id);
    return { ok: true, updated: updated, record: feePlan || recordData };
  }
  let notificationWarning = "";
  if (type === "attendance" && recordData.status === "Absent" &&
      (!oldRecord || String(oldRecord.status || "").toLowerCase() !== "absent")) {
    const student = readSchoolRecords_("student").find(item =>
      String(item.studentId || "").trim() === String(recordData.studentId || "").trim()
    );
    if (student) notificationWarning = notifyAttendanceAbsences_(
      [{ student: student, date: recordData.date }],
      actor,
      body.portalUrl
    );
  }
  if (type === "employee" && !accountSaved) {
    if (String(recordData.dashboardStatus || "").toLowerCase() === "active") {
      syncAccountFromSchoolRecord_(type, recordData, body.previousEmail);
    }
  } else if (type === "parent" && !accountSaved) {
    syncAccountFromSchoolRecord_(type, recordData, body.previousEmail);
  }
  const responseRecord = type === "campus"
    ? { ...recordData, pinConfigured: Boolean(campusPinState_(tenant.tenantId, recordData.id)) }
    : recordData;
  return { ok: true, updated: updated, record: responseRecord, ...(notificationWarning ? { notificationWarning: notificationWarning } : {}) };
}

function saveAttendanceRegister_(body) {
  const actor = body.authUser;
  const tenant = body.schoolTenant || {};
  const role = String(actor && (actor.role || actor.type) || "").toLowerCase();
  const teacherRecord = role === "employee" ? readSchoolRecords_("employee").find(employee =>
    String(employee.email || "").trim().toLowerCase() === String(actor.email || "").trim().toLowerCase()
  ) : null;
  const isTeacher = role === "employee" &&
    String(teacherRecord && teacherRecord.role || "").toLowerCase().includes("teacher");
  if (role !== "admin" && !isTeacher) {
    return { ok: false, error: "Only administrators and assigned teachers can save attendance." };
  }

  const className = String(body.className || "").trim();
  const campusId = isTeacher
    ? effectiveCampusId_(teacherRecord || actor)
    : validatedCampusId_(tenant, body.campusId || DEFAULT_CAMPUS_ID);
  if (!campusId) return { ok: false, error: "Select an active campus in this school workspace." };
  const days = Array.isArray(body.days)
    ? body.days
    : body.date && Array.isArray(body.students)
      ? [{ date: body.date, students: body.students }]
      : [];
  const timezone = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), timezone, "yyyy-MM-dd");
  if (isTeacher && className !== String(teacherRecord.className || "").trim()) {
    return { ok: false, error: "Teachers can only save attendance for their assigned class." };
  }
  if (!className || className.length > 100 || !Array.isArray(days) || days.length < 1 || days.length > 5) {
    return { ok: false, error: "Choose a class and at least one valid attendance day to save." };
  }

  const statuses = {
    present: "Present",
    absent: "Absent",
    late: "Late",
    excused: "Excused"
  };
  const requestedDays = [];
  const seenDates = new Set();
  for (const day of days) {
    const date = String(day && day.date || "").trim();
    if (!isWeekdayDateKey_(date) ||
        date > today || seenDates.has(date)) {
      return { ok: false, error: "The attendance register contains an invalid, weekend, future, or duplicate date." };
    }
    seenDates.add(date);
    if (!Array.isArray(day.students) || day.students.length < 1 || day.students.length > 500) {
      return { ok: false, error: `The register for ${date} must contain between 1 and 500 students.` };
    }
    const requested = new Map();
    for (const entry of day.students) {
      const studentId = String(entry && entry.studentId || "").trim();
      const status = String(entry && entry.status || "").trim().toLowerCase();
      if (!studentId || !Object.prototype.hasOwnProperty.call(statuses, status) || requested.has(studentId)) {
        return { ok: false, error: `The register for ${date} contains an invalid or duplicate student entry.` };
      }
      const expectedStatus = Object.prototype.hasOwnProperty.call(entry || {}, "expectedStatus")
        ? String(entry.expectedStatus || "").trim().toLowerCase()
        : null;
      if (expectedStatus !== null && expectedStatus && !Object.prototype.hasOwnProperty.call(statuses, expectedStatus)) {
        return { ok: false, error: `The register for ${date} contains an invalid previous status.` };
      }
      requested.set(studentId, {
        status: statuses[status],
        expectedStatus: expectedStatus === null ? null : expectedStatus ? statuses[expectedStatus] : ""
      });
    }
    requestedDays.push({ date: date, students: requested });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  const savedRecords = [];
  const absenceNotifications = [];
  let response;
  try {
    const students = readSchoolRecords_("student");
    const classStudents = new Map(students
      .filter(student =>
        String(student.className || "").trim() === className &&
        String(student.studentId || "").trim() &&
        isActiveStudent_(student) &&
        effectiveCampusId_(student) === campusId
      )
      .map(student => [String(student.studentId).trim(), student]));
    if (!classStudents.size) {
      return { ok: false, error: "This class has no students with valid student IDs." };
    }
    for (const day of requestedDays) {
      if (day.students.size !== classStudents.size ||
          [...classStudents.keys()].some(studentId => !day.students.has(studentId))) {
        return { ok: false, error: `Mark an attendance status for every student in ${className} on ${day.date}.` };
      }
    }

    const headersRequired = ["id", "date", "studentId", "studentName", "className", "campusId", "status", "present", "absent", "notes"];
    const sheet = ensureHeaders_(RECORD_SHEETS.attendance, headersRequired);
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(header => String(header).trim());
    const studentIdColumn = headers.indexOf("studentId");
    const dateColumn = headers.indexOf("date");
    const idColumn = headers.indexOf("id");
    const existingRows = values.slice(1);
    const rowByStudentDate = new Map();
    existingRows.forEach((row, index) => {
      const studentId = String(row[studentIdColumn] || "").trim();
      const rowDate = row[dateColumn] instanceof Date
        ? Utilities.formatDate(row[dateColumn], timezone, "yyyy-MM-dd")
        : String(row[dateColumn] || "").trim().slice(0, 10);
      if (studentId && rowDate && !rowByStudentDate.has(`${studentId}|${rowDate}`)) {
        rowByStudentDate.set(`${studentId}|${rowDate}`, index);
      }
    });

    const conflicts = [];
    requestedDays.forEach(day => {
      day.students.forEach((entry, studentId) => {
        if (entry.expectedStatus === null) return;
        const rowIndex = rowByStudentDate.get(`${studentId}|${day.date}`);
        const row = rowIndex === undefined ? null : existingRows[rowIndex];
        const currentStatus = row && headers.indexOf("status") >= 0
          ? String(row[headers.indexOf("status")] || "")
          : "";
        if (currentStatus !== entry.expectedStatus && currentStatus !== entry.status) {
          conflicts.push({
            date: day.date,
            studentId: studentId,
            expectedStatus: entry.expectedStatus,
            currentStatus: currentStatus,
            requestedStatus: entry.status
          });
        }
      });
    });
    if (conflicts.length) {
      return { ok: true, records: [], conflicts: conflicts };
    }

    const appendedRows = [];
    requestedDays.forEach(day => {
      day.students.forEach((entry, studentId) => {
        const status = entry.status;
        const student = classStudents.get(studentId);
        const key = `${studentId}|${day.date}`;
        const rowIndex = rowByStudentDate.get(key);
        const previousRow = rowIndex === undefined ? null : existingRows[rowIndex];
        const previousStatus = previousRow && headers.indexOf("status") >= 0
          ? String(previousRow[headers.indexOf("status")] || "").toLowerCase()
          : "";
        const present = status === "Present";
        const absent = status === "Absent";
        if (absent && previousStatus !== "absent") {
          absenceNotifications.push({ student: student, date: day.date });
        }
        const record = {
          id: previousRow && previousRow[idColumn] || `attendance-${studentId}-${day.date}`,
          date: day.date,
          studentId: studentId,
          studentName: `${student.firstName || ""} ${student.lastName || ""}`.trim(),
          className: className,
          campusId: campusId,
          status: status,
          present: present ? 1 : 0,
          absent: absent ? 1 : 0,
          notes: previousRow ? previousRow[headers.indexOf("notes")] || "" : ""
        };
        const row = headers.map((header, index) =>
          Object.prototype.hasOwnProperty.call(record, header)
            ? record[header]
            : previousRow ? previousRow[index] : ""
        );
        if (rowIndex === undefined) {
          appendedRows.push(row);
          rowByStudentDate.set(key, existingRows.length + appendedRows.length - 1);
        } else {
          existingRows[rowIndex] = row;
        }
        savedRecords.push(record);
      });
    });

    if (existingRows.length) {
      setSafeSheetValues_(sheet.getRange(2, 1, existingRows.length, headers.length), existingRows);
    }
    if (appendedRows.length) {
      writeSafeRows_(sheet, existingRows.length + 2, appendedRows);
    }
    logAudit_(actor, "saved attendance register", "attendance", className, null, {
      className: className,
      dates: requestedDays.map(day => day.date),
      records: savedRecords.length
    });
    response = { ok: true, records: savedRecords };
  } finally {
    lock.releaseLock();
  }
  const notificationWarning = notifyAttendanceAbsences_(absenceNotifications, actor, body.portalUrl);
  if (notificationWarning) response.notificationWarning = notificationWarning;
  return response;
}

function saveEmployeeAttendance_(body) {
  const actor = body.authUser || {};
  const role = String(actor.role || actor.type || "").trim().toLowerCase();
  const actorEmail = String(actor.email || "").trim().toLowerCase();
  const operation = String(body.operation || "").trim();
  if (!actorEmail || !["admin", "employee"].includes(role)) {
    return { ok: false, error: "Only an employee or school administrator can manage employee attendance." };
  }
  if (!["checkIn", "checkOut", "markAbsent"].includes(operation)) {
    return { ok: false, error: "Choose a valid employee attendance action." };
  }
  if ((operation === "markAbsent" && role !== "admin") ||
      (["checkIn", "checkOut"].includes(operation) && role !== "employee")) {
    return { ok: false, error: "You do not have permission to perform this attendance action." };
  }

  const tenant = body.schoolTenant || {};
  const timezone = Session.getScriptTimeZone();
  const now = new Date();
  const today = Utilities.formatDate(now, timezone, "yyyy-MM-dd");
  const date = operation === "markAbsent" ? String(body.date || "").trim() : today;
  const dateParts = date.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const parsedDate = dateParts
    ? new Date(Date.UTC(Number(dateParts[1]), Number(dateParts[2]) - 1, Number(dateParts[3])))
    : null;
  if (!parsedDate ||
      parsedDate.getUTCFullYear() !== Number(dateParts[1]) ||
      parsedDate.getUTCMonth() !== Number(dateParts[2]) - 1 ||
      parsedDate.getUTCDate() !== Number(dateParts[3]) ||
      date > today) {
    return { ok: false, error: "Choose a valid employee attendance date that is not in the future." };
  }

  const employees = readSchoolRecords_("employee");
  const targetEmail = operation === "markAbsent"
    ? String(body.employeeEmail || "").trim().toLowerCase()
    : actorEmail;
  const employee = employees.find(record =>
    String(record.email || "").trim().toLowerCase() === targetEmail
  );
  const employeeAccount = targetEmail ? findUserByEmail_(targetEmail) : null;
  if (!employee || !employeeAccount || !(employeeAccount.passwordHash || employeeAccount.password) ||
      String(employeeAccount.role || employeeAccount.type || "").trim().toLowerCase() !== "employee") {
    return { ok: false, error: "An active employee account and profile are required to record attendance." };
  }

  const campusId = role === "admin"
    ? validatedCampusId_(tenant, body.campusId || DEFAULT_CAMPUS_ID)
    : effectiveCampusId_(employee);
  if (!campusId) return { ok: false, error: "Select an active campus in this school workspace." };
  if (effectiveCampusId_(employee) !== campusId) {
    return { ok: false, error: "That employee is not assigned to the selected campus." };
  }
  if (operation === "checkIn" &&
      !["present", "late"].includes(String(body.status || "").trim().toLowerCase())) {
    return { ok: false, error: "Choose Present or Late before checking in." };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (["checkIn", "checkOut"].includes(operation) &&
        !isRegisteredEmployeeAttendanceKiosk_(tenant, campusId, body.kioskToken)) {
      return { ok: false, error: "Check in or check out using your school's registered attendance kiosk." };
    }
    const matches = readSchoolRecords_("employeeAttendance").filter(record =>
      String(record.employeeEmail || "").trim().toLowerCase() === targetEmail &&
      String(record.date || "").slice(0, 10) === date
    );
    if (matches.length > 1) {
      return { ok: false, error: "Duplicate attendance records exist for this employee and date. Ask the administrator to resolve them." };
    }

    const existing = matches[0] || null;
    if (operation === "checkIn" && existing && existing.clockInAt) {
      return { ok: false, error: "You have already checked in today." };
    }
    if (operation === "checkOut" && (!existing || !existing.clockInAt)) {
      return { ok: false, error: "Check in before recording your check-out." };
    }
    if (operation === "checkOut" && existing.clockOutAt) {
      return { ok: false, error: "You have already checked out today." };
    }
    if (operation === "markAbsent" && existing &&
        (existing.clockInAt || ["present", "late"].includes(String(existing.status || "").toLowerCase()))) {
      return { ok: false, error: "An employee who has checked in cannot be marked absent." };
    }
    if (operation === "markAbsent" && existing &&
        String(existing.status || "").toLowerCase() === "absent") {
      return { ok: true, record: existing, unchanged: true };
    }

    const before = existing ? { ...existing } : null;
    const record = {
      ...(existing || {}),
      id: existing ? existing.id : Utilities.getUuid(),
      employeeEmail: targetEmail,
      employeeName: String(employee.name || employeeAccount.name || targetEmail),
      date: date,
      campusId: campusId,
      status: operation === "markAbsent"
        ? "Absent"
        : operation === "checkIn"
          ? String(body.status).trim().toLowerCase() === "late" ? "Late" : "Present"
          : existing.status,
      clockInAt: operation === "checkIn" ? now.toISOString() : String(existing && existing.clockInAt || ""),
      clockOutAt: operation === "checkOut" ? now.toISOString() : String(existing && existing.clockOutAt || ""),
      clockInTime: operation === "checkIn"
        ? Utilities.formatDate(now, timezone, "HH:mm")
        : String(existing && existing.clockInTime || ""),
      clockOutTime: operation === "checkOut"
        ? Utilities.formatDate(now, timezone, "HH:mm")
        : String(existing && existing.clockOutTime || ""),
      markedAbsentByEmail: operation === "markAbsent"
        ? actorEmail
        : operation === "checkIn" ? "" : String(existing && existing.markedAbsentByEmail || ""),
      updatedAt: now.toISOString()
    };
    if (!existing) record.createdAt = now.toISOString();
    writeSchoolRecord_("employeeAttendance", record);
    logAudit_(actor, `employee attendance ${operation}`, "employeeAttendance", record.id, before, record,
      `${record.employeeName} · ${date} · ${record.status}`);
    return { ok: true, record: record };
  } finally {
    lock.releaseLock();
  }
}

function employeeAttendanceKioskPropertyKey_(tenantId, campusId) {
  const tenant = String(tenantId || "").trim();
  const campus = String(campusId || "").trim();
  if (!tenant || !campus) return "";
  return `EMPLOYEE_ATTENDANCE_KIOSK:${hashAdmissionToken_(`${tenant}:${campus}`)}`;
}

function isRegisteredEmployeeAttendanceKiosk_(tenant, campusId, token) {
  const value = String(token || "").trim();
  if (value.length < 32 || value.length > 120) return false;
  const propertyKey = employeeAttendanceKioskPropertyKey_(tenant && tenant.tenantId, campusId);
  if (!propertyKey) return false;
  const storedHash = PropertiesService.getScriptProperties().getProperty(propertyKey);
  return Boolean(storedHash) && hashAdmissionToken_(value) === storedHash;
}

function manageEmployeeAttendanceKiosk_(body, operation) {
  const actor = body.authUser || {};
  if (String(actor.role || actor.type || "").trim().toLowerCase() !== "admin") {
    return { ok: false, error: "Only a school administrator can manage the attendance kiosk." };
  }
  if (!["register", "revoke"].includes(operation)) {
    return { ok: false, error: "Choose a valid attendance kiosk action." };
  }
  const tenant = body.schoolTenant || {};
  const campusId = validatedCampusId_(tenant, body.campusId || DEFAULT_CAMPUS_ID);
  const propertyKey = employeeAttendanceKioskPropertyKey_(tenant.tenantId, campusId);
  if (!propertyKey) return { ok: false, error: "Select an active campus before managing its attendance kiosk." };
  const properties = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (operation === "revoke") {
      properties.deleteProperty(propertyKey);
      logAudit_(actor, "revoked employee attendance kiosk", "employeeAttendanceKiosk", campusId, null, null);
      return { ok: true, campusId: campusId };
    }
    const token = newAdmissionToken_();
    properties.setProperty(propertyKey, hashAdmissionToken_(token));
    logAudit_(actor, "registered employee attendance kiosk", "employeeAttendanceKiosk", campusId, null, null);
    return { ok: true, campusId: campusId, kioskToken: token };
  } finally {
    lock.releaseLock();
  }
}

function saveEmployeeLeave_(body) {
  const actor = body.authUser || {};
  const actorEmail = String(actor.email || "").trim().toLowerCase();
  if (String(actor.role || actor.type || "").toLowerCase() !== "employee" || !actorEmail) {
    return { ok: false, error: "Only an employee can submit a leave request." };
  }
  const employee = readSchoolRecords_("employee").find(record =>
    String(record.email || "").trim().toLowerCase() === actorEmail
  );
  if (!employee) return { ok: false, error: "Your employee profile could not be found. Contact the school administrator." };

  const startDate = String(body.startDate || "").trim();
  const endDate = String(body.endDate || "").trim();
  const reason = String(body.reason || "").trim();
  if (!isValidDateKey_(startDate) || !isValidDateKey_(endDate) || endDate < startDate) {
    return { ok: false, error: "Enter valid leave dates with the end date on or after the start date." };
  }
  if (!reason || reason.length > 2000) {
    return { ok: false, error: "Enter a leave reason up to 2,000 characters." };
  }
  const durationDays = (Date.parse(`${endDate}T00:00:00Z`) - Date.parse(`${startDate}T00:00:00Z`)) / 86400000;
  if (durationDays > 365) return { ok: false, error: "A leave request cannot exceed 366 calendar days." };

  const now = new Date().toISOString();
  const request = {
    id: Utilities.getUuid(),
    employeeEmail: actorEmail,
    employeeName: String(employee.name || actor.name || actorEmail),
    startDate: startDate,
    endDate: endDate,
    reason: reason,
    status: "Pending",
    campusId: effectiveCampusId_(employee),
    requestedAt: now,
    reviewedAt: "",
    reviewedByEmail: "",
    reviewNotes: ""
  };
  writeSchoolRecord_("employeeLeave", request);
  logAudit_(actor, "submitted employee leave request", "employeeLeave", request.id, null, request);

  const tenant = body.schoolTenant || {};
  const adminEmail = String(tenant.adminEmail || "").trim().toLowerCase();
  if (adminEmail) {
    const sheet = ensureHeaders_(MESSAGES, messageSheetHeaders_());
    appendRecord_(sheet, {
      id: Utilities.getUuid(),
      audience: "Specific",
      recipient: adminEmail,
      subject: `Leave request from ${request.employeeName}`,
      message: `${request.employeeName} requested leave from ${startDate} to ${endDate}.\n\nReason: ${reason}`,
      sendDate: new Date(),
      sendTime: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm"),
      sentBy: request.employeeName,
      sentByEmail: actorEmail,
      createdAt: new Date(),
      actionView: "employeeLeave",
      replyToId: "",
      recipientType: "admin",
      recipientId: adminEmail,
      editedAt: "",
      sentByRole: "employee",
      messageType: "direct"
    });
  }
  return { ok: true, request: request };
}

function reviewEmployeeLeave_(body) {
  const actor = body.authUser || {};
  const actorEmail = String(actor.email || "").trim().toLowerCase();
  if (String(actor.role || actor.type || "").toLowerCase() !== "admin" || !actorEmail) {
    return { ok: false, error: "Only a school administrator can review leave requests." };
  }
  const status = String(body.status || "").trim();
  if (!["Approved", "Declined"].includes(status)) {
    return { ok: false, error: "Choose Approve or Decline for this leave request." };
  }
  const requestId = String(body.requestId || "").trim();
  const request = readSchoolRecords_("employeeLeave").find(record => String(record.id) === requestId);
  if (!request) return { ok: false, error: "This leave request is no longer available. Refresh and try again." };
  if (String(request.status || "") !== "Pending") {
    return { ok: false, error: "This leave request has already been reviewed. Refresh the list." };
  }
  const tenant = body.schoolTenant || {};
  const campusId = validatedCampusId_(tenant, body.campusId || DEFAULT_CAMPUS_ID);
  if (!campusId || effectiveCampusId_(request) !== campusId) {
    return { ok: false, error: "This leave request is not in the active campus." };
  }
  const before = { ...request };
  request.status = status;
  request.reviewedAt = new Date().toISOString();
  request.reviewedByEmail = actorEmail;
  request.reviewNotes = String(body.reviewNotes || "").trim().slice(0, 1000);
  writeSchoolRecord_("employeeLeave", request);
  logAudit_(actor, `${status.toLowerCase()} employee leave request`, "employeeLeave", request.id, before, request);
  appendInboxMessage_(
    request.employeeEmail,
    `Leave request ${status.toLowerCase()}`,
    `Your leave request for ${request.startDate} to ${request.endDate} was ${status.toLowerCase()} by the school administrator.${request.reviewNotes ? `\n\nNote: ${request.reviewNotes}` : ""}`,
    "School",
    "employeeLeave"
  );
  return { ok: true, request: request };
}

function saveTimetableRecord_(actor, record, body) {
  const actorRole = String(actor.role || actor.type || "").toLowerCase();
  const actorEmail = String(actor.email || "").trim().toLowerCase();
  const isAdmin = actorRole === "admin";
  const employee = actorRole === "employee"
    ? readSchoolRecords_("employee").find(item =>
      String(item.email || "").trim().toLowerCase() === actorEmail
    )
    : null;
  const isTeacher = Boolean(employee &&
    String(employee.role || employee.jobRole || "").toLowerCase().includes("teacher"));
  if (!isAdmin && !isTeacher) {
    return { ok: false, error: "Only a school administrator or assigned teacher can manage timetable lessons." };
  }
  const days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
  const day = String(record.day || "").trim();
  let className = String(record.className || "").trim();
  const subject = String(record.subject || "").trim();
  const startTime = String(record.startTime || "").trim();
  const endTime = String(record.endTime || "").trim();
  const room = String(record.room || "").trim();
  const id = String(record.id || "").trim();
  const validTime = value => /^([01]\d|2[0-3]):[0-5]\d$/.test(value);
  if (isTeacher) {
    const assignedClass = String(employee.className || "").trim();
    const assignedSubjects = String(employee.subjects || "").split(/[,;\n]/)
      .map(value => value.trim().toLowerCase()).filter(Boolean);
    if (!assignedClass) {
      return { ok: false, error: "Ask your administrator to assign your class before entering timetable lessons." };
    }
    if (className.toLowerCase() !== assignedClass.toLowerCase()) {
      return { ok: false, error: "You can only enter timetable lessons for your assigned class." };
    }
    if (!assignedSubjects.includes(subject.toLowerCase())) {
      return { ok: false, error: "You can only enter timetable lessons for your assigned subjects." };
    }
    className = assignedClass;
    record.className = assignedClass;
    record.teacherEmail = actorEmail;
    record.campusId = effectiveCampusId_(employee);
  }
  if (!id || !days.includes(day) || !className || !subject || !validTime(startTime) || !validTime(endTime) || startTime >= endTime) {
    return { ok: false, error: "Enter a valid day, class, subject, and start/end times. End time must be after start time." };
  }
  record.day = day;
  record.className = className.slice(0, 100);
  record.subject = subject.slice(0, 120);
  record.room = room.slice(0, 100);
  record.startTime = startTime;
  record.endTime = endTime;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let previousRecord = null;
  let response = null;
  try {
    const employees = readSchoolRecords_("employee");
    const teachers = employees.filter(employee =>
      String(employee.role || employee.jobRole || "").toLowerCase().includes("teacher") &&
      String(employee.email || "").trim() &&
      effectiveCampusId_(employee) === effectiveCampusId_(record)
    );
    const savedLessons = readSchoolRecords_("timetable");
    previousRecord = savedLessons.find(item => String(item.id) === id) || null;
    if (isTeacher && previousRecord &&
        (String(previousRecord.teacherEmail || "").trim().toLowerCase() !== actorEmail ||
         String(previousRecord.createdByEmail || "").trim().toLowerCase() !== actorEmail)) {
      return { ok: false, error: "You can only edit timetable lessons that you created for your account." };
    }
    const isUpdate = Boolean(previousRecord);
    const lessons = savedLessons.filter(item => String(item.id) !== id);
    const overlaps = lesson => String(lesson.day) === day &&
      startTime < String(lesson.endTime || "") &&
      endTime > String(lesson.startTime || "");
    const unavailable = teacher => lessons.some(lesson =>
      String(lesson.teacherEmail || "").toLowerCase() === String(teacher.email).toLowerCase() &&
      overlaps(lesson)
    );
    const requestedEmail = isTeacher ? actorEmail : String(record.teacherEmail || "").trim().toLowerCase();
    let assigned = requestedEmail
      ? teachers.find(teacher => String(teacher.email).trim().toLowerCase() === requestedEmail)
      : null;
    if (requestedEmail && !assigned) return { ok: false, error: "Select a valid teacher from the school employee records." };
    if (!assigned) {
      const eligible = teachers
        .filter(teacher => !teacher.className || String(teacher.className).trim().toLowerCase() === className.toLowerCase())
        .filter(teacher => !teacher.subjects || String(teacher.subjects).toLowerCase().includes(subject.toLowerCase()))
        .filter(teacher => !unavailable(teacher))
        .sort((left, right) => {
          const leftMatch = String(left.className || "").trim().toLowerCase() === className.toLowerCase() ? 0 : 1;
          const rightMatch = String(right.className || "").trim().toLowerCase() === className.toLowerCase() ? 0 : 1;
          return leftMatch - rightMatch || String(left.name || "").localeCompare(String(right.name || ""));
        })[0];
      assigned = eligible || null;
      if (!assigned) return { ok: false, error: "No available teacher matches this class and time. Assign a teacher or resolve existing schedule conflicts." };
    }

    const conflict = lessons.find(lesson => overlaps(lesson) && (
      String(lesson.teacherEmail || "").trim().toLowerCase() === String(assigned.email).trim().toLowerCase() ||
      effectiveCampusId_(lesson) === effectiveCampusId_(record) && (
        String(lesson.className || "").trim().toLowerCase() === className.toLowerCase() ||
        (room && String(lesson.room || "").trim().toLowerCase() === room.toLowerCase())
      )
    ));
    if (conflict) {
      const reason = String(conflict.teacherEmail || "").trim().toLowerCase() === String(assigned.email).trim().toLowerCase()
        ? `Teacher ${assigned.name || assigned.email}`
        : String(conflict.className || "").trim().toLowerCase() === className.toLowerCase()
        ? `Class ${className}`
        : `Room ${room}`;
      return { ok: false, error: `${reason} already has a lesson at that time (${conflict.day} ${conflict.startTime}–${conflict.endTime}).` };
    }

    record.teacherEmail = String(assigned.email).trim().toLowerCase();
    record.teacherName = String(assigned.name || assigned.email).trim();
    const now = new Date().toISOString();
    record.updatedAt = now;
    if (isTeacher) {
      record.createdByEmail = previousRecord && previousRecord.createdByEmail || actorEmail;
      record.createdAt = previousRecord && previousRecord.createdAt || now;
      record.updatedByEmail = actorEmail;
    }
    const updated = writeSchoolRecord_("timetable", record);
    const messages = ensureHeaders_(MESSAGES, [
      "id", "audience", "recipient", "subject", "message",
      "sendDate", "sentBy", "sentByEmail", "createdAt", "actionView",
      "replyToId", "recipientType", "recipientId", "editedAt", "sentByRole", "messageType"
    ]);
    appendRecord_(messages, {
      id: Utilities.getUuid(),
      audience: "Specific employee",
      recipient: record.teacherEmail,
      subject: `Timetable ${isUpdate ? "updated" : "scheduled"} · ${record.subject}`,
      message: `${record.day}, ${record.startTime}–${record.endTime} · ${record.className}${record.room ? ` · Room ${record.room}` : ""}`,
      sendDate: `${record.day} ${record.startTime}`,
      sentBy: actor.name || "School Admin",
      sentByEmail: actor.email || "",
      createdAt: new Date(),
      actionView: "timetable",
      sentByRole: "system",
      messageType: "system"
    });
    logAudit_(actor, updated ? "updated timetable" : "created timetable", "timetable", id, previousRecord, record);
    response = { ok: true, updated: updated, record: record };
  } finally {
    lock.releaseLock();
  }
  try {
    const notificationWarning = notifyTimetableParents_(record, previousRecord, actor, body && body.portalUrl);
    if (notificationWarning) response.notificationWarning = notificationWarning;
  } catch (error) {
    console.error(`Timetable was saved, but parent notifications could not be processed: ${error.message}`);
    response.notificationWarning = "Timetable saved, but parent email and inbox delivery status could not be confirmed.";
  }
  return response;
}

function notifyTimetableParents_(record, previousRecord, actor, portalUrl) {
  const changedFields = ["day", "startTime", "endTime", "className", "subject", "room", "teacherEmail"];
  if (previousRecord && !changedFields.some(field =>
    String(previousRecord[field] || "") !== String(record[field] || "")
  )) return "";
  const students = readSchoolRecords_("student").filter(student =>
    isActiveStudent_(student) &&
    String(student.className || "").trim() === String(record.className || "").trim() &&
    effectiveCampusId_(student) === effectiveCampusId_(record)
  );
  const parents = readSchoolRecords_("parent");
  const recipients = new Map();
  students.forEach(student => parentContactsForStudent_(student, parents).forEach(parent => {
    if (!recipients.has(parent.email)) recipients.set(parent.email, { name: parent.name, students: [] });
    recipients.get(parent.email).students.push(`${student.firstName || ""} ${student.lastName || ""}`.trim());
  }));
  if (!recipients.size) return "";
  const safePortalUrl = /^https:\/\/[a-z0-9.-]+(?::\d+)?(?:\/[^\s]*)?$/i.test(String(portalUrl || "").trim())
    ? String(portalUrl).trim()
    : "";
  const isUpdate = Boolean(previousRecord);
  const detail = `${record.day}, ${record.startTime}–${record.endTime} · ${record.subject}${record.room ? ` · Room ${record.room}` : ""}`;
  let emailSent = 0;
  let optedOut = 0;
  let failed = 0;
  let inboxFailed = 0;
  recipients.forEach((recipient, email) => {
    const children = [...new Set(recipient.students)].filter(Boolean).join(", ");
    const message = `The ${record.className} timetable was ${isUpdate ? "updated" : "scheduled"} for ${detail}.${children ? ` This applies to ${children}.` : ""}`;
    try {
      appendInboxMessage_(email, `Timetable ${isUpdate ? "updated" : "scheduled"} · ${record.subject}`, message, actor.name || "School", "inbox");
    } catch (error) {
      inboxFailed++;
      console.error(`Could not save a timetable update for ${email}: ${error.message}`);
    }
    if (!parentEmailPreferenceEnabled_(email, "emailTimetableNotifications")) {
      optedOut++;
      return;
    }
    try {
      MailApp.sendEmail({
        to: email,
        subject: `Timetable ${isUpdate ? "updated" : "scheduled"} · ${record.className}`,
        body: `${message}\nSign in to your parent portal inbox for this update.${safePortalUrl ? `\n${safePortalUrl}` : ""}`
      });
      emailSent++;
    } catch (error) {
      failed++;
      console.error(`Timetable was saved, but parent email delivery failed for ${email}: ${error.message}`);
    }
  });
  if (inboxFailed || failed || optedOut) {
    return `Timetable saved. Email sent to ${emailSent} parent(s); ${optedOut} have email updates off and ${failed} email(s) failed. ${inboxFailed ? `${inboxFailed} parent inbox message(s) also failed.` : "Inbox fallbacks were saved where available."}`;
  }
  return `Timetable saved and emailed to ${emailSent} parent(s); inbox copies are available.`;
}

function deleteRecord_(body) {
  const actor = body.authUser;
  const type = String(body.type || "");
  const id = String(body.id || "");
  if (!actor || String(actor.role || "").toLowerCase() !== "admin") {
    return { ok: false, error: "Only a school administrator can delete records." };
  }
  if (!RECORD_SHEETS[type] || !id) return { ok: false, error: "The record to delete is invalid." };
  if (type === "campus") {
    if (id === DEFAULT_CAMPUS_ID) return { ok: false, error: "The main campus cannot be deleted." };
    const assignedRecords = Object.keys(RECORD_SHEETS)
      .filter(recordType => !["campus", "parent", "feePlan"].includes(recordType))
      .some(recordType => readSchoolRecords_(recordType).some(record =>
        effectiveCampusId_(record) === id
      ));
    if (assignedRecords) {
      return { ok: false, error: "This campus still has school records assigned to it. Move or remove those records before deleting the campus." };
    }
  }
  const recordSheet = ensureHeaders_(RECORD_SHEETS[type], ["id"]);
  const recordValues = recordSheet.getDataRange().getValues();
  const recordIdColumn = recordValues[0].indexOf("id");
  const recordRowIndex = recordValues.findIndex((row, index) => index > 0 && String(row[recordIdColumn]) === id);
  let record = null;
  const duplicateFeePlans = [];
  if (type === "feeInvoice") {
    const invoice = readSchoolRecords_("feeInvoice").find(item => String(item.id) === id);
    if (invoice && readRows_(PAYMENTS).concat(readSchoolRecords_("fee")).some(payment => String(payment.invoiceId || "") === id)) {
      return { ok: false, error: "An invoice with recorded payments cannot be deleted. Keep it for the financial audit trail." };
    }
  }
  if (recordRowIndex >= 0) {
    record = rowToObject_(recordValues[0], recordValues[recordRowIndex]);
    if (String(actor.role || actor.type || "").toLowerCase() === "admin" &&
        !isSchoolOwnerAdmin_(actor, body.schoolTenant || {})) {
      if (type === "parent") {
        return { ok: false, error: "Only the school owner can delete a parent account shared across campuses." };
      }
      let recordCampusId = effectiveCampusId_(record);
      if (["fee", "feeInvoice", "gradeEntry", "reportCard", "attendance", "followUp"].includes(type) &&
          record.studentId) {
        const student = readSchoolRecords_("student").find(item =>
          String(item.studentId || "").trim() === String(record.studentId || "").trim()
        );
        if (!student) return { ok: false, error: "The linked student could not be found; this record cannot be deleted." };
        recordCampusId = effectiveCampusId_(student);
      }
      if (recordCampusId !== String(actor.campusId || "").trim()) {
        return { ok: false, error: "You are not authorized to delete records for another campus." };
      }
    }
    recordSheet.deleteRow(recordRowIndex + 1);
    if (type === "feePlan") {
      const feePlanName = String(record.name || "").trim().toLowerCase();
      const currentValues = recordSheet.getDataRange().getValues();
      const headers = currentValues[0].map(header => String(header).trim());
      const nameColumn = headers.indexOf("name");
      if (nameColumn >= 0 && feePlanName) {
        const duplicateRowIndexes = currentValues
          .map((row, index) => ({ row: row, index: index }))
          .filter(item => item.index > 0 &&
            String(item.row[nameColumn] || "").trim().toLowerCase() === feePlanName)
          .sort((a, b) => b.index - a.index);
        duplicateRowIndexes.forEach(item => {
          duplicateFeePlans.push(rowToObject_(headers, item.row));
          recordSheet.deleteRow(item.index + 1);
        });
      }
    }
  } else if (type === "fee") {
    const paymentSheet = ensureHeaders_(PAYMENTS, ["id"]);
    const paymentValues = paymentSheet.getDataRange().getValues();
    const paymentIdColumn = paymentValues[0].indexOf("id");
    const paymentRowIndex = paymentValues.findIndex((row, index) => index > 0 && String(row[paymentIdColumn]) === id);
    if (paymentRowIndex >= 0) {
      record = rowToObject_(paymentValues[0], paymentValues[paymentRowIndex]);
      if (String(actor.role || actor.type || "").toLowerCase() === "admin" &&
          !isSchoolOwnerAdmin_(actor, body.schoolTenant || {})) {
        const student = readSchoolRecords_("student").find(item =>
          String(item.studentId || "").trim() === String(record.studentId || "").trim()
        );
        if (!student || effectiveCampusId_(student) !== String(actor.campusId || "").trim()) {
          return { ok: false, error: "You are not authorized to delete payments for another campus." };
        }
      }
      paymentSheet.deleteRow(paymentRowIndex + 1);
    }
  }
  if (!record) return { ok: false, error: "Record not found. Refresh and try again." };
  logAudit_(actor, "deleted record", type, id, record, null);
  if (duplicateFeePlans.length) {
    logAudit_(
      actor,
      "deleted duplicate school-wide fee types",
      "feePlan",
      id,
      duplicateFeePlans,
      null,
      `Deleted ${duplicateFeePlans.length} duplicate campus fee type record(s) with the same school-wide name.`
    );
  }
  if (type === "fee" && record.invoiceId) updateFeeInvoiceStatus_(record.invoiceId, actor);
  if (type === "employee" || type === "parent") deleteSchoolUser_(record.email, type);
  if (type === "student") {
    const tenant = body.schoolTenant || {};
    return {
      ok: true,
      studentUsage: {
        activeStudents: activeStudentCount_(readSchoolRecords_("student")),
        studentLimit: studentLimitForPlan_(tenant.plan),
        campusCount: readSchoolRecords_("campus").length
      }
    };
  }
  return { ok: true };
}

function saveUser_(body) {
  const actor = body.authUser;
  if (!actor || String(actor.role || "").toLowerCase() !== "admin") {
    return { ok: false, error: "Only an administrator can create accounts." };
  }

  const user = body.user;
  if (!user || !user.email || !user.name) {
    return { ok: false, error: "Name and email are required." };
  }

  const sheet = ensureHeaders_(USERS, [
    "id", "name", "email", "password", "passwordHash", "passwordSalt", "role", "type",
    "jobRole", "department", "className", "childStudentId", "profilePicture", "sessionVersion", "campusId",
    "phone", "phoneCountry", "address", "relationship", "occupation",
    "emailAbsenceNotifications", "emailPaymentNotifications",
    "emailTimetableNotifications", "emailReportCardNotifications"
  ]);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(header => String(header).trim());
  const emailColumn = headers.indexOf("email");
  if (emailColumn < 0) {
    throw new Error("The Users sheet is missing its email column. Run setup() and try again.");
  }
  const email = String(user.email).toLowerCase().trim();
  const previousEmail = String(body.previousEmail || email).toLowerCase().trim();
  const isAccountUpdate = Boolean(String(body.previousEmail || "").trim());
  const password = String(user.password || "");
  const requestedRole = String(user.role || user.type || "employee").toLowerCase();
  if (!["employee", "parent"].includes(requestedRole)) {
    return { ok: false, error: "Only employee and parent accounts can be managed through this action." };
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { ok: false, error: "Enter a valid email address." };
  }
  if (requestedRole === "employee" && password) {
    return { ok: false, error: "Employees must set their own password using a school invitation link." };
  }
  if (password && password.length < 12) {
    return { ok: false, error: "New employee and parent passwords must be at least 12 characters." };
  }
  const allUsers = values.slice(1).map(row =>
    Object.fromEntries(headers.map((header, index) => [header, row[index]]))
  );
  const conflictingUser = allUsers.find(row =>
    String(row.email || "").toLowerCase().trim() === email &&
    (!isAccountUpdate || String(row.email || "").toLowerCase().trim() !== previousEmail)
  );
  if (conflictingUser) {
    return { ok: false, error: "That email address already belongs to another school account." };
  }
  const existingUser = allUsers.find(row =>
    String(row.email || "").toLowerCase().trim() === email ||
    String(row.email || "").toLowerCase().trim() === previousEmail
  );
  if (!password && !existingUser) {
    return { ok: false, error: "A password of at least 12 characters is required for a new account." };
  }
  const salt = password ? Utilities.getUuid() : existingUser.passwordSalt;
  const rowValues = {
    id: existingUser && existingUser.id || user.id || email,
    name: user.name,
    email: email,
    password: "",
    passwordHash: password ? hashPassword_(password, salt) : existingUser.passwordHash,
    passwordSalt: salt,
    sessionVersion: (Number(existingUser && existingUser.sessionVersion) || 0) + (password && existingUser ? 1 : 0),
    role: user.role || "employee",
    type: user.type || user.role || "employee",
    jobRole: user.jobRole || "",
    department: user.department || "",
    className: user.className || "",
    campusId: user.campusId || existingUser && existingUser.campusId || "",
    childStudentId: user.childStudentId || "",
    profilePicture: user.profilePicture || existingUser && existingUser.profilePicture || "",
    phone: user.phone || existingUser && existingUser.phone || "",
    phoneCountry: user.phoneCountry || existingUser && existingUser.phoneCountry || "",
    address: user.address || existingUser && existingUser.address || "",
    relationship: user.relationship || existingUser && existingUser.relationship || "",
    occupation: user.occupation || existingUser && existingUser.occupation || "",
    emailAbsenceNotifications: existingUser && String(existingUser.emailAbsenceNotifications).toLowerCase() === "true",
    emailPaymentNotifications: existingUser && String(existingUser.emailPaymentNotifications).toLowerCase() === "true",
    emailTimetableNotifications: existingUser && String(existingUser.emailTimetableNotifications).toLowerCase() === "true",
    emailReportCardNotifications: existingUser && String(existingUser.emailReportCardNotifications).toLowerCase() === "true",
  };

  let updated = false;
  for (let row = 1; row < values.length; row++) {
    if ([email, previousEmail].includes(String(values[row][emailColumn] || "").toLowerCase().trim())) {
      writeRecord_(sheet, row + 1, rowValues, headers);
      updated = true;
    }
  }

  if (previousEmail !== email && existingUser) {
    deleteSchoolUser_(previousEmail, String(existingUser.role || user.role || ""));
  }
  if (updated) {
    return { ok: true, updated: true };
  }

  appendRecord_(sheet, rowValues);
  return { ok: true, created: true };
}

function updateProfile_(body) {
  const actor = body.authUser;
  const picture = String(body.profilePicture || "");
  const hasPicture = /^data:image\/(?:jpeg|png);base64,/i.test(picture) && picture.length <= 45000;
  const hasProfileDetails = body.profileDetails && typeof body.profileDetails === "object" && !Array.isArray(body.profileDetails);
  const hasBackground = Object.prototype.hasOwnProperty.call(body, "portalBackgroundImage");
  const background = String(body.portalBackgroundImage || "");
  const hasPortalLogo = Object.prototype.hasOwnProperty.call(body, "portalLogo");
  const portalLogo = String(body.portalLogo || "");
  const hasBrandColor = Object.prototype.hasOwnProperty.call(body, "brandColor");
  const brandColor = String(body.brandColor || "");
  const role = String(actor && (actor.role || actor.type) || "").toLowerCase();
  const isAdmin = role === "admin";
  const tenant = body.schoolTenant || {};
  if (!actor || (!hasPicture && !hasProfileDetails && !hasBackground && !hasPortalLogo && !hasBrandColor)) {
    return { ok: false, error: "A signed-in user and profile or portal setting are required." };
  }
  if (isAdmin && (hasPicture || hasPortalLogo || hasBrandColor || hasBackground) &&
      !isSchoolOwnerAdmin_(actor, tenant)) {
    return { ok: false, error: "Only the main school administrator can change school-wide profile and branding settings." };
  }
  if (picture && !hasPicture) return { ok: false, error: "The profile image must be a JPG or PNG image under 45,000 characters." };
  if (portalLogo && (!/^data:image\/(?:jpeg|png);base64,/i.test(portalLogo) || portalLogo.length > 45000)) {
    return { ok: false, error: "The school portal logo must be a JPG or PNG image under 45,000 characters." };
  }
  let profileDetails = null;
  let profileRecordType = "";
  if (hasProfileDetails) {
    if (!["parent", "employee"].includes(role)) {
      return { ok: false, error: "Only parents and employees can edit contact details from their profile." };
    }
    const input = body.profileDetails;
    const name = String(input.name || "").trim();
    const phone = String(input.phone || "").trim();
    const phoneCountry = String(input.phoneCountry || "").trim().toUpperCase();
    const address = String(input.address || "").trim();
    if (!name || name.length > 150) return { ok: false, error: "Enter your full name (up to 150 characters)." };
    if (phone.length > 40) return { ok: false, error: "Phone number must be 40 characters or fewer." };
    if (phoneCountry && !/^[A-Z]{2}$/.test(phoneCountry)) return { ok: false, error: "Choose a valid phone country code." };
    if (address.length > 500) return { ok: false, error: "Address must be 500 characters or fewer." };
    profileDetails = { name, phone, phoneCountry, address };
    if (role === "parent") {
      const relationship = String(input.relationship || "").trim();
      const occupation = String(input.occupation || "").trim();
      if (relationship.length > 100) return { ok: false, error: "Relationship must be 100 characters or fewer." };
      if (occupation.length > 120) return { ok: false, error: "Occupation must be 120 characters or fewer." };
      Object.assign(profileDetails, { relationship, occupation });
      profileRecordType = "parent";
    } else {
      profileRecordType = "employee";
    }
  }
  if (hasBrandColor && (!isAdmin || !(hasPicture || hasPortalLogo) || !/^#[0-9a-f]{6}$/i.test(brandColor))) {
    return { ok: false, error: "A school brand color must be a valid hex color selected with an administrator profile picture or school logo." };
  }
  if (hasBackground && !isAdmin) return { ok: false, error: "Only a school administrator can change the portal background." };
  if (hasBackground && background && (!/^data:image\/(?:jpeg|png);base64,/i.test(background) || background.length > 45000)) {
    return { ok: false, error: "The portal background must be a JPG or PNG image under 45,000 characters." };
  }

  let updatedProfile = null;
  if (hasPicture || profileDetails) {
    const profileHeaders = [
      "id", "name", "email", "password", "passwordHash", "passwordSalt", "role", "type",
      "jobRole", "department", "className", "childStudentId", "profilePicture",
      "phone", "phoneCountry", "address", "relationship", "occupation"
    ];
    const sheet = ensureHeaders_(USERS, profileHeaders);
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(header => String(header).trim());
    const emailColumn = headers.indexOf("email");
    const rowIndex = values.findIndex((row, index) =>
      index > 0 && String(row[emailColumn] || "").trim().toLowerCase() === String(actor.email || "").trim().toLowerCase()
    );
    if (rowIndex < 1) return { ok: false, error: "User account was not found." };
    const roleColumn = headers.indexOf("role");
    const typeColumn = headers.indexOf("type");
    const storedRole = String(values[rowIndex][roleColumn] || values[rowIndex][typeColumn] || "").toLowerCase();
    if (storedRole !== role) return { ok: false, error: "Your account role could not be verified. Sign in again and retry." };

    let profileRecord = null;
    if (profileDetails) {
      profileRecord = readSchoolRecords_(profileRecordType).find(record =>
        String(record.email || "").trim().toLowerCase() === String(actor.email || "").trim().toLowerCase()
      );
      if (!profileRecord) return { ok: false, error: "Your school profile could not be found. Contact your school administrator." };
    }

    if (hasPicture) sheet.getRange(rowIndex + 1, headers.indexOf("profilePicture") + 1).setValue(picture);
    if (profileDetails) {
      Object.keys(profileDetails).forEach(field => sheet.getRange(rowIndex + 1, headers.indexOf(field) + 1).setValue(profileDetails[field]));
      writeSchoolRecord_(profileRecordType, { ...profileRecord, ...profileDetails, updatedAt: new Date().toISOString() });
      updatedProfile = profileDetails;
    }
  }

  const updates = {};
  if (isAdmin && hasPicture) updates.schoolProfilePicture = picture;
  if (isAdmin && hasPortalLogo) {
    updates.portalLogo = portalLogo;
    updates.brandingSetupCompleted = true;
  }
  if (hasBrandColor) updates.brandColor = brandColor.toLowerCase();
  if (hasBackground) updates.portalBackgroundImage = background;
  if (Object.keys(updates).length) updateTenant_(tenant.schoolCode || tenant.tenantId, updates);
  return {
    ok: true,
    ...(hasPicture ? { profilePicture: picture } : {}),
    ...(updatedProfile ? { profileDetails: updatedProfile } : {}),
    ...(isAdmin ? {
      portalLogo: hasPortalLogo ? portalLogo : tenant.portalLogo || "",
      schoolProfilePicture: hasPicture ? picture : tenant.schoolProfilePicture || actor.profilePicture || ""
    } : {}),
    ...(hasBrandColor ? { brandColor: brandColor.toLowerCase() } : {}),
    ...(isAdmin && hasPicture ? { brandingSetupCompleted: true } : {}),
    portalBackgroundImage: hasBackground ? background : tenant.portalBackgroundImage || ""
  };
}

function completeBrandingSetup_(body) {
  const tenant = body.schoolTenant || {};
  if (!isSchoolOwnerAdmin_(body.authUser, tenant)) {
    return { ok: false, error: "Only the main school administrator can complete school branding setup." };
  }
  const updated = updateTenant_(tenant.schoolCode || tenant.tenantId, { brandingSetupCompleted: true });
  return { ok: true, brandingSetupCompleted: true, brandColor: updated.brandColor || "" };
}

function schoolPaystackPropertyKey_(tenantId) {
  const schoolCode = normalizeSchoolCode_(tenantId);
  if (!schoolCode) throw new Error("School account could not be identified.");
  return `SCHOOL_PAYSTACK_SECRET_${schoolCode.toUpperCase()}`;
}

function getSchoolPaystackSettings_(body, tenantId) {
  if (String(body.authUser && body.authUser.role || "").toLowerCase() !== "admin") {
    return { ok: false, error: "Only a school administrator can manage Paystack settings." };
  }
  const secret = String(PropertiesService.getScriptProperties().getProperty(schoolPaystackPropertyKey_(tenantId)) || "");
  return {
    ok: true,
    connected: Boolean(secret),
    lastFour: secret.slice(-4),
    mode: secret.indexOf("sk_test_") === 0 ? "test" : secret ? "live" : ""
  };
}

function saveSchoolPaystackSettings_(body, tenantId) {
  if (String(body.authUser && body.authUser.role || "").toLowerCase() !== "admin") {
    return { ok: false, error: "Only a school administrator can manage Paystack settings." };
  }
  const properties = PropertiesService.getScriptProperties();
  const propertyKey = schoolPaystackPropertyKey_(tenantId);
  if (body.remove === true) {
    properties.deleteProperty(propertyKey);
    return { ok: true, connected: false, lastFour: "", mode: "" };
  }
  const secret = String(body.secretKey || "").trim();
  if (!/^sk_(test|live)_[A-Za-z0-9]+$/.test(secret)) {
    return { ok: false, error: "Enter a valid Paystack Secret Key beginning with sk_test_ or sk_live_." };
  }
  const response = UrlFetchApp.fetch("https://api.paystack.co/bank?country=nigeria", {
    method: "get",
    headers: { Authorization: `Bearer ${secret}` },
    muteHttpExceptions: true
  });
  let result;
  try {
    result = JSON.parse(response.getContentText());
  } catch (error) {
    throw new Error("Paystack returned an unreadable response while checking the Secret Key.");
  }
  if (response.getResponseCode() >= 300 || !result.status) {
    throw new Error(result.message || "Paystack could not verify this Secret Key. Check the key and try again.");
  }
  properties.setProperty(propertyKey, secret);
  return {
    ok: true,
    connected: true,
    lastFour: secret.slice(-4),
    mode: secret.indexOf("sk_test_") === 0 ? "test" : "live"
  };
}

function schoolFeeReservations_(attempts, invoiceId, exceptReference, now) {
  return (attempts || []).reduce((total, attempt) => {
    if (String(attempt.invoiceId || "") !== String(invoiceId || "") ||
        String(attempt.reference || "") === String(exceptReference || "") ||
        !schoolFeeReservationIsActive_(attempt, now)) return total;
    return total + Math.max(0, Number(attempt.amount) || 0);
  }, 0);
}

function schoolFeeReservationIsActive_(attempt, now) {
  if (String(attempt.status || "").toLowerCase() !== "initiated") return false;
  const createdAt = new Date(attempt.createdAt || "").getTime();
  return !Number.isFinite(createdAt) || now - createdAt <= 6 * 60 * 60 * 1000;
}

function startSchoolFeePayment_(body, tenant) {
  const payer = body.authUser || {};
  if (String(payer.role || "").toLowerCase() !== "parent") {
    return { ok: false, error: "Only a signed-in parent can start a school fee payment." };
  }
  const secret = String(PropertiesService.getScriptProperties().getProperty(
    schoolPaystackPropertyKey_(tenant.tenantId)
  ) || "");
  if (!secret) return { ok: false, error: "This school has not connected its Paystack account yet. Contact the school office." };

  const invoiceId = String(body.invoiceId || "").trim();
  const invoice = readSchoolRecords_("feeInvoice").find(record => String(record.id) === invoiceId);
  if (!invoice) return { ok: false, error: "Select a valid fee invoice from the school." };
  const studentId = String(invoice.studentId || "").trim();
  const students = readSchoolRecords_("student");
  const linkedStudentIds = linkedStudentIdsForParent_(payer, students);
  const student = students.find(record =>
    String(record.studentId || "").trim() === studentId && linkedStudentIds.has(studentId)
  );
  if (!student) return { ok: false, error: "Select a student linked to your parent account." };
  const amount = Number(body.amount);
  if (!Number.isFinite(amount) || amount <= 0) return { ok: false, error: "Enter a payment greater than zero." };
  const studentName = `${student.firstName || ""} ${student.lastName || ""}`.trim();
  const parentEmail = String(payer.email || "").trim().toLowerCase();
  if (!studentName || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(parentEmail)) {
    return { ok: false, error: "The parent or student details are incomplete. Contact the school." };
  }

  const returnUrl = String(body.returnUrl || "").trim().replace(/[?#].*$/, "");
  if (!/^https?:\/\/[^/?#]+(?:\/[^?#]*)?$/i.test(returnUrl)) {
    return { ok: false, error: "The school website return address is invalid. Reopen the school portal and try again." };
  }
  const attempts = ensureHeaders_(SCHOOL_FEE_ATTEMPTS, [
    "reference", "tenantId", "studentId", "studentName", "parentName",
    "parentEmail", "feePlanId", "invoiceId", "paidFor", "amount", "academicYear",
    "term", "status", "createdAt", "paidAt", "portalUrl"
  ]);
  const preflightAttempts = readRows_(SCHOOL_FEE_ATTEMPTS);
  const initialBalance = feeInvoiceBalance_(invoice).balance;
  const initiallyReserved = schoolFeeReservations_(preflightAttempts, invoiceId, "", Date.now());
  if (amount <= initialBalance + 0.000001 &&
      amount > Math.max(0, initialBalance - initiallyReserved) + 0.000001) {
    const now = Date.now();
    const unresolvedAttempts = preflightAttempts.filter(attempt =>
      String(attempt.invoiceId || "") === invoiceId &&
      schoolFeeReservationIsActive_(attempt, now) &&
      attempt.reference &&
      String(attempt.reference).indexOf("initializing-") !== 0
    );
    unresolvedAttempts.forEach(attempt => verifySchoolFeePayment_(attempt.reference, tenant.tenantId));
  }
  let latestInvoice, reservationRecord;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    latestInvoice = readSchoolRecords_("feeInvoice").find(record => String(record.id) === invoiceId);
    if (!latestInvoice) return { ok: false, error: "This fee invoice is no longer available." };
    if (String(latestInvoice.studentId || "").trim() !== studentId) {
      return { ok: false, error: "This fee invoice's student link changed. Refresh the parent portal and try again." };
    }
    const latestBalance = feeInvoiceBalance_(latestInvoice).balance;
    if (amount > latestBalance + 0.000001) {
      return { ok: false, error: `Enter a payment no greater than the remaining invoice balance of NGN ${latestBalance.toLocaleString("en-NG")}.` };
    }
    const existingAttempts = readRows_(SCHOOL_FEE_ATTEMPTS);
    const now = Date.now();
    const supersedeReferences = new Set(
      body.confirmPendingRetry === true && Array.isArray(body.supersedeReferences)
        ? body.supersedeReferences.map(reference => String(reference || ""))
        : []
    );
    if (body.confirmPendingRetry === true) {
      existingAttempts.forEach((attempt, index) => {
        const reference = String(attempt.reference || "");
        if (String(attempt.invoiceId || "") !== invoiceId ||
            !schoolFeeReservationIsActive_(attempt, now) ||
            !reference ||
            reference.indexOf("initializing-") === 0 ||
            !supersedeReferences.has(reference) ||
            String(attempt.parentEmail || "").trim().toLowerCase() !== parentEmail ||
            normalizeSchoolCode_(attempt.tenantId) !== normalizeSchoolCode_(tenant.tenantId)) return;
        writeRecord_(attempts, index + 2, { ...attempt, status: "superseded" });
        existingAttempts[index] = { ...attempt, status: "superseded" };
      });
    }
    const reservedAmount = schoolFeeReservations_(existingAttempts, invoiceId, "", now);
    const availableBalance = Math.max(0, latestBalance - reservedAmount);
    if (amount > availableBalance + 0.000001) {
      const activeAttempts = existingAttempts.filter(attempt =>
        String(attempt.invoiceId || "") === invoiceId &&
        schoolFeeReservationIsActive_(attempt, now)
      );
      const retryReferences = activeAttempts.filter(attempt =>
        String(attempt.reference || "").indexOf("initializing-") !== 0 &&
        String(attempt.parentEmail || "").trim().toLowerCase() === parentEmail &&
        normalizeSchoolCode_(attempt.tenantId) === normalizeSchoolCode_(tenant.tenantId)
      ).map(attempt => String(attempt.reference || ""));
      const retryRequired = !body.confirmPendingRetry && retryReferences.length > 0;
      const unresolvedReservation = activeAttempts.length > 0;
      if (unresolvedReservation) {
        return {
          ok: false,
          retryRequired: retryRequired,
          ...(retryRequired ? { retryReferences: retryReferences } : {}),
          error: `A previous Paystack checkout is still pending. NGN ${reservedAmount.toLocaleString("en-NG")} remains reserved to prevent duplicate payment. Available checkout amount: NGN ${availableBalance.toLocaleString("en-NG")}.`
        };
      }
      return { ok: false, error: `Existing payment attempts reserve part of this invoice. The available checkout amount is NGN ${availableBalance.toLocaleString("en-NG")}.` };
    }

    reservationRecord = {
      tenantId: tenant.tenantId,
      studentId: studentId,
      studentName: studentName,
      parentName: payer.name || "",
      parentEmail: parentEmail,
      feePlanId: latestInvoice.feePlanId,
      invoiceId: latestInvoice.id,
      paidFor: latestInvoice.feeName,
      amount: amount,
      academicYear: latestInvoice.academicYear,
      term: latestInvoice.term,
      portalUrl: admissionsLink_(returnUrl, `school=${encodeURIComponent(tenant.schoolCode || tenant.tenantId)}`),
      reference: `initializing-${Utilities.getUuid()}`,
      status: "initiated",
      createdAt: new Date()
    };
    appendRecord_(attempts, reservationRecord);
  } finally {
    lock.releaseLock();
  }

  let result;
  let response;
  try {
    response = UrlFetchApp.fetch("https://api.paystack.co/transaction/initialize", {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: `Bearer ${secret}` },
      payload: JSON.stringify({
        email: parentEmail,
        amount: Math.round(amount * 100),
        currency: "NGN",
        callback_url: `${returnUrl}?feePaymentReturn=1&tenant=${encodeURIComponent(tenant.tenantId)}`,
        metadata: {
          schoolCode: tenant.schoolCode || tenant.tenantId,
          feePlanId: latestInvoice.feePlanId,
          invoiceId: latestInvoice.id,
          studentId: studentId,
          parentEmail: parentEmail,
          academicYear: latestInvoice.academicYear,
          term: latestInvoice.term
        }
      }),
      muteHttpExceptions: true
    });
    try {
      result = JSON.parse(response.getContentText());
    } catch (error) {
      throw new Error("Paystack returned an unreadable response while starting checkout.");
    }
    if (response.getResponseCode() >= 300 || !result.status) {
      throw new Error(result.message || "Could not start Paystack checkout.");
    }
    if (!result.data || !result.data.authorization_url || !result.data.reference) {
      throw new Error("Paystack did not return a complete checkout response.");
    }
  } catch (error) {
    if (response && (response.getResponseCode() >= 300 || result && result.status === false)) {
      const releaseLock = LockService.getScriptLock();
      let acquired = false;
      try {
        releaseLock.waitLock(10000);
        acquired = true;
        const rows = readRows_(SCHOOL_FEE_ATTEMPTS);
        const rowIndex = rows.findIndex(row => String(row.reference || "") === reservationRecord.reference);
        if (rowIndex >= 0 && String(rows[rowIndex].status || "").toLowerCase() === "initiated") {
          writeRecord_(attempts, rowIndex + 2, { ...rows[rowIndex], status: "failed" });
        }
      } catch (releaseError) {
        throw new Error(`${error.message} The invoice reservation could not be released: ${releaseError.message}`);
      } finally {
        if (acquired) releaseLock.releaseLock();
      }
      throw error;
    }
    throw new Error(`${error.message} Checkout status is unknown, so the invoice balance remains reserved for up to six hours. Do not start another payment; the reservation will expire automatically.`);
  }

  const pendingPayment = { ...reservationRecord, reference: result.data.reference };
  const finalizeLock = LockService.getScriptLock();
  let finalizeLockAcquired = false;
  try {
    finalizeLock.waitLock(10000);
    finalizeLockAcquired = true;
    const rows = readRows_(SCHOOL_FEE_ATTEMPTS);
    const rowIndex = rows.findIndex(row => String(row.reference || "") === reservationRecord.reference);
    if (rowIndex < 0 || String(rows[rowIndex].status || "").toLowerCase() !== "initiated") {
      throw new Error("The invoice reservation could not be found.");
    }
    writeRecord_(attempts, rowIndex + 2, pendingPayment);
    CacheService.getScriptCache().put(`schoolfee:${result.data.reference}`, JSON.stringify(pendingPayment), 21600);
  } catch (error) {
    throw new Error(`Paystack created checkout ${result.data.reference}, but its reservation could not be finalized. Do not start another payment; contact the school office. ${error.message}`);
  } finally {
    if (finalizeLockAcquired) finalizeLock.releaseLock();
  }
  return { ok: true, checkoutUrl: result.data.authorization_url };
}

function markRecordedSchoolFeeAttempt_(attemptsSheet, reference, payment) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const attempts = readRows_(SCHOOL_FEE_ATTEMPTS);
    const index = attempts.findIndex(record => String(record.reference || "") === String(reference || ""));
    if (index < 0 || !["initiated", "superseded"].includes(String(attempts[index].status || "").toLowerCase())) return;
    writeRecord_(attemptsSheet, index + 2, {
      ...attempts[index],
      status: payment && payment.invoiceId ? "paid" : "paid_unapplied"
    });
  } finally {
    lock.releaseLock();
  }
}

function verifySchoolFeePayment_(reference, tenantId) {
  const paymentReference = String(reference || "").trim();
  if (!paymentReference) return { ok: false, error: "Paystack payment reference is missing." };
  const requestedTenant = findTenant_(tenantId);
  if (!requestedTenant) return { ok: false, error: "School account could not be found." };
  requestSpreadsheetId = String(requestedTenant.spreadsheetId || "");
  if (!requestSpreadsheetId) return { ok: false, error: "The school's payment records are unavailable. Contact the school." };
  const attemptsSheet = ensureHeaders_(SCHOOL_FEE_ATTEMPTS, [
    "reference", "tenantId", "studentId", "studentName", "parentName",
    "parentEmail", "feePlanId", "invoiceId", "paidFor", "amount", "academicYear",
    "term", "status", "createdAt", "paidAt", "portalUrl"
  ]);
  const attempts = readRows_(SCHOOL_FEE_ATTEMPTS);
  const attemptIndex = attempts.findIndex(record => String(record.reference || "") === paymentReference);
  if (attemptIndex < 0) return { ok: false, error: "This school fee payment could not be found. Contact the school before trying again." };
  const pending = attempts[attemptIndex];
  const cache = CacheService.getScriptCache();
  const cacheKey = `schoolfee:${paymentReference}`;
  if (normalizeSchoolCode_(pending.tenantId) !== normalizeSchoolCode_(tenantId)) {
    return { ok: false, error: "Payment reference does not match this school." };
  }
  const sheet = ensureHeaders_(PAYMENTS, [
    "id", "studentId", "studentName", "parentName", "parentEmail",
    "amount", "paidFor", "feePlanId", "invoiceId", "paymentMethod", "reference",
    "paymentDate", "academicYear", "term", "createdAt"
  ]);
  const existingPayment = readRows_(PAYMENTS).find(record => String(record.reference || "") === paymentReference);
  if (existingPayment) {
    if (["initiated", "superseded"].includes(String(pending.status || "").toLowerCase())) {
      markRecordedSchoolFeeAttempt_(attemptsSheet, paymentReference, existingPayment);
      cache.remove(cacheKey);
    }
    return { ok: true, message: "This fee payment was already recorded.", payment: existingPayment };
  }
  const tenant = findTenant_(pending.tenantId);
  if (!tenant) return { ok: false, error: "School account could not be found." };
  requestSpreadsheetId = String(tenant.spreadsheetId || "");
  if (!requestSpreadsheetId) return { ok: false, error: "The school's payment records are unavailable. Contact the school." };

  const secret = String(PropertiesService.getScriptProperties().getProperty(
    schoolPaystackPropertyKey_(tenant.tenantId)
  ) || "");
  if (!secret) return { ok: false, error: "This school's Paystack connection is unavailable. Contact the school." };
  const response = UrlFetchApp.fetch(
    `https://api.paystack.co/transaction/verify/${encodeURIComponent(paymentReference)}`,
    { method: "get", headers: { Authorization: `Bearer ${secret}` }, muteHttpExceptions: true }
  );
  let result;
  try {
    result = JSON.parse(response.getContentText());
  } catch (error) {
    throw new Error("Paystack returned an unreadable response while verifying payment.");
  }
  const transaction = result.data || {};
  if (response.getResponseCode() >= 300 || !result.status) {
    return { ok: false, error: "Paystack has not confirmed this fee payment." };
  }
  if (String(transaction.status || "").toLowerCase() !== "success") {
    const terminalStatus = String(transaction.status || "").toLowerCase();
    if (["failed", "abandoned", "cancelled", "canceled"].includes(terminalStatus)) {
      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        const latestAttempts = readRows_(SCHOOL_FEE_ATTEMPTS);
        const latestIndex = latestAttempts.findIndex(record =>
          String(record.reference || "") === paymentReference
        );
        if (latestIndex >= 0 && ["initiated", "superseded"].includes(String(latestAttempts[latestIndex].status || "").toLowerCase())) {
          const headers = attemptsSheet.getRange(1, 1, 1, attemptsSheet.getLastColumn())
            .getValues()[0].map(value => String(value).trim());
          const statusColumn = headers.indexOf("status");
          if (statusColumn < 0) throw new Error("The school fee attempts sheet is missing its status column.");
          attemptsSheet.getRange(latestIndex + 2, statusColumn + 1).setValue(terminalStatus);
        }
      } finally {
        lock.releaseLock();
      }
      cache.remove(cacheKey);
      return {
        ok: false,
        error: "Paystack reports that this checkout was abandoned, canceled, or failed. The invoice balance has been released; start a new payment to try again."
      };
    }
    return { ok: false, error: "Paystack has not confirmed this fee payment." };
  }
  if (String(transaction.reference || "") !== paymentReference) {
    return { ok: false, error: "The confirmed payment reference does not match this checkout." };
  }
  if (Number(transaction.amount) !== Math.round(Number(pending.amount) * 100) ||
      String(transaction.currency || "").toUpperCase() !== "NGN") {
    return { ok: false, error: "The confirmed payment amount does not match the selected school fee." };
  }
  const customerEmail = String(transaction.customer && transaction.customer.email || "").trim().toLowerCase();
  if (customerEmail && customerEmail !== String(pending.parentEmail || "").toLowerCase()) {
    return { ok: false, error: "The confirmed Paystack payer does not match this payment." };
  }
  const metadata = transaction.metadata || {};
  if (metadata.schoolCode && normalizeSchoolCode_(metadata.schoolCode) !== normalizeSchoolCode_(tenant.schoolCode || tenant.tenantId)) {
    return { ok: false, error: "The confirmed payment does not match this school." };
  }
  if (metadata.invoiceId && String(metadata.invoiceId) !== String(pending.invoiceId || "")) {
    return { ok: false, error: "The confirmed payment does not match this fee invoice." };
  }
  if (metadata.studentId && String(metadata.studentId) !== String(pending.studentId || "")) {
    return { ok: false, error: "The confirmed payment does not match this student." };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let recordedPayment;
  let paymentAppliedToInvoice = true;
  try {
    const latestAttempts = readRows_(SCHOOL_FEE_ATTEMPTS);
    const latestAttemptIndex = latestAttempts.findIndex(record => String(record.reference || "") === paymentReference);
    if (latestAttemptIndex < 0) return { ok: false, error: "This school fee payment could not be found. Contact the school before trying again." };
    const latestPending = latestAttempts[latestAttemptIndex];
    const duplicatePayment = readRows_(PAYMENTS).find(record => String(record.reference || "") === paymentReference);
    if (duplicatePayment) {
      if (["initiated", "superseded"].includes(String(latestPending.status || "").toLowerCase())) {
        writeRecord_(attemptsSheet, latestAttemptIndex + 2, {
          ...latestPending,
          status: duplicatePayment.invoiceId ? "paid" : "paid_unapplied"
        });
      }
      cache.remove(cacheKey);
      return { ok: true, message: "This fee payment was already recorded.", payment: duplicatePayment };
    }
    const latestInvoice = latestPending.invoiceId
      ? readSchoolRecords_("feeInvoice").find(record => String(record.id) === String(latestPending.invoiceId))
      : null;
    if (latestInvoice) {
      const currentBalance = feeInvoiceBalance_(latestInvoice).balance;
      const otherReservations = schoolFeeReservations_(latestAttempts, latestPending.invoiceId, paymentReference, Date.now());
      paymentAppliedToInvoice = Number(latestPending.amount) <= Math.max(0, currentBalance - otherReservations) + 0.000001;
    } else {
      paymentAppliedToInvoice = !latestPending.invoiceId;
    }
    const paymentDate = transaction.paid_at || new Date();
    recordedPayment = {
      id: Utilities.getUuid(),
      studentId: latestPending.studentId,
      studentName: latestPending.studentName,
      parentName: latestPending.parentName,
      parentEmail: latestPending.parentEmail,
      amount: Number(latestPending.amount),
      paidFor: latestPending.paidFor,
      feePlanId: latestPending.feePlanId,
      invoiceId: paymentAppliedToInvoice ? latestPending.invoiceId || "" : "",
      paymentMethod: "Paystack",
      reference: paymentReference,
      paymentDate: paymentDate,
      academicYear: latestPending.academicYear,
      term: latestPending.term,
      status: "paid",
      createdAt: new Date()
    };
    appendRecord_(sheet, recordedPayment);
    writeRecord_(attemptsSheet, latestAttemptIndex + 2, {
      ...latestPending,
      reference: paymentReference,
      status: paymentAppliedToInvoice ? "paid" : "paid_unapplied",
      createdAt: latestPending.createdAt,
      paidAt: paymentDate
    });
    cache.remove(cacheKey);
  } finally {
    lock.releaseLock();
  }
  let notificationWarning = "";
  let updatedInvoice = null;
  try {
    if (recordedPayment.invoiceId) updatedInvoice = updateFeeInvoiceStatus_(
      recordedPayment.invoiceId,
      { name: pending.parentName || "Parent", email: pending.parentEmail }
    );
    notificationWarning = notifyFeePayment_(recordedPayment, pending.portalUrl).join(" ");
    logAudit_(
      { name: pending.parentName || "Parent", email: pending.parentEmail },
      "verified Paystack payment",
      "payment",
      recordedPayment.id,
      null,
      recordedPayment,
      `Verified NGN ${Number(pending.amount).toLocaleString("en-NG")} for invoice ${pending.invoiceId || "unlinked"}.`
    );
  } catch (error) {
    console.error(`Fee payment ${paymentReference} was recorded, but notifications or audit logging failed: ${error.message}`);
    notificationWarning = "Payment was recorded, but some receipt notifications could not be completed.";
  }
  if (!paymentAppliedToInvoice) {
    const reviewMessage = "Payment was received but not applied to the invoice because other payments or pending checkouts used its remaining balance. Contact the school office to reconcile it.";
    notificationWarning = [notificationWarning, reviewMessage].filter(Boolean).join(" ");
  }
  return {
    ok: true,
    message: notificationWarning || "Payment verified and recorded successfully.",
    payment: recordedPayment,
    ...(updatedInvoice ? { invoice: {
      id: updatedInvoice.id,
      paidAmount: updatedInvoice.paidAmount,
      balance: updatedInvoice.balance,
      status: updatedInvoice.status
    } } : {}),
    ...(notificationWarning ? { warning: notificationWarning } : {})
  };
}

function announcementRecipients_(audience, recipient) {
  const accounts = readRows_(USERS);
  const recipients = [];
  const seenEmails = new Set();
  const addRecipient = (email, name) => {
    const normalizedEmail = String(email || "").trim().toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizedEmail) || seenEmails.has(normalizedEmail)) return;
    seenEmails.add(normalizedEmail);
    recipients.push({ email: normalizedEmail, name: String(name || "") });
  };
  const accountRole = account => String(account.role || account.type || "").toLowerCase();

  if (audience === "Everyone") {
    accounts.filter(account => ["admin", "employee", "parent"].includes(accountRole(account)))
      .forEach(account => addRecipient(account.email, account.name));
    ["employee", "parent"].forEach(type => {
      readSchoolRecords_(type).forEach(record => addRecipient(record.email, record.name));
    });
  } else if (audience === "Specific employee" || audience === "Specific parent or guardian") {
    const role = audience === "Specific employee" ? "employee" : "parent";
    const email = String(recipient || "").trim().toLowerCase();
    const account = accounts.find(item => accountRole(item) === role &&
      String(item.email || "").trim().toLowerCase() === email);
    const record = readSchoolRecords_(role).find(item =>
      String(item.email || "").trim().toLowerCase() === email
    );
    if (account || record) addRecipient(email, account && account.name || record && record.name);
  } else if (audience === "Specific student") {
    const studentId = String(recipient || "").trim();
    const student = readSchoolRecords_("student").find(item =>
      String(item.studentId || "").trim() === studentId
    );
    if (student) {
      parentContactsForStudent_(student, readSchoolRecords_("parent"))
        .forEach(parent => addRecipient(parent.email, parent.name));
    }
  }

  return recipients;
}

function sendAnnouncementEmails_(body, sender, schoolName, recipients) {
  const subject = String(body.subject || "School announcement")
    .replace(/[\r\n\x00-\x1f\x7f]/g, " ").trim().slice(0, 200) || "School announcement";
  const message = String(body.message || "");
  const senderName = String(sender.name || "School Admin");
  const school = String(schoolName || "School")
    .replace(/[\r\n\x00-\x1f\x7f]/g, " ").trim().slice(0, 100) || "School";
  const portalUrl = /^https:\/\/[a-z0-9.-]+(?::\d+)?(?:\/[^\s]*)?$/i.test(String(body.portalUrl || "").trim())
    ? String(body.portalUrl).trim()
    : "";
  const sendDateTime = [body.sendDate, body.sendTime].filter(Boolean).join(" ");
  const plainBody = `${subject}\n\n${message}\n\nSent by ${senderName} at ${school}.${sendDateTime ? `\n${sendDateTime}` : ""}${portalUrl ? `\n\nOpen your school portal: ${portalUrl}` : ""}`;
  const htmlBody = `<div style="background:#f3f6fb;padding:30px 12px;font-family:Arial,sans-serif;color:#172033"><div style="max-width:560px;margin:0 auto;background:#fff;border:1px solid #e4e8ef;border-radius:16px;overflow:hidden"><div style="height:7px;background:linear-gradient(90deg,#3762ea,#17b890)"></div><div style="padding:28px"><div style="font-size:11px;font-weight:bold;letter-spacing:2px;color:#3762ea">${admissionsHtml_(school)}</div><h2 style="margin:22px 0 14px;font-size:23px">${admissionsHtml_(subject)}</h2><p style="font-size:14px;line-height:1.75;white-space:pre-wrap">${admissionsHtml_(message)}</p>${portalUrl ? `<p><a href="${admissionsHtml_(portalUrl)}">Open your school portal</a></p>` : ""}<p style="border-top:1px solid #e4e8ef;padding-top:16px;margin:24px 0 0;color:#758096;font-size:11px">Sent by ${admissionsHtml_(senderName)}${sendDateTime ? ` · ${admissionsHtml_(sendDateTime)}` : ""}</p></div></div></div>`;
  let sent = 0;
  let failed = 0;

  recipients.forEach(recipient => {
    try {
      MailApp.sendEmail({
        to: recipient.email,
        subject: subject,
        name: school,
        body: plainBody,
        htmlBody: htmlBody
      });
      sent++;
    } catch (error) {
      failed++;
      console.error(`Announcement email could not be sent to ${recipient.email}: ${error.message}`);
    }
  });

  return { sent: sent, failed: failed };
}

function saveAnnouncement_(body) {
  const sender = body.authUser;
  if (!sender || String(sender.role || "").toLowerCase() !== "admin") {
    return { ok: false, error: "Only an administrator can send announcements." };
  }

  if (!body.subject || !body.message) {
    return { ok: false, error: "Subject and message are required." };
  }

  const audience = String(body.audience || "Everyone");
  const recipients = announcementRecipients_(audience, body.recipient);
  const schoolName = body.schoolTenant && body.schoolTenant.schoolName || sender.schoolName || "School";
  const sheet = ensureHeaders_(MESSAGES, [
    "id", "audience", "recipient", "subject", "message",
    "sendDate", "sendTime", "sentBy", "sentByEmail", "createdAt", "actionView",
    "replyToId", "recipientType", "recipientId", "editedAt", "sentByRole", "messageType"
  ]);
  const allowedViews = ["dashboard", "inbox", "attendance", "timetable", "fees", "parentFees", "parentReportCards"];
  const actionView = allowedViews.includes(String(body.actionView || "")) ? String(body.actionView) : "";

  const savedMessage = {
    id: Utilities.getUuid(),
    audience: audience,
    recipient: body.recipient || "",
    subject: body.subject,
    message: body.message,
    sendDate: body.sendDate || "",
    sendTime: body.sendTime || "",
    sentBy: sender.name || "School Admin",
    sentByEmail: sender.email || "",
    createdAt: new Date(),
    actionView: actionView,
    replyToId: "",
    recipientType: "",
    recipientId: "",
    editedAt: "",
    sentByRole: "admin",
    messageType: "announcement"
  };
  appendRecord_(sheet, savedMessage);

  const emailResult = sendAnnouncementEmails_(body, sender, schoolName, recipients);
  savedMessage.createdAt = new Date();
  const savedRow = findMessageRow_(savedMessage.id);
  if (savedRow) writeRecord_(savedRow.sheet, savedRow.rowNumber, savedMessage, savedRow.headers);
  else console.error(`Announcement ${savedMessage.id} was not found after saving; its edit window could not be refreshed.`);
  const notificationWarning = recipients.length === 0
    ? "Announcement saved, but no valid email recipients were found."
    : emailResult.failed
      ? `Announcement saved in the school inbox, but email delivery failed for ${emailResult.failed} of ${recipients.length} recipients.`
      : "";
  return {
    ok: true,
    message: savedMessage,
    emailSent: emailResult.sent,
    emailFailed: emailResult.failed,
    ...(notificationWarning ? { notificationWarning: notificationWarning } : {})
  };
}

function reportCardPdfBlob_(record, schoolName) {
  const document = DocumentApp.create(`Report Card - ${record.studentName || record.studentId || "Student"}`);
  const fileId = document.getId();
  try {
    const body = document.getBody();
    body.clear();
    body.appendParagraph(String(schoolName || "School")).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph("Student Report Card").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(`Student: ${record.studentName || "Student"} · ID: ${record.studentId || "—"} · Class: ${record.className || "—"}`);
    body.appendParagraph(`${record.term || "Term"} · ${record.academicYear || ""}`);
    if (record.overallGrade) body.appendParagraph(`Overall grade: ${record.overallGrade}`);
    const results = String(record.subjectsAndResults || "").split(/\r?\n/).map(line => line.trim()).filter(Boolean)
      .map(line => {
        const match = line.match(/^(.+?)\s*(?::|\s+[–—-]\s+)\s*(.+)$/);
        return match ? [match[1].trim(), match[2].trim()] : [line, "—"];
      });
    body.appendTable([["Subject", "Result"], ...results]);
    if (record.teacherComment) body.appendParagraph(`Teacher's comment: ${record.teacherComment}`);
    if (record.principalComment) body.appendParagraph(`Principal's comment: ${record.principalComment}`);
    body.appendParagraph("Confidential student record").setItalic(true);
    document.saveAndClose();
    const pdf = DriveApp.getFileById(fileId).getAs(MimeType.PDF);
    pdf.setName(`Report Card - ${String(record.studentName || "Student").replace(/[\\/:*?"<>|]/g, "-")}.pdf`);
    DriveApp.getFileById(fileId).setTrashed(true);
    return pdf;
  } catch (error) {
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch (cleanupError) {
      console.error(`Temporary report card document ${fileId} could not be removed: ${cleanupError.message}`);
    }
    throw error;
  }
}

function sendReportCard_(body) {
  const actor = body.authUser || {};
  if (String(actor.role || "").toLowerCase() !== "admin") {
    return { ok: false, error: "Only a school administrator can send report cards." };
  }
  const recordId = String(body.reportCardId || "").trim();
  const record = readSchoolRecords_("reportCard").find(item => String(item.id) === recordId);
  if (!record) return { ok: false, error: "Report card not found. Refresh and try again." };
  if (record.sentAt) return { ok: false, error: "This report card has already been sent." };
  const parentEmail = String(record.parentEmail || "").trim().toLowerCase();
  const student = readSchoolRecords_("student").find(item =>
    String(item.studentId || "").trim() === String(record.studentId || "").trim()
  );
  const parent = readSchoolRecords_("parent").find(item =>
    String(item.email || "").trim().toLowerCase() === parentEmail
  );
  if (!student || !parent || !parentEmail ||
      !linkedStudentIdsForParent_(parent, [student]).has(String(student && student.studentId || "").trim())) {
    return { ok: false, error: "The report card must have a valid linked student and parent recipient." };
  }
  const portalUrl = /^https:\/\/[a-z0-9.-]+(?::\d+)?(?:\/[^\s]*)?$/i.test(String(body.portalUrl || "").trim())
    ? String(body.portalUrl).trim()
    : "";
  let inboxSaved = false;
  try {
    appendInboxMessage_(
      parentEmail,
      `Report card: ${record.studentName || "Student"} · ${record.term || ""} ${record.academicYear || ""}`.trim(),
      `Your child's report card is ready. Sign in to the parent portal and open Report Cards to view it. A PDF copy is attached to this email.`,
      actor.name || "School",
      "parentReportCards"
    );
    inboxSaved = true;
  } catch (error) {
    console.error(`Report card ${recordId} was prepared, but the parent inbox fallback failed: ${error.message}`);
  }
  if (!parentEmailPreferenceEnabled_(parentEmail, "emailReportCardNotifications")) {
    if (!inboxSaved) {
      return { ok: false, error: "The parent has email updates turned off, and the report card could not be saved to their inbox." };
    }
    const delivery = markReportCardSent_(record, actor);
    return {
      ok: true,
      emailSent: false,
      emailStatus: "opted_out",
      sentAt: delivery.sentAt,
      notificationWarning: [
        "The report card is in the parent inbox. Report-card email updates are turned off for this parent.",
        delivery.warning
      ].filter(Boolean).join(" ")
    };
  }
  try {
    const pdf = reportCardPdfBlob_(record, body.schoolTenant && body.schoolTenant.schoolName || actor.schoolName || "School");
    const schoolName = body.schoolTenant && body.schoolTenant.schoolName || actor.schoolName || "School";
    const subject = `Report card · ${record.studentName || "Student"} · ${record.term || ""} ${record.academicYear || ""}`.trim();
    const message = `Dear ${parent.name || "Parent/Guardian"},\n\nYour child's report card for ${record.term || "this term"} ${record.academicYear || ""} is attached as a PDF. You can also view it in the parent portal under Report Cards.${portalUrl ? `\n\nOpen your school portal: ${portalUrl}` : ""}\n\nStudent: ${record.studentName || "Student"}\nClass: ${record.className || "—"}\nOverall grade: ${record.overallGrade || "—"}\n\nRegards,\n${schoolName}`;
    const htmlBody = `<p>Dear ${admissionsHtml_(parent.name || "Parent/Guardian")},</p><p>Your child's report card for <strong>${admissionsHtml_(`${record.term || "this term"} ${record.academicYear || ""}`)}</strong> is attached as a PDF. You can also view it in the parent portal under Report Cards.</p>${portalUrl ? `<p><a href="${admissionsHtml_(portalUrl)}">Open your school portal</a></p>` : ""}<p><strong>Student:</strong> ${admissionsHtml_(record.studentName || "Student")}<br><strong>Class:</strong> ${admissionsHtml_(record.className || "—")}<br><strong>Overall grade:</strong> ${admissionsHtml_(record.overallGrade || "—")}</p><p>Regards,<br>${admissionsHtml_(schoolName)}</p>`;
    MailApp.sendEmail({
      to: parentEmail,
      subject: subject,
      name: schoolName,
      body: message,
      htmlBody: htmlBody,
      attachments: [pdf]
    });
    const delivery = markReportCardSent_(record, actor);
    return {
      ok: true,
      emailSent: true,
      emailStatus: "sent",
      sentAt: delivery.sentAt,
      notificationWarning: [
        !inboxSaved ? "The report card email was sent, but the parent inbox fallback could not be saved." : "",
        delivery.warning
      ].filter(Boolean).join(" ") || undefined
    };
  } catch (error) {
    console.error(`Report card ${recordId} was saved to the parent inbox, but PDF email delivery failed: ${error.message}`);
    if (inboxSaved) {
      const delivery = markReportCardSent_(record, actor);
      return {
        ok: true,
        emailSent: false,
        emailStatus: "failed",
        sentAt: delivery.sentAt,
        notificationWarning: [
          `The parent inbox was updated; report-card PDF email delivery failed: ${error.message}`,
          delivery.warning
        ].filter(Boolean).join(" ")
      };
    }
    return {
      ok: false,
      error: `Report-card email delivery failed, and the parent inbox could not be updated: ${error.message}`
    };
  }
}

function markReportCardSent_(record, actor) {
  const sentAt = new Date().toISOString();
  record.sentAt = sentAt;
  let warning = "";
  try {
    writeSchoolRecord_("reportCard", record);
  } catch (error) {
    warning = `Delivery succeeded, but the sent status could not be saved: ${error.message}. Avoid sending again until the records refresh.`;
    console.error(`Report card ${record.id} was delivered, but its sent status could not be saved: ${error.message}`);
  }
  try {
    logAudit_(actor, "sent report card", "reportCard", record.id, null, { parentEmail: record.parentEmail }, `Sent report card for ${record.studentName || record.studentId}.`);
  } catch (error) {
    console.error(`Report card ${record.id} was delivered, but the audit entry could not be saved: ${error.message}`);
  }
  return { sentAt: sentAt, warning: warning };
}

function getMessages_(body) {
  const user = body.authUser || {};
  const rows = readRows_(MESSAGES);
  const email = String(user.email || "").toLowerCase();
  const role = String(user.role || user.type || "").toLowerCase();
  const isAdmin = role === "admin";
  const linkedStudentIds = role === "parent"
    ? linkedStudentIdsForParent_(user, readSchoolRecords_("student"))
    : new Set();

  const messages = rows.filter(message =>
    message.audience === "Everyone" ||
    (role === "parent" && String(message.audience || "") === "Specific student" &&
      linkedStudentIds.has(String(message.recipient || "").trim())) ||
    String(message.recipient).toLowerCase() === email ||
    String(message.recipient) === String(user.id || "") ||
    String(message.sentByEmail || "").trim().toLowerCase() === email ||
    isAdmin && Boolean(message.sentByEmail) &&
      ["Everyone", "Specific", "Specific employee", "Specific parent or guardian", "Specific student"].includes(String(message.audience || ""))
  );
  if (messages.some(message => !message.sentByRole && message.sentByEmail)) {
    const authorRoles = new Map(readRows_(USERS).map(account => [
      String(account.email || "").trim().toLowerCase(),
      String(account.role || account.type || "").trim().toLowerCase()
    ]));
    messages.forEach(message => {
      if (!message.sentByRole) {
        message.sentByRole = authorRoles.get(String(message.sentByEmail || "").trim().toLowerCase()) || "";
      }
    });
  }

  if (!isAdmin) return { ok: true, messages: messages };

  const messageIds = new Set(messages.map(message => String(message.id || "")).filter(Boolean));
  const readReceipts = messageIds.size
    ? readRows_(MESSAGE_READS).filter(receipt => messageIds.has(String(receipt.messageId || "")))
    : [];
  ensureHeaders_(NOTIFICATION_CLEARS, ["messageId", "userEmail", "clearedAt"]);
  const clearedNotificationIds = readRows_(NOTIFICATION_CLEARS)
    .filter(row => String(row.userEmail || "").trim().toLowerCase() === email)
    .map(row => String(row.messageId || ""));
  return {
    ok: true,
    messages: messages,
    readReceipts: readReceipts,
    clearedNotificationIds: clearedNotificationIds
  };
}

function markMessagesRead_(body) {
  const user = body.authUser || {};
  const readerType = String(user.role || user.type || "").toLowerCase();
  if (!["employee", "parent"].includes(readerType)) {
    return { ok: false, error: "Only employees and parents can mark received messages as opened." };
  }

  const messageIds = Array.isArray(body.messageIds)
    ? [...new Set(body.messageIds.map(id => String(id || "").trim()).filter(Boolean))].slice(0, 200)
    : [];
  if (!messageIds.length) return { ok: true, recorded: 0 };

  const accessibleIds = new Set(
    getMessages_({ authUser: user }).messages
      .map(message => String(message.id || ""))
      .filter(Boolean)
  );
  const requestedIds = messageIds.filter(id => accessibleIds.has(id));
  if (!requestedIds.length) return { ok: true, recorded: 0 };

  const readerEmail = String(user.email || "").trim().toLowerCase();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = ensureHeaders_(MESSAGE_READS, [
      "messageId", "readerEmail", "readerName", "readerType", "readAt"
    ]);
    const existing = new Set(readRows_(MESSAGE_READS)
      .filter(receipt => String(receipt.readerEmail || "").toLowerCase() === readerEmail)
      .map(receipt => String(receipt.messageId || "")));
    let recorded = 0;
    requestedIds.forEach(messageId => {
      if (existing.has(messageId)) return;
      appendRecord_(sheet, {
        messageId: messageId,
        readerEmail: readerEmail,
        readerName: String(user.name || readerEmail),
        readerType: readerType,
        readAt: new Date()
      });
      existing.add(messageId);
      recorded++;
    });
    return { ok: true, recorded: recorded };
  } finally {
    lock.releaseLock();
  }
}

function messageSheetHeaders_() {
  return [
    "id", "audience", "recipient", "subject", "message", "sendDate", "sendTime",
    "sentBy", "sentByEmail", "createdAt", "actionView", "replyToId",
    "recipientType", "recipientId", "editedAt", "sentByRole", "messageType"
  ];
}

function findMessageRow_(messageId) {
  const sheet = ensureHeaders_(MESSAGES, messageSheetHeaders_());
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(header => String(header).trim());
  const idColumn = headers.indexOf("id");
  if (idColumn < 0) return null;
  for (let index = 1; index < values.length; index++) {
    if (String(values[index][idColumn] || "") !== messageId) continue;
    const record = {};
    headers.forEach((header, column) => { record[header] = values[index][column]; });
    return { sheet: sheet, headers: headers, rowNumber: index + 1, record: record };
  }
  return null;
}

function messageCreatedAtMillis_(value) {
  const date = value instanceof Date ? value : new Date(value);
  const milliseconds = date.getTime();
  return Number.isFinite(milliseconds) ? milliseconds : null;
}

function messageSchoolAdmin_(body) {
  const actor = body.authUser || {};
  const actorRole = String(actor.role || actor.type || "").trim().toLowerCase();
  const actorEmail = String(actor.email || "").trim().toLowerCase();
  if (actorRole !== "parent" || !actorEmail) {
    return { ok: false, error: "Only a parent account can message the school administrator." };
  }
  const subject = String(body.subject || "").trim();
  const message = String(body.message || "").trim();
  if (!subject || subject.length > 200 || message.length < 1 || message.length > 5000) {
    return { ok: false, error: "Enter a subject (up to 200 characters) and a message (up to 5,000 characters)." };
  }
  const tenant = body.schoolTenant || {};
  const adminEmail = String(tenant.adminEmail || "").trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(adminEmail)) {
    return { ok: false, error: "The school administrator's message inbox is unavailable. Contact the school directly." };
  }
  const adminAccount = findUserByEmail_(adminEmail);
  const now = new Date();
  const record = {
    id: Utilities.getUuid(),
    audience: "Specific",
    recipient: adminEmail,
    subject: subject,
    message: message,
    sendDate: now,
    sendTime: Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm"),
    sentBy: String(actor.name || actorEmail),
    sentByEmail: actorEmail,
    createdAt: now,
    actionView: "",
    replyToId: "",
    recipientType: "admin",
    recipientId: String(adminAccount && adminAccount.id || adminEmail),
    editedAt: "",
    sentByRole: "parent",
    messageType: "direct"
  };
  appendRecord_(ensureHeaders_(MESSAGES, messageSheetHeaders_()), record);
  return { ok: true, message: record };
}

function replyToMessage_(body) {
  const actor = body.authUser || {};
  const actorRole = String(actor.role || actor.type || "").toLowerCase();
  if (!["admin", "employee", "parent"].includes(actorRole)) {
    return { ok: false, error: "This account cannot reply to school messages." };
  }
  const original = findMessageRow_(String(body.messageId || "").trim());
  if (!original) return { ok: false, error: "The message is no longer available." };

  const actorEmail = String(actor.email || "").trim().toLowerCase();
  const messageRecipient = String(original.record.recipient || "").trim().toLowerCase();
  const actorIds = [actorEmail, String(actor.id || "").trim().toLowerCase()].filter(Boolean);
  const addressedToActor = String(original.record.audience || "") !== "Everyone" &&
    Boolean(messageRecipient) && actorIds.includes(messageRecipient);
  const originalAuthor = findUserByEmail_(original.record.sentByEmail);
  const originalAuthorRole = String(originalAuthor && (originalAuthor.role || originalAuthor.type) || "").toLowerCase();
  const canReply = actorRole === "admin"
    ? addressedToActor && ["employee", "parent"].includes(originalAuthorRole)
    : addressedToActor && originalAuthorRole === "admin";
  if (!canReply) {
    return { ok: false, error: "You can only reply to a message directly addressed to your account." };
  }

  const message = String(body.message || "").trim();
  if (!message || message.length > 5000) {
    return { ok: false, error: "Your reply must be between 1 and 5,000 characters." };
  }
  const defaultSubject = /^re:/i.test(String(original.record.subject || ""))
    ? String(original.record.subject)
    : `Re: ${String(original.record.subject || "School message")}`;
  const subject = String(body.subject || defaultSubject).trim();
  if (!subject || subject.length > 200) {
    return { ok: false, error: "The reply subject must be between 1 and 200 characters." };
  }

  const recipient = String(original.record.sentByEmail).trim().toLowerCase();
  const recipientType = originalAuthorRole;
  const now = new Date();
  const reply = {
    id: Utilities.getUuid(),
    audience: "Specific",
    recipient: recipient,
    subject: subject,
    message: message,
    sendDate: now,
    sendTime: Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm"),
    sentBy: String(actor.name || actorEmail),
    sentByEmail: actorEmail,
    createdAt: now,
    actionView: "",
    replyToId: String(original.record.id),
    recipientType: recipientType,
    recipientId: String(originalAuthor.id || recipient),
    editedAt: "",
    sentByRole: actorRole,
    messageType: "reply"
  };
  appendRecord_(ensureHeaders_(MESSAGES, messageSheetHeaders_()), reply);
  return { ok: true, message: reply };
}

function editMessage_(body) {
  const actor = body.authUser || {};
  const actorRole = String(actor.role || actor.type || "").toLowerCase();
  const actorEmail = String(actor.email || "").trim().toLowerCase();
  if (!["admin", "employee", "parent"].includes(actorRole) || !actorEmail) {
    return { ok: false, error: "This account cannot edit school messages." };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const item = findMessageRow_(String(body.messageId || "").trim());
    if (!item) return { ok: false, error: "The message is no longer available." };
    if (String(item.record.sentByEmail || "").trim().toLowerCase() !== actorEmail) {
      return { ok: false, error: "You can only edit a message you sent." };
    }
    const createdAt = messageCreatedAtMillis_(item.record.createdAt);
    const age = createdAt === null ? Infinity : Date.now() - createdAt;
    if (item.record.messageType === "system" || age < 0 || age > 10000) {
      return { ok: false, error: "Messages can only be edited within 10 seconds of being sent." };
    }

    const subject = String(body.subject || "").trim();
    const message = String(body.message || "").trim();
    if (!subject || subject.length > 200 || !message || message.length > 5000) {
      return { ok: false, error: "Enter a subject (up to 200 characters) and message (up to 5,000 characters)." };
    }
    item.record.subject = subject;
    item.record.message = message;
    item.record.editedAt = new Date();
    writeRecord_(item.sheet, item.rowNumber, item.record, item.headers);
    return { ok: true, message: item.record };
  } finally {
    lock.releaseLock();
  }
}

function removeMessageReferences_(messageId) {
  [
    { sheetName: MESSAGE_READS, headers: ["messageId", "readerEmail", "readerName", "readerType", "readAt"] },
    { sheetName: NOTIFICATION_CLEARS, headers: ["messageId", "userEmail", "clearedAt"] }
  ].forEach(definition => {
    const sheet = ensureHeaders_(definition.sheetName, definition.headers);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(value => String(value).trim());
    const messageIdColumn = headers.indexOf("messageId");
    if (messageIdColumn < 0) return;
    const values = sheet.getDataRange().getValues();
    for (let index = values.length - 1; index > 0; index--) {
      if (String(values[index][messageIdColumn] || "") === messageId) sheet.deleteRow(index + 1);
    }
  });
}

function deleteMessage_(body) {
  const actor = body.authUser || {};
  if (String(actor.role || actor.type || "").toLowerCase() !== "admin") {
    return { ok: false, error: "Only a school administrator can delete messages." };
  }
  const messageId = String(body.messageId || "").trim();
  if (!messageId) return { ok: false, error: "Choose a message to delete." };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const item = findMessageRow_(messageId);
    if (!item) return { ok: false, error: "The message is no longer available." };
    item.sheet.deleteRow(item.rowNumber);
    removeMessageReferences_(messageId);
    return { ok: true, deleted: 1, messageId: messageId };
  } finally {
    lock.releaseLock();
  }
}

function clearNotifications_(body) {
  const actor = body.authUser || {};
  if (String(actor.role || actor.type || "").toLowerCase() !== "admin") {
    return { ok: false, error: "Only a school administrator can clear notifications." };
  }
  const email = String(actor.email || "").trim().toLowerCase();
  if (!email) return { ok: false, error: "The administrator account has no email address." };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const messageIds = [...new Set(getMessages_({ authUser: actor }).messages
      .map(message => String(message.id || "").trim())
      .filter(Boolean))];
    const sheet = ensureHeaders_(NOTIFICATION_CLEARS, ["messageId", "userEmail", "clearedAt"]);
    const cleared = readRows_(NOTIFICATION_CLEARS);
    const existing = new Set(cleared
      .filter(row => String(row.userEmail || "").trim().toLowerCase() === email)
      .map(row => String(row.messageId || "")));
    let added = 0;
    messageIds.forEach(messageId => {
      if (existing.has(messageId)) return;
      appendRecord_(sheet, { messageId: messageId, userEmail: email, clearedAt: new Date() });
      existing.add(messageId);
      added++;
    });
    return { ok: true, cleared: added, clearedMessageIds: [...existing] };
  } finally {
    lock.releaseLock();
  }
}

function getSheet_(name, headers) {
  const spreadsheet = requestSpreadsheetId
    ? openSpreadsheetById_(requestSpreadsheetId)
    : SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName(name);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }

  if (sheet.getLastRow() === 0 && headers.length) {
    sheet.appendRow(headers);
  }

  return sheet;
}

function findUserByEmail_(email) {
  const normalizedEmail = String(email || "").toLowerCase().trim();
  if (!normalizedEmail) return null;
  return readRows_(USERS).filter(row =>
    String(row.email || "").toLowerCase().trim() === normalizedEmail
  ).pop() || null;
}

function getSessionContext_(token) {
  const raw = CacheService.getScriptCache().get(`session:${String(token || "")}`);
  if (!raw) return null;
  try {
    const session = JSON.parse(raw);
    if (!session.email || !session.spreadsheetId || !session.tenantId) return null;
    requestSpreadsheetId = session.spreadsheetId;
    const account = findUserByEmail_(session.email);
    if (!account || (Number(session.authVersion) || 0) !== (Number(account.sessionVersion) || 0)) return null;
    return { session: session, account: account };
  } catch (error) {
    console.error(`Invalid session cache entry: ${error.message}`);
    return null;
  }
}

function getSession_(token) {
  const context = getSessionContext_(token);
  return context ? context.session : null;
}

function openSpreadsheetById_(spreadsheetId) {
  const id = String(spreadsheetId || "");
  if (!id) throw new Error("A spreadsheet ID is required.");
  if (!requestSpreadsheetCache) requestSpreadsheetCache = new Map();
  if (!requestSpreadsheetCache.has(id)) {
    requestSpreadsheetCache.set(id, SpreadsheetApp.openById(id));
  }
  return requestSpreadsheetCache.get(id);
}

function hashPassword_(password, salt) {
  let digest = `${salt}:${password}`;
  for (let round = 0; round < 12000; round++) {
    digest = Utilities.base64EncodeWebSafe(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, digest, Utilities.Charset.UTF_8)
    );
  }
  return digest;
}

function cacheKey_(value) {
  return Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value), Utilities.Charset.UTF_8)
  ).slice(0, 40);
}

function ensureSchoolSchema_(spreadsheetId) {
  requestSpreadsheetId = spreadsheetId;
  ensureHeaders_(USERS, [
    "id", "name", "email", "password", "passwordHash", "passwordSalt", "role", "type",
    "jobRole", "department", "className", "childStudentId", "profilePicture"
  ]);
  ensureHeaders_(MESSAGES, messageSheetHeaders_());
  ensureHeaders_(PAYMENTS, ["id", "studentId", "studentName", "parentName", "parentEmail", "amount", "paidFor", "paymentMethod", "reference", "paymentDate", "createdAt"]);
  ensureHeaders_(RECORDS, ["id", "type", "dataJson", "updatedBy", "updatedAt"]);
  Object.keys(RECORD_SHEETS).forEach(type => ensureHeaders_(RECORD_SHEETS[type], ["id"]));
  SYSTEM_SHEETS.forEach(name => ensureHeaders_(name, ["message"]));
  const spreadsheet = openSpreadsheetById_(spreadsheetId);
  const oldMessages = spreadsheet.getSheetByName("Messages");
  const inbox = spreadsheet.getSheetByName(MESSAGES);
  if (oldMessages && oldMessages.getLastRow() > 1 && inbox.getLastRow() <= 1) {
    const oldValues = oldMessages.getDataRange().getValues();
    const oldHeaders = oldValues[0].map(String);
    const newHeaders = inbox.getRange(1, 1, 1, inbox.getLastColumn()).getValues()[0].map(String);
    const migrated = oldValues.slice(1).map(row => {
      const oldRecord = rowToObject_(oldHeaders, row);
      return newHeaders.map(header => oldRecord[header] === undefined ? "" : oldRecord[header]);
    });
    if (migrated.length) writeSafeRows_(inbox, inbox.getLastRow() + 1, migrated);
  }
}

function disableSchoolSpreadsheetTriggers() {
  const schoolHandlers = new Set(["onSchoolSheetEdit_", "onSchoolSheetChange_"]);
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (!schoolHandlers.has(trigger.getHandlerFunction())) return;
    ScriptApp.deleteTrigger(trigger);
    removed += 1;
  });
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty("PENDING_SCHOOL_TRIGGER_SPREADSHEET_IDS");
  return `Removed ${removed} school spreadsheet trigger(s). Manage school records through Clasten.`;
}

function authorizeSchoolTriggerManagement() {
  return disableSchoolSpreadsheetTriggers();
}

function sheetRecordType_(sheetName) {
  return Object.keys(RECORD_SHEETS).find(type => RECORD_SHEETS[type] === sheetName) || "";
}

function rowToObject_(headers, row) {
  const record = {};
  headers.forEach((header, index) => {
    if (!header) return;
    let value = row[index];
    if (typeof value === "string" && /^[\[{]/.test(value.trim())) {
      try { value = JSON.parse(value); } catch (error) {
        console.warn(`Could not parse JSON value in school sheet column "${header}": ${error.message}`);
      }
    }
    record[header] = value;
  });
  return record;
}

function isValidDateKey_(value) {
  const parts = String(value || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!parts) return false;
  const year = Number(parts[1]);
  const month = Number(parts[2]);
  const day = Number(parts[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day;
}

function isWeekdayDateKey_(value) {
  const parts = String(value || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!parts) return false;
  const year = Number(parts[1]);
  const month = Number(parts[2]);
  const day = Number(parts[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day &&
    date.getUTCDay() >= 1 &&
    date.getUTCDay() <= 5;
}

function readSchoolRecords_(type) {
  const sheet = ensureHeaders_(RECORD_SHEETS[type], ["id"]);
  const values = readSheetValues_(sheet);
  if (values.length < 2) return [];
  const headers = values[0].map(header => String(header).trim());
  const idColumn = headers.indexOf("id");
  return values.slice(1)
    .filter(row => String(row[idColumn] || "").trim())
    .map(row => {
      const record = rowToObject_(headers, row);
      if (type === "attendance" && record.date instanceof Date) {
        record.date = Utilities.formatDate(record.date, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      return record;
    });
}

function ensureDefaultCampus_() {
  let campuses = readSchoolRecords_("campus");
  if (campuses.some(campus => String(campus.id || "") === DEFAULT_CAMPUS_ID)) {
    return campuses;
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    campuses = readSchoolRecords_("campus");
    if (!campuses.some(campus => String(campus.id || "") === DEFAULT_CAMPUS_ID)) {
      writeSchoolRecord_("campus", {
        id: DEFAULT_CAMPUS_ID,
        name: "Main campus",
        address: "",
        isDefault: true,
        createdAt: new Date().toISOString()
      });
      campuses = readSchoolRecords_("campus");
    }
    return campuses;
  } finally {
    lock.releaseLock();
  }
}

function validatedCampusId_(tenant, requestedCampusId) {
  const campuses = ensureDefaultCampus_();
  if (!planSupportsFeature_(tenant && tenant.plan, "multiCampus")) return DEFAULT_CAMPUS_ID;
  const campusId = String(requestedCampusId || DEFAULT_CAMPUS_ID).trim();
  return campuses.some(campus => String(campus.id || "") === campusId) ? campusId : "";
}

function effectiveCampusId_(record) {
  return String(record && record.campusId || DEFAULT_CAMPUS_ID).trim() || DEFAULT_CAMPUS_ID;
}

function schoolFeePlans_(feePlans, actor) {
  let normalized = normalizeSchoolFeePlans_(feePlans || []);
  if (!actor || !normalized.updates.length) return normalized.plans;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    normalized = normalizeSchoolFeePlans_(readSchoolRecords_("feePlan"));
    normalized.updates.forEach(update => {
      writeSchoolRecord_("feePlan", update.record);
      logAudit_(
        actor,
        "standardized school-wide fee type",
        "feePlan",
        update.original.id,
        update.original,
        update.record,
        `${update.sharedName} now uses the Main campus amount across all campuses.`
      );
    });
    return normalized.plans;
  } finally {
    lock.releaseLock();
  }
}

function normalizeSchoolFeePlans_(plans) {
  const groups = new Map();
  plans.forEach(plan => {
    const name = String(plan.name || "").trim();
    const key = name ? name.toLowerCase() : `id:${String(plan.id || "")}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(plan);
  });
  const sharedPlans = [], updates = [];
  groups.forEach(group => {
    const canonical = group.find(plan => effectiveCampusId_(plan) === DEFAULT_CAMPUS_ID) || group[0];
    const sharedName = String(canonical.name || "").trim();
    const sharedAmount = canonical.amount;
    const sharedDescription = String(canonical.description || "").trim();
    const sharedActive = canonical.active !== false &&
      String(canonical.active || "Active").toLowerCase() !== "inactive";
    const canonicalPlan = {
      ...canonical,
      name: sharedName,
      amount: sharedAmount,
      description: sharedDescription,
      active: sharedActive
    };
    delete canonicalPlan.campusId;
    sharedPlans.push(canonicalPlan);
    group.forEach(plan => {
      const record = {
        ...plan,
        name: sharedName,
        amount: sharedAmount,
        description: sharedDescription,
        active: sharedActive
      };
      delete record.campusId;
      const changed = String(plan.name || "") !== sharedName ||
        String(plan.amount) !== String(sharedAmount) ||
        String(plan.description || "") !== sharedDescription ||
        (plan.active !== false && String(plan.active || "Active").toLowerCase() !== "inactive") !== sharedActive ||
        Boolean(plan.campusId);
      if (changed) updates.push({ original: plan, record: record, sharedName: sharedName });
    });
  });
  return { plans: sharedPlans, updates: updates };
}

function studentLimitForPlan_(plan) {
  const planId = String(plan || "").toLowerCase();
  if (planId === "standard") return 200;
  return planId === "enterprise" ? 0 : 100;
}

function isActiveStudent_(student) {
  return !["archived", "withdrawn", "inactive"].includes(String(student && student.status || "active").trim().toLowerCase());
}

function activeStudentCount_(students) {
  return (students || []).filter(isActiveStudent_).length;
}

function studentLimitError_(plan, activeCount) {
  const limit = studentLimitForPlan_(plan);
  if (!limit || activeCount < limit) return "";
  const planName = String(plan || "school").toUpperCase();
  return `The ${planName} plan allows up to ${limit} active students. Archive or withdraw a student, or upgrade the school plan to enroll another student.`;
}

function nextStudentId_(students) {
  const year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy");
  const prefix = `SCH-${year}-`;
  const sequence = students.reduce((maximum, student) => {
    const match = String(student.studentId || "").match(new RegExp(`^SCH-${year}-(\\d+)$`));
    return match ? Math.max(maximum, Number(match[1])) : maximum;
  }, 0) + 1;
  return `${prefix}${String(sequence).padStart(4, "0")}`;
}

function writeSchoolRecord_(type, record) {
  if (!RECORD_SHEETS[type] || !record || !record.id) {
    throw new Error("Cannot save this record without a valid record type and ID.");
  }
  const safeRecord = { ...record };
  delete safeRecord.password;
  delete safeRecord.passwordHash;
  delete safeRecord.passwordSalt;
  delete safeRecord.passport;
  const sheet = ensureHeaders_(RECORD_SHEETS[type], ["id"]);
  const values = sheet.getDataRange().getValues();
  const existingHeaders = values[0].map(String);
  const newHeaders = Object.keys(safeRecord).filter(key => !existingHeaders.includes(key));
  if (newHeaders.length) {
    setSafeSheetValues_(sheet.getRange(1, existingHeaders.length + 1, 1, newHeaders.length), [newHeaders]);
    existingHeaders.push(...newHeaders);
  }
  const idColumn = existingHeaders.indexOf("id");
  const rowIndex = values.findIndex((row, index) =>
    index > 0 && String(row[idColumn]) === String(safeRecord.id)
  );
  const updated = rowIndex >= 0;
  if (updated && type === "parent") {
    ["activationTokenHash", "activationTokenExpiresAt"].forEach(key => {
      if (!Object.prototype.hasOwnProperty.call(safeRecord, key)) {
        const column = existingHeaders.indexOf(key);
        safeRecord[key] = column >= 0 ? values[rowIndex][column] || "" : "";
      }
    });
  }
  writeRecord_(sheet, updated ? rowIndex + 1 : values.length + 1, safeRecord, existingHeaders);
  return updated;
}

function syncAccountFromSchoolRecord_(type, record, previousEmail) {
  const email = String(record.email || "").toLowerCase().trim();
  const previous = String(previousEmail || "").toLowerCase().trim();
  if (!email) {
    if (previous) deleteSchoolUser_(previous, type);
    return;
  }
  const sheet = ensureHeaders_(USERS, [
    "id", "name", "email", "password", "passwordHash", "passwordSalt", "role", "type",
    "jobRole", "department", "className", "childStudentId", "profilePicture", "sessionVersion"
  ]);
  const rows = readRows_(USERS);
  const index = rows.findIndex(row => {
    const accountEmail = String(row.email || "").toLowerCase().trim();
    return accountEmail === email || (previous && accountEmail === previous);
  });
  const existing = index >= 0 ? rows[index] : null;
  const conflictingAccount = rows.find((row, rowIndex) =>
    rowIndex !== index && String(row.email || "").toLowerCase().trim() === email
  );
  if (conflictingAccount) throw new Error(`The email ${email} is already assigned to another school account.`);
  const role = type;
  const password = type === "employee" ? "" : String(record.password || "");
  if (type === "employee" && !existing) return;
  if (!existing && password.length < 12) {
    throw new Error(`Set a password of at least 12 characters to create this ${type} login.`);
  }
  const salt = password ? Utilities.getUuid() : existing && existing.passwordSalt;
  const account = {
    id: existing && existing.id || email,
    name: record.name || "",
    email: email,
    password: "",
    passwordHash: password ? hashPassword_(password, salt) : existing && existing.passwordHash || "",
    passwordSalt: salt || "",
    sessionVersion: (Number(existing && existing.sessionVersion) || 0) + (password && existing ? 1 : 0),
    role: role,
    type: role,
    jobRole: type === "employee" ? record.role || "" : "",
    department: type === "employee" ? record.department || "" : "",
    className: type === "employee" ? record.className || "" : "",
    childStudentId: type === "parent" ? record.childStudentId || "" : "",
    profilePicture: existing && existing.profilePicture || ""
  };
  if (index >= 0) writeRecord_(sheet, index + 2, account);
  else appendRecord_(sheet, account);
}

function deleteSchoolUser_(email, role) {
  const normalizedEmail = String(email || "").toLowerCase().trim();
  if (!normalizedEmail) return;
  const sheet = ensureHeaders_(USERS, ["id", "name", "email", "password", "passwordHash", "passwordSalt", "role", "type"]);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0].map(String);
  const emailColumn = headers.indexOf("email");
  const roleColumn = headers.indexOf("role");
  for (let index = rows.length - 1; index > 0; index--) {
    if (String(rows[index][emailColumn] || "").toLowerCase().trim() === normalizedEmail &&
        String(rows[index][roleColumn] || "").toLowerCase() === role) {
      sheet.deleteRow(index + 1);
    }
  }
}

function onSchoolSheetEdit_(event) {
  if (!event || !event.range || !event.source) return;
  requestSpreadsheetId = event.source.getId();
  const sheet = event.range.getSheet();
  const type = sheetRecordType_(sheet.getName());
  if (!type || event.range.getRow() === 1) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const idColumn = headers.indexOf("id") + 1;
  if (idColumn < 1) return;
  const rowNumber = event.range.getRow();
  if (rowNumber > sheet.getLastRow()) return;
  let id = String(sheet.getRange(rowNumber, idColumn).getValue() || "").trim();
  if (!id) {
    id = `${type}-${Utilities.getUuid()}`;
    sheet.getRange(rowNumber, idColumn).setValue(id);
  }
  const record = rowToObject_(headers, sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0]);
  if (type === "student" && !String(record.studentId || "").trim()) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      record.studentId = nextStudentId_(readSchoolRecords_("student"));
      let studentIdColumn = headers.indexOf("studentId") + 1;
      if (studentIdColumn < 1) {
        studentIdColumn = sheet.getLastColumn() + 1;
        sheet.getRange(1, studentIdColumn).setValue("studentId");
      }
      sheet.getRange(rowNumber, studentIdColumn).setValue(record.studentId);
    } finally {
      lock.releaseLock();
    }
  } else if (type === "attendance") {
    const student = readSchoolRecords_("student").find(item =>
      String(item.studentId || "").trim() === String(record.studentId || "").trim()
    );
    if (student) {
      record.studentName = `${student.firstName || ""} ${student.lastName || ""}`.trim();
      record.className = String(student.className || "");
      const status = String(record.status || "").trim().toLowerCase();
      if (["present", "absent", "late", "excused"].includes(status)) {
        record.status = status.charAt(0).toUpperCase() + status.slice(1);
        record.present = status === "present" ? 1 : 0;
        record.absent = status === "absent" ? 1 : 0;
      }
      writeSchoolRecord_(type, record);
    }
  } else if (type === "employee" || type === "parent") {
    const emailColumn = headers.indexOf("email") + 1;
    const previousEmail = emailColumn > 0 && event.range.getColumn() === emailColumn
      ? String(event.oldValue || "").toLowerCase().trim()
      : "";
    syncAccountFromSchoolRecord_(type, record, previousEmail);
  }
}

function onSchoolSheetChange_(event) {
  if (event && event.source) requestSpreadsheetId = event.source.getId();
  if (event && event.changeType === "REMOVE_ROW") {
    syncSchoolAccountsFromRoster_();
  }
}

function syncSchoolAccountsFromRoster_() {
  const activeEmails = {};
  ["employee", "parent"].forEach(type => {
    readSchoolRecords_(type).forEach(record => {
      const email = String(record.email || "").toLowerCase().trim();
      if (email) activeEmails[`${type}:${email}`] = true;
    });
  });
  const users = ensureHeaders_(USERS, [
    "id", "name", "email", "password", "passwordHash", "passwordSalt", "role", "type",
    "jobRole", "department", "className", "childStudentId", "profilePicture"
  ]);
  const rows = users.getDataRange().getValues();
  const headers = rows[0].map(String);
  const emailColumn = headers.indexOf("email");
  const roleColumn = headers.indexOf("role");
  for (let index = rows.length - 1; index > 0; index--) {
    const role = String(rows[index][roleColumn] || "").toLowerCase();
    const email = String(rows[index][emailColumn] || "").toLowerCase().trim();
    if (["employee", "parent"].includes(role) && email && !activeEmails[`${role}:${email}`]) {
      users.deleteRow(index + 1);
    }
  }
}

function registrySheet_() {
  const registryId = String(PropertiesService.getScriptProperties().getProperty("MASTER_SPREADSHEET_ID") || "");
  if (!registryId) throw new Error("Run setup() to configure the school registry.");
  return ensureHeadersIn_(openSpreadsheetById_(registryId), TENANTS, [
    "tenantId", "schoolName", "schoolCode", "spreadsheetId", "adminEmail",
    "plan", "billingPeriod", "subscriptionStatus", "trialEndsAt", "paidThrough",
    "createdAt", "portalLogo", "schoolProfilePicture", "portalBackgroundImage", "brandColor", "brandingSetupCompleted",
    "termsVersion", "termsAcceptedAt"
  ]);
}

function ensureHeadersIn_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  const existing = sheet.getLastColumn()
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    : [];
  headers.forEach(header => {
    if (!existing.includes(header)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      existing.push(header);
    }
  });
  return sheet;
}

function normalizeSchoolCode_(value) {
  return String(value || "").toLowerCase().trim().replace(/[^a-z0-9-]/g, "").replace(/-+/g, "-").replace(/^-|-$/g, "");
}

function readRowsFrom_(sheet) {
  const values = readSheetValues_(sheet);
  if (values.length < 2) return [];
  return values.slice(1).map(row => {
    const record = {};
    values[0].forEach((header, index) => record[header] = row[index]);
    return record;
  });
}

function findTenant_(schoolCode) {
  const normalizedCode = normalizeSchoolCode_(schoolCode);
  if (!normalizedCode) return null;
  const tenant = tenantRows_().find(row =>
    normalizeSchoolCode_(row.schoolCode) === normalizedCode
  );
  return tenant ? { ...tenant } : null;
}

function findTenantBySpreadsheetId_(spreadsheetId) {
  const id = String(spreadsheetId || "");
  if (!id) return null;
  const tenant = tenantRows_().find(row =>
    String(row.spreadsheetId || "") === id
  );
  return tenant ? { ...tenant } : null;
}

function tenantRows_() {
  if (!requestTenantRows) requestTenantRows = readRowsFrom_(registrySheet_());
  return requestTenantRows;
}

function upsertTenant_(tenant) {
  const sheet = registrySheet_();
  const rows = readRowsFrom_(sheet);
  const rowIndex = rows.findIndex(row =>
    normalizeSchoolCode_(row.schoolCode) === normalizeSchoolCode_(tenant.schoolCode)
  );
  const preserved = rowIndex >= 0 ? rows[rowIndex] : {};
  const record = rowIndex >= 0
    ? {
      ...tenant,
      plan: preserved.plan || tenant.plan,
      billingPeriod: preserved.billingPeriod || tenant.billingPeriod,
      subscriptionStatus: preserved.subscriptionStatus || tenant.subscriptionStatus,
      trialEndsAt: preserved.trialEndsAt || tenant.trialEndsAt,
      paidThrough: preserved.paidThrough || tenant.paidThrough,
      createdAt: preserved.createdAt || tenant.createdAt
    }
    : tenant;
  writeRecord_(sheet, rowIndex >= 0 ? rowIndex + 2 : sheet.getLastRow() + 1, record);
  requestTenantRows = null;
}

function updateTenant_(schoolCode, updates) {
  const sheet = registrySheet_();
  const rows = readRowsFrom_(sheet);
  const index = rows.findIndex(row =>
    normalizeSchoolCode_(row.schoolCode) === normalizeSchoolCode_(schoolCode)
  );
  if (index < 0) throw new Error("School account could not be found.");
  const record = { ...rows[index], ...updates };
  writeRecord_(sheet, index + 2, record);
  requestTenantRows = null;
  return record;
}

function isSubscriptionUsable_(tenant) {
  const status = String(tenant.subscriptionStatus || "").toLowerCase();
  if (status === "active") return !tenant.paidThrough || new Date(tenant.paidThrough).getTime() >= Date.now();
  if (status === "trial") return new Date(tenant.trialEndsAt).getTime() >= Date.now();
  return false;
}

function signupRequest_(body) {
  const termsVersion = String(body.termsVersion || "").trim();
  if (body.termsAccepted !== true || termsVersion !== TERMS_VERSION) {
    return { ok: false, error: "Please review and accept the current Clasten terms before signing up." };
  }
  const schoolName = String(body.schoolName || "").trim();
  const adminName = String(body.adminName || "").trim();
  const email = String(body.email || "").trim().toLowerCase();
  const password = String(body.password || "");
  const plan = String(body.plan || "standard").toLowerCase();
  const signupMode = String(body.signupMode || "trial").toLowerCase();
  let billingPeriod = String(body.billingPeriod || "annual");
  if (billingPeriod === "six_months") billingPeriod = "per_term";
  if (schoolName.length < 2 || adminName.length < 2 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { ok: false, error: "Enter the school name, administrator name, and a valid email address." };
  }
  if (password.length < 12) return { ok: false, error: "Choose a password with at least 12 characters." };
  if (!["basic", "standard", "enterprise"].includes(plan)) return { ok: false, error: "Choose a valid school plan." };
  if (!["annual", "per_term"].includes(billingPeriod)) return { ok: false, error: "Choose a valid billing period." };
  if (signupMode === "paid") {
    const selectedPlan = getPlans_().find(item => item.id === plan);
    const price = billingPeriod === "per_term" ? selectedPlan && selectedPlan.perTermPrice : selectedPlan && selectedPlan.annualPrice;
    if (!price || price <= 0) return { ok: false, error: "Pricing for this plan and billing period is not configured yet. Choose another option or start a free trial." };
  }
  const requestedCode = normalizeSchoolCode_(body.schoolCode || schoolName);
  const schoolCode = requestedCode.length < 3
    ? ""
    : `${requestedCode}-${Utilities.getUuid().slice(0, 4).replace(/-/g, "").toLowerCase()}`;
  if (!schoolCode) return { ok: false, error: "Enter a school code with at least three letters or numbers." };
  if (findTenant_(schoolCode)) return { ok: false, error: "That school code is already in use. Choose another." };
  const cache = CacheService.getScriptCache();
  const key = `signup:${cacheKey_(email)}`;
  if (cache.get(key)) return { ok: false, error: "A verification email was already sent. Please wait a few minutes before requesting another." };
  const signupDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
  const counterKey = "SCHOOL_SIGNUP_DAILY_COUNTER";
  const properties = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  lock.waitLock(3000);
  try {
    if (cache.get(key)) return { ok: false, error: "A verification email was already sent. Please wait a few minutes before requesting another." };
    const storedCounter = properties.getProperty(counterKey);
    let dailyCount = 0;
    if (storedCounter) {
      let counter;
      try {
        counter = JSON.parse(storedCounter);
      } catch (error) {
        throw new Error(`The school signup counter could not be read: ${error.message}`);
      }
      if (counter.date === signupDate) dailyCount = Number(counter.count) || 0;
    }
    dailyCount = Math.max(dailyCount, Number(cache.get(`signup-daily:${signupDate}`) || 0));
    if (dailyCount >= 10) return { ok: false, error: "New school signups are temporarily paused. Please contact Clasten support." };
    properties.setProperty(counterKey, JSON.stringify({ date: signupDate, count: dailyCount + 1 }));
    cache.put(key, "1", 180);
  } finally {
    lock.releaseLock();
  }
  const code = String(Math.floor(100000 + Math.random() * 900000));
  try {
    cache.put(`verify:${key}`, JSON.stringify({
      schoolName: schoolName, adminName: adminName, email: email,
      password: password, schoolCode: schoolCode, code: code, plan: plan,
      signupMode: signupMode, billingPeriod: billingPeriod,
      termsVersion: TERMS_VERSION, termsAcceptedAt: new Date()
    }), 900);
    sendAccountCodeEmail_(
      email,
      "Verify your Clasten school account",
      "Use this code to verify your school account.",
      code
    );
  } catch (error) {
    cache.remove(`verify:${key}`);
    cache.remove(key);
    throw error;
  }
  return { ok: true, schoolCode: schoolCode, message: "We sent a verification code to your email address." };
}

function verifySignup_(body) {
  const email = String(body.email || "").trim().toLowerCase();
  const key = `signup:${cacheKey_(email)}`;
  const cache = CacheService.getScriptCache();
  const pending = cache.get(`verify:${key}`);
  if (!pending) return { ok: false, error: "The verification code expired. Start signup again." };
  const signup = JSON.parse(pending);
  const attemptsKey = `verify-attempts:${key}`;
  const attempts = Number(cache.get(attemptsKey) || 0);
  if (attempts >= 5) return { ok: false, error: "Too many incorrect verification attempts. Start signup again." };
  if (String(body.code || "").trim() !== signup.code) {
    cache.put(attemptsKey, String(attempts + 1), 900);
    return { ok: false, error: "That verification code is incorrect." };
  }
  const properties = PropertiesService.getScriptProperties();
  const provisioningKey = `school-provisioning:${cacheKey_(signup.schoolCode)}`;
  const provisioningToken = Utilities.getUuid();
  let existingTenant, claimedProvisioning = false;
  const provisioningLock = LockService.getScriptLock();
  provisioningLock.waitLock(5000);
  try {
    existingTenant = findTenant_(signup.schoolCode);
    if (!existingTenant) {
      const currentReservation = properties.getProperty(provisioningKey);
      if (currentReservation) {
        let expiresAt = 0;
        try {
          expiresAt = Number(JSON.parse(currentReservation).expiresAt) || 0;
        } catch (error) {
          console.warn(`Invalid school provisioning reservation: ${error.message}`);
        }
        if (expiresAt > Date.now()) {
          return { ok: false, error: "School workspace setup is already in progress. Wait a few minutes, then submit the same verification code again." };
        }
      }
      properties.setProperty(provisioningKey, JSON.stringify({
        token: provisioningToken,
        expiresAt: Date.now() + 10 * 60 * 1000
      }));
      claimedProvisioning = true;
    }
  } finally {
    provisioningLock.releaseLock();
  }
  if (existingTenant) {
    if (String(existingTenant.adminEmail || "").toLowerCase() !== email) {
      return { ok: false, error: "That school code is now in use. Contact support to complete setup." };
    }
    updateTenant_(signup.schoolCode, {
      termsVersion: signup.termsVersion,
      termsAcceptedAt: signup.termsAcceptedAt
    });
    const setupWarning = SCHOOL_DATA_EDIT_NOTICE;
    cache.remove(`verify:${key}`);
    cache.remove(attemptsKey);
    cache.remove(key);
    return {
      ok: true,
      schoolCode: signup.schoolCode,
      message: signup.signupMode === "paid"
        ? "Your school account was already created. Sign in to pay for your selected plan and open your school workspace."
        : "Your school account was already created. Sign in with your school code and administrator account.",
      signupMode: signup.signupMode || "trial",
      setupWarning: setupWarning
    };
  }

  const spreadsheet = SpreadsheetApp.create(`${signup.schoolName} - Clasten`);
  ensureSchoolSchema_(spreadsheet.getId());
  const salt = Utilities.getUuid();
  appendRecord_(ensureHeaders_(USERS, [
    "id", "name", "email", "password", "passwordHash", "passwordSalt", "role", "type",
    "jobRole", "department", "className", "childStudentId", "profilePicture"
  ]), {
    id: email, name: signup.adminName, email: email, password: "",
    passwordHash: hashPassword_(signup.password, salt), passwordSalt: salt,
    role: "admin", type: "employee", jobRole: "", department: "",
    className: "", childStudentId: "", profilePicture: ""
  });
  upsertTenant_({
    tenantId: signup.schoolCode, schoolName: signup.schoolName,
    schoolCode: signup.schoolCode, spreadsheetId: spreadsheet.getId(),
    adminEmail: email, plan: signup.plan, billingPeriod: signup.billingPeriod || "annual",
    subscriptionStatus: signup.signupMode === "paid" ? "pending_payment" : "trial",
    trialEndsAt: signup.signupMode === "paid" ? "" : new Date(Date.now() + 14 * 24 * 60 * 60 * 1000),
    paidThrough: "", createdAt: new Date(),
    termsVersion: signup.termsVersion, termsAcceptedAt: signup.termsAcceptedAt
  });
  if (claimedProvisioning) {
    const reservation = JSON.parse(properties.getProperty(provisioningKey) || "{}");
    if (reservation.token === provisioningToken) properties.deleteProperty(provisioningKey);
  }
  const setupWarning = SCHOOL_DATA_EDIT_NOTICE;
  cache.remove(`verify:${key}`);
  cache.remove(attemptsKey);
  cache.remove(key);
  return {
    ok: true,
    schoolCode: signup.schoolCode,
    message: signup.signupMode === "paid"
      ? "School account created. Sign in to pay for your selected plan and open your school workspace."
      : "School account created. Sign in with your school code and administrator account.",
    signupMode: signup.signupMode || "trial",
    setupWarning: setupWarning
  };
}

function requestPasswordReset_(body) {
  const tenant = findTenant_(body.schoolCode);
  const email = String(body.email || "").trim().toLowerCase();
  const generic = "If the account exists, a password reset code has been emailed.";
  if (!tenant || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: true, message: generic };
  const key = `reset:${tenant.tenantId}:${cacheKey_(email)}`;
  const cache = CacheService.getScriptCache();
  if (cache.get(`${key}:rate`)) return { ok: true, message: generic };
  requestSpreadsheetId = String(tenant.spreadsheetId);
  const account = findUserByEmail_(email);
  if (account) {
    const code = String(Math.floor(100000 + Math.random() * 900000));
    cache.put(key, code, 900);
    try {
      sendAccountCodeEmail_(
        email,
        "Reset your Clasten password",
        "Use this one-time code to reset your Clasten password.",
        code
      );
    } catch (error) {
      cache.remove(key);
      throw error;
    }
  }
  cache.put(`${key}:rate`, "1", 180);
  return { ok: true, message: generic };
}

function sendAccountCodeEmail_(email, subject, introduction, code) {
  const expiry = "This code expires in 15 minutes. If you did not request it, you can ignore this email.";
  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      name: "Clasten",
      body: `${introduction}\n\n${code}\n\n${expiry}`,
      htmlBody: `<div style="font-family:Arial,sans-serif;color:#172033;max-width:520px;margin:0 auto;padding:24px"><div style="font-size:12px;font-weight:bold;letter-spacing:2px;color:#3762ea">CLASTEN</div><h2 style="margin:20px 0 8px">${introduction}</h2><p>Enter this one-time code in the Clasten page:</p><div style="display:inline-block;padding:14px 22px;background:#f0f3ff;border-radius:10px;font-size:28px;font-weight:bold;letter-spacing:6px;color:#2447c8">${code}</div><p style="color:#758096;font-size:13px;margin-top:22px">${expiry}</p></div>`
    });
  } catch (error) {
    console.error(`Could not send account code to ${email}: ${error.message}`);
    throw new Error("The school server could not send the email. Ask the Apps Script owner to run authorizeMailService() and check its daily email quota, then request a new code.");
  }
}

function resetPassword_(body) {
  const tenant = findTenant_(body.schoolCode);
  const email = String(body.email || "").trim().toLowerCase();
  const password = String(body.password || "");
  const key = tenant ? `reset:${tenant.tenantId}:${cacheKey_(email)}` : "";
  const cache = CacheService.getScriptCache();
  const attemptsKey = key ? `${key}:attempts` : "";
  const attempts = attemptsKey ? Number(cache.get(attemptsKey) || 0) : 0;
  if (attempts >= 5) return { ok: false, error: "Too many incorrect reset attempts. Request a new code." };
  if (!tenant || !key || cache.get(key) !== String(body.code || "").trim()) {
    if (attemptsKey) cache.put(attemptsKey, String(attempts + 1), 900);
    return { ok: false, error: "The school, email, or reset code is invalid or expired." };
  }
  if (password.length < 12) return { ok: false, error: "Choose a password with at least 12 characters." };
  requestSpreadsheetId = String(tenant.spreadsheetId);
  const sheet = ensureHeaders_(USERS, [
    "id", "name", "email", "password", "passwordHash", "passwordSalt", "role", "type",
    "jobRole", "department", "className", "childStudentId", "profilePicture", "sessionVersion"
  ]);
  const rows = readRows_(USERS);
  const index = rows.findIndex(row => String(row.email || "").toLowerCase().trim() === email);
  if (index < 0) return { ok: false, error: "The school, email, or reset code is invalid or expired." };
  const salt = Utilities.getUuid();
  const account = {
    ...rows[index],
    password: "",
    passwordHash: hashPassword_(password, salt),
    passwordSalt: salt,
    sessionVersion: (Number(rows[index].sessionVersion) || 0) + 1
  };
  writeRecord_(sheet, index + 2, account);
  cache.remove(key);
  cache.remove(attemptsKey);
  return { ok: true, message: "Password updated. You can now sign in." };
}

function getPlans_(options) {
  const properties = PropertiesService.getScriptProperties();
  const annualDiscountEligible = !options || options.annualDiscountEligible !== false;
  const annualDiscountPercent = key => {
    const planKey = String(key || "").toUpperCase();
    if (!["STANDARD", "ENTERPRISE"].includes(planKey)) return 0;
    const propertyKey = `${planKey}_ANNUAL_DISCOUNT_PERCENT`;
    const configuredPercent = properties.getProperty(propertyKey);
    if (configuredPercent === null || String(configuredPercent).trim() === "") return 20;
    const percent = Number(String(configuredPercent).trim());
    if (!Number.isInteger(percent) || percent < 0 || percent > 99) {
      throw new Error(`${propertyKey} must be a whole percentage from 0 to 99.`);
    }
    return percent;
  };
  const configuredPrice = key => {
    const price = Number(properties.getProperty(key) || 0);
    return Number.isFinite(price) && price > 0 ? price : 0;
  };
  const planPrices = key => {
    const perTermPriceKey = `${key}_PER_TERM_PRICE`;
    const perTermPriceProperty = properties.getProperty(perTermPriceKey);
    const hasPerTermPrice = perTermPriceProperty !== null &&
      String(perTermPriceProperty).trim() !== "";
    const perTermPrice = hasPerTermPrice
      ? configuredPrice(perTermPriceKey)
      : configuredPrice(`${key}_6_MONTH_PRICE`);
    const annualPriceKey = `${key}_12_MONTH_PRICE`;
    const annualPriceProperty = properties.getProperty(annualPriceKey);
    const configuredAnnualPrice = configuredPrice(annualPriceKey);
    const hasConfiguredAnnualPrice = annualPriceProperty !== null &&
      String(annualPriceProperty).trim() !== "";
    const configuredDiscountPercent = annualDiscountEligible ? annualDiscountPercent(key) : 0;
    const discountRate = configuredDiscountPercent / 100;
    const annualOriginalPrice = hasConfiguredAnnualPrice ? configuredAnnualPrice : 0;
    const annualPrice = annualOriginalPrice > 0
      ? Math.round(annualOriginalPrice * (1 - discountRate) * 100) / 100
      : 0;
    const annualSavings = discountRate > 0 && annualOriginalPrice > 0 && annualPrice > 0
      ? Math.max(0, Math.round((annualOriginalPrice - annualPrice) * 100) / 100)
      : 0;
    return {
      perTermPrice: perTermPrice,
      annualOriginalPrice: annualOriginalPrice,
      annualPrice: annualPrice,
      annualSavings: annualSavings,
      annualDiscountPercent: annualSavings > 0 ? configuredDiscountPercent : 0
    };
  };
  const commonFeatures = [
    "Student, employee, and parent records",
    "Attendance tracking",
    "Academic, term, and school calendar management",
    "Parent dashboards and linked student information",
    "Admissions applications, placements, and enrollment"
  ];
  const feeFeatures = commonFeatures.slice();
  feeFeatures[1] = "Attendance tracking, fee invoices, and Paystack payments";
  return [
    {
      id: "basic",
      name: "Basic",
      description: "Core school operations for one campus, with space for up to 100 active students.",
      features: commonFeatures.concat(["One campus", "Up to 100 active students"]),
      ...planPrices("BASIC")
    },
    {
      id: "standard",
      name: "Standard",
      description: "Online fee management and Paystack payments for one campus, with space for up to 200 active students.",
      features: feeFeatures.concat(["One campus", "Up to 200 active students"]),
      ...planPrices("STANDARD")
    },
    {
      id: "enterprise",
      name: "Enterprise",
      description: "For growing school groups: unlimited active students, multiple campuses, report cards, school-wide announcements, advanced CBT exams, and printable ID cards under one school subscription.",
      features: feeFeatures.concat([
        "Multiple campuses",
        "Unlimited active students",
        "Student report cards",
        "School-wide announcements",
        "CBT exams with question banks, timed attempts, scoring, written marking, and controlled result release",
        "Printable and reissuable student and employee ID cards"
      ]),
      ...planPrices("ENTERPRISE")
    }
  ];
}

function planSupportsFeature_(plan, feature) {
  const planId = String(plan || "").toLowerCase();
  if (feature === "fees") return ["standard", "enterprise"].includes(planId);
  if (feature === "multiCampus") return planId === "enterprise";
  if (["reportCards", "announcements", "cbt", "idCards"].includes(feature)) return planId === "enterprise";
  return false;
}

function subscriptionPaymentsSheet_() {
  const registryId = String(PropertiesService.getScriptProperties().getProperty("MASTER_SPREADSHEET_ID") || "");
  if (!registryId) throw new Error("Run setup() to configure the school registry.");
  return ensureHeadersIn_(openSpreadsheetById_(registryId), SUBSCRIPTION_PAYMENTS, [
    "reference", "schoolCode", "adminEmail", "plan", "billingPeriod",
    "amountNaira", "currency", "paidAt", "paidThrough", "paymentChannel",
    "ownerEmailStatus", "scholaroEmailStatus", "annualDiscountApplied"
  ]);
}

function hasUsedAnnualDiscount_(tenantId, paymentRows) {
  const rows = Array.isArray(paymentRows) ? paymentRows : readRowsFrom_(subscriptionPaymentsSheet_());
  return rows.some(payment =>
    normalizeSchoolCode_(payment.schoolCode) === normalizeSchoolCode_(tenantId) &&
    ["standard", "enterprise"].includes(String(payment.plan || "").toLowerCase()) &&
    String(payment.billingPeriod || "").toLowerCase() === "annual"
  );
}

function annualDiscountReservationKey_(tenantId) {
  return `annual-discount:${cacheKey_(normalizeSchoolCode_(tenantId))}`;
}

function clearAnnualDiscountReservation_(tenantId, token) {
  if (!token) return;
  const cache = CacheService.getScriptCache();
  const key = annualDiscountReservationKey_(tenantId);
  if (cache.get(key) === token) cache.remove(key);
}

function startSubscription_(body, tenantId) {
  const properties = PropertiesService.getScriptProperties();
  const secret = String(properties.getProperty("PAYSTACK_SECRET_KEY") || "");
  if (!secret) throw new Error("Subscription checkout is not configured. Contact Clasten support.");
  const scholaroBillingEmail = String(properties.getProperty("SCHOLARO_BILLING_EMAIL") || "").trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(scholaroBillingEmail)) {
    throw new Error("Subscription payment notifications are not configured. Contact Clasten support.");
  }
  const planId = String(body.plan || "").toLowerCase();
  let period = String(body.period || "");
  if (period === "six_months") period = "per_term";
  if (!["basic", "standard", "enterprise"].includes(planId) || !["per_term", "annual"].includes(period)) {
    throw new Error("Choose a valid subscription plan and billing period.");
  }
  const tenant = findTenant_(tenantId);
  if (!tenant) throw new Error("School account could not be found. Sign in again and retry payment.");
  if (String(tenant.subscriptionStatus || "").toLowerCase() === "active" &&
      String(tenant.plan || "").toLowerCase() === planId &&
      isSubscriptionUsable_(tenant)) {
    throw new Error("This subscription plan is already active. Choose a different plan to avoid another payment.");
  }
  const returnUrl = String(body.returnUrl || "").trim().replace(/[?#].*$/, "");
  if (!/^https?:\/\/[^/?#]+(?:\/[^?#]*)?$/i.test(returnUrl)) {
    throw new Error("The school website return address is invalid. Reopen the school website and try again.");
  }
  let plan, amount, annualDiscountApplied = false, reservationToken = "";
  const paymentLock = LockService.getScriptLock();
  paymentLock.waitLock(10000);
  try {
    let annualDiscountEligible = true;
    if (period === "annual" && ["standard", "enterprise"].includes(planId)) {
      const paymentRows = readRowsFrom_(subscriptionPaymentsSheet_());
      annualDiscountEligible = !hasUsedAnnualDiscount_(tenantId, paymentRows);
      if (annualDiscountEligible &&
          CacheService.getScriptCache().get(annualDiscountReservationKey_(tenantId))) {
        throw new Error("A first-year annual discount checkout is already in progress for this school. Complete it or wait for it to expire before trying again.");
      }
    }
    plan = getPlans_({ annualDiscountEligible: annualDiscountEligible }).find(item => item.id === planId);
    amount = period === "per_term" ? plan && plan.perTermPrice : plan && plan.annualPrice;
    if (!amount || amount <= 0) {
      throw new Error("Pricing for this plan has not been configured yet. Contact Clasten support.");
    }
    annualDiscountApplied = period === "annual" &&
      ["standard", "enterprise"].includes(planId) &&
      Number(plan.annualDiscountPercent) > 0 &&
      Number(plan.annualSavings) > 0;
    if (annualDiscountApplied) {
      reservationToken = Utilities.getUuid();
      CacheService.getScriptCache().put(
        annualDiscountReservationKey_(tenantId),
        reservationToken,
        21600
      );
    }
  } finally {
    paymentLock.releaseLock();
  }
  try {
    const response = UrlFetchApp.fetch("https://api.paystack.co/transaction/initialize", {
      method: "post", contentType: "application/json",
      headers: { Authorization: `Bearer ${secret}` },
      payload: JSON.stringify({
        email: tenant.adminEmail, amount: Math.round(amount * 100), currency: "NGN",
        callback_url: `${returnUrl}?paymentReturn=1&tenant=${encodeURIComponent(tenantId)}`,
        metadata: { tenantId: tenantId, plan: planId, period: period }
      }),
      muteHttpExceptions: true
    });
    const result = JSON.parse(response.getContentText());
    if (response.getResponseCode() >= 300 || !result.status || !result.data || !result.data.authorization_url) {
      throw new Error(result.message || "Could not start Paystack checkout.");
    }
    CacheService.getScriptCache().put(`pay:${result.data.reference}`, JSON.stringify({
      tenantId: tenantId, plan: planId, period: period,
      amountNaira: amount, annualDiscountApplied: annualDiscountApplied,
      annualDiscountReservationToken: reservationToken
    }), 21600);
    return { ok: true, checkoutUrl: result.data.authorization_url };
  } catch (error) {
    clearAnnualDiscountReservation_(tenantId, reservationToken);
    throw error;
  }
}

function verifySubscriptionPayment_(reference, tenantId) {
  const paymentReference = String(reference || "");
  if (!paymentReference) return { ok: false, error: "Payment reference is missing or expired." };
  const cache = CacheService.getScriptCache();
  const billingSheet = subscriptionPaymentsSheet_();
  const paymentRows = readRowsFrom_(billingSheet);
  const recordedPayment = paymentRows.find(row => String(row.reference || "") === paymentReference);
  if (recordedPayment) {
    if (normalizeSchoolCode_(recordedPayment.schoolCode) !== normalizeSchoolCode_(tenantId)) {
      return { ok: false, error: "Payment reference does not match this school." };
    }
    const recordedTenant = findTenant_(tenantId);
    if (!recordedTenant) return { ok: false, error: "School account could not be found." };
    const recordedPaidThrough = recordedPayment.paidThrough || "";
    if (!recordedPaidThrough) return { ok: false, error: "The recorded payment is missing its subscription end date. Contact Clasten support." };
    if (String(recordedTenant.plan || "") !== String(recordedPayment.plan || "") ||
        String(recordedTenant.billingPeriod || "") !== String(recordedPayment.billingPeriod || "") ||
        String(recordedTenant.subscriptionStatus || "").toLowerCase() !== "active" ||
        String(recordedTenant.paidThrough || "") !== String(recordedPaidThrough)) {
      updateTenant_(tenantId, {
        plan: recordedPayment.plan,
        billingPeriod: recordedPayment.billingPeriod,
        subscriptionStatus: "active",
        paidThrough: recordedPaidThrough
      });
    }
    const recordedPending = cache.get(`pay:${paymentReference}`);
    if (recordedPending) {
      clearAnnualDiscountReservation_(
        tenantId,
        JSON.parse(recordedPending).annualDiscountReservationToken
      );
    }
    cache.remove(`pay:${paymentReference}`);
    return { ok: true, message: "This subscription payment was already applied." };
  }
  const cached = cache.get(`pay:${paymentReference}`);
  if (!cached) return { ok: false, error: "Payment reference is missing or expired." };
  const pending = JSON.parse(cached);
  if (normalizeSchoolCode_(pending.tenantId) !== normalizeSchoolCode_(tenantId)) {
    return { ok: false, error: "Payment reference does not match this school." };
  }
  const secret = String(PropertiesService.getScriptProperties().getProperty("PAYSTACK_SECRET_KEY") || "");
  if (!secret) return { ok: false, error: "Payment verification is not configured." };
  const response = UrlFetchApp.fetch(`https://api.paystack.co/transaction/verify/${encodeURIComponent(paymentReference)}`, {
    method: "get", headers: { Authorization: `Bearer ${secret}` }, muteHttpExceptions: true
  });
  const result = JSON.parse(response.getContentText());
  if (response.getResponseCode() >= 300 || !result.status || result.data.status !== "success") {
    return { ok: false, error: "Paystack has not confirmed this payment." };
  }
  const cachedExpected = Number(pending.amountNaira);
  let expected = Number.isFinite(cachedExpected) && cachedExpected > 0 ? cachedExpected : 0;
  if (!expected) {
    const plan = getPlans_().find(item => item.id === pending.plan);
    expected = ["per_term", "six_months"].includes(pending.period)
      ? plan && plan.perTermPrice
      : plan && plan.annualPrice;
  }
  if (!expected || Number(result.data.amount) !== Math.round(expected * 100) || result.data.currency !== "NGN") {
    return { ok: false, error: "The verified payment amount does not match the selected subscription." };
  }
  let tenant, start;
  const paymentLock = LockService.getScriptLock();
  paymentLock.waitLock(10000);
  try {
    const latestRows = readRowsFrom_(billingSheet);
    if (latestRows.some(row => String(row.reference || "") === paymentReference)) {
      cache.remove(`pay:${paymentReference}`);
      clearAnnualDiscountReservation_(tenantId, pending.annualDiscountReservationToken);
      return { ok: true, message: "This subscription payment was already applied." };
    }
    tenant = findTenant_(tenantId);
    start = tenant.paidThrough && new Date(tenant.paidThrough).getTime() > Date.now()
      ? new Date(tenant.paidThrough)
      : new Date();
    const durationMonths = pending.period === "per_term" ? 4 : pending.period === "six_months" ? 6 : 12;
    start.setMonth(start.getMonth() + durationMonths);
    appendRecord_(billingSheet, {
      reference: paymentReference, schoolCode: tenantId, adminEmail: tenant.adminEmail,
      plan: pending.plan, billingPeriod: pending.period,
      amountNaira: Number(result.data.amount) / 100, currency: result.data.currency,
      paidAt: result.data.paid_at || new Date(), paidThrough: start,
      paymentChannel: result.data.channel || "",
      ownerEmailStatus: "Pending", scholaroEmailStatus: "Pending",
      annualDiscountApplied: pending.annualDiscountApplied === true
    });
    updateTenant_(tenantId, {
      plan: pending.plan, billingPeriod: pending.period,
      subscriptionStatus: "active", paidThrough: start
    });
  } finally {
    paymentLock.releaseLock();
  }
  clearAnnualDiscountReservation_(tenantId, pending.annualDiscountReservationToken);
  cache.remove(`pay:${paymentReference}`);
  const payment = {
    reference: paymentReference,
    schoolCode: tenantId,
    schoolName: tenant.schoolName || tenantId,
    adminEmail: tenant.adminEmail,
    plan: pending.plan,
    billingPeriod: pending.period,
    amountNaira: Number(result.data.amount) / 100,
    currency: result.data.currency,
    paidAt: result.data.paid_at || new Date(),
    paidThrough: start,
    paymentChannel: result.data.channel || ""
  };
  const emailNotifications = sendSubscriptionPaymentNotifications_(billingSheet, payment);
  return {
    ok: true, message: "Subscription payment verified.", paidThrough: start,
    emailNotifications: emailNotifications
  };
}

function sendSubscriptionPaymentNotifications_(billingSheet, payment) {
  const values = billingSheet.getDataRange().getValues();
  if (values.length < 2) throw new Error("The subscription payment was saved, but its email receipt row could not be found.");
  const headers = values[0].map(header => String(header).trim());
  const referenceColumn = headers.indexOf("reference");
  const rowIndex = values.findIndex((row, index) =>
    index > 0 && String(row[referenceColumn] || "") === String(payment.reference)
  );
  if (rowIndex < 1) throw new Error("The subscription payment was saved, but its email receipt row could not be found.");

  const scholaroEmail = String(PropertiesService.getScriptProperties().getProperty("SCHOLARO_BILLING_EMAIL") || "").trim();
  const recipients = [
    { email: String(payment.adminEmail || "").trim(), statusColumn: "ownerEmailStatus", audience: "school" },
    { email: scholaroEmail, statusColumn: "scholaroEmailStatus", audience: "scholaro" }
  ];
  const emailNotifications = {};
  recipients.forEach(recipient => {
    const statusColumn = headers.indexOf(recipient.statusColumn);
    if (statusColumn < 0) throw new Error(`The subscription payment sheet is missing ${recipient.statusColumn}. Run setup() and try again.`);
    if (String(values[rowIndex][statusColumn] || "") === "Sent") {
      emailNotifications[recipient.audience] = "Sent";
      return;
    }
    if (!recipient.email) {
      const error = recipient.audience === "scholaro"
        ? "SCHOLARO_BILLING_EMAIL is not configured in Script Properties."
        : "The registered school administrator email is missing.";
      setSafeSheetValues_(billingSheet.getRange(rowIndex + 1, statusColumn + 1), [[`Failed: ${error}`]]);
      console.error(`Subscription ${payment.reference}: ${error}`);
      emailNotifications[recipient.audience] = "Failed";
      return;
    }

    try {
      sendSubscriptionReceiptEmail_(recipient.email, payment, recipient.audience);
      setSafeSheetValues_(billingSheet.getRange(rowIndex + 1, statusColumn + 1), [["Sent"]]);
      emailNotifications[recipient.audience] = "Sent";
    } catch (error) {
      setSafeSheetValues_(billingSheet.getRange(rowIndex + 1, statusColumn + 1), [[`Failed: ${error.message}`]]);
      console.error(`Could not send ${recipient.audience} receipt for ${payment.reference}: ${error.message}`);
      emailNotifications[recipient.audience] = "Failed";
    }
  });
  return emailNotifications;
}

function sendSubscriptionReceiptEmail_(recipient, payment, audience) {
  const html = value => escapeEmailHtml_(value == null || value === "" ? "—" : value);
  const date = value => {
    const parsed = value instanceof Date ? value : new Date(value);
    return Number.isNaN(parsed.getTime()) ? html(value) : html(Utilities.formatDate(parsed, Session.getScriptTimeZone(), "dd MMM yyyy, HH:mm"));
  };
  const planName = String(payment.plan || "standard").replace(/^\w/, letter => letter.toUpperCase());
  const periodName = ["per_term", "six_months"].includes(payment.billingPeriod) ? "Per term" : "1 year";
  const amount = `${payment.currency || "NGN"} ${Number(payment.amountNaira || 0).toLocaleString("en-NG", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  const subject = `Subscription payment confirmed · ${payment.schoolName} · ${payment.reference}`;
  const heading = audience === "scholaro" ? "A school subscription has been paid" : "Your school subscription is active";
  const greeting = audience === "scholaro"
    ? "A verified subscription payment was received. Here are the transaction details for your records."
    : `Hello ${html(payment.schoolName)} team, your Paystack payment is confirmed and your school subscription has been updated.`;
  const rows = [
    ["School", payment.schoolName],
    ["School code", payment.schoolCode],
    ["Administrator email", payment.adminEmail],
    ["Plan", `${planName} · ${periodName}`],
    ["Amount paid", amount],
    ["Payment date", date(payment.paidAt)],
    ["Payment method", payment.paymentChannel || "Paystack"],
    ["Transaction reference", payment.reference],
    ["Subscription active through", date(payment.paidThrough)]
  ];
  const detailRows = rows.map(([label, value]) =>
    `<tr><td style="padding:11px 12px;border-bottom:1px solid #e8edf5;color:#6b7890;font-size:12px;width:42%">${html(label)}</td><td style="padding:11px 12px;border-bottom:1px solid #e8edf5;color:#17233b;font-size:13px;font-weight:600;word-break:break-word">${html(value)}</td></tr>`
  ).join("");
  const supportEmail = String(PropertiesService.getScriptProperties().getProperty("ADMIN_EMAIL") || "").trim();
  const supportLine = supportEmail
    ? `Need help? Contact <a href="mailto:${html(supportEmail)}" style="color:#315bc7">${html(supportEmail)}</a>.`
    : "Keep this email as your payment confirmation and subscription record.";
  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    name: "Clasten",
    body: `${heading}\n\n${greeting}\n\n${rows.map(([label, value]) => `${label}: ${value}`).join("\n")}\n\n${supportEmail ? `Support: ${supportEmail}` : "Keep this email for your records."}`,
    htmlBody: `<div style="margin:0;padding:30px 12px;background:#f3f6fb;font-family:Arial,sans-serif;color:#17233b"><div style="max-width:620px;margin:0 auto;background:#ffffff;border:1px solid #e2e8f1;border-radius:16px;overflow:hidden"><div style="height:7px;background:linear-gradient(90deg,#315bc7,#16a581)"></div><div style="padding:30px 32px"><div style="font-size:12px;font-weight:800;letter-spacing:2px;color:#315bc7">CLASTEN</div><div style="display:inline-block;margin-top:22px;padding:7px 10px;background:#e9f8f2;border-radius:20px;color:#138563;font-size:11px;font-weight:700">PAYMENT CONFIRMED</div><h1 style="margin:13px 0 8px;font-size:24px;line-height:1.25;color:#17233b">${html(heading)}</h1><p style="margin:0 0 22px;color:#66738a;font-size:14px;line-height:1.6">${greeting}</p><div style="margin-bottom:20px;padding:17px 18px;background:#f5f7fc;border-radius:12px"><div style="font-size:11px;color:#738097;text-transform:uppercase;letter-spacing:1px">Amount received</div><div style="margin-top:5px;font-size:26px;font-weight:800;color:#2447c8">${html(amount)}</div><div style="margin-top:5px;font-size:12px;color:#66738a">${html(planName)} plan · ${html(periodName)}</div></div><table role="presentation" style="width:100%;border-collapse:collapse">${detailRows}</table><div style="margin-top:22px;padding-top:16px;border-top:1px solid #e8edf5;color:#758096;font-size:12px;line-height:1.6">${supportLine}</div></div><div style="padding:15px 32px;background:#fafbfd;color:#9aa4b4;font-size:11px">Automated subscription receipt · Clasten</div></div></div>`
  });
}

function escapeEmailHtml_(value) {
  return String(value).replace(/[&<>"']/g, character => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  })[character]);
}

function ensureHeaders_(name, headers) {
  const sheet = getSheet_(name, headers);
  if (!headers.length) {
    return sheet;
  }

  const existing = sheet.getLastColumn()
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    : [];
  headers.forEach(header => {
    if (!existing.includes(header)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      existing.push(header);
    }
  });
  return sheet;
}

function writeRecord_(sheet, rowNumber, record, headers = null) {
  headers = headers || sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = headers.map(header => Object.prototype.hasOwnProperty.call(record, header) ? record[header] : "");
  writeSafeRows_(sheet, rowNumber, [values]);
}

function appendRecord_(sheet, record) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = headers.map(header => Object.prototype.hasOwnProperty.call(record, header) ? record[header] : "");
  writeSafeRows_(sheet, sheet.getLastRow() + 1, [values]);
}

function writeSafeRows_(sheet, startRow, rows) {
  if (!rows.length) return;
  const endRow = startRow + rows.length - 1;
  if (endRow > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), endRow - sheet.getMaxRows());
  setSafeSheetValues_(sheet.getRange(startRow, 1, rows.length, rows[0].length), rows);
}

function setSafeSheetValues_(range, values) {
  const safeValues = values.map(row => row.slice());
  const literalCells = [];
  values.forEach((row, rowIndex) => row.forEach((value, columnIndex) => {
    if (typeof value === "string" && /^[\s\uFEFF]*[=+\-@]/.test(value)) {
      safeValues[rowIndex][columnIndex] = "";
      literalCells.push({ row: rowIndex + 1, column: columnIndex + 1, value: value });
    }
  }));
  range.setValues(safeValues);
  literalCells.forEach(cell => range.getCell(cell.row, cell.column).setRichTextValue(
    SpreadsheetApp.newRichTextValue().setText(cell.value).build()
  ));
}

function readRows_(sheetName) {
  const sheet = getSheet_(sheetName, []);
  const values = readSheetValues_(sheet);

  if (values.length < 2) {
    return [];
  }

  const headers = values[0];
  return values.slice(1).map(row => {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = row[index];
    });
    return record;
  });
}

function readSheetValues_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  return lastRow && lastColumn
    ? sheet.getRange(1, 1, lastRow, lastColumn).getValues()
    : [];
}

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
