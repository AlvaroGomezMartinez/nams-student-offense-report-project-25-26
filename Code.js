/**
 *
 *                                  nams-student-offense-report-project-25-26
 *
 * The script below supports NAMS's student office referral workflow.
 * Teachers fill out a Google Form
 * (https://docs.google.com/forms/d/e/1FAIpQLSfQLr8mRLxVpFahi3xKUYiEExblbU5_6S_jOlBphM3Kfl8cPg/viewform)
 * when they want to create an offense report. The submission of the Google Form populates in the
 * 2025-2026 NAMS Student Offense Report sheet that is monitored by the NAMS administration.
 * Administration will process the student offense report and provide information in the sheet.
 * Administration will then indicate that they are ready to send an email to the teacher who filled out
 * the form with the action they took and the consequences that the student received.
 *
 * Google Apps Script Developer: Alvaro Gomez, Academic Technology Coach, 210-397-9408
 *
*/

var rows = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
var headers = rows.shift();
var checkBoxRange = SpreadsheetApp.getActive().getRange("Q2:Q");
var stuInfraction = SpreadsheetApp.getActive().getRange("J2:N");
var d = new Date();
var today = Utilities.formatDate(d, "GMT-06:00", "MMM d, ''yy h:mm a");

function wrapText() {
  stuInfraction.setWrap(true);
}

/**
 * Adds a user interface to the spreadsheet that gives the administrator
 * the option to send emails to the teachers. When the checkbox is checked,
 * it will run the processSelectedRows function.
*/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Teacher Followup Emails")
    .addItem("Send Emails Now", "processSelectedRows")
    .addToUi();
  //Looks for cells without a checkbox in column Q and inserts one if missing.
  rows.forEach(function (row) {
    if (row[16] === "") {
      checkBoxRange.insertCheckboxes();
    }
    stuInfraction.setWrap(true);
  });
}

/**
 * Checks whether a row from Column Q is checked and also if there's a
 * date in Column R.
 * If checked and no date, then it returns the data's row and creates the
 * const "User" with the email data.
 * If checked and has a date, then it skips the row.
 * If not checked and no date, then it skips the row.
*/
function processSelectedRows() {
  let currentRowIndex = -1; // Initialize variable to track the row where an error occurred

  try {
    rows.forEach(function (row, i) {
      currentRowIndex = i + 2; // +2 accounts for header and 0-indexing
      if (
        row[16] === true &&
        (row[17] ? String(row[17]).trim() === "" : true) &&
        row[5] != ""
      ) {
        const user = {
          administrator: row[14],
          date: row[6],
          followup: row[15],
          id: row[2],
          sendemail: row[16],
          sentDate: row[17],
          staffSubmittingReferral: row[5],
          student: row[1],
          timestamp: row[0],
        };

        const teacherEmailList = {
          "Aguilar, R": {
            Email: "russell.aguilar@nisd.net",
            Salutation: "Mr. ",
          },
          Atoui: {
            Email: "atlanta.atoui@nisd.net",
            Salutation: "Mrs.",
          },
          Bowery: {
            Email: "melissa.bowery@nisd.net",
            Salutation: "Mrs. ",
          },
          "Cantu, S": {
            Email: "sandy.cantu@nisd.net",
            Salutation: "Mrs. ",
          },
          Casanova: {
            Email: "henry.casanova@nisd.net",
            Salutation: "Mr. ",
          },
          Coyle: {
            Email: "deborah.coyle@nisd.net",
            Salutation: "Mrs. ",
          },
          "De Leon, U": {
            Email: "ulices.deleon@nisd.net",
            Salutation: "Mr. ",
          },
          "Deleon, R": {
            Email: "rebeca.deleon@nisd.net",
            Salutation: "Mrs. ",
          },
          Farias: {
            Email: "michelle.farias@nisd.net",
            Salutation: "Mrs. ",
          },
          "Franco, G": {
            Email: "george.franco01@nisd.net",
            Salutation: "Mr.",
          },
          Garcia: {
            Email: "danny.garcia@nisd.net",
            Salutation: "Mr. ",
          },
          Goff: {
            Email: "steven.goff@nisd.net",
            Salutation: "Mr. ",
          },
          Gomez: {
            Email: "alvaro.gomez@nisd.net",
            Salutation: "Mr.",
          },
          Gonzales: {
            Email: "zina.gonzales@nisd.net",
            Salutation: "Dr.",
          },
          Hernandez: {
            Email: "david.hernandez@nisd.net",
            Salutation: "Mr. ",
          },
          Hutton: {
            Email: "rebekah.hutton@nisd.net",
            Salutation: "Mrs. ",
          },
          Idrogo: {
            Email: "valerie.idrogo@nisd.net",
            Salutation: "Mrs. ",
          },
          Jasso: {
            Email: "nadia.jasso@nisd.net",
            Salutation: "Mrs. ",
          },
          Marquez: {
            Email: "monica.marquez@nisd.net",
            Salutation: "Mrs. ",
          },
          Ollendieck: {
            Email: "reggie.ollendieck@nisd.net",
            Salutation: "Mr. ",
          },
          Paez: {
            Email: "john.paez@nisd.net",
            Salutation: "Mr. ",
          },
          Ramon: {
            Email: "israel.ramon@nisd.net",
            Salutation: "Mr. ",
          },
          Tellez: {
            Email: "lisa.tellez@nisd.net",
            Salutation: "Mrs. ",
          },
          Trevino: {
            Email: "marcos.trevino@nisd.net",
            Salutation: "Mr. ",
          },
          Wine: {
            Email: "stephanie.wine@nisd.net",
            Salutation: "Mrs. ",
          },
          Yeager: {
            Email: "sheila.yeager@nisd.net",
            Salutation: "Mrs. ",
          },
        };

        const adminList = {
          Ollendieck: "Mr. ",
          "Cantu, S": "Mrs. ",
          Yeager: "Mrs. ",
        };

        var teacherName = user.staffSubmittingReferral;
        var adminName = user.administrator;
        var adminTitle = adminList[adminName];
        var adminFollowUp = user.followup;
        var to = teacherEmailList[teacherName].Email;
        var teacherTitle = teacherEmailList[teacherName].Salutation;
        var studentName = user.student;
        var studentId = user.id;
        var rawReferralDate = new Date(user.date);
        var referralDate = Utilities.formatDate(
          rawReferralDate,
          SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
          "MMM d, yyyy"
        );
        var subject = "NAMS Adminstrator Followup: Student Offense Report";
        var message = `${teacherTitle} ${teacherName},\n\nThis is a follow-up to the office referral you submitted for: ${studentName} (${studentId}) on ${referralDate}.\n\nThe student was seen by ${adminTitle} ${adminName}, and the following action was taken:\n${adminFollowUp}\n\nThis email is provided for your records.\n\nSincerely,\nNAMS Admin`;

        MailApp.sendEmail(to, subject, message);

        if (row[17] ? String(row[17]).trim() === "" : true) {
          var emailSent = [];
          emailSent.push("R" + (i + 2));

          var timezonOffset = "-06:00"; // Adjusts the time for standard time
          var correctedDate = Utilities.formatDate(
            new Date(),
            "GMT" + timezonOffset,
            "MMM d, ''yy h:mm a"
          );

          SpreadsheetApp.getActiveSheet()
            .getRangeList(emailSent)
            .setValue(correctedDate);
          SpreadsheetApp.flush();
        }
      }
    });
  } catch (error) {
    MailApp.sendEmail({
      to: "alvaro.gomez@nisd.net",
      subject: "Error occurred on the 2025-2026 NAMS Student Referral Form",
      htmlBody:
        "An error occurred on row " +
        currentRowIndex +
        ": " +
        error.message,
    });

    var ui = SpreadsheetApp.getUi();
    ui.alert(
      "An error occurred while sending. Al was notified automatically right now to troubleshoot the error. You can check the last column to see which emails went out to teachers. If the cell doesn't have a date & time, the email didn't go out to that particular teacher. This is the error that Al will look at. You can click \"OK\" below to close this message."
    );

    Logger.log("An error occurred on row " + currentRowIndex + ": " + error.message);
  }
}
