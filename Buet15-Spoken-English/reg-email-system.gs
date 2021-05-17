/*
  Library Info.
  ------------------------------------------------------------------------
  Script Id: 1PvbQZoE804UdpINDPeuTJeu-vrYFd7gzkZNnZCmsqGHgwECpzKh-aoZC
  Version: Latest (9)
  Author: Shams-E-Shifat
  Email: sshifat022@gmail.com
  ------------------------------------------------------------------------
 */
const SPREADSHEET_LINK = 'paste-spreadsheet-url';
const RESPONSE_SHEET_NAME = 'sheet-name-where-the-response-is-stored';
const EMAIL_CONTAINING_COLUMN_NAME = 'Email';

const EMAIL_BODY_CONTAINING_DOC_LINK = 'paste-email-body-containing-doc-url';
const EMAIL_SUBJECT = 'subject-of-email';
const EMAIL_SENDER_NAME = "Email-sender-name";

// Run this function only the first time.
function runMe(){
  SpokenEnglishPlatform.submit('execute');
}

// No need to edit this function.
function execute(){
  SpokenEnglishPlatform.SendEmail(
    SPREADSHEET_LINK,
    RESPONSE_SHEET_NAME,
    EMAIL_CONTAINING_COLUMN_NAME,
    EMAIL_BODY_CONTAINING_DOC_LINK,
    EMAIL_SUBJECT,
    EMAIL_SENDER_NAME);
}
