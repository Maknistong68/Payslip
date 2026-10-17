/*
 * Payslip Generator - settings for this copy of the app.
 * microsoftClientId: the "Application (client) ID" of the free Microsoft app registration that turns on
 * "Sign in with Microsoft" on the Send step (README, "Sign in with Microsoft: one-time setup"). It is a
 * public ID, not a secret. Empty: the button is hidden and everything else works as before.
 */
window.PAYSLIP_CONFIG = window.PAYSLIP_CONFIG || {
  microsoftClientId: ''
};
