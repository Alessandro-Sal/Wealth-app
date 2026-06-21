/**
 * FIX (2.17): Email centralizzata per tutte le notifiche/report.
 *
 * Prima il destinatario era cablato come "alessandro.saladino01@gmail.com" in 8 punti
 * (App_AI, App_MarketMail, App_MondayBriefing, App_DividendsMail, App_BudgetGuardian,
 * App_Reporting, App_SniperAlert): in un fork le email finivano all'autore originale.
 *
 * Ordine di risoluzione:
 *   1. Script Property "NOTIFY_EMAIL" (se impostata in Impostazioni progetto -> Proprietà script)
 *   2. Email dell'utente che ha fatto il deploy (Session.getActiveUser().getEmail())
 *
 * @return {string} L'indirizzo email a cui inviare notifiche e report.
 */
function getNotifyEmail() {
  try {
    const prop = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL');
    if (prop && String(prop).trim() !== '') return String(prop).trim();
  } catch (e) {
    // ignora: si passa al fallback
  }
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (e) {
    return '';
  }
}
