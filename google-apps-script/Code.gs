/**
 * KlickMaroc — Google Apps Script backend (free)
 * Deploy as Web App: Execute as Me, Access Anyone.
 * Set a Script Property named STORE_EMAIL to the merchant email.
 * Never publish customer details in the GitHub repository.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const data = JSON.parse((e.postData && e.postData.contents) || '{}');
    const required = ['name','phone','city','address','order_details'];
    if (required.some(k => !String(data[k] || '').trim())) throw new Error('Missing fields');
    if (!/^0[5-7][0-9]{8}$/.test(String(data.phone).replace(/\s/g,''))) throw new Error('Invalid phone');
    const email = PropertiesService.getScriptProperties().getProperty('STORE_EMAIL');
    if (!email) throw new Error('STORE_EMAIL not configured');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Bind this script to a Google Sheet');
    let sheet = ss.getSheetByName('Commandes');
    if (!sheet) {
      sheet = ss.insertSheet('Commandes');
      sheet.appendRow(['Date','Référence','Nom','Téléphone','Ville','Adresse','Notes','Produits','Sous-total DH','Livraison DH','Total DH','Paiement','Statut']);
    }
    const id = 'KM-' + Utilities.getUuid().slice(0,8).toUpperCase();
    const safe = v => "'" + String(v == null ? '' : v).slice(0,3000).replace(/[\r\n]/g,' ');
    const row = [new Date(),id,safe(data.name),safe(data.phone),safe(data.city),safe(data.address),safe(data.notes),safe(data.order_details),Number(data.subtotal)||0,data.shipping==null?'À confirmer':Number(data.shipping),data.total==null?'À confirmer':Number(data.total),'Paiement à la livraison','Nouveau'];
    sheet.appendRow(row);
    const body = ['Nouvelle commande KlickMaroc','Référence : '+id,'Nom : '+data.name,'Téléphone : '+data.phone,'Ville : '+data.city,'Adresse : '+data.address,'Notes : '+(data.notes||''),'','Produits :',data.order_details,'','Sous-total : '+data.subtotal+' DH','Livraison : '+(data.shipping==null?'À confirmer':data.shipping+' DH'),'Total : '+(data.total==null?'À confirmer':data.total+' DH'),'Paiement à la livraison'].join('\n');
    try { MailApp.sendEmail(email, 'Nouvelle commande KlickMaroc '+id, body); } catch(mailError) { console.error('Order saved, email failed: '+mailError); }
    return ContentService.createTextOutput(JSON.stringify({success:true,reference:id})).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    console.error(error);
    return ContentService.createTextOutput(JSON.stringify({success:false,message:'Order not accepted'})).setMimeType(ContentService.MimeType.JSON);
  } finally { lock.releaseLock(); }
}
