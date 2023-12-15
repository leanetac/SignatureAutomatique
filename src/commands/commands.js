Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Récupération de la signature de l'utilisateur sur la plateforme de contact
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {

	let _settings = Office.context.roamingSettings;

	// Récupération de la signature
	jQuery.ajax({
		url: 'https://vcard.thinkad.club/api/index.php',
		method: 'POST',
		jsonp: true,
		data: {'mail': Office.context.mailbox.userProfile.emailAddress},
		success: (resp) => {
			// Sauvegarde de la signature dans les paramètres de l'Add-in
			_settings.set('tmplSignature', 'signatureVcard');
			_settings.set('signatureContent', resp);
			_settings.saveAsync();
			// Insertion de la signature dans le mail
			insert_signature(resp);
		}
	})
  

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 * Suppression de la signature
 * @param event {Office.AddinCommands.Event}
 */
function deleteSignature(event) {
	// On supprime le template et le contenu de la signature des paramètres de l'Add-in
	let _settings = Office.context.roamingSettings;
	_settings.remove('tmplSignature');
	_settings.remove('signatureContent');
	_settings.saveAsync();

	// On enlève la signature du mail
	Office.context.mailbox.item.body.setAsync
		(
			"",
			{ coercionType: Office.CoercionType.Html },
			function (asyncResult) { console.log("Suppression de la signature provenant de la plateforme de contact"); }
		);
}


function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

/**
 * Insertion de la signature dans le corps du mail (appointment)
 * @param {any} str Contenu de la signature
 */
function set_body(str)
{
  Office.context.mailbox.item.body.setAsync
  (
	get_cal_offset() + str,
	{coercionType: Office.CoercionType.Html},
	function (asyncResult)
	{
	  console.log("set_body - " + JSON.stringify(asyncResult));
	}
  );
}

/**
 * Insertion d'une signature
 * @param {any} str Contenu de la signature (text, HTML)
 */
function set_signature(str)
{
  Office.context.mailbox.item.body.setSignatureAsync
  (
	str,
	{coercionType: Office.CoercionType.Html},
	function (asyncResult)
	{
	  console.log("set_signature - " + JSON.stringify(asyncResult));
	}
  );
}

/**
 * Définition du lieu d'insertion de la signature
 * @param {any} str Contenu de la signature
 */
function insert_signature(str)
{
  if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Appointment)
  {
	set_body(str);
  }
  else
  {
	set_signature(str);
  }
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;