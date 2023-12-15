// Fonction s'exécutant lors de la composition d'un nouveau message (newMail, reply)
function onNewMessageComposeHandler(event) {
    // On récupère les paramètres de l'Add-in spécifiques à l'utilisateur
    let _settings = Office.context.roamingSettings;
    let tmplSignature = _settings.get('tmplSignature');
    let signatureContent = _settings.get('signatureContent');

    if (tmplSignature == "signatureVcard" ) {
        // Si une signature a déjà été enregistrée, on l'applique
        set_signature(signatureContent);
    } else {
        // Sinon on récupère la signature sur la plateforme de contact
        var xhr = new XMLHttpRequest();

        xhr.addEventListener("readystatechange", function () {
            if (this.readyState === 4 && this.status === 200) {
                set_signature(this.response);
            }
        });

        xhr.open("POST", "https://vcard.thinkad.club/api/index.php");
        xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

        xhr.send("mail=" + Office.context.mailbox.userProfile.emailAddress);

    }
}

// Ajout d'une signature au mail
function set_signature(str) {
    Office.context.mailbox.item.body.setSignatureAsync
    (
        str,
        { coercionType: Office.CoercionType.Html },
        function (asyncResult) {
            console.log("set_signature - " + JSON.stringify(asyncResult));
        }
    );
}

if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}