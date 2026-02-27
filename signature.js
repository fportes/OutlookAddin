Office.onReady(() => {
    if (Office.context.mailbox.item) {
        insertSignatureIfMissing();
    }
});

function insertSignatureIfMissing() {
    const signatureMarker = "SummitCareSignatureMarker";

    Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Html,
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {

                if (!result.value.includes(signatureMarker)) {

                    const signatureHtml = `
                        <div id="SummitCareSignatureMarker">
                            <br>
                            <strong>${Office.context.mailbox.userProfile.displayName}</strong><br>
                            ${Office.context.mailbox.userProfile.emailAddress}<br>
                            SummitCare Management LLC
                            <br><br>
                        </div>
                    `;

                    Office.context.mailbox.item.body.prependAsync(
                        signatureHtml,
                        { coercionType: Office.CoercionType.Html }
                    );
                }
            }
        }
    );
}
