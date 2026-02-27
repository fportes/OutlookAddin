Office.onReady(() => {});

function onMessageComposeHandler(event) {
    insertSignature().then(() => {
        event.completed();
    });
}

async function insertSignature() {

    const marker = "SummitCareSignatureMarker";

    return new Promise((resolve) => {

        Office.context.mailbox.item.body.getAsync(
            Office.CoercionType.Html,
            async (result) => {

                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    resolve();
                    return;
                }

                if (result.value.includes(marker)) {
                    resolve();
                    return;
                }

                try {

                    const token = await OfficeRuntime.auth.getAccessToken({
                        allowSignInPrompt: true
                    });

                    const response = await fetch(
                        "https://graph.microsoft.com/v1.0/me",
                        {
                            headers: { Authorization: `Bearer ${token}` }
                        }
                    );

                    const user = await response.json();

                    const signature = `
                        <div id="${marker}" style="font-family:Arial; font-size:12px;">
                            <br>
                            <strong>${user.displayName || ""}</strong><br>
                            ${user.jobTitle || ""}<br>
                            ${user.businessPhones?.[0] || ""}<br>
                            ${user.mobilePhone || ""}<br>
                            ${user.mail || ""}<br>
                            SummitCare Management LLC
                            <br><br>
                        </div>
                    `;

                    Office.context.mailbox.item.body.prependAsync(
                        signature,
                        { coercionType: Office.CoercionType.Html },
                        () => resolve()
                    );

                } catch {
                    resolve();
                }
            }
        );
    });
}
