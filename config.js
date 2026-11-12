'use strict';
(async () => {
  const status = document.getElementById('status');
  try {
    await microsoftTeams.app.initialize();
    microsoftTeams.pages.config.registerOnSaveHandler(async event => {
      try {
        const page = new URL('info.html', document.location.href);
        page.search = '';
        page.hash = '';
        await microsoftTeams.pages.config.setConfig({
          entityId: 'gerbsch-raeder-wareneingang',
          suggestedDisplayName: 'Räder-Wareneingang',
          contentUrl: page.href,
          websiteUrl: page.href
        });
        event.notifySuccess();
      } catch (_) {
        event.notifyFailure('Die Einrichtung konnte nicht gespeichert werden.');
      }
    });
    microsoftTeams.pages.config.setValidityState(true);
    status.textContent = 'Bereit. Bitte die Installation in Teams überprüfen.';
  } catch (_) {
    status.textContent = 'Diese Seite muss innerhalb von Microsoft Teams geöffnet werden.';
  }
})();
