import { bootstrapApplication } from '@angular/platform-browser';
import { appConfig } from './app/app.config';
import { AppComponent } from './app/app.component';
import { AccountManager } from './sso/authConfig';
import { makeGraphRequest } from './sso/msgraph-helper';

const accountManager = new AccountManager();

Office.onReady(async (info) => {
  console.log(`Add-in is ready for ${info.host} on ${info.platform}`);

  // Initialize when Office is ready.
  if (info.host === Office.HostType.Outlook) {
    try {
      // Initialize MSAL and ensure it's ready before proceeding.
      await accountManager.initialize();
      console.log("AccountManager initialized successfully.");
      
      // Proceed with SSO only after successful initialization.
      await SSO();
    } catch (error) {
      console.error("Failed to initialize AccountManager:", error);
    }
  }

});

// SSO
async function SSO() {
  const accessToken = await accountManager.ssoGetAccessToken(["user.read"]);
  
  const response: { displayName: string; mail: string } = await makeGraphRequest(accessToken, "/me", "");
  console.log(`Hello, ${response.displayName}!`);

  if(accessToken) {
    bootstrapApplication(AppComponent, appConfig)
    .catch((err) => console.error(err));
  }
}