import { Component } from '@angular/core';

@Component({
  selector: 'app-extract-data',
  standalone: true,
  imports: [],
  templateUrl: './extract-data.component.html',
  styleUrls: ['./extract-data.component.css']
})
export class ExtractDataComponent {
  subject: string | null = null;
  body: string | null = null;
  isLoading: boolean = false;

  async extractEmailData(): Promise<void> {
    if (Office && Office.context && Office.context.mailbox) {
      const item = Office.context.mailbox.item;

      if (item && item.itemType === Office.MailboxEnums.ItemType.Message) {
        const message = item as Office.MessageRead;

        this.isLoading = true;

        this.subject = message.subject || "No subject";

        try {
          this.body = await this.getBodyAsync(message);
        } catch (error) {
          console.error("Failed to get email body:", error);
          this.body = "Failed to retrieve email data";
        }

        this.isLoading = false;
      } else {
        this.subject = "No email item is selected";
        this.body = null;
      }
    } else {
      console.error("Office.js is not available");
      this.subject = "Office.js is not available";
      this.body = null;
    }
  }

  // Helper method to wrap getAsync in a Promise
  private getBodyAsync(message: Office.MessageRead): Promise<string> {
    return new Promise((resolve, reject) => {
      message.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value as string);
        } else {
          reject(result.error.message);
        }
      });
    });
  }
}
