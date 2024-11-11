import { Component } from '@angular/core';

@Component({
  selector: 'app-extract-data',
  standalone: true,
  imports: [],
  templateUrl: './extract-data.component.html',
  styleUrls: ['./extract-data.component.css']
})
export class ExtractDataComponent {
  loading: boolean = false;
  subject: string = '';
  body: string = '';
  imagePreviewSrc: string | null = null;
  noAttachments: boolean = true;

  ngOnInit(): void {
    Office.onReady((info: any) => {
      if (info.host === Office.HostType.Outlook) {
        // Office is ready to interact with
      }
    });
  }

  extractEmailData() {
    this.loading = true;
    this.imagePreviewSrc = null;

    const item = Office.context.mailbox.item;

    if(item){
      this.subject = item.subject || 'No subject';
      this.getEmailBody(item).then((body) => {
        this.body = body || 'No body content';
        return this.retrieveFirstImageAttachment(item);
      }).then(() => {
        this.loading = false;
      }).catch((error) => {
        console.error("Error extracting email data:", error);
        this.loading = false;
      });
    }
  }

  getEmailBody(item: any): Promise<string> {
    return new Promise((resolve, reject) => {
      if (item.body) {
        item.body.getAsync(Office.CoercionType.Text, (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value);
          } else {
            reject("Failed to retrieve body content");
          }
        });
      } else {
        resolve('No body content');
      }
    });
  }

  retrieveFirstImageAttachment(item: any): Promise<void> {
    return new Promise((resolve, reject) => {
      const attachments = item.attachments;

      if (attachments && attachments.length > 0) {
        const firstImageAttachment = attachments.find((att: any) => att.contentType && att.contentType.startsWith("image/"));

        if (firstImageAttachment) {
          item.getAttachmentContentAsync(firstImageAttachment.id, (result: any) => {
            if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
              this.displayAttachment(result.value.content, result.value.format);
              this.noAttachments = false;
              resolve();
            } else {
              reject("Failed to retrieve image attachment content");
            }
          });
        } else {
          resolve();
        }
      } else {
        resolve();
      }
    });
  }

  displayAttachment(content: string, format: string) {
    if (format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
      this.imagePreviewSrc = `data:image/png;base64,${content}`;
    } else if (format === Office.MailboxEnums.AttachmentContentFormat.Url) {
      this.imagePreviewSrc = content;
    }
  }
}
