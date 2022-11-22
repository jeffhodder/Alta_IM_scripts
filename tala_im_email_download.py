import imaplib, email
import os

user = 'j.adjaho@altatrading.com'
password = 'table20!longerchair'
imap_url = 'imap-mail.outlook.com'
attachment_dir = '\\\\altfps\\arcadiagroup$\Midoffice\Tala IM'

def get_attachments(msg):
    for part in msg.walk():
        if part.get_content_maintype == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fileName = part.get_filename()

        if bool(fileName):
            filePath = os.path.join(attachment_dir, fileName)
            with open(filePath, 'wb') as f:
                f.write(part.get_payload(decode=True))

def connect(self, username, password):
    if(self.hostname == 'imap.outlook.com'):
            imap_server = "outlook.office365.com"
            self.server = self.transport(imap_server, self.port)
            self.server = imaplib.IMAP4_SSL(imap_server)
            (retcode, capabilities) = self.server.login(user,password)
            self.server.select('AutomationFolder/Tala_IM')
    else:
            typ, msg = self.server.login(user, password)
            if self.folder:
                self.server.select(self.folder)
            else:
                self.server.select()


# con = imaplib.IMAP4_SSL(imap_url)
# con.login(user, password)
# # print(con.list())
# con.select('AutomationFolder/Tala_IM')

print(self.server.select('AutomationFolder/Tala_IM'))
# selecting the first email
email_id_raw = str(self.server.select('AutomationFolder/Tala_IM'))
email_id = email_id_raw[10:-3]
email_id_bytes = bytes(email_id, encoding='utf8')

# print(email_id_bytes)
result, data = self.server.fetch(email_id_bytes, '(RFC822)')
raw = email.message_from_bytes(data[0][1])
get_attachments(raw)

print('done')
