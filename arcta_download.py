import imaplib, email
import os

user = 'j.adjaho@altatrading.com'
password = 'table40!longerchair'
imap_url = 'imap-mail.outlook.com'
attachment_dir = '\\\\altfps\\arcadiagroup$\Midoffice\ARCTAfiles'

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

con = imaplib.IMAP4_SSL(imap_url)
con.login(user, password)
# print(con.list())
con.select('ARCTA_all')

print(con.select('ARCTA_all'))
# selecting the first email
email_id_raw = str(con.select('ARCTA_all'))
email_id = email_id_raw[10:-3]
print(email_id)


for i in range(1,int(email_id)):
    email_id_bytes = bytes(str(i), encoding='utf8')
    # print(email_id_bytes)
    result, data = con.fetch(email_id_bytes, '(RFC822)')
    raw = email.message_from_bytes(data[0][1])
    get_attachments(raw)

print('done')
