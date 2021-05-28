import imaplib, email, os, re, time, sys
from pathlib import Path
from copy import copy


def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)


class Email:
    def __init__(
        self,
        id,
        subject,
        _from,
        recipient,
        body=None,
        seen=False,
        attachments=[],
        fromInbox=None,
    ):
        self.subject = subject
        self._from = _from
        self.recipient = recipient
        self.body = body
        self.attachments = attachments
        self.id = id
        self.Inbox = fromInbox
        self.seen = seen

    def mark_read(self):
        timestamp = time.strftime("%H:%M:%S")
        if self.seen:
            print(f"[{timestamp}] INFO: mail {self.id} is already marked as read")
        self.Inbox.reconnect()
        self.Inbox.imap.select("INBOX")
        self.Inbox.imap.store(self.id, "+FLAGS", "\SEEN")
        print(f"[{timestamp}] MailID {self.id} from {self._from} marked as read")
        self.seen = True


class Inbox:
    def __init__(self, user, password, imapServer):
        self.user = user
        self.password = password
        self.imapServer = imapServer

        self.imap = imaplib.IMAP4_SSL(imapServer)
        try:
            loginResponse = self.imap.login(user, password)
        except:
            raise Exception(
                "Critical: Unhandled error when trying to login to IMAP server."
            )

    def reconnect(self):
        self.imap = imaplib.IMAP4_SSL(self.imapServer)
        self.imap.login(self.user, self.password)

    def check_inbox(
        self, folder="INBOX", unread_only=True, last=None, attachmentsDir="attachments"
    ):
        response, nEmails = self.imap.select(folder)
        nEmails = int(nEmails[0])
        if unread_only:
            response, unreadEmailIDs = self.imap.search(None, "UNSEEN")
            emailRange = unreadEmailIDs[0].decode("utf-8").split()
        else:
            if last:
                lastEmailToCheck = nEmails - last
            else:
                lastEmailToCheck = 1
            emailRange = range(nEmails, lastEmailToCheck, -1)

        def _fetch_flags(mailID):
            rawFlags = self.imap.fetch(mailID, "FLAGS")
            decodedFlags = rawFlags[1][0].decode("utf-8")
            flagsPattern = re.compile(r"[\d]+\s\(FLAGS\s\(([\\\s\w]+)\)\)")
            if flagsPattern.match(decodedFlags):
                return flagsPattern.match(decodedFlags)[1].split()
            else:
                return []

        foundEmails = []
        for mailID in emailRange:
            # Leer cabecera del correo
            response, rawEmail = self.imap.fetch(str(mailID), "(RFC822)")
            for content in rawEmail:
                if isinstance(content, tuple):
                    rawEmail = content
            parsedEmail = email.message_from_bytes(rawEmail[1])
            rawSubject, encoding = email.header.decode_header(
                parsedEmail.get("Subject")
            )[0]
            if isinstance(rawEmail, bytes):
                subject = rawSubject.decode(encoding)
            else:
                subject = rawSubject
            _from = email.utils.parseaddr(parsedEmail["FROM"])[1]
            rawRecipients = email.header.decode_header(parsedEmail.get("To"))[0]
            if isinstance(rawRecipients, bytes):
                recipients = rawRecipients.decode(encoding)
            else:
                recipients = rawRecipients
            recipientList = [rec.strip() for rec in recipients[0].split(",")]
            # Crear objeto Email
            currentEmail = Email(
                mailID, subject, _from, recipientList, fromInbox=self, attachments=[]
            )
            if "\\Seen" in _fetch_flags(mailID):
                currentEmail.seen = True
            # Leer contenido del correo
            # if the email message is multipart
            body = None
            if parsedEmail.is_multipart():
                # iterate over email parts
                for part in parsedEmail.walk():
                    # extract content type of email
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    try:
                        # get the email body
                        body = part.get_payload(decode=True).decode()
                    except:
                        pass
                    if "attachment" in content_disposition:
                        # download attachment
                        filename = part.get_filename()
                        if filename:
                            year_month = time.strftime("%Y-%m")
                            day = time.strftime("%d-%a")
                            folder_name = os.path.join(
                                attachmentsDir, year_month, clean(_from), day, mailID
                            )
                            Path(folder_name).mkdir(parents=True, exist_ok=True)
                            filepath = os.path.join(folder_name, filename)
                            currentEmail.attachments.append(filepath)
                            # download attachment and save it
                            open(filepath, "wb").write(part.get_payload(decode=True))
            else:
                # extract content type of email
                content_type = parsedEmail.get_content_type()
                # get the email body
                body = parsedEmail.get_payload(decode=True).decode()
            if body:
                currentEmail.body = body
            foundEmails.append(copy(currentEmail))
            timestamp = time.strftime("%H:%M:%S")
            print(
                f'[{timestamp}] Fetching email from {_from} subject: "{subject}" id: {mailID}',
                file=sys.stdout,
            )
        return foundEmails
