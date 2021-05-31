import imaplib, email, os, re, time, sys
import smtplib, ssl
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
        emailObject=None,
    ):
        self.subject = subject
        self._from = _from
        self.recipient = recipient
        self.body = body
        self.attachments = attachments
        self.id = id
        self.Inbox = fromInbox
        self.seen = seen
        self.emailObject = emailObject

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

    def logout(self):
        self.imap.close()
        self.imap.logout()

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
                    break
            else:
                print(
                    "[{timestamp}] MailID {mailID} contains no RFC822 tag. Can't read.".format(
                        timestamp=time.strftime("%H:%M:%S"), mailID=mailID
                    )
                )
                continue
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
                mailID,
                subject,
                _from,
                recipientList,
                fromInbox=self,
                attachments=[],
                emailObject=parsedEmail,
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
                        filename, encoding = email.header.decode_header(
                            part.get_filename()
                        )[0]
                        if encoding is not None:
                            filename = filename.decode(encoding)
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


class SMTPHandler:
    def __init__(self, user: str, passwd: str, server: str, port: int = 465):
        self.user = user
        self.passwd = passwd
        self.server = server
        self.port = port
        pass

    def send_text_mail(self, mail, replyingTo=None):
        message = message = email.message.EmailMessage()
        message["From"] = "Sistema de gestión de despachos Logica Express <{f}>".format(
            f=mail._from
        )
        if replyingTo:
            message["Subject"] = "RE: " + replyingTo["Subject"].replace(
                "Re: ", ""
            ).replace("RE: ", "")
            message["References"] = replyingTo[
                "Message-ID"
            ]  # +replyingTo["References"].strip()
            message["In-Reply-To"] = replyingTo["Message-ID"]
            message["Thread-Topic"] = replyingTo["Thread-Topic"]
            message["Thread-Index"] = replyingTo["Thread-Index"]
        else:
            message["Subject"] = mail.subject
        message["To"] = mail.recipient
        message.set_content(mail.body)
        SSLContext = ssl.create_default_context()
        with smtplib.SMTP_SSL(
            self.server, self.port, context=SSLContext
        ) as smtpConnection:
            smtpConnection.login(self.user, self.passwd)
            response = smtpConnection.sendmail(
                mail._from, mail.recipient, message.as_string()
            )
        return response


def send_confirmation_mail(
    reportData: list,
    to: str,
    _from: str,
    subject: str,
    outboxHandler: SMTPHandler,
    replyingTo=None,
) -> bool:
    transactionEmail = Email(
        id=99, subject="Re:{subject}".format(subject=subject), _from=_from, recipient=to
    )
    transactionEmail.body = build_text_report(reportData)
    print(outboxHandler.send_text_mail(transactionEmail, replyingTo))
    return True


def build_text_report(reportRawData: list) -> str:
    taintedImport = False
    finalReport = ""
    for fileReportData in reportRawData:
        fileResultMsg = ""
        if len(fileReportData["general_issues"]) > 0:
            fileResultMsg += '¡Atención! Hubo algunos problemas al importar tus despachos del archivo "{filename}":\n\n'.format(
                filename=fileReportData["filename"]
            )
            taintedImport = True
            for warning in fileReportData["general_issues"]:
                fileResultMsg += "- {warning}\n".format(warning=warning)
            fileResultMsg += "\n"
        else:
            fileResultMsg += 'Recibimos el archivo "{filename}" con despachos!\n\n'.format(
                filename=fileReportData["filename"]
            )
        successfulImports = []
        issuesImports = []
        failedImports = []
        for dispatch, errorCode, warnings in fileReportData["dispatches"]:
            if errorCode == 0:
                successfulImports.append((dispatch, warnings))
            elif errorCode == 2:
                failedImports.append((dispatch, warnings))
            elif errorCode == 1:
                issuesImports.append((dispatch, warnings))
            else:
                raise Exception(f"DispatchID {dispatch.id} code unknown: {errorCode}")
        if len(failedImports) > 0:
            fileResultMsg += "Los siguientes despachos no pudieron ser importados:\n"
            for failedDispatch in failedImports:
                fileResultMsg += " - {dispatchID}:\n".format(
                    dispatchID=failedDispatch[0]
                )
                for warningMsg in failedDispatch[1]:
                    fileResultMsg += "\t {msg}\n".format(msg=warningMsg)
                fileResultMsg += "\n"
                taintedImport = True
        if len(issuesImports) > 0:
            fileResultMsg += (
                "Los siguientes despachos fueron importados con observaciones:\n"
            )
            for issuesDispatch in issuesImports:
                fileResultMsg += " - {dispatchID}:\n".format(
                    dispatchID=issuesDispatch[0].id
                )
                for warningMsg in issuesDispatch[1]:
                    fileResultMsg += "\t {msg}\n".format(msg=warningMsg)
                fileResultMsg += "\n"
                taintedImport = True
        if len(successfulImports) > 0:
            fileResultMsg += (
                "Los siguientes despachos fueron importados exitosamente:\n"
            )
            for successfulDispatch in successfulImports:
                fileResultMsg += " - {dispatchID}:\n".format(
                    dispatchID=successfulDispatch[0].id
                )
                fileResultMsg += "\n"

        finalReport += fileResultMsg
    if taintedImport:
        finalReport += "Puedes responder a este correo adjuntando un archivo de despachos corregido hasta el final del día.\n\n"
    finalReport += """Gracias por preferir Logica Express!
    Para cualquier problema, puedes contactarnos a soporte@logicaexpress.cl"""
    return finalReport
