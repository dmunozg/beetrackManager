from beetrack.beetrack_api import BeetrackAPI
from beetrack.beetrack_objects import Item, Dispatch
from beetrack import xls_import, mail_handler
from time import sleep
from dotenv import load_dotenv
import json, os, time, sys
import pandas as pd

BASE_URL = "https://app.beetrack.com/api/external/v1"
ALLOWED_FILES_EXTENSIONS = [
    "xlsx",
]
# TODO
# Esto hay que reemplazarlo por una tabla SQL
ALLOWED_CLIENTS_DF = pd.DataFrame(
    [
        [
            "lkeeler.flipout@gmail.com",
            "Pruebas",
            "TEST",
            "Reibo 3619, Puente Alto",
            True,  # Allow override
        ],
        [
            "matias@logicaexpress.cl",
            "Pruebas",
            "TEST",
            "Cerro Loma Larga 3610, Puente Alto",
            True,  # Allow override
        ],
    ],
    columns=[
        "allowedEmail",
        "clientName",
        "codePrefix",
        "pickupAddress",
        "allowOverride",
    ],
)


def check_if_allowed(filename):
    extension = filename.split(".")[-1]
    if extension in ALLOWED_FILES_EXTENSIONS:
        return True
    else:
        return False


def user_overrides(emailAddress):
    overridesList = (
        ALLOWED_CLIENTS_DF["allowedEmail"]
        .where(ALLOWED_CLIENTS_DF["allowOverride"] == True)
        .dropna()
        .tolist()
    )
    if emailAddress in overridesList:
        return True
    else:
        return False


load_dotenv()
LogicaAPI = BeetrackAPI(os.getenv("BEETRACK_APIKEY"), BASE_URL)
MailOutbox = mail_handler.SMTPHandler(
    user=os.getenv("IMAP_USER"),
    passwd=os.getenv("IMAP_PASSWD"),
    server=os.getenv("SMTP_SERVER"),
    port=os.getenv("SMTP_PORT"),
)
MailInbox = mail_handler.Inbox(
    os.getenv("IMAP_USER"), os.getenv("IMAP_PASSWD"), os.getenv("IMAP_SERVER")
)

fetchedEmails = MailInbox.check_inbox()
if len(fetchedEmails) == 0:
    timestamp = time.strftime("%H:%M:%S")
    print(f"[{timestamp}] No new mail found. Exiting.", file=sys.stdout)
    MailInbox.logout()
    sys.exit(0)
for email in fetchedEmails:
    if email._from in ALLOWED_CLIENTS_DF["allowedEmail"].to_list():
        timestamp = time.strftime("%H:%M:%S")
        print(
            f"[{timestamp}] MailID {email.id} found from allowed client: {email._from}"
        )
    else:
        continue
    if any([check_if_allowed(att) for att in email.attachments]):
        print(f"MailID {email.id} contains a valid file", file=sys.stdout)
    else:
        print(f"MailID {email.id} doesn't contain any valid file. Skipping.")
        email.mark_read()
        continue
    clientName = (
        ALLOWED_CLIENTS_DF["clientName"]
        .where(ALLOWED_CLIENTS_DF["allowedEmail"] == email._from)
        .dropna()
        .iloc[0]
    )
    pickupAddress = (
        ALLOWED_CLIENTS_DF["pickupAddress"]
        .where(ALLOWED_CLIENTS_DF["allowedEmail"] == email._from)
        .dropna()
        .iloc[0]
    )
    if user_overrides(email._from):
        print(
            "[{timestamp}] Sender can override data. Reading mail body".format(
                timestamp=time.strftime("%H:%M:%S")
            )
        )
        print(email.body)
        allowOverride = True
    else:
        allowOverride = False
    timestamp = time.strftime("%H:%M:%S")
    reports = []
    for attachment in email.attachments:
        if not check_if_allowed(attachment):
            continue
        else:
            print(
                f"[{timestamp}] Will scan file for dispatches for client {clientName} to be picked up from {pickupAddress}",
                file=sys.stdout,
            )
        foundDispatchesData, warnings = xls_import.xlsx_to_dispatches(
            attachment, clientName, pickupAddress
        )
        reportData = {
            "filename": os.path.basename(attachment),
            "general_issues": warnings,
            "dispatches": foundDispatchesData,
        }
        reports.append(reportData)
        if len(foundDispatchesData) == 0:
            print(
                "[{timestamp}] Could not parse any dispatches in file {filename}:".format(
                    timestamp=time.strftime("%H:%M:%S"),
                    filename=os.path.basename(attachment),
                ),
                file=sys.stderr,
            )
            print(warnings, sep="\n", file=sys.stderr)
        elif len(warnings) > 0:
            print(
                "[{timestamp}] There were errors importing dispatches from file {filename}:".format(
                    timestamp=time.strftime("%H:%M:%S"),
                    filename=os.path.basename(attachment),
                ),
                file=sys.stderr,
            )
            print(warnings, sep="\n", file=sys.stderr)
        else:
            pass
        for newDispatch, errorCode, warnings in foundDispatchesData:
            if errorCode == 2:
                print(
                    "[{timestamp}] Failed to import dispatch {dispatchCode}:".format(
                        timestamp=time.strftime("%H:%M:%S"),
                        dispatchCode=newDispatch.id,
                    ),
                    file=sys.stderr,
                )
                print(*warnings, sep="\n", file=sys.stderr)
                continue
            else:
                response = LogicaAPI.create_dispatch(newDispatch.dump_dict())
                if os.getenv("DEBUG"):
                    with open("{id}.json".format(id=newDispatch.id), "w") as jsonFile:
                        json.dump(
                            newDispatch.dump_dict(),
                            jsonFile,
                            indent=4,
                            ensure_ascii=False,
                        )
                print(response)
            if errorCode == 1:
                print(
                    "[{timestamp}] Dispatch {dispatchCode} was imported with issues:".format(
                        timestamp=time.strftime("%H:%M:%S"),
                        dispatchCode=newDispatch.id,
                    ),
                    file=sys.stderr,
                )
                print(*warnings, sep="\n", file=sys.stderr)
            elif errorCode == 0:
                print(
                    "[{timestamp}] Dispatch {dispatchCode} was imported successfully".format(
                        timestamp=time.strftime("%H:%M:%S"),
                        dispatchCode=newDispatch.id,
                    ),
                    file=sys.stdout,
                )
            else:
                print("CRITICAL: Unknown error code.")
    print(
        "[{timestamp}] Sending transactional email to {recipient}".format(
            timestamp=time.strftime("%H:%M:%S"), recipient=email._from
        ),
    )
    mail_handler.send_confirmation_mail(
        reports,
        _from=os.getenv("IMAP_USER"),
        to=email._from,
        subject=email.subject,
        outboxHandler=MailOutbox,
    )
    email.mark_read()
MailInbox.logout()
print(
    "[{timestamp}] Checked all mails. Going to sleep.".format(
        timestamp=time.strftime("%H:%M:%S")
    )
)
